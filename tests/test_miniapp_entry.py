import asyncio
import ast
import hashlib
import hmac
import json
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, Mock
from urllib.parse import urlencode, urlparse
from test_presentation_support import load_functions

web = SimpleNamespace(json_response=lambda data, status=200, headers=None:
                      SimpleNamespace(text=json.dumps(data), status=status, headers=headers))


class MiniAppEntryTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.user = {'name': 'Old', 'credits': 70, 'lang': 'ru', 'history': ['Example']}
        self.handler = AsyncMock()
        self.bot = SimpleNamespace(id=10, send_message=AsyncMock(), set_chat_menu_button=AsyncMock())
        self.env = load_functions(['verify_telegram_init_data', 'miniapp_verified_user',
            'miniapp_user_record', 'api_profile_handler', 'api_action_handler',
            'run_miniapp_action', 'accept_miniapp_action', '_cors_headers', 'native_miniapp_url', 'sync_menu_button'], {
            'BOT_TOKEN': 'unit-test-token', 'INIT_DATA_MAX_AGE_SECONDS': 604800,
            'hashlib': hashlib, 'hmac': hmac, 'time': time, 'asyncio': asyncio,
            'parse_qsl': __import__('urllib.parse', fromlist=['parse_qsl']).parse_qsl,
            '_aiohttp_web': web, 'MINIAPP_URL': 'https://example.test/index.html',
            'WEB_BASE_URL': 'https://example.test', 'urlparse': urlparse, 'urlencode': urlencode,
            'MINIAPP_VERSION': 18, 'tr': lambda k, lang: 'Открыть меню',
            'MenuButtonWebApp': lambda **kw: SimpleNamespace(**kw), 'WebAppInfo': lambda **kw: SimpleNamespace(**kw),
            'get_user': lambda uid: self.user, 'save_users': Mock(), 'user_lang': lambda uid: 'ru',
            'ADMIN_IDS': [], 'bot': self.bot, 'dp': SimpleNamespace(storage=None),
            'FSMContext': lambda **kw: kw, 'StorageKey': lambda **kw: kw,
            '_FakeMessageForApi': lambda uid, name: SimpleNamespace(from_user=SimpleNamespace(id=uid, first_name=name)),
            'ensure_answerable': lambda m: m, '_handle_miniapp_payload': self.handler,
            '_miniapp_action_tasks': {}, '_miniapp_receipts': {}, 'MAX_CONCURRENT_GENERATIONS': 20,
        })
        tree = ast.parse((Path(__file__).resolve().parents[1] / 'main.py').read_text())
        n = next(n for n in tree.body if isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'MINIAPP_ACTIONS' for t in n.targets))
        self.env['MINIAPP_ACTIONS'] = ast.literal_eval(n.value)

    def signed(self, uid=123, age=0):
        p = {'user': json.dumps({'id': uid, 'first_name': 'Тигран'}), 'auth_date': str(int(time.time())-age)}
        secret = hmac.new(b'WebAppData', b'unit-test-token', hashlib.sha256).digest()
        p['hash'] = hmac.new(secret, '\n'.join(f'{k}={v}' for k,v in sorted(p.items())).encode(), hashlib.sha256).hexdigest()
        return urlencode(p)

    async def call(self, name, body):
        return await self.env[name](SimpleNamespace(json=AsyncMock(return_value=body)))

    async def drain(self):
        tasks = list(self.env['_miniapp_action_tasks'].values())
        if tasks: await asyncio.gather(*tasks)

    async def test_signed_profile_refreshes_name_balance_history(self):
        r = await self.call('api_profile_handler', {'initData': self.signed(), 'user_id': 999})
        self.assertEqual(r.status, 200)
        profile = json.loads(r.text)['profile']
        self.assertEqual(profile['name'], 'Тигран')
        self.assertEqual(profile['credits'], 70)
        self.assertEqual(profile['history'], ['Example'])
        self.assertFalse(profile['unlimited'])
        self.assertEqual(r.headers['Cache-Control'], 'no-store')
        self.assertEqual(self.user['name'], 'Тигран')

    async def test_rejects_forged_expired_and_invalid_requests(self):
        for signed in ['', self.signed()+'bad', self.signed(age=700000), self.signed(uid='123')]:
            self.assertEqual((await self.call('api_profile_handler', {'initData': signed})).status, 401)
            self.assertEqual((await self.call('api_action_handler', {'initData': signed, 'action': 'help'})).status, 401)
        for body in [[], {'payload': []}, {'initData': self.signed(), 'action': []}, {'initData': self.signed(), 'action': 'invalid'}]:
            self.assertEqual((await self.call('api_action_handler', body)).status, 400)
        self.handler.assert_not_awaited()

    async def test_accepts_promptly_deduplicates_and_dispatches_same_payload(self):
        release = asyncio.Event()
        self.handler.side_effect = lambda *args: None
        async def hold(*args): await release.wait()
        self.handler.side_effect = hold
        body = {'initData': self.signed(), 'action': 'gen_presentation', 'requestId': 'test-1', 'payload': {'topic': 'Clouds', 'slides': 8, 'style': 'designer_soft'}}
        self.assertEqual((await self.call('api_action_handler', body)).status, 202)
        self.assertEqual((await self.call('api_action_handler', body)).status, 202)
        self.assertEqual((await self.call('api_action_handler', {**body, 'requestId': 'test-2'})).status, 409)
        self.assertEqual((await self.call('api_action_handler', {**body, 'payload': {'topic': 'Changed'}})).status, 409)
        release.set(); await self.drain()
        self.handler.assert_awaited_once()
        self.assertEqual(self.handler.await_args.args[2], body['payload'])
        self.assertEqual(self.handler.await_args.args[0].from_user.id, 123)
        self.assertEqual((await self.call('api_action_handler', body)).status, 202)
        self.assertEqual(self.handler.await_count, 1)

    async def test_error_notifies_chat_and_unlocks(self):
        self.handler.side_effect = RuntimeError('fixture failure')
        await self.call('api_action_handler', {'initData': self.signed(), 'action': 'help'})
        await self.drain()
        self.bot.send_message.assert_awaited_once()
        self.assertEqual(self.env['_miniapp_action_tasks'], {})

    async def test_menu_url_has_no_stale_personal_data(self):
        await self.env['sync_menu_button'](123, self.user, 'ru')
        button = self.bot.set_chat_menu_button.await_args.kwargs['menu_button']
        self.assertIn('entry=menu', button.web_app.url)
        self.assertNotIn('credits=', button.web_app.url)
        self.assertNotIn('name=', button.web_app.url)
        await self.env['sync_menu_button'](None, {}, 'ru')
        self.assertIsNone(self.bot.set_chat_menu_button.await_args.kwargs['chat_id'])
