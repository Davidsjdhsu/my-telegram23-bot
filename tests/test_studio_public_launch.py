"""Public rollout still requires signed identities and isolates accounts."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, patch

from aiohttp import web
from studio_private import COOKIE, PILOT_UID, is_pilot_user, menu_url
import test_studio_private as fixtures
from test_studio_private import HEADERS, ORIGIN


class PublicLaunchTests(unittest.IsolatedAsyncioTestCase):
    signed = fixtures.StudioTests.signed
    login = fixtures.StudioTests.login

    async def asyncSetUp(self):
        self.public = patch('studio_private.PUBLIC_ACCESS', True)
        self.public.start()
        self.addCleanup(self.public.stop)
        await fixtures.StudioTests.asyncSetUp(self)
        self.users[123] = {'credits': 100}
        self.users[456] = {'credits': 200}

    async def asyncTearDown(self):
        await fixtures.StudioTests.asyncTearDown(self)

    async def test_signed_non_tester_uses_own_profile_and_payment(self):
        self.studio.profile_callback = AsyncMock(side_effect=lambda user: self.users[user['id']])
        await self.login(123)
        r = await self.client.get('/studio/profile?user_id=456')
        self.assertEqual((await r.json())['profile']['credits'], 100)
        for changes in ({'confirmedCost': 0}, {'confirmedCost': 36}):
            deck = fixtures.StudioTests.deck(self) | changes
            r = await self.client.post('/studio/export', headers=HEADERS, json=deck)
            self.assertEqual(r.status, 409)
        self.send.assert_not_awaited()
        self.assertEqual(self.users[123]['credits'], 100)
        self.studio.menu_action = AsyncMock(return_value=web.json_response({'ok': True}))
        r = await self.client.post('/studio/action', headers=HEADERS,
                                  json={'action': 'history', 'payload': {'user_id': 456}})
        self.assertEqual(r.status, 200)
        self.assertEqual(self.studio.menu_action.call_args.args[0]['id'], 123)

    async def test_unsigned_forged_stale_wrong_origin_still_rejected(self):
        for init in ['', self.signed(123) + 'bad', self.signed(123, age=3602),
                     self.signed(123, age=-100), self.signed('123'), self.signed(0)]:
            r = await self.client.post('/studio/session', headers=HEADERS,
                                       json={'initData': init, 'user_id': PILOT_UID})
            self.assertEqual(r.status, 403)
        self.assertEqual((await self.client.get('/studio/profile')).status, 401)
        await self.login(123)
        r = await self.client.post('/studio/action', headers={'Origin': 'https://evil.test'},
                                  json={'action': 'history'})
        self.assertEqual(r.status, 403)

    async def test_public_menu_and_legacy_link(self):
        for uid in (None, 123, PILOT_UID):
            self.assertEqual(menu_url(ORIGIN, uid, 'old'), ORIGIN + '/studio/')
        for uid in (True, '123', 0, -1):
            self.assertFalse(is_pilot_user(uid))
        r = await self.client.get('/index.html?entry=menu&lang=en', allow_redirects=False)
        self.assertEqual(r.status, 302)
        self.assertEqual(r.headers['Location'], '/studio/?entry=menu&lang=en')
        self.assertEqual(r.headers['Cache-Control'], 'no-store')
        for path in ('/studio/', '/studio/templates/', '/studio/constructor/'):
            if path != '/studio/':
                await self.login(123)
            r = await self.client.get(path)
            self.assertEqual(r.status, 200)
            self.assertNotIn('личный тест', (await r.text()).lower())

    async def test_more_than_pilot_session_limit_and_restart(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.studio.state_root = Path(tmp)
            await self.login(123)
            first = next(iter(self.studio.sessions))
            for uid in range(1000, 1040):
                await self.login(uid)
            self.assertIn(first, self.studio.sessions)
            self.studio.sessions.clear()
            self.client.session.cookie_jar.update_cookies({COOKIE: first})
            self.studio.profile_callback = AsyncMock(side_effect=lambda user: self.users[user['id']])
            self.assertEqual((await self.client.get('/studio/profile')).status, 200)
            self.assertEqual(self.studio.profile_callback.call_args.args[0]['id'], 123)

    async def test_rollback_revokes_non_testers_not_admins(self):
        await self.login(123)
        with patch('studio_private.PUBLIC_ACCESS', False):
            self.assertEqual((await self.client.get('/studio/profile')).status, 401)
            self.assertEqual(menu_url(ORIGIN, None, 'old'), 'old')
            self.assertEqual(menu_url(ORIGIN, 123, 'old'), 'old')
            self.assertEqual(menu_url(ORIGIN, PILOT_UID, 'old'), ORIGIN + '/studio/')
            r = await self.client.post('/studio/session', headers=HEADERS,
                                      json={'initData': self.signed(123)})
            self.assertEqual(r.status, 403)
