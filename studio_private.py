"""Owner-only Mini App pilot. No credit mutations and no public feature switch."""
import asyncio
import hashlib
import json
import mimetypes
import secrets
import time
import zipfile
from pathlib import Path
from urllib.parse import parse_qsl

from aiohttp import web

PILOT_UID = 465823470
VERSION = "studio-menu-pilot-2"
COOKIE = "alix_studio"
BASE = Path(__file__).resolve().parent


def menu_url(base, uid, fallback):
    return base.rstrip('/') + '/studio/' if uid == PILOT_UID and base.startswith('https://') else fallback


BOOTSTRAP = '''<!doctype html><html lang="ru"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>AlixDoc · Тест</title><body style="background:#191b20;color:#e8dfcf;font:18px system-ui;padding:30px"><h1>Alix<span style="color:#c4aa7d">Doc</span></h1><p id="s">Открываю личный тест…</p><a href="/index.html" style="color:#c4aa7d">Прежнее меню</a><script src="/telegram-web-app.js"></script><script>
(async()=>{try{const tg=window.Telegram?.WebApp;tg?.ready();tg?.expand();const r=await fetch('/studio/session',{method:'POST',headers:{'Content-Type':'application/json','X-Studio-Request':'1'},body:JSON.stringify({initData:tg?.initData||''})});if(!r.ok)throw Error('Тест доступен только владельцу. Откройте его кнопкой «Меню» в Telegram.');location.replace('/studio/home/');}catch(e){document.getElementById('s').textContent=e.message;}})();
</script></body></html>'''


class Studio:
    def __init__(self, verify, send_file, action, origin, profile=None, menu_action=None):
        self.verify, self.send_file, self.action = verify, send_file, action
        self.profile_callback, self.menu_action = profile, menu_action
        self.forms_cache = None
        self.origin = origin.rstrip('/')
        self.sessions = {}
        self.busy = False
        self.receipts = {}
        self.archives = [zipfile.ZipFile(p) for p in sorted(BASE.glob('studio-assets-*.zip'))]

    def authorized(self, request):
        s = self.sessions.get(request.cookies.get(COOKIE, ''))
        if not s or s['expires'] < time.monotonic():
            raise web.HTTPUnauthorized(text='Откройте «Меню» заново в Telegram.')
        return s

    def check_write(self, request):
        if request.headers.get('X-Studio-Request') != '1' or request.headers.get('Origin') != self.origin:
            raise web.HTTPForbidden()

    async def body(self, request, limit):
        data = bytearray()
        async for chunk in request.content.iter_chunked(65536):
            data.extend(chunk)
            if len(data) > limit:
                raise web.HTTPRequestEntityTooLarge(max_size=limit, actual_size=len(data))
        try:
            value = json.loads(data)
            if not isinstance(value, dict):
                raise ValueError()
            return value
        except (ValueError, UnicodeError):
            raise web.HTTPBadRequest(text='Некорректный запрос')

    async def entry(self, request):
        return web.Response(text=BOOTSTRAP, content_type='text/html', headers={'Cache-Control': 'no-store'})

    async def session(self, request):
        self.check_write(request)
        body = await self.body(request, 16384)
        init = body.get('initData', '')
        if not isinstance(init, str):
            raise web.HTTPForbidden()
        # The shared verifier authenticates Telegram's signature; this pilot also
        # requires an explicit, recent auth_date (never trusts query user IDs).
        try:
            pairs = dict(parse_qsl(init, strict_parsing=True))
            age = time.time() - int(pairs.get('auth_date', '0'))
        except (ValueError, TypeError):
            raise web.HTTPForbidden()
        user = self.verify(init)
        if not user or user.get('id') != PILOT_UID or not -30 <= age <= 3600:
            raise web.HTTPForbidden()
        now = time.monotonic()
        self.sessions = {k: v for k, v in self.sessions.items() if v['expires'] > now}
        if len(self.sessions) >= 16:
            self.sessions.pop(next(iter(self.sessions)))
        token = secrets.token_urlsafe(32)
        self.sessions[token] = {'expires': now + 7200, 'user': user}
        response = web.json_response({'ok': True, 'version': VERSION}, headers={'Cache-Control':'no-store'})
        response.set_cookie(COOKIE, token, max_age=7200, secure=True, httponly=True, samesite='Strict', path='/studio/')
        return response

    async def profile(self, request):
        s = self.authorized(request)
        if not self.profile_callback:
            raise web.HTTPServiceUnavailable()
        return web.json_response({'ok': True, 'profile': await self.profile_callback(s['user'])}, headers={'Cache-Control':'no-store'})

    async def menu_request(self, request):
        s = self.authorized(request)
        self.check_write(request)
        body = await self.body(request, 16384)
        action, payload = body.get('action'), body.get('payload', {})
        allowed = {'gen_word','gen_excel','topup','history','help','plan','mode_buttons','mode_voice','chat','image','collab','word','excel','template'}
        if not isinstance(action, str) or action not in allowed or not isinstance(payload, dict):
            raise web.HTTPBadRequest()
        if not self.menu_action:
            raise web.HTTPServiceUnavailable()
        return await self.menu_action(s['user'], action, payload, body.get('requestId'))

    async def forms(self, request):
        self.authorized(request)
        if self.forms_cache is None:
            from studio_menu import forms_html
            self.forms_cache = forms_html(BASE)
        return web.Response(text=self.forms_cache, content_type='text/html', headers={'Cache-Control':'no-store'})

    async def static(self, request):
        self.authorized(request)
        name = request.match_info['path']
        if not name or name.endswith('/'):
            name += 'index.html'
        if '..' in name.split('/') or '\\' in name or name.startswith('/'):
            raise web.HTTPNotFound()
        path = (BASE / 'studio_ui' / name).resolve()
        root = (BASE / 'studio_ui').resolve()
        data = None
        if root in path.parents and path.is_file():
            data = path.read_bytes()
        else:
            for archive in self.archives:
                try:
                    data = archive.read(name)
                    break
                except KeyError:
                    pass
        if data is None:
            raise web.HTTPNotFound()
        return web.Response(body=data, content_type=mimetypes.guess_type(name)[0] or 'application/octet-stream', headers={
            'Cache-Control': 'private, no-cache', 'X-Content-Type-Options':'nosniff',
            'Referrer-Policy':'same-origin',
        })

    async def generate(self, request):
        s = self.authorized(request)
        self.check_write(request)
        body = await self.body(request, 16384)
        topic = body.get('topic')
        kind, style, key, slides = (body.get(k) for k in ('kind', 'style', 'key', 'slides'))
        if not isinstance(topic, str) or not topic.strip() or len(topic)>500 or type(slides) is not int or slides not in (4,6,8,10,12,16):
            raise web.HTTPBadRequest()
        catalog = json.loads((BASE/'studio_ui/templates/catalog.json').read_text())
        group = 'standard' if kind == 'standard' else style if kind == 'designer' and style in ('strict','soft') else None
        if not group or not any(r['group']==group and r['key']==key for r in catalog):
            raise web.HTTPBadRequest()
        payload = {'topic':topic.strip(), 'slides':slides, 'style':key if kind=='standard' else 'designer_'+style, 'photo_mode':'ai'}
        if kind == 'designer':
            payload['studio_template'] = key
        # Existing action endpoint owns concurrency and request-id deduplication.
        return await self.action(s['user'], payload, body.get('requestId'))

    async def export(self, request):
        self.authorized(request)
        self.check_write(request)
        if self.busy:
            return web.json_response({'ok':False,'error':'Предыдущий файл ещё готовится.'},status=409)
        self.busy = True
        try:
            async with asyncio.timeout(90):
                body = await self.body(request, 16_000_000)
                fingerprint = hashlib.sha256(json.dumps(body, sort_keys=True).encode()).hexdigest()
                if self.receipts.get(fingerprint, 0) > time.monotonic()-600:
                    return web.json_response({'ok':True,'duplicate':True})
                from studio_export import build_pptx
                data = await asyncio.to_thread(build_pptx, body)
                # Only the fixed authenticated owner receives files. No client chat_id.
                await self.send_file(PILOT_UID, data)
                self.receipts = {k:v for k,v in self.receipts.items() if v>time.monotonic()-600}
                self.receipts[fingerprint] = time.monotonic()
                return web.json_response({'ok':True})
        except (ValueError, TypeError, KeyError):
            return web.json_response({'ok':False,'error':'Проверьте проект: неверный элемент или слишком большой файл.'},status=400)
        except TimeoutError:
            return web.json_response({'ok':False,'error':'Истекло время ожидания. Проверьте чат перед повторной отправкой.'},status=504)
        except web.HTTPException:
            raise
        except Exception as exc:
            print('studio_export_error', type(exc).__name__, flush=True)
            return web.json_response({'ok':False,'error':'Не удалось отправить файл. Проверьте чат и попробуйте снова.'},status=503)
        finally:
            self.busy = False

    def register(self, app):
        app.router.add_get('/studio/', self.entry)
        app.router.add_post('/studio/session', self.session)
        app.router.add_post('/studio/generate', self.generate)
        app.router.add_post('/studio/export', self.export)
        app.router.add_get('/studio/profile', self.profile)
        app.router.add_post('/studio/action', self.menu_request)
        app.router.add_get('/studio/forms/', self.forms)
        app.router.add_get('/studio/{path:.*}', self.static)
