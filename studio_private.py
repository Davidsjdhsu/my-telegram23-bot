"""Allowlisted Mini App pilot. Paid constructor; no public feature switch."""
import asyncio
import hashlib
import json
import mimetypes
import os
import re
import secrets
import time
import zipfile
from pathlib import Path
from urllib.parse import parse_qsl

from aiohttp import web

PILOT_UID = 465823470
def pilot_ids(extra):
    """Additional testers get Studio access only, never admin privileges."""
    ids = {PILOT_UID, 909828109}
    for value in extra.split(','):
        value = value.strip()
        if re.fullmatch(r"[0-9]{1,16}", value) and int(value) > 0:
            ids.add(int(value))
    return frozenset(ids)


PILOT_UIDS = pilot_ids(os.getenv("STUDIO_ADDITIONAL_TESTER_IDS", ""))
VERSION = "studio-mobile-pilot-3"
COOKIE = "alix_studio"
BASE = Path(__file__).resolve().parent


def is_pilot_user(uid):
    return type(uid) is int and uid in PILOT_UIDS


def menu_url(base, uid, fallback):
    return base.rstrip('/') + '/studio/' if is_pilot_user(uid) and base.startswith('https://') else fallback


# Only the generic entry shell is public. Account data, forms, catalog and
# generation endpoints retain the signed Telegram tester session requirement.
PUBLIC_SHELL = {
    'home/style.css', 'home/app.js', 'shared/mobile.css', 'shared/mobile.js',
    'templates/fonts.css', 'shared/preferences.js', 'shared/en.json', 'shared/billing.js', 'shared/billing.css',
}
PUBLIC_SHELL.update('templates/' + name for name in re.findall(
    r'url\((fonts/[a-f0-9]+\.woff2)\)',
    (BASE / 'studio_ui/templates/fonts.css').read_text(),
))


class Studio:
    def __init__(self, verify, send_file, action, origin, profile=None, menu_action=None, outline_planner=None, outline_submit=None, settings=None, billing=None, state_root=None, presentation_orders=None, available=None):
        self.verify, self.send_file, self.action = verify, send_file, action
        self.profile_callback, self.menu_action = profile, menu_action
        self.settings_callback = settings
        self.billing = billing
        self.available = available or (lambda uid: billing.get_user(uid).get('credits', 0) if billing else 0)
        self.state_root = Path(state_root) if state_root else None
        self.presentation_orders = presentation_orders
        if self.state_root:
            self.state_root.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.forms_cache = None
        self.outline = None
        if outline_planner and outline_submit:
            from studio_outline import OutlinePreview
            self.outline = OutlinePreview(self, outline_planner, outline_submit)
        self.origin = origin.rstrip('/')
        self.sessions = {}
        self.busy = False
        self.import_busy = False
        self.receipts = {}
        self.archives = [zipfile.ZipFile(p) for p in sorted(BASE.glob('studio-assets-*.zip'))]

    def shortage(self, uid, required):
        from credit_shortage import error
        return error(required, self.available(uid))

    def authorized(self, request):
        s = self.sessions.get(request.cookies.get(COOKIE, ''))
        if not s and self.state_root:
            token = request.cookies.get(COOKIE, '')
            if isinstance(token, str) and 20 <= len(token) <= 100:
                try:
                    saved = json.loads((self.state_root/'sessions.json').read_text())
                    record = saved.get(hashlib.sha256(token.encode()).hexdigest())
                    if record and record['expires'] > time.time() and is_pilot_user(record['user'].get('id')):
                        s = {'expires': time.monotonic() + record['expires'] - time.time(), 'user': record['user']}
                        self.sessions[token] = s
                except (OSError, ValueError, KeyError, TypeError):
                    pass
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
        return web.Response(text=(BASE / 'studio_ui/home/index.html').read_text(), content_type='text/html', headers={'Cache-Control': 'no-store'})

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
        if not user or not is_pilot_user(user.get('id')) or not -30 <= age <= 3600:
            raise web.HTTPForbidden()
        now = time.monotonic()
        self.sessions = {k: v for k, v in self.sessions.items() if v['expires'] > now}
        if len(self.sessions) >= 16:
            self.sessions.pop(next(iter(self.sessions)))
        token = secrets.token_urlsafe(32)
        self.sessions[token] = {'expires': now + 7200, 'user': user}
        if self.state_root:
            from studio_word import atomic
            path = self.state_root/'sessions.json'
            try: saved = json.loads(path.read_text())
            except (OSError, ValueError): saved = {}
            saved = {k:v for k,v in saved.items() if v['expires'] > time.time()}
            saved[hashlib.sha256(token.encode()).hexdigest()] = {'expires': time.time()+7200, 'user': user}
            saved = dict(sorted(saved.items(), key=lambda item:item[1]['expires'])[-32:])
            atomic(path, saved)
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

    async def settings(self, request):
        user = self.authorized(request)['user']
        self.check_write(request)
        from studio_settings import preferences
        try:
            values = preferences(await self.body(request, 1024))
        except ValueError:
            raise web.HTTPBadRequest()
        if not self.settings_callback:
            raise web.HTTPServiceUnavailable()
        await self.settings_callback(user, values)
        return web.json_response({'ok': True, **values}, headers={'Cache-Control': 'no-store'})

    async def forms(self, request):
        self.authorized(request)
        if self.forms_cache is None:
            from studio_menu import forms_html
            self.forms_cache = forms_html(BASE)
        return web.Response(text=self.forms_cache, content_type='text/html', headers={'Cache-Control':'no-store'})

    async def static(self, request):
        name = request.match_info['path']
        if name not in PUBLIC_SHELL:
            self.authorized(request)
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

    async def import_presentation(self, request):
        self.authorized(request)
        self.check_write(request)
        if self.import_busy:
            return web.json_response({'ok':False,'error':'Другой файл ещё открывается. Повторите через несколько секунд.'},status=409)
        self.import_busy = True
        try:
            from studio_import import MAX_UPLOAD, import_pptx
            raw = bytearray()
            async for chunk in request.content.iter_chunked(65536):
                raw.extend(chunk)
                if len(raw) > MAX_UPLOAD:
                    raise ValueError('Выберите PPTX до 10 МБ.')
            from urllib.parse import unquote
            name = unquote(request.headers.get('X-Presentation-Name', 'Презентация.pptx'))[:180]
            result = await asyncio.to_thread(import_pptx, bytes(raw), name)
            return web.json_response(result, headers={'Cache-Control':'no-store'})
        except ValueError as error:
            return web.json_response({'ok':False,'error':str(error)}, status=400)
        except Exception as error:
            print('studio_import_error', type(error).__name__, flush=True)
            return web.json_response({'ok':False,'error':'Не удалось открыть этот PPTX. Исходный файл и текущий проект не изменены.'},status=400)
        finally:
            self.import_busy = False

    async def export(self, request):
        uid = self.authorized(request)['user']['id']
        self.check_write(request)
        if self.busy:
            return web.json_response({'ok':False,'error':'Предыдущий файл ещё готовится.'},status=409)
        self.busy = True
        try:
            async with asyncio.timeout(90):
                body = await self.body(request, 16_000_000)
                from constructor_pricing import quote, VERSION
                price = quote(body)
                if body.get('pricingVersion') != VERSION or type(body.get('confirmedCost')) is not int or body['confirmedCost'] != price['cost']:
                    return web.json_response({'ok':False,'error':'Цена изменилась. Обновите меню и подтвердите новую стоимость.', **price},status=409)
                if not self.billing:
                    raise web.HTTPServiceUnavailable(text='Оплата конструктора временно недоступна.')
                from studio_export import build_pptx
                data = await asyncio.to_thread(build_pptx, body)
                def canonical(value):
                    if isinstance(value, dict):
                        return {k: canonical(v) for k,v in value.items() if k not in ('id','name','confirmedCost','pricingVersion')}
                    if isinstance(value, list):
                        return [canonical(v) for v in value]
                    return value
                fingerprint = hashlib.sha256(json.dumps(canonical(body), sort_keys=True).encode()).hexdigest()
                result = await self.billing.deliver(uid, fingerprint, price['cost'], data, self.send_file)
                return web.json_response(result)
        except (ValueError, TypeError, KeyError):
            return web.json_response({'ok':False,'error':'Проверьте проект: неверный элемент или слишком большой файл.'},status=400)
        except TimeoutError:
            return web.json_response({'ok':False,'error':'Истекло время ожидания. Проверьте чат перед повторной отправкой.'},status=504)
        except web.HTTPException as exc:
            if exc.status == 402 and exc.content_type == 'application/json':
                return web.json_response(json.loads(exc.text), status=402)
            return web.json_response({"ok":False,"error":exc.text}, status=exc.status)
        except Exception as exc:
            print('studio_export_error', type(exc).__name__, flush=True)
            return web.json_response({'ok':False,'error':'Не удалось отправить файл. Проверьте чат и попробуйте снова.'},status=503)
        finally:
            self.busy = False

    def register(self, app):
        from studio_palette_preview import register as register_palette_preview
        register_palette_preview(self, app)
        if self.outline:
            self.outline.register(app)
        app.router.add_get('/studio/', self.entry)
        app.router.add_post('/studio/session', self.session)
        app.router.add_post('/studio/generate', self.generate)
        app.router.add_post('/studio/export', self.export)
        app.router.add_post('/studio/import-pptx', self.import_presentation)
        app.router.add_get('/studio/profile', self.profile)
        app.router.add_post('/studio/settings', self.settings)
        app.router.add_post('/studio/action', self.menu_request)
        app.router.add_get('/studio/forms/', self.forms)
        app.router.add_get('/studio/{path:.*}', self.static)
