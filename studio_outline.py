"""Bounded private outline cache. Content/images/billing start only on confirm."""
import asyncio
import copy
import hashlib
import json
import secrets
import time
from aiohttp import web
from presentation_pricing import presentation_cost


class OutlinePreview:
    def __init__(self, studio, planner, submit):
        self.studio, self.planner, self.submit = studio, planner, submit
        self.rows = {}
        self.path = studio.state_root/'outlines.json' if getattr(studio, 'state_root', None) else None
        if self.path:
            try:
                self.rows = json.loads(self.path.read_text())
                self.rows = {k:r for k,r in self.rows.items() if r['expires'] > time.time()}
                for row in self.rows.values():
                    if row['status'] in ('building', 'submitting'):
                        row.update(status='interrupted', error='Сервер перезапустился. Заказ сохранён, можно продолжить.')
                    elif row['status'] == 'planning':
                        row.update(status='failed', error='Подготовка плана прервалась. Составьте план ещё раз.')
                self.save()
            except (OSError, ValueError, KeyError): self.rows = {}

    def save(self):
        if self.path:
            from studio_word import atomic
            atomic(self.path, {k:{f:v for f,v in r.items() if f!='task'} for k,r in self.rows.items()})

    def sync_order(self, row):
        service = getattr(self.studio, 'presentation_orders', None)
        if service and row.get('data') and row['status'] not in ('ready','planning','failed'):
            key = service.key(row['uid'], row['data'])
            path = service.root/key/'order.json'
            try:
                order = json.loads(path.read_text())
                row.update(status=order['status'], stage=order.get('stage'), charged=order.get('charged'))
            except (OSError, ValueError): pass


    def response(self, row):
        self.sync_order(row)
        value = {'stage': row.get('stage'), 'charged': row.get('charged'), 'ok': True, 'id': row['id'], 'status': row['status'], 'cost': row['cost']}
        if row['status'] in ('ready', 'building', 'delivering', 'completed', 'build_failed', 'interrupted', 'delivery_unknown'):
            plan = row['data']['pres_plan']
            value.update(title=plan['title'], audience=plan['audience'],
                         titles=[s['title'] for s in plan['slides']], palette=row['palette'])
        if row.get('error'):
            value['error'] = row['error']
        return web.json_response(value, headers={'Cache-Control': 'no-store'})

    async def create(self, request):
        user = self.studio.authorized(request)['user']
        self.studio.check_write(request)
        body = await self.studio.body(request, 16384)
        topic, key, count = (body.get(k) for k in ('topic', 'key', 'slides'))
        from studio_private import BASE
        catalog = json.loads((BASE/'studio_ui/templates/catalog.json').read_text())
        kind, style = body.get('kind'), body.get('style')
        group = 'standard' if kind == 'standard' else style if kind == 'designer' and style in ('strict', 'soft') else None
        if (not group or not isinstance(topic, str) or not 1 <= len(topic.strip()) <= 500
                or type(count) is not int or count not in (4, 6, 8, 10, 12, 16)
                or not any(r['group'] == group and r['key'] == key for r in catalog)):
            raise web.HTTPBadRequest()
        payload = {'topic': topic.strip(), 'key': key, 'slides': count, 'kind': kind}
        if kind == 'designer':
            payload['style'] = style
        from studio_settings import CONTENT_LANGUAGES, PALETTES
        if 'content_lang' in body:
            if not isinstance(body['content_lang'], str) or body['content_lang'] not in CONTENT_LANGUAGES:
                raise web.HTTPBadRequest()
            payload['content_lang'] = body['content_lang']
        if body.get('palette') is not None:
            if kind != 'designer' or not isinstance(body['palette'], str) or body['palette'] not in PALETTES:
                raise web.HTTPBadRequest()
            payload['palette'] = body['palette']
        if kind == 'designer':
            from presentation_options import options
            try:
                payload.update(options(body))
            except (TypeError, ValueError):
                raise web.HTTPBadRequest()
        elif any(k in body for k in ('background_mode', 'diagram_layout')):
            raise web.HTTPBadRequest()
        fingerprint = hashlib.sha256(json.dumps(payload, sort_keys=True).encode()).hexdigest()
        now = time.time()
        self.rows = {k: r for k, r in self.rows.items() if r['expires'] > now}
        for row in self.rows.values():
            if row['uid'] == user['id']:
                if row['fingerprint'] == fingerprint and row['status'] != 'failed':
                    return self.response(row)
                if row['status'] in ('planning', 'submitting', 'building'):
                    return web.json_response({'ok': False, 'error': 'Дождитесь текущего плана.'}, status=409)
                if now - row['created'] < 15:
                    return web.json_response({'ok': False, 'error': 'Подождите несколько секунд перед новым планом.'}, status=429)
        if len(self.rows) >= 12:
            finished = [k for k,r in self.rows.items() if r['status'] in ('completed','failed','ready')]
            if not finished:
                raise web.HTTPTooManyRequests(text='Дождитесь завершения текущих сборок.')
            oldest = min(finished, key=lambda k: self.rows[k]['created'])
            del self.rows[oldest]
        cost = presentation_cost(count, 'designer_'+style if kind == 'designer' else key)
        if self.studio.profile_callback:
            profile = await self.studio.profile_callback(user)
            if not profile.get('unlimited') and profile.get('available', profile.get('credits', 0)) < cost:
                raise self.studio.shortage(user['id'], cost)
        token = secrets.token_urlsafe(24)
        row = dict(id=token, uid=user['id'], fingerprint=fingerprint, created=now, expires=now+86400,
                   status='planning', cost=presentation_cost(count, 'designer_'+style if kind == 'designer' else key), error=None)
        self.rows[token] = row
        self.save()
        row['task'] = asyncio.create_task(self.prepare(row, user, payload))
        return self.response(row)

    async def prepare(self, row, user, payload):
        try:
            async with asyncio.timeout(75):
                data, palette = await self.planner(user, payload)
            if payload['kind'] == 'designer':
                from presentation_options import options
                from presentation_support import plan_key
                chosen = options(payload)
                data['studio_design_options'] = chosen
                data['pres_plan']['design'].update(chosen)
                data['pres_plan_key'] = plan_key(data)
                palette = {**palette, **chosen}
            row.update(data=copy.deepcopy(data), palette=palette, status='ready')
        except web.HTTPException as error:
            row.update(status='failed', error=error.text)
        except Exception:
            row.update(status='failed', error='Не удалось подготовить план. Попробуйте ещё раз. Кредиты не списаны.')

        finally:
            self.save()

    def get_row(self, request):
        user = self.studio.authorized(request)['user']
        row = self.rows.get(request.match_info['outline_id'])
        if not row or row['uid'] != user['id'] or row['expires'] <= time.time():
            raise web.HTTPNotFound(text='План устарел. Составьте его заново.')
        return user, row

    async def status(self, request):
        _, row = self.get_row(request)
        return self.response(row)

    async def confirm(self, request):
        user, row = self.get_row(request)
        self.studio.check_write(request)
        self.sync_order(row)
        # The client cannot override the confirmed plan, recipient, settings or cost.
        if row['status'] in ('building', 'delivering', 'completed'):
            return self.response(row)
        if row['status'] == 'delivery_unknown':
            raise web.HTTPConflict(text='Проверьте файлы в чате и обратитесь в поддержку. Повторной отправки и списания не будет.')
        if row['status'] not in ('ready', 'build_failed', 'interrupted'):
            return web.json_response({'ok': False, 'error': 'Дождитесь готового плана.'}, status=409)
        row['status'] = 'submitting'
        self.save()
        try:
            task = await self.submit(user, copy.deepcopy(row['data']))
            row.update(status='building', error=None, expires=time.time()+86400)
            if isinstance(task, asyncio.Future):
                task.add_done_callback(lambda done: self.build_finished(row, done))
            else:
                # Do not claim readiness when an adapter has no completion signal.
                row.update(status='build_failed', error='Не удалось получить статус сборки. Проверьте сообщения бота.')
        except web.HTTPException:
            row['status'] = 'ready'
            raise
        except Exception:
            row['status'] = 'ready'
            return web.json_response({'ok': False, 'error': 'Не удалось запустить сборку. Повторите попытку.'}, status=503)
        self.save()
        return self.response(row)

    def build_finished(self, row, task):
        try:
            result = task.result()
            success = isinstance(result, dict) and result.get('delivered') is True
        except (Exception, asyncio.CancelledError):
            success = False
        row.update(status='completed' if success else 'build_failed', expires=time.time()+86400,
                   error=None if success else 'Сборка не завершилась успешно. Подробности и доступные действия — в чате бота. Проверьте полученные файлы перед повторной попыткой.')

        self.sync_order(row)
        self.save()

    async def latest(self, request):
        uid=self.studio.authorized(request)['user']['id']
        rows=[r for r in self.rows.values() if r['uid']==uid and r['expires']>time.time()]
        if not rows: return web.json_response({'ok':True,'empty':True})
        return self.response(max(rows,key=lambda r:r['created']))

    async def cleanup(self, app):
        tasks = [r['task'] for r in self.rows.values() if r.get('task') and not r['task'].done()]
        for task in tasks:
            task.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)

    def register(self, app):
        app.router.add_post('/studio/outline', self.create)
        app.router.add_get('/studio/outline/latest', self.latest)
        app.router.add_get('/studio/outline/{outline_id}', self.status)
        app.router.add_post('/studio/outline/{outline_id}/confirm', self.confirm)
        app.on_cleanup.append(self.cleanup)
