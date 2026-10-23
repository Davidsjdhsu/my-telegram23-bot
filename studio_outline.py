"""Bounded owner-only outline cache. Content/images/billing start only on confirm."""
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

    def response(self, row):
        value = {'ok': True, 'id': row['id'], 'status': row['status'], 'cost': row['cost']}
        if row['status'] in ('ready', 'building', 'completed', 'build_failed'):
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
        if (body.get('kind') != 'standard' or not isinstance(topic, str) or not 1 <= len(topic.strip()) <= 500
                or type(count) is not int or count not in (4, 6, 8, 10, 12, 16)
                or not any(r['group'] == 'standard' and r['key'] == key for r in catalog)):
            raise web.HTTPBadRequest()
        payload = {'topic': topic.strip(), 'key': key, 'slides': count}
        fingerprint = hashlib.sha256(json.dumps(payload, sort_keys=True).encode()).hexdigest()
        now = time.monotonic()
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
            oldest = min(self.rows, key=lambda k: self.rows[k]['created'])
            del self.rows[oldest]
        token = secrets.token_urlsafe(24)
        row = dict(id=token, uid=user['id'], fingerprint=fingerprint, created=now, expires=now+3600,
                   status='planning', cost=presentation_cost(count), error=None)
        self.rows[token] = row
        row['task'] = asyncio.create_task(self.prepare(row, user, payload))
        return self.response(row)

    async def prepare(self, row, user, payload):
        try:
            async with asyncio.timeout(75):
                data, palette = await self.planner(user, payload)
            row.update(data=copy.deepcopy(data), palette=palette, status='ready')
        except web.HTTPException as error:
            row.update(status='failed', error=error.text)
        except Exception:
            row.update(status='failed', error='Не удалось подготовить план. Попробуйте ещё раз. Кредиты не списаны.')

    def get_row(self, request):
        user = self.studio.authorized(request)['user']
        row = self.rows.get(request.match_info['outline_id'])
        if not row or row['uid'] != user['id'] or row['expires'] <= time.monotonic():
            raise web.HTTPNotFound(text='План устарел. Составьте его заново.')
        return user, row

    async def status(self, request):
        _, row = self.get_row(request)
        return self.response(row)

    async def confirm(self, request):
        user, row = self.get_row(request)
        self.studio.check_write(request)
        # The client cannot override the confirmed plan, recipient, settings or cost.
        if row['status'] in ('building', 'completed', 'build_failed'):
            return self.response(row)
        if row['status'] != 'ready':
            return web.json_response({'ok': False, 'error': 'Дождитесь готового плана.'}, status=409)
        row['status'] = 'submitting'
        try:
            task = await self.submit(user, copy.deepcopy(row['data']))
            row.update(status='building', error=None, expires=time.monotonic()+7200)
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
        return self.response(row)

    def build_finished(self, row, task):
        try:
            result = task.result()
            success = isinstance(result, dict) and result.get('delivered') is True
        except (Exception, asyncio.CancelledError):
            success = False
        row.update(status='completed' if success else 'build_failed', expires=time.monotonic()+7200,
                   error=None if success else 'Сборка не завершилась успешно. Подробности и доступные действия — в чате бота. Проверьте полученные файлы перед повторной попыткой.')

    async def cleanup(self, app):
        tasks = [r['task'] for r in self.rows.values() if not r['task'].done()]
        for task in tasks:
            task.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)

    def register(self, app):
        app.router.add_post('/studio/outline', self.create)
        app.router.add_get('/studio/outline/{outline_id}', self.status)
        app.router.add_post('/studio/outline/{outline_id}/confirm', self.confirm)
        app.on_cleanup.append(self.cleanup)
