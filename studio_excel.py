"""Private Excel jobs: explicit quote, persisted file and cautious delivery recovery."""
import asyncio
import hashlib
import json
from pathlib import Path
import secrets
import shutil
import time
from aiohttp import web
from studio_word import atomic, WordJob, WordWorkflow
from studio_excel_catalog import normalize, catalog
from service_pricing import excel_price, VERSION

class ExcelWorkflow(WordWorkflow):
    def __init__(self, studio, root, build, send, charge, afford, parse, cost):
        super().__init__(studio, root, build, None, send, charge, afford)
        self.parse_callback, self.cost = parse, cost
        self.parsing = set()

    async def specs(self, request):
        self.studio.authorized(request)
        return web.json_response(catalog(self.cost), headers={'Cache-Control':'no-store'})

    async def draft(self, request):
        uid = self.studio.authorized(request)['user']['id']
        directory = self.root/'drafts'; directory.mkdir(exist_ok=True, mode=0o700)
        path = directory/(str(uid)+'.json')
        if request.method == 'POST':
            self.studio.check_write(request)
            try: value = normalize(await self.studio.body(request, 64000), draft=True)
            except ValueError as e: raise web.HTTPBadRequest(text=str(e))
            atomic(path, {'created':time.time(), 'payload':value})
            return web.json_response({'ok':True})
        try:
            value = json.loads(path.read_text())
            if time.time()-value['created'] > 86400: value=None; path.unlink(missing_ok=True)
        except (OSError, ValueError): value=None
        latest=None
        for p in self.root.glob('*/order.json'):
            try: row=json.loads(p.read_text())
            except (OSError, ValueError): continue
            if row['uid']==uid and (time.time()-row['created']<=86400 or row['status']=='delivery_unknown') and (latest is None or row['created']>latest['created']): latest=row
        return web.json_response({'ok':True,'draft':value,'order':self.public(WordJob(self,latest)) if latest else None}, headers={'Cache-Control':'no-store'})

    async def parse(self, request):
        uid=self.studio.authorized(request)['user']['id']; self.studio.check_write(request)
        try: value=normalize(await self.studio.body(request, 64000), draft=True)
        except ValueError as e: raise web.HTTPBadRequest(text=str(e))
        if not value['source']: raise web.HTTPBadRequest(text='Вставьте позиции, количество и цены.')
        if uid in self.parsing: raise web.HTTPConflict(text='Уже разбираем ваш текст.')
        directory=self.root/'imports'; directory.mkdir(exist_ok=True, mode=0o700)
        path=directory/(str(uid)+'.json'); now=time.time()
        try: old=json.loads(path.read_text())
        except (OSError, ValueError): old={}
        key=hashlib.sha256(value['source'].encode()).hexdigest()
        if old.get('key')==key and old.get('result') and now-old.get('time',0)<86400:
            return web.json_response({'ok':True, **old['result']}, headers={'Cache-Control':'no-store'})
        attempts=[t for t in old.get('attempts',[]) if now-t<86400]
        if len(attempts)>=20 or (attempts and now-attempts[-1]<15) or self.parsing:
            raise web.HTTPTooManyRequests(text='Разбор текста пока занят. Заполните строки вручную или попробуйте позже.')
        self.parsing.add(uid)
        record={'time':now,'key':key,'attempts':attempts+[now]}; atomic(path,record)
        try:
            async with asyncio.timeout(100): result=await self.parse_callback(value['source'])
            if not isinstance(result,dict) or not result.get('rows'): raise ValueError('Не удалось выделить позиции. Добавьте их вручную.')
            checked=normalize({**value,'title':result.get('title') or value['title'] or 'Смета','rows':result['rows']},draft=True)
            record['result']={'title':checked['title'],'rows':checked['rows']}; atomic(path,record)
            return web.json_response({'ok':True, **record['result']},headers={'Cache-Control':'no-store'})
        except ValueError as e: raise web.HTTPBadRequest(text=str(e))
        except TimeoutError: raise web.HTTPServiceUnavailable(text='Разбор занял слишком много времени. Можно заполнить строки вручную.')
        except Exception: raise web.HTTPServiceUnavailable(text='Не удалось разобрать текст. Исходные данные сохранены в форме.')
        finally: self.parsing.discard(uid)

    async def create(self, request):
        uid=self.studio.authorized(request)['user']['id']; self.studio.check_write(request)
        try: payload=normalize(await self.studio.body(request,64000))
        except ValueError as e: raise web.HTTPBadRequest(text=str(e))
        identity={k:v for k,v in payload.items() if k!='source' or payload['kind']!='expense_estimate'}
        fingerprint=hashlib.sha256((VERSION+json.dumps(identity,sort_keys=True)).encode()).hexdigest()
        rows=[]
        for path in self.root.glob('*/order.json'):
            try: r=json.loads(path.read_text())
            except (OSError,ValueError): continue
            if time.time()-r['created']>86400 and r['id'] not in self.tasks and r['status']!='delivery_unknown':
                shutil.rmtree(path.parent); continue
            rows.append(r)
            if r['uid']==uid:
                if r['fingerprint']==fingerprint: return self.response(WordJob(self,r))
                if r['id'] in self.tasks or r['status']=='delivery_unknown':
                    raise web.HTTPConflict(text='Сначала завершите предыдущую таблицу или уточните отправку в поддержке.')
        if len(rows)>=40: raise web.HTTPTooManyRequests(text='Достигнут лимит сохранённых заявок. Попробуйте позже.')
        token=secrets.token_urlsafe(24); (self.root/token).mkdir(mode=0o700)
        job=WordJob(self,dict(id=token,uid=uid,created=time.time(),fingerprint=fingerprint,payload=payload,
            cost=excel_price(payload['kind']),actual_cost=excel_price(payload['kind']),status='ready',stage='Проверьте таблицу и стоимость',title=payload['title'],sent=[]))
        job.save(); return self.response(job)

    async def confirm(self, request):
        user=self.studio.authorized(request)['user']; self.studio.check_write(request)
        job=self.get(user['id'],request.match_info['word_id']); r=job.row
        body=await self.studio.body(request,1024)
        if r['id'] in self.tasks or r['status']=='completed': return self.response(job)
        if r['status']=='delivery_unknown': raise web.HTTPConflict(text='Отправка или оплата требуют проверки. Проверьте чат и обратитесь в поддержку. Повторного списания нет.')
        if r['status'] not in ('ready','interrupted','failed'): raise web.HTTPConflict()
        if type(body.get('confirmedCost')) is not int or body['confirmedCost']!=r['cost']:
            raise web.HTTPConflict(text='Подтвердите актуальную стоимость.')
        if not r.get('billing_pending') and not self.afford(user['id'],r['cost']): raise self.studio.shortage(user['id'],r['cost'])
        if any(v[0]==user['id'] for v in self.tasks.values()): raise web.HTTPConflict(text='Предыдущая таблица ещё собирается.')
        r.update(status='building',stage='Собираем таблицу с формулами',notice=None); job.save()
        task=asyncio.create_task(self.run(user,job)); self.tasks[r['id']]=(user['id'],task)
        return self.response(job)

    async def run(self, user, job):
        try:
            async with asyncio.timeout(1200):
                if (job.directory/'document.xlsx').exists(): await self.deliver(job)
                else: await self.build(user,job)
            if job.row['status'] in ('building','delivering'):
                job.row.update(status='delivery_unknown' if job.row.get('sending') else 'interrupted',stage='Сборка не завершена. Данные сохранены.');job.save()
        except (Exception,asyncio.CancelledError) as e:
            print('studio_excel_error',type(e).__name__,flush=True)
            notice=str(e) if isinstance(e,ValueError) else e.text if isinstance(e,web.HTTPException) else 'Сборка прервалась. Данные сохранены, можно продолжить.'
            job.row.update(status='delivery_unknown' if job.row.get('sending') else 'interrupted',notice=notice); job.save()
        finally: self.tasks.pop(job.row['id'],None)

    async def deliver(self, job):
        r=job.row
        if not r.get('billing_pending') and not self.afford(r['uid'],r['cost']): raise web.HTTPPaymentRequired(text='Недостаточно кредитов.')
        if not r['sent']:
            r.update(status='delivering',stage='Отправляем Excel в чат',sending='xlsx'); job.save()
            await self.send(r['uid'],job.directory/'document.xlsx',r['title'])
            r['sent']=['xlsx']; r.pop('sending',None); job.save()
        r['billing_pending']=True; job.save()
        r['charged']=await self.charge(r['uid'],r['id'],r['cost'],r['title'])
        r.pop('billing_pending',None); r.update(status='completed',stage='Таблица готова',notice=None); job.save()

    def register(self,app):
        app.router.add_get('/studio/excel/catalog',self.specs)
        app.router.add_get('/studio/excel/draft',self.draft)
        app.router.add_post('/studio/excel/draft',self.draft)
        app.router.add_post('/studio/excel/parse',self.parse)
        app.router.add_post('/studio/excel/orders',self.create)
        app.router.add_get('/studio/excel/orders/{word_id}',self.status)
        app.router.add_post('/studio/excel/orders/{word_id}/confirm',self.confirm)
        app.on_cleanup.append(self.cleanup)
