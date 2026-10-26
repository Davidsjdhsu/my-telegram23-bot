"""Private Word workflow: authenticated previews, durable jobs and delivery receipts."""
import asyncio
import copy
import hashlib
import json
import os
from pathlib import Path
import secrets
import shutil
import tempfile
import time
from service_pricing import VERSION
from aiohttp import web
from studio_word_catalog import catalog, normalize, price, SPECS

ACTIVE={'building','delivering'}

def atomic(path,value):
    tmp=path.with_suffix('.tmp')
    with open(tmp,'w',encoding='utf-8') as f:
        json.dump(value,f,ensure_ascii=False);f.flush();os.fsync(f.fileno())
    os.chmod(tmp,0o600);os.replace(tmp,path)

class WordState:
    def __init__(self,job):self.job=job
    async def get_data(self):return copy.deepcopy(self.job.row.get('data',{}))
    async def set_data(self,data):self.job.row['data']=copy.deepcopy(data);self.job.save()
    async def update_data(self,**data):self.job.row.setdefault('data',{}).update(copy.deepcopy(data));self.job.save()
    async def set_state(self,state):self.job.row['fsm']=str(state);self.job.save()
    async def clear(self):self.job.row['data']={};self.job.save()

class WordJob:
    def __init__(self,service,row):self.service,self.row=service,row
    @property
    def directory(self):return self.service.root/self.row['id']
    def save(self):atomic(self.directory/'order.json',self.row)
    def stage(self,text):self.row['stage']=text;self.save()
    def choice(self,cost,notice):
        self.row.update(status='choice',actual_cost=cost,notice=notice,stage='Нужно ваше решение');self.save()
    async def deliver(self,docx,pdf,cost,title):
        if cost>self.row['cost']:raise ValueError('Стоимость превышает подтверждённую.')
        self.row.update(actual_cost=cost,title=title)
        for ext,source in [('docx',docx),('pdf',pdf)]:
            shutil.copyfile(source,self.directory/('document.'+ext))
        self.row.update(status='delivering',stage='Отправляем Word и PDF');self.save()
        return await self.service.deliver(self)

class WordWorkflow:
    def __init__(self,studio,root,build,preview,send,charge,afford):
        self.studio=studio;self.root=Path(root);self.root.mkdir(parents=True,exist_ok=True,mode=0o700)
        self.build,self.preview_callback,self.send,self.charge,self.afford=build,preview,send,charge,afford
        self.tasks={};self.preview_users=set();self.preview_cache={}
        # Delivery intent survives restart. Never guess whether Telegram received it.
        for p in self.root.glob('*/order.json'):
            try:
                row=json.loads(p.read_text())
                if row['status'] in ACTIVE:
                    row.update(status='delivery_unknown' if row.get('sending') else 'interrupted',stage='Работа прервалась. Можно продолжить.',restart_notice=True)
                    atomic(p,row)
            except (OSError,ValueError,KeyError):pass

    async def notify_restarts(self, notify):
        for path in self.root.glob('*/order.json'):
            try:
                row=json.loads(path.read_text())
                if not row.get('restart_notice'): continue
                text='Сборка файла прервалась при перезапуске. Данные сохранены. Откройте соответствующий раздел Word или Excel в меню и нажмите «Продолжить сборку».'
                if row['status']=='delivery_unknown': text='Связь прервалась при отправке файла. Проверьте чат и обратитесь в поддержку. Автоматического повторного списания не будет.'
                await notify(row['uid'],text)
                row.pop('restart_notice',None);atomic(path,row)
            except Exception: continue

    def public(self,job):
        r=job.row
        return {k:r.get(k) for k in ('id','status','stage','cost','actual_cost','notice','charged','title')}|{'ok':True}

    def response(self,job):
        return web.json_response(self.public(job),headers={'Cache-Control':'no-store'})

    def get(self,uid,token):
        if not isinstance(token,str) or not __import__('re').fullmatch(r'[A-Za-z0-9_-]{24,60}',token):raise web.HTTPNotFound()
        try:row=json.loads((self.root/token/'order.json').read_text())
        except (OSError,ValueError):raise web.HTTPNotFound()
        if row['uid']!=uid:raise web.HTTPNotFound()
        return WordJob(self,row)

    async def specs(self,request):
        self.studio.authorized(request)
        return web.json_response({'ok':True,'documents':catalog()},headers={'Cache-Control':'no-store'})

    async def preview(self,request):
        uid=self.studio.authorized(request)['user']['id'];self.studio.check_write(request)
        try:payload=normalize(await self.studio.body(request,64000),preview=True)
        except ValueError as e:raise web.HTTPBadRequest(text=str(e))
        key=hashlib.sha256(json.dumps(payload,sort_keys=True).encode()).hexdigest()
        cached=self.preview_cache.get(uid)
        if cached and cached[0]==key and time.monotonic()-cached[2]<300:
            return web.Response(body=cached[1],content_type='image/png',headers={'Cache-Control':'no-store'})
        if uid in self.preview_users or self.preview_users or self.tasks:
            raise web.HTTPTooManyRequests(text='Предпросмотр занят. Повторите через несколько секунд.')
        self.preview_users.add(uid)
        try:
            async with asyncio.timeout(75):image=await self.preview_callback(payload)
            self.preview_cache[uid]=(key,image,time.monotonic())
            return web.Response(body=image,content_type='image/png',headers={'Cache-Control':'no-store'})
        except (ValueError,RuntimeError) as e:raise web.HTTPBadRequest(text=str(e))
        except TimeoutError:raise web.HTTPServiceUnavailable(text='Предпросмотр не успел загрузиться. Попробуйте ещё раз.')
        finally:self.preview_users.discard(uid)

    async def draft(self,request):
        uid=self.studio.authorized(request)['user']['id']
        directory=self.root/'drafts';directory.mkdir(exist_ok=True,mode=0o700);path=directory/(str(uid)+'.json')
        if request.method=='POST':
            self.studio.check_write(request)
            try:value=normalize(await self.studio.body(request,64000),preview=True)
            except ValueError as e:raise web.HTTPBadRequest(text=str(e))
            atomic(path,{'created':time.time(),'payload':value})
            return web.json_response({'ok':True})
        try:
            value=json.loads(path.read_text())
            if time.time()-value['created']>86400:
                value=None;path.unlink(missing_ok=True)
        except (OSError,ValueError):value=None
        latest=None
        for p in self.root.glob('*/order.json'):
            try:r=json.loads(p.read_text())
            except (OSError,ValueError):continue
            if r['uid']==uid and (time.time()-r['created']<=86400 or r['status']=='delivery_unknown') and (latest is None or r['created']>latest['created']):latest=r
        return web.json_response({'ok':True,'draft':value,'order':self.public(WordJob(self,latest)) if latest else None},headers={'Cache-Control':'no-store'})

    async def create(self,request):
        uid=self.studio.authorized(request)['user']['id'];self.studio.check_write(request)
        try:payload=normalize(await self.studio.body(request,64000))
        except ValueError as e:raise web.HTTPBadRequest(text=str(e))
        fingerprint=hashlib.sha256((VERSION+json.dumps(payload,sort_keys=True)).encode()).hexdigest()
        # Small pilot: bounded persisted order cache; no personal data in public assets.
        rows=[]
        for path in self.root.glob('*/order.json'):
            try:r=json.loads(path.read_text())
            except (OSError,ValueError):continue
            if time.time()-r['created']>86400 and r['id'] not in self.tasks and r['status']!='delivery_unknown':
                shutil.rmtree(path.parent);continue
            rows.append(r)
            if r['uid']==uid:
                if r['fingerprint']==fingerprint:return self.response(WordJob(self,r))
                if r['id'] in self.tasks or r['status']=='delivery_unknown':raise web.HTTPConflict(text='Сначала завершите предыдущую работу или уточните отправку в поддержке.')
        if len(rows)>=40:raise web.HTTPTooManyRequests(text='Достигнут лимит сохранённых заявок. Попробуйте позже.')
        token=secrets.token_urlsafe(24);(self.root/token).mkdir(mode=0o700)
        job=WordJob(self,dict(id=token,uid=uid,created=time.time(),fingerprint=fingerprint,payload=payload,
            cost=price(payload),status='ready',stage='Проверьте данные и стоимость',title=payload['topic'] or SPECS[payload['kind']]['title'],sent=[]))
        job.save();return self.response(job)

    async def status(self,request):
        uid=self.studio.authorized(request)['user']['id']
        return self.response(self.get(uid,request.match_info['word_id']))

    async def confirm(self,request):
        user=self.studio.authorized(request)['user'];self.studio.check_write(request)
        job=self.get(user['id'],request.match_info['word_id']);r=job.row
        body=await self.studio.body(request,1024);action=body.get('action','build')
        if r['id'] in self.tasks or r['status']=='completed':return self.response(job)
        if r['status']=='delivery_unknown':raise web.HTTPConflict(text='Статус отправки неизвестен. Проверьте чат и обратитесь в поддержку. Повторного списания нет.')
        if r['status']=='choice':
            if action not in ('accept','refine'):raise web.HTTPBadRequest(text='Выберите доработку или получение текущего варианта.')
        elif action!='build' or r['status'] not in ('ready','failed','interrupted'):raise web.HTTPBadRequest()
        expected=r.get('actual_cost') if action=='accept' else r['cost']
        if type(body.get('confirmedCost')) is not int or body['confirmedCost']!=expected:raise web.HTTPConflict(text='Подтвердите актуальную стоимость.')
        if not r.get('billing_pending') and not self.afford(user['id'],expected):raise web.HTTPPaymentRequired(text='Недостаточно кредитов. Пополните баланс в профиле.')
        if any(value[0]==user['id'] for value in self.tasks.values()):raise web.HTTPConflict(text='Предыдущий документ ещё собирается.')
        r.update(status='building',stage='Готовим содержание',notice=None);job.save()
        task=asyncio.create_task(self.run(user,job,action));self.tasks[r['id']]=(user['id'],task)
        return self.response(job)

    async def run(self,user,job,action):
        try:
            async with asyncio.timeout(1200):
                if job.row.get('sent') or all((job.directory/('document.'+ext)).exists() for ext in ('docx','pdf')):
                    await self.deliver(job)
                else:await self.build(user,job,action)
            if job.row['status'] in ACTIVE:
                job.row.update(status='delivery_unknown' if job.row.get('sending') else 'failed',notice=job.row.get('notice') or 'Не удалось завершить сборку. Данные сохранены. Кредиты не списаны.');job.save()
        except (Exception,asyncio.CancelledError) as e:
            # Only exception classes in logs; field values and documents stay private.
            print('studio_word_error',type(e).__name__,flush=True)
            notice=str(e) if isinstance(e,ValueError) else e.text if isinstance(e,web.HTTPException) else 'Работа прервалась. Данные сохранены. Можно продолжить.'
            job.row.update(status='delivery_unknown' if job.row.get('sending') else 'interrupted',notice=notice);job.save()
        finally:self.tasks.pop(job.row['id'],None)

    async def deliver(self,job):
        r=job.row;cost=r['actual_cost']
        if not r.get('billing_pending') and len(r['sent'])<2 and not self.afford(r['uid'],cost):raise web.HTTPPaymentRequired()
        for ext in ('docx','pdf'):
            if ext in r['sent']:continue
            r['sending']=ext;job.save()
            await self.send(r['uid'],job.directory/('document.'+ext),r['title'],ext)
            r['sent'].append(ext);r.pop('sending',None);job.save()
        # Paid marker and balance live in the same durable user record.
        r['billing_pending']=True;job.save()
        r['charged']=await self.charge(r['uid'],r['id'],cost,r['title'])
        r.pop('billing_pending',None)
        r.update(status='completed',stage='Документ готов',notice=None);job.save()
        return {'delivered':True}

    async def cleanup(self,app):
        tasks=[v[1] for v in self.tasks.values()]
        for task in tasks:task.cancel()
        await asyncio.gather(*tasks,return_exceptions=True)

    def register(self,app):
        app.router.add_get('/studio/word/catalog',self.specs)
        app.router.add_get('/studio/word/draft',self.draft)
        app.router.add_post('/studio/word/draft',self.draft)
        app.router.add_post('/studio/word/preview',self.preview)
        app.router.add_post('/studio/word/orders',self.create)
        app.router.add_get('/studio/word/orders/{word_id}',self.status)
        app.router.add_post('/studio/word/orders/{word_id}/confirm',self.confirm)
        app.on_cleanup.append(self.cleanup)
