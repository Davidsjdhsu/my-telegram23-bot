"""Connect the private Word workflow to the existing document builders."""
import asyncio
import copy
from pathlib import Path
import shutil
import tempfile
from types import SimpleNamespace
from aiohttp import web
from pypdf import PdfReader
from studio_word import WordState, WordWorkflow
from studio_word_catalog import SPECS, render_preview

class QuietMessage:
    def __init__(self,uid,job):
        self.from_user=SimpleNamespace(id=uid);self.chat=SimpleNamespace(id=uid);self.job=job
    async def answer(self,text,**kwargs):
        self.job.row['notice']=str(text)[:2000];self.job.save()
        return SimpleNamespace(delete=self.noop)
    async def noop(self,*a,**kw):pass
    async def answer_document(self,*a,**kw):raise RuntimeError('Use durable Word delivery')


def attach(studio,app,api):
    async def preview(payload):
        with tempfile.TemporaryDirectory(prefix='studio-word-preview-') as d:
            docx,pdf=Path(d)/'preview.docx',Path(d)/'preview.pdf'
            await api['run_heavy'](render_preview,docx,payload,api['build_word'])
            await api['convert_word_pdf'](str(docx),str(pdf),timeout=40,queue_timeout=20)
            if SPECS[payload['kind']]['cover'] and len(PdfReader(pdf).pages)!=1:
                raise ValueError('Титульный не помещается на одну страницу. Сократите длинные поля или тему.')
            command=shutil.which('pdftoppm')
            if not command:raise RuntimeError('Предпросмотр временно недоступен.')
            p=await asyncio.create_subprocess_exec(command,'-f','1','-singlefile','-scale-to','1100','-png',str(pdf),str(Path(d)/'page'),stdout=asyncio.subprocess.DEVNULL,stderr=asyncio.subprocess.DEVNULL)
            try:await asyncio.wait_for(p.wait(),20)
            finally:
                if p.returncode is None:p.kill();await p.wait()
            if p.returncode:raise RuntimeError('Не удалось подготовить предпросмотр.')
            return (Path(d)/'page.png').read_bytes()

    async def build(user,job,action):
        uid=user['id'];payload=job.row['payload'];kind=payload['kind'];fields=payload['fields']
        if api['get_user'](uid).get('busy') or uid in api['_miniapp_action_tasks']:
            raise web.HTTPConflict(text='Предыдущая работа ещё выполняется.')
        api['miniapp_user_record'](user)
        if SPECS[kind]['cover'] and not job.row.get('cover_checked'):
            job.stage('Проверяем заполнение титульного листа')
            await preview(payload)
            job.row['cover_checked']=True;job.save()
        if SPECS[kind]['fixed']:
            ok,reason=api['start_job'](uid)
            if not ok:raise web.HTTPConflict(text=reason)
            try:
                job.stage('Заполняем выбранный документ')
                with tempfile.TemporaryDirectory(prefix='studio-word-build-') as directory:
                    docx,pdf=Path(directory)/'filled.docx',Path(directory)/'filled.pdf'
                    await api['run_heavy'](render_preview,docx,payload,api['build_word'])
                    job.stage('Проверяем Word и PDF')
                    await api['convert_word_pdf'](str(docx),str(pdf))
                    await job.deliver(docx,pdf,job.row['cost'],SPECS[kind]['title'])
            finally:api['finish_job'](uid)
            return
        state=WordState(job)
        if not job.row.get('data'):
            data=dict(word_kind=kind,word_size='short' if payload['pages']<=3 else 'long',word_pages=payload['pages'],
                topic=payload['topic'],user_text=payload['content'],extra='',content_lang=payload['content_lang'],
                extra_used=0,car_sale=None,fixed_word=None,studio_word_fields=fields)
            from academic_templates import title_request_key
            data['academic_cover_allow_blanks']=title_request_key(data)
            await state.set_data(data)
        await api['word_build'](QuietMessage(uid,job),state,quality_action=action if action in ('accept','refine') else None,studio_job=job)

    async def send(uid,path,title,ext):
        filename=api['safe_filename'](title,fallback='Документ')+'.'+ext
        message=await api['bot'].send_document(uid,api['FSInputFile'](str(path),filename=filename),caption='Ваш документ · '+ext.upper())
        if api.get('_file_library'): api['_file_library'].record(uid,message,title,ext,str(path.parent))
        return message

    from order_billing import charge_callback
    charge=charge_callback(api, 'studio_word_payments', 'Word')

    root=Path(api['USERS_FILE']).parent/'studio-word'
    service=WordWorkflow(studio,root,build,preview,send,charge,api['can_afford'])
    service.register(app)
    return service
