import ast,asyncio,json,tempfile,unittest,os
from pathlib import Path
from datetime import datetime
from types import SimpleNamespace
from unittest.mock import AsyncMock,Mock
from aiohttp import web
from openpyxl import load_workbook
from openpyxl.utils.protection import hash_password
from studio_excel_runtime import attach
from studio_word import WordJob
from studio_excel_catalog import normalize
from test_office import office_env,COLORS
from test_excel_startup import MODEL
from test_excel_study import MOTION
from test_presentation_support import load_functions

class AllKinds(unittest.IsolatedAsyncioTestCase):
 async def test_existing_builders_receive_saved_data_password_and_deliver_once(self):
  fixtures={'price_list':[{'name':'Товар','qty':2,'price':150,'discount':10}],
    'family_budget':[{'name':'Доход','type':'income','amount':1000},{'name':'Расход','type':'expense','amount':200}],
    'project_budget':[{'name':'Доход','type':'income','amount':1000},{'name':'Расход','type':'expense','amount':200}],
    'calc_table':MOTION,'startup_model':MODEL}
  for kind,fixture in fixtures.items():
   with self.subTest(kind=kind),tempfile.TemporaryDirectory() as d:
    user={'credits':100,'history':[],'generations':0};sent=[]
    async def send(uid,file,**kw):
     wb=load_workbook(file);self.assertTrue(all(s.protection.password==hash_password('Test-2026') for s in wb));wb.close();sent.append(file)
    async def heavy(fn,*a,**kw):return fn(*a,**kw)
    parsed=dict(MODEL,title='Тест') if kind=='startup_model' else {'title':'Тест','rows':fixture}
    env=office_env();env.update(USERS_FILE=d+'/users.json',get_user=lambda uid:user,can_afford=lambda uid,c:user['credits']>=c,
      miniapp_user_record=lambda u:user,_miniapp_action_tasks={},start_job=Mock(return_value=(True,'')),finish_job=Mock(),run_heavy=heavy,
      bot=SimpleNamespace(send_document=send),FSInputFile=lambda p,**kw:p,safe_filename=lambda v,**kw:v,
      save_users_now=AsyncMock(return_value=True),ADMIN_IDS=set(),note_success=Mock(),CREDIT_COSTS={'excel':5},
      ask_grok=AsyncMock(return_value=json.dumps(parsed)),extract_json=json.loads,user_lang=lambda uid:'ru',
      tr=lambda *a:'Собираем таблицу',pick_theme=lambda t:('default',COLORS),grok_json_lang_instruction=lambda lang:'',
      content_gen_lang=lambda *a,**kw:'ru',datetime=datetime,os=os)
    tree=ast.parse(Path('main.py').read_text())
    for node in tree.body:
     if isinstance(node,ast.Assign) and any(isinstance(t,ast.Name) and t.id in ('EXCEL_ROW_SCHEMA','EXCEL_MIN_ROWS_AI','EXCEL_KIND_DESC') for t in node.targets):exec(compile(ast.Module(body=[node],type_ignores=[]),'fixture','exec'),env)
    env=load_functions(['excel_build'],env)
    service=attach(SimpleNamespace(),web.Application(),env)
    payload=normalize({'kind':kind,'title':'Тест','source':'Мои данные','excel_protect':True,'excel_password':'Test-2026'})
    (service.root/'one').mkdir();job=WordJob(service,dict(id='one',uid=-903099,payload=payload,cost=5,sent=[],status='building'))
    job.save();await service.run({'id':-903099},job)
    self.assertEqual(job.row['status'],'completed');self.assertEqual(user['credits'],95);self.assertEqual(len(sent),1)
    await service.run({'id':-903099},job);self.assertEqual(user['credits'],95);self.assertEqual(len(sent),1)
    env['ask_grok'].assert_awaited_once()
    Path('/tmp/xl_-903099.xlsx').unlink(missing_ok=True)
