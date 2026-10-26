import asyncio
import copy
import json
import time
import unittest
from types import SimpleNamespace
from unittest.mock import AsyncMock, Mock
from aiohttp import web, CookieJar
from aiohttp.test_utils import TestClient, TestServer
from studio_private import Studio, COOKIE, PILOT_UID
from test_presentation_support import load_functions
from presentation_support import plan_key
from presentation_pricing import presentation_cost

HEADERS={'Origin':'https://example.test','X-Studio-Request':'1'}
PAYLOAD={'topic':'Как работает кофе','kind':'standard','key':'coffee','slides':4}
PLAN={'design':{},'title':'Кофе','audience':'широкая аудитория','slides':[{'title':x,'number':i+1,'point':'Internal outline point'} for i,x in enumerate(['Кофе','Зачем пьют кофе','Где выращивают','Путь до чашки'])]}
DATA={'slides':4,'topic':'Как работает кофе','theme_name':'coffee','pres_plan':PLAN,'pres_plan_key':'saved-key'}

class OutlineTests(unittest.IsolatedAsyncioTestCase):
 async def asyncSetUp(self):
  self.planner=AsyncMock(return_value=(copy.deepcopy(DATA),{'bg':'#49362b','ink':'#f2e3ce'}));self.build=asyncio.get_running_loop().create_future();self.submit=AsyncMock(return_value=self.build);self.send=AsyncMock()
  self.studio=Studio(lambda _:None,self.send,AsyncMock(),'https://example.test',outline_planner=self.planner,outline_submit=self.submit)
  self.studio.sessions['owner']={'user':{'id':PILOT_UID},'expires':time.monotonic()+600}
  app=web.Application();self.studio.register(app);self.client=TestClient(TestServer(app),cookie_jar=CookieJar(unsafe=True));await self.client.start_server()
  self.client.session.cookie_jar.update_cookies({COOKIE:'owner'})
 async def asyncTearDown(self):
  await self.client.close()
  for z in self.studio.archives:z.close()
 async def create(self):
  r=await self.client.post('/studio/outline',json=PAYLOAD,headers=HEADERS);self.assertEqual(r.status,200);row=await r.json();await self.studio.outline.rows[row['id']]['task'];return row['id']
 async def test_preview_only_titles_and_reuses_same_outline(self):
  token=await self.create();r=await self.client.get('/studio/outline/'+token);body=await r.json()
  self.assertEqual(body['titles'],[s['title'] for s in PLAN['slides']]);self.assertEqual(body['cost'],64)
  self.assertNotIn('Internal outline point',json.dumps(body));self.submit.assert_not_awaited();self.send.assert_not_awaited()
  again=await self.client.post('/studio/outline',json=PAYLOAD,headers=HEADERS);self.assertEqual((await again.json())['id'],token);self.planner.assert_awaited_once()
 async def test_confirmation_uses_server_plan_and_cannot_repeat(self):
  token=await self.create();url='/studio/outline/'+token+'/confirm'
  for _ in range(2):
   r=await self.client.post(url,json={'cost':0,'recipient':123,'pres_plan':{}},headers=HEADERS);self.assertEqual(r.status,200);self.assertEqual((await r.json())['status'],'building')
  self.submit.assert_awaited_once();user,data=self.submit.call_args.args;self.assertEqual(user['id'],PILOT_UID);self.assertEqual(data,DATA)
 async def test_inflight_double_confirm_does_not_resubmit(self):
  token=await self.create();entered=asyncio.Event();release=asyncio.Event()
  async def slow(*args):entered.set();await release.wait()
  self.submit.side_effect=slow;url='/studio/outline/'+token+'/confirm'
  first=asyncio.create_task(self.client.post(url,json={},headers=HEADERS));await entered.wait()
  r=await self.client.post(url,json={},headers=HEADERS);self.assertEqual(r.status,409);release.set();self.assertEqual((await first).status,200);self.submit.assert_awaited_once()
 async def test_build_completion_is_observed_and_duplicate_confirm_never_restarts(self):
  token=await self.create();url='/studio/outline/'+token
  await self.client.post(url+'/confirm',json={},headers=HEADERS)
  self.assertEqual((await (await self.client.get(url)).json())['status'],'building')
  self.build.set_result({'delivered':True});await asyncio.sleep(0)
  done=await (await self.client.get(url)).json();self.assertEqual(done['status'],'completed');self.assertEqual(len(done['titles']),4)
  again=await self.client.post(url+'/confirm',json={},headers=HEADERS);self.assertEqual((await again.json())['status'],'completed');self.submit.assert_awaited_once()
 async def test_handled_build_failure_never_looks_successful(self):
  token=await self.create();url='/studio/outline/'+token
  await self.client.post(url+'/confirm',json={},headers=HEADERS)
  self.build.set_result(None);await asyncio.sleep(0)
  result=await (await self.client.get(url)).json();self.assertEqual(result['status'],'build_failed');self.assertTrue(result['error'])
  await self.client.post(url+'/confirm',json={},headers=HEADERS);self.assertEqual(self.submit.await_count,2)
 async def test_exception_and_cancelled_build_are_not_success(self):
  token=await self.create();row=self.studio.outline.rows[token]
  for cancelled in [False,True]:
   future=asyncio.get_running_loop().create_future()
   if cancelled:future.cancel()
   else:future.set_exception(RuntimeError('private details'))
   self.studio.outline.build_finished(row,future);self.assertEqual(row['status'],'build_failed');self.assertNotIn('private details',row['error'])
 async def test_balance_failure_retriable_and_expiry(self):
  token=await self.create();url='/studio/outline/'+token
  self.submit.side_effect=web.HTTPPaymentRequired(text='Недостаточно кредитов')
  self.assertEqual((await self.client.post(url+'/confirm',json={},headers=HEADERS)).status,402)
  self.assertEqual(self.studio.outline.rows[token]['status'],'ready');self.submit.side_effect=None
  self.studio.outline.rows[token]['expires']=0
  self.assertEqual((await self.client.post(url+'/confirm',json={},headers=HEADERS)).status,404)
 async def test_auth_origin_validation_and_no_client_price(self):
  token=await self.create()
  self.assertEqual((await self.client.post('/studio/outline/'+token+'/confirm',json={},headers={'Origin':'https://evil.test','X-Studio-Request':'1'})).status,403)
  for change in [{'kind':'designer'},{'slides':True},{'key':'not-a-template'},{'topic':''}]:
   self.assertEqual((await self.client.post('/studio/outline',json={**PAYLOAD,**change},headers=HEADERS)).status,400)
  self.studio.sessions['owner']['user']['id']=123
  self.assertEqual((await self.client.get('/studio/outline/'+token)).status,404)
  self.client.session.cookie_jar.clear()
  self.assertEqual((await self.client.post('/studio/outline',json=PAYLOAD,headers=HEADERS)).status,401)
  self.assertEqual((await self.client.get('/studio/outline/'+token)).status,401)
 async def test_designer_preview_price_and_saved_settings(self):
  for style in ['strict','soft']:
   self.studio.outline.rows.clear()
   payload={**PAYLOAD,'kind':'designer','style':style,'key':'d16','slides':6}
   r=await self.client.post('/studio/outline',json=payload,headers=HEADERS);self.assertEqual(r.status,200)
   body=await r.json();self.assertEqual(body['cost'],192)
   row=self.studio.outline.rows[body['id']];await row['task']
   self.assertEqual(self.planner.call_args.args[1],{**payload,'background_mode':'solid','diagram_layout':'classic'})
   r=await self.client.post('/studio/outline/'+body['id']+'/confirm',json={'cost':0},headers=HEADERS)
   self.assertEqual((await r.json())['cost'],192)
  for change in [{'style':'invalid'},{'key':'coffee'},{'key':'d999'}]:
   r=await self.client.post('/studio/outline',json={**payload,**change},headers=HEADERS);self.assertEqual(r.status,400)
 async def test_tester_plans_and_confirmation_are_isolated(self):
  owner_token=await self.create()
  self.studio.sessions['tester']={'user':{'id':909828109},'expires':time.monotonic()+600}
  self.client.session.cookie_jar.update_cookies({COOKIE:'tester'})
  self.assertEqual((await self.client.get('/studio/outline/'+owner_token)).status,404)
  self.assertEqual((await self.client.post('/studio/outline/'+owner_token+'/confirm',json={},headers=HEADERS)).status,404)
  tester_token=await self.create();self.assertNotEqual(tester_token,owner_token)
  self.assertEqual(self.planner.call_args.args[0]['id'],909828109)
  r=await self.client.post('/studio/outline/'+tester_token+'/confirm',json={},headers=HEADERS)
  self.assertEqual(r.status,200);self.assertEqual(self.submit.call_args.args[0]['id'],909828109)
  self.client.session.cookie_jar.update_cookies({COOKIE:'owner'})
  self.assertEqual((await self.client.get('/studio/outline/'+tester_token)).status,404)
 async def test_failed_planner_no_content_or_submission(self):
  self.planner.side_effect=ValueError('private provider diagnostic')
  token=await self.create();r=await self.client.get('/studio/outline/'+token);body=await r.json()
  self.assertEqual(body['status'],'failed');self.assertNotIn('private provider',json.dumps(body));self.submit.assert_not_awaited()

class AdapterTests(unittest.IsolatedAsyncioTestCase):
 async def test_real_builder_receives_saved_outline_without_replanning(self):
  state=SimpleNamespace(set_data=AsyncMock(),set_state=AsyncMock());builder=AsyncMock(return_value={'delivered':True});tasks={PILOT_UID:object()}
  env={'miniapp_user_record':Mock(),'FSMContext':Mock(return_value=state),'StorageKey':Mock(),'dp':SimpleNamespace(storage=object()),'bot':SimpleNamespace(id=1,send_message=AsyncMock()),'Form':SimpleNamespace(waiting_confirm='confirmed'),'_build_presentation':builder,'ensure_answerable':lambda m:m,'_FakeMessageForApi':lambda uid,name:SimpleNamespace(from_user=SimpleNamespace(id=uid)),'_miniapp_action_tasks':tasks}
  run=load_functions(['run_studio_preview'],env)['run_studio_preview'];result=await run({'id':PILOT_UID},DATA);self.assertEqual(result,{'delivered':True})
  state.set_data.assert_awaited_once_with(DATA);builder.assert_awaited_once();self.assertFalse(tasks)
 async def test_submit_rechecks_balance_and_existing_job_before_start(self):
  job=AsyncMock();tasks={};user={'id':PILOT_UID};afford=Mock(return_value=False)
  env={'get_user':lambda uid:{},'_miniapp_action_tasks':tasks,'MAX_CONCURRENT_GENERATIONS':2,'can_afford':afford,'presentation_cost':presentation_cost,'_aiohttp_web':web,'asyncio':asyncio,'run_studio_preview':job}
  submit=load_functions(['studio_submit_preview'],env)['studio_submit_preview']
  with self.assertRaises(web.HTTPPaymentRequired):await submit(user,DATA)
  self.assertFalse(tasks);afford.return_value=True;tasks[PILOT_UID]=object()
  with self.assertRaises(web.HTTPConflict):await submit(user,DATA)
  job.assert_not_awaited()
 async def test_planner_only_creates_plan_and_binds_settings(self):
  import uuid
  creator=AsyncMock(return_value=PLAN);color={'bg':(73,54,43),'ink':(242,227,206)}
  env={'get_user':lambda uid:{},'_miniapp_action_tasks':{},'user_lang':lambda uid:'ru','uuid':uuid,'can_afford':lambda *a:True,'presentation_cost':presentation_cost,'_create_presentation_plan':creator,'plan_key':plan_key,'format_plan':lambda *a:'summary','apply_design':lambda color,plan:color,'THEMES':{'coffee':color,'default':color}}
  fn=load_functions(['studio_plan_preview'],env)['studio_plan_preview'];data,palette=await fn({'id':PILOT_UID},PAYLOAD)
  self.assertEqual(data['pres_plan'],PLAN);self.assertEqual(data['pres_plan_key'],plan_key(data));self.assertEqual(data['theme_name'],'coffee');self.assertEqual(palette['bg'],'#49362b');creator.assert_awaited_once()

 async def test_designer_planner_preserves_template_style_and_price(self):
  import uuid
  from presentation_profiles import choose_design
  for style in ['strict','soft']:
   creator=AsyncMock(return_value=PLAN);color={'bg':(20,41,59),'ink':(242,227,206)}
   afford=Mock(return_value=True)
   env={'get_user':lambda uid:{},'_miniapp_action_tasks':{},'user_lang':lambda uid:'ru','uuid':uuid,'can_afford':afford,'presentation_cost':presentation_cost,'_create_presentation_plan':creator,'plan_key':plan_key,'format_plan':lambda *a:'summary','apply_design':lambda color,plan:color,'THEMES':{'default':color}}
   fn=load_functions(['studio_plan_preview'],env)['studio_plan_preview']
   data,palette=await fn({'id':909828109},{**PAYLOAD,'kind':'designer','style':style,'key':'d16','slides':6})
   self.assertEqual(data['theme_name'],'designer_'+style);self.assertEqual(choose_design(data)['template_id'],'d16')
   afford.assert_called_once_with(909828109,192);self.assertEqual(data['pres_plan_key'],plan_key(data))

class DesignerChoiceDeliveryTests(OutlineTests):
 async def test_choices_reach_submission_and_final_plan_at_same_price(self):
  payload={**PAYLOAD,'kind':'designer','style':'soft','key':'d42','background_mode':'gradient','diagram_layout':'cards'}
  r=await self.client.post('/studio/outline',json=payload,headers=HEADERS)
  token=(await r.json())['id'];row=self.studio.outline.rows[token];await row['task']
  self.assertEqual(row['status'],'ready')
  self.assertEqual(row['data']['pres_plan']['design']['diagram_layout'],'cards')
  self.assertEqual(row['data']['studio_design_options']['background_mode'],'gradient')
  self.assertEqual(row['data']['pres_plan_key'],plan_key(row['data']))
  await self.client.post('/studio/outline/'+token+'/confirm',headers=HEADERS,json={'diagram_layout':'classic','background_mode':'solid'})
  self.assertEqual(self.submit.call_args.args[1]['studio_design_options']['diagram_layout'],'cards')
  self.assertEqual(row['cost'],128)
 async def test_unknown_choices_rejected(self):
  for field in ('diagram_layout','background_mode'):
   r=await self.client.post('/studio/outline',json={**PAYLOAD,'kind':'designer','style':'strict','key':'d16',field:'bad'},headers=HEADERS)
   self.assertEqual(r.status,400)
