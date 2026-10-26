import ast, asyncio, copy, json, time, unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, Mock
from aiohttp import web, CookieJar
from aiohttp.test_utils import TestClient, TestServer
from studio_private import Studio, COOKIE, PILOT_UID
from studio_settings import PALETTES, preferences, content_language
from presentation_profiles import choose_design, resolved_profile, relative_luminance
from presentation_style_direction import has_custom_palette
from presentation_support import plan_key
from test_studio_outline import load_functions, PAYLOAD, PLAN, HEADERS

class SettingsTests(unittest.IsolatedAsyncioTestCase):
 async def asyncSetUp(self):
  self.callback=AsyncMock();self.studio=Studio(lambda _:None,AsyncMock(),AsyncMock(),'https://example.test',settings=self.callback)
  self.studio.sessions['owner']={'user':{'id':PILOT_UID},'expires':time.monotonic()+600}
  app=web.Application();self.studio.register(app);self.client=TestClient(TestServer(app),cookie_jar=CookieJar(unsafe=True));await self.client.start_server();self.client.session.cookie_jar.update_cookies({COOKIE:'owner'})
 async def asyncTearDown(self):
  await self.client.close()
  for z in self.studio.archives:z.close()
 async def test_preferences_only_for_signed_owner_and_allowed_values(self):
  values={'ui_lang':'en','content_lang':'de'}
  r=await self.client.post('/studio/settings',json=values,headers=HEADERS);self.assertEqual(r.status,200);self.callback.assert_awaited_once_with({'id':PILOT_UID},values)
  for body in [{'ui_lang':'en'},{**values,'content_lang':'xx'},{**values,'ui_lang':[]},{**values,'uid':99}]:
   self.assertEqual((await self.client.post('/studio/settings',json=body,headers=HEADERS)).status,400)
  self.assertEqual((await self.client.post('/studio/settings',json=values,headers={**HEADERS,'Origin':'https://other.test'})).status,403)
  self.client.session.cookie_jar.clear();self.assertEqual((await self.client.post('/studio/settings',json=values,headers=HEADERS)).status,401)
  self.callback.assert_awaited_once()
 async def test_independent_content_language_and_palette_enter_saved_plan(self):
  import uuid
  creator=AsyncMock(return_value=copy.deepcopy(PLAN));color={'bg':(12,24,32),'ink':(240,240,240)}
  env={'get_user':lambda uid:{'studio_content_lang':'fr'},'_miniapp_action_tasks':{},'user_lang':lambda uid:'ru','uuid':uuid,'can_afford':lambda *a:True,'presentation_cost':lambda *a:60,'_create_presentation_plan':creator,'plan_key':plan_key,'format_plan':lambda *a:'summary','apply_design':lambda color,plan:color,'THEMES':{'default':color}}
  fn=load_functions(['studio_plan_preview'],env)['studio_plan_preview']
  data,preview=await fn({'id':PILOT_UID},{**PAYLOAD,'kind':'designer','style':'soft','key':'d16','content_lang':'en','palette':'nature'})
  self.assertEqual(data['content_lang'],'en');self.assertEqual(data['studio_palette'],'nature');self.assertEqual(data['extra'],'Шаблон №16');self.assertEqual(preview['bg'],PALETTES['nature']['bg']);self.assertEqual(creator.call_args.args[1],'ru')
  _,other=await fn({'id':PILOT_UID},PAYLOAD);self.assertEqual(creator.call_args.args[0]['content_lang'],'fr')

class PaletteTests(unittest.TestCase):
 def test_all_standard_palettes_reach_designer_scene_and_keep_fonts(self):
  catalog=json.loads(Path('studio_ui/templates/catalog.json').read_text())
  self.assertEqual(set(PALETTES),{r['key'] for r in catalog if r['group']=='standard'})
  for style in ['designer_strict','designer_soft']:
   for key,palette in PALETTES.items():
    base={'topic':'Кофейня','extra':'Шаблон №16','theme_name':style}
    original=choose_design(base);design=choose_design({**base,'studio_palette':key});resolved=resolved_profile(design)
    self.assertEqual(resolved['bg'],palette['bg'],key);self.assertEqual(resolved['ink'],palette['ink']);self.assertEqual(design['template_id'],original['template_id']);self.assertEqual(resolved['font'],resolved_profile(original)['font']);self.assertTrue(has_custom_palette(design))
    a,b=relative_luminance(resolved['bg']),relative_luminance(resolved['ink']);self.assertGreaterEqual((max(a,b)+.05)/(min(a,b)+.05),4.5)
    self.assertNotEqual(plan_key(base),plan_key({**base,'studio_palette':key}))
    from presentation_designer import build_scene
    for number in (1,2):
     scene=build_scene({'title':'Проверка цвета','point':'Образец','content':'Короткий текст','_designer_design':design},None,number,4,style=style)
     self.assertEqual(scene['background'].lower(),palette['bg'].lower(),(style,key,number))
 def test_language_rejects_unexpected_types(self):
  for value in [None,[],{},123,'bad']:
   self.assertEqual(content_language(value,'ru'),'ru')

class PayloadLanguageTests(unittest.IsolatedAsyncioTestCase):
 async def test_word_and_excel_receive_selected_language_without_changing_ui(self):
  from excel_styles import normalize_excel_appearance
  env={'Message':object,'FSMContext':object,'ensure_answerable':lambda m:m,'user_lang':lambda uid:'ru','get_user':lambda uid:{'studio_content_lang':'fr'},'can_afford':lambda *a:True,'_clip_text':lambda v,n:str(v or '')[:n], 'CREDIT_COSTS':{'excel':5},'WORD_PAGE_CHOICE_KINDS':{'essay','doc'},'WORD_DEFAULT_PAGES':2,'normalize_word_pages':lambda v:int(v),'word_cost':lambda n, kind='doc':10,'thin_long_doc_warning':lambda *a:None,'word_build':AsyncMock(),'excel_build':AsyncMock(),'normalize_excel_appearance':normalize_excel_appearance}
  fn=load_functions(['_handle_miniapp_payload'],env)['_handle_miniapp_payload']
  for uid,expected in [(PILOT_UID,'en'),(987,'ru')]:
   for action in ['gen_word','gen_excel']:
    state=SimpleNamespace(update_data=AsyncMock());m=SimpleNamespace(from_user=SimpleNamespace(id=uid))
    await fn(m,state,{'content':'Test topic','kind':'essay' if action=='gen_word' else 'calc_table','content_lang':'en','pages':2},action)
    self.assertEqual(state.update_data.call_args.kwargs['content_lang'],expected)
