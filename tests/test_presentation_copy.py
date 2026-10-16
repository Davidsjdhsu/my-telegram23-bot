import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock
from PIL import Image
import presentation_copy as pc
import presentation_visuals as pv
import presentation_quality as pq
import presentation_support as ps
import test_presentation_support as f

class ContextTests(unittest.IsolatedAsyncioTestCase):
    def test_personal_request_is_not_every_upload(self):
        for topic in ['Мои путешествия: места, люди и впечатления','Наша свадьба','My family album']:
            self.assertTrue(pc.needs_photo_context({'topic':topic,'user_photos':['a']}))
        for topic in ['Солнечная энергетика, мои фото','Устройство двигателя','Наши солнечные панели']:
            self.assertFalse(pc.needs_photo_context({'topic':topic,'user_photos':['a']}))
        for extra in [{'pres_photo_context':''},{'user_text':'В Казани мы гуляли у кремля.'}]:
            self.assertFalse(pc.needs_photo_context(dict(topic='Мои путешествия',user_photos=['a'],**extra)))

    async def test_question_before_job_and_charge(self):
        env,state,sent=f.BuildTests().build_env(3)
        state.data.update(topic='Мои путешествия',user_photos=['a']);env['ask_photo_context']=AsyncMock()
        await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-1)),state)
        env['ask_photo_context'].assert_awaited_once();env['start_job'].assert_not_called()
        env['ask_grok'].assert_not_awaited();env['spend_credits'].assert_not_called();self.assertFalse(sent)

    async def test_answer_and_caption_only_continue(self):
        for text,only in [('Это снимки из разных лет.',False),(pc.PHOTO_ONLY_LABELS['ru'],True)]:
            env=f.load_functions(['receive_photo_context'],{'user_lang':lambda _:'ru','_build_presentation':AsyncMock()})
            state=f.State({'topic':'Мои путешествия','user_photos':['a']})
            await env['receive_photo_context'](SimpleNamespace(text=text,from_user=SimpleNamespace(id=-1),answer=AsyncMock()),state)
            self.assertEqual(state.data['photo_album_only'],only)
            self.assertEqual(state.data['pres_photo_context'],'' if only else text)
            self.assertFalse(pc.needs_photo_context(state.data));env['_build_presentation'].assert_awaited_once()

    async def test_bad_context_keeps_state(self):
        env=f.load_functions(['receive_photo_context'],{'user_lang':lambda _:'ru','_build_presentation':AsyncMock()})
        for text in (None,'x'*3001):
            state=f.State({'user_photos':['a']})
            await env['receive_photo_context'](SimpleNamespace(text=text,from_user=SimpleNamespace(id=-1),answer=AsyncMock()),state)
            self.assertEqual(state.data,{'user_photos':['a']})
        env['_build_presentation'].assert_not_awaited()

class SourceTests(unittest.IsolatedAsyncioTestCase):
    async def test_photos_before_copy_and_reused_in_all_styles(self):
        for style in ('minimal','designer_soft','designer_strict'):
            with tempfile.TemporaryDirectory() as td:
                path=Path(td)/'photo.jpg';Image.new('RGB',(800,600),'navy').save(path)
                env,state,sent=f.BuildTests().build_env(3)
                state.data.update(theme_name=style,user_photos=[str(path)],pres_photo_captions={str(path):'слайд 2, без обрезки'})
                state.data['pres_plan_key']=ps.plan_key(state.data);content=f.make_content(3)
                for slide in content['slides']:slide['designer']={'layout':'auto','blocks':[]}
                events=[]
                async def vision(*a,**k):
                    events.append('photo')
                    return json.dumps({'description':'На крыше установлены солнечные панели.','slides':[2],
                                       'relevant':True,'has_text':False,'has_chart':False,'defective':False})
                async def generate(prompt,**kw):
                    events.append('copy');self.assertEqual(events[0],'photo')
                    self.assertIn('На крыше установлены солнечные панели.',prompt);return json.dumps(content)
                env['ask_presentation_vision']=AsyncMock(side_effect=vision);env['ask_grok']=AsyncMock(side_effect=generate)
                await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-1)),state)
                self.assertEqual(len(sent),2)
                self.assertEqual(sum('Подбери слайды' in c.args[1] for c in env['ask_presentation_vision'].await_args_list),1)
                sources=env['review_content'].await_args.args[0]['photo_sources']
                self.assertEqual(sources[0]['slide'],2);self.assertNotIn(td,json.dumps(sources))

    async def test_failed_description_keeps_photo_without_fabricated_evidence(self):
        with tempfile.TemporaryDirectory() as td:
            p=Path(td)/'photo.png';Image.new('RGB',(500,300),'green').save(p)
            inspect=AsyncMock(side_effect=TimeoutError());slides=[{'title':'Портрет','_designer_style':'designer_soft','_skip_ai_image':True}]
            owned=await pv.prepare_own_images('Портрет',slides,[str(p)],{str(p):'слайд 1'},td,inspect,json.loads,describe=True)
            facts=pc.photo_sources(owned[0]);self.assertFalse(facts[0]['description_available']);self.assertEqual(facts[0]['visible'],'')
            self.assertFalse(owned[0][0].get('safe_frames'))
            generate=AsyncMock();assets,_=await pv.prepare_images('Портрет',slides,[str(p)],{},td,generate,inspect,json.loads,owned=owned)
            self.assertEqual(assets[0]['origin'],'user');self.assertEqual(inspect.await_count,1);generate.assert_not_awaited()

    async def test_album_replans_from_photos_without_ai_memories(self):
        with tempfile.TemporaryDirectory() as td:
            p=Path(td)/'photo.png';Image.new('RGB',(500,300),'green').save(p)
            env,state,sent=f.BuildTests().build_env(3)
            state.data.update(topic='Мои путешествия',theme_name='designer_soft',user_photos=[str(p)],pres_photo_context='',photo_album_only=True)
            state.data['pres_plan_key']=ps.plan_key(state.data)
            env['ask_presentation_vision']=AsyncMock(return_value=json.dumps({'description':'Белый автомобиль у дома.','slides':[2]}))
            content=f.make_content(3)
            for slide in content['slides']:slide['designer']={'layout':'photo','blocks':[]}
            env['ask_grok']=AsyncMock(return_value=json.dumps(content))
            await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-1)),state)
            self.assertEqual(len(sent),2);env['generate_image'].assert_not_awaited()
            self.assertEqual(env['_create_presentation_plan'].await_count,1)
            self.assertEqual(env['_create_presentation_plan'].await_args.args[0]['photo_sources'][0]['visible'],'Белый автомобиль у дома.')
            self.assertIn('НЕ обязан иметь conclusion',env['ask_grok'].await_args.args[0])

class EditorialTests(unittest.IsolatedAsyncioTestCase):
    async def test_empty_and_invented_copy_never_become_advisory(self):
        for problem in ('empty','ungrounded'):
            issue={'slide':2,'problem':problem,'fix':'Use the actual photograph and provided facts only.'}
            self.assertEqual(pq.parse_review({'issues':[issue]},{2}),[issue])
            env,state,sent=f.BuildTests().build_env(3);env['review_content']=AsyncMock(return_value=[issue])
            await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-1)),state)
            self.assertFalse(sent);env['spend_credits'].assert_not_called();env['generate_image'].assert_not_awaited()
            self.assertEqual(env['ask_grok'].await_count,3);self.assertEqual(state.current,'confirm')

    def test_album_does_not_require_fake_conclusion(self):
        content=f.make_content(3);content['_photo_story']=True
        content['slides'][0]['content']='Белый автомобиль у дома.';content['slides'][1]['content']='Ребёнок в игрушечной машине.'
        for s in content['slides']:s['designer']={'layout':'photo','blocks':[]}
        self.assertEqual(pq.local_issues(content,True),[])
        del content['_photo_story'];self.assertEqual(pq.local_issues(content,True)[0]['problem'],'conclusion')

    def test_writer_and_editor_receive_same_evidence(self):
        data={'slides':3,'topic':'Мои путешествия','pres_photo_context':'Поездка в Казань в июле.',
              'photo_sources':[{'slide':2,'visible':'Ребёнок в игрушечной машине.','caption':'Сын на прогулке'}]}
        for prompt in (ps.plan_prompt(data,''),ps.content_prompt(data,f.make_plan(3),''),pq.review_prompt(data,f.make_plan(3),f.make_content(3))):
            for text in ('Ребёнок в игрушечной машине.','Поездка в Казань в июле.','Сын на прогулке','Не приписывай пользователю'):
                self.assertIn(text,prompt)
