import ast
import json
import re
import unittest
from pathlib import Path
from unittest.mock import AsyncMock
import presentation_visuals as pv
import presentation_designer as pd
import test_presentation_support as fixtures


class QualityTests(unittest.TestCase):
    def style(self,text):
        root=ast.parse(Path('main.py').read_text());ns={'re':re}
        nodes=[n for n in root.body if (isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='STYLE_KEYWORDS' for t in n.targets)) or (isinstance(n,ast.FunctionDef) and n.name=='extract_style')]
        exec(compile(ast.Module(body=nodes,type_ignores=[]),'style','exec'),ns)
        return ns['extract_style'](text)

    def test_designer_inflections_override_topic(self):
        for word in ['дизайнерскую','дизайнерской','дизайнерский','дизайнерские']:
            self.assertEqual(self.style(f'{word} презентация про технологии'), 'designer_soft')
        self.assertEqual(self.style('Мягкие скругления, технологии отелей'), 'designer_soft')
        self.assertEqual(self.style('Строгую дизайнерскую презентацию про технологии'),'designer_strict')
        self.assertEqual(self.style('Без дизайнерского стиля, технологии'),'tech')

    def test_image_ratio_matches_layout_instead_of_slide_number(self):
        self.assertEqual(pv.ordinary_image_ratio({'visual_layout':'wide'},4),'16:9')
        self.assertEqual(pv.ordinary_image_ratio({'visual_layout':'focus'},2),'3:4')
        self.assertEqual(pv.ordinary_image_ratio({},2),'4:3')
        self.assertEqual(pv.ordinary_image_ratio({},5),'4:3')

    def test_headline_shrinks_before_breaking_words(self):
        node=pd.text_node('Коралловые рифы',(0,0,440,244),'#000000',86,42,role='title')
        self.assertIn('Коралловые',node['text'].splitlines())
        self.assertNotIn('Кораллов\n',node['text'])

    def test_requested_dark_designer_background(self):
        scene=pd.build_scene({'title':'Умный номер','content':'Гость выбирает освещение.',
          '_designer_background':'dark','_designer_accent':'#70D5CF'},None,2,8)
        self.assertEqual(scene['background'],'#080B10')
        title=next(n for n in scene['nodes'] if n.get('role')=='title')
        self.assertEqual(title['color'],'#FFFFFF')


class RetryTests(unittest.IsolatedAsyncioTestCase):
    async def test_transient_plan_failure_retries_and_validates(self):
        model=AsyncMock(side_effect=['ERROR',json.dumps(fixtures.make_plan())])
        env=fixtures.load_functions(['_create_presentation_plan'],{
          'ask_grok':model,'grok_failed':lambda x:x=='ERROR','extract_json':json.loads,
          'content_gen_lang':lambda *a:'ru','grok_json_lang_instruction':lambda *a:''})
        result=await env['_create_presentation_plan']({'slides':8},'ru')
        self.assertEqual(len(result['slides']),8)
        self.assertEqual(model.await_count,2)
