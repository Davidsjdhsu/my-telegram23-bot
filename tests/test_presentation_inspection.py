import sys
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock,Mock
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from PIL import Image
import presentation_visuals as pv
import test_presentation_support as f
GOOD={'relevant':True,'has_text':False,'has_chart':False,'defective':False}

class InspectionTests(unittest.IsolatedAsyncioTestCase):
    async def test_presentation_inspector_requests_low_effort_json_without_sdk_retries(self):
        response=SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=json.dumps(GOOD)),finish_reason='stop')])
        create=AsyncMock(return_value=response)
        options=Mock(return_value=SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=create))))
        prepare=Mock(return_value=('fixture','image/jpeg'))
        async def heavy(fn,*args):return fn(*args)
        env=f.load_functions(['ask_presentation_vision'],{'VISION_MODEL':'grok-4.6','client':SimpleNamespace(with_options=options),'run_heavy':heavy,'_prepare_vision_image':prepare})
        self.assertEqual(json.loads(await env['ask_presentation_vision']('photo.jpg','check',800)),GOOD)
        options.assert_called_once_with(timeout=60.0,max_retries=0)
        kwargs=create.await_args.kwargs
        self.assertEqual(kwargs['extra_body'],{'reasoning_effort':'low'})
        self.assertEqual(kwargs['response_format'],{'type':'json_object'})
        self.assertGreaterEqual(kwargs['max_tokens'],2048)
        prepare.assert_called_once_with('photo.jpg',1024)
        response.choices[0].finish_reason='length'
        with self.assertRaises(ValueError):await env['ask_presentation_vision']('photo.jpg','check')

    async def test_rejected_scene_uses_specific_repair_before_next_review(self):
        with tempfile.TemporaryDirectory() as td:
            prompts=[]
            async def gen(prompt,path,**kw):
                prompts.append(prompt);Image.new('RGB',(80,60),(len(prompts)*30,50,60)).save(path);return True
            inspect=AsyncMock(side_effect=[json.dumps({**GOOD,'has_text':True,'replacement_scene':'A whale swimming underwater in open blue water'}),json.dumps({**GOOD,'relevant':False,'replacement_scene':'A humpback whale with long pectoral fins underwater'}),json.dumps(GOOD)])
            assets,w=await pv.prepare_images('Ocean',[{'title':'Whales','image_prompt':'A whale'}],[],{},td,gen,inspect,json.loads)
            self.assertEqual(len(prompts),3)
            self.assertTrue(prompts[1].startswith('A whale swimming'))
            self.assertTrue(prompts[2].startswith('A humpback whale'))
            self.assertEqual(assets[0]['origin'],'ai')
            self.assertEqual(w['ai_omitted'],0)

    def test_image_prompt_never_includes_full_slide_text_or_json(self):
        p=pv.image_prompt('Ocean',{'title':'Текст заголовка','content':'СЕКРЕТНЫЙ ТЕКСТ СЛАЙДА', 'image_prompt':'A coral reef underwater'})
        self.assertTrue(p.startswith('A coral reef underwater'))
        self.assertNotIn('СЕКРЕТНЫЙ',p)
        self.assertNotIn('Текст заголовка',p)
        self.assertNotIn('"scene":',p)
        self.assertNotIn('"meaning":',p)
