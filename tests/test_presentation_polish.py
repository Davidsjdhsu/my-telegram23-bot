import tempfile
import unittest
from pathlib import Path
from PIL import Image
import presentation_designer as pd
import presentation_profiles as pp
import presentation_visuals as pv

class PresentationPolishTests(unittest.TestCase):
    def test_cover_title_preserved_and_wide_for_long_topics(self):
        titles=['Бутик-отель в Тоскане: отдых среди виноградников',
                'Современная логистика: как роботы помогают сотрудникам на складе',
                'Круговорот воды в природе: испарение, конденсация, осадки и сток']
        with tempfile.TemporaryDirectory() as td:
            path=Path(td)/'photo.png';Image.new('RGB',(1600,900)).save(path)
            asset={'path':str(path),'width':1600,'height':900,'origin':'ai'}
            for family in pp.PROFILES:
                for title in titles:
                    d=pp.choose_design({'topic':title});d['family']=family
                    item={'title':title,'point':'Короткое объяснение темы для выбранной аудитории','_designer_design':d}
                    pd.plan_designer_images([item],'designer_soft')
                    self.assertEqual(item['_designer_aspect_ratio'],'16:9')
                    scene=pd.build_scene(item,asset,1,8)
                    actual=next(n for n in scene['nodes'] if n.get('role')=='title')
                    self.assertEqual(' '.join(actual['text'].split()),title)
                    self.assertGreaterEqual(actual['box'][2],990)

    def test_closing_text_preserved_for_every_family(self):
        for family in pp.PROFILES:
            d=pp.choose_design({'topic':'Тема'});d['family']=family
            for body in ['Главная мысль и конкретный вывод.',('Это содержательное объяснение результата. '*11).strip()]:
                scene=pd.build_scene({'title':'Что следует из рассмотренных примеров','content':body,'_designer_design':d},None,8,8)
                title=next(n for n in scene['nodes'] if n.get('role')=='title')
                text=next(n for n in scene['nodes'] if n.get('role')=='body')
                if len(body)<300:self.assertLess(title['box'][0]+title['box'][2],text['box'][0])
                self.assertEqual(' '.join(text['text'].split()),body)

    def test_image_checker_prioritizes_slide_text_and_preserves_supporting_examples(self):
        prompt=pv.review_prompt('Отель',{'title':'Интерьеры','content':'Светлые комнаты с деревянной мебелью','image_prompt':'Outdoor terrace'})
        self.assertIn('title и content, а не intended_scene',prompt)
        self.assertIn('relevant=false',prompt)
        self.assertIn('невидимый механизм',prompt)
        self.assertIn('Светлые комнаты',prompt)
