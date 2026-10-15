import copy
import tempfile
import unittest
from pathlib import Path
from PIL import Image
import presentation_profiles as pp
import presentation_support as ps
import presentation_designer as pd
import test_presentation_support as f

BRIEF = {'topic':'Бутик-отель в Тоскане','theme_name':'designer_soft',
         'extra':'Светлый кремовый фон, оливковые и терракотовые акценты. Концепция вымышленного бутик-отеля.'}

class TuscanyRegressionTests(unittest.TestCase):
    def test_requested_colors_are_visible_and_readable(self):
        d=pp.choose_design(BRIEF)
        self.assertEqual(d['accent'],'#B65E43');self.assertEqual(d['secondary'],'#707A41')
        for n in (1,2):
            scene=pd.build_scene({'title':'Тосканские интерьеры','content':'Камень и дерево в интерьере.','_designer_design':d},None,n,8)
            title=next(x for x in scene['nodes'] if x.get('role')=='title')
            self.assertEqual(title['color'],pp.resolved_profile(d)['accent' if n==1 else 'secondary_ink'])
            a,b=pp.relative_luminance(title['color']),pp.relative_luminance(scene['background'])
            self.assertGreaterEqual((max(a,b)+.05)/(min(a,b)+.05),4.5)

    def test_fictional_scope_is_specific(self):
        self.assertTrue(ps.fictional_project(BRIEF))
        self.assertFalse(ps.fictional_project({'extra':'Реальный отель. Без вымышленных цифр и отзывов.'}))
        self.assertTrue(ps.fictional_project({'extra':'A fictional boutique hotel'}))

    def test_fictional_plan_cannot_invent_booking(self):
        plan=f.make_plan();plan['slides'][-1]['title']='Выберите даты и забронируйте'
        with self.assertRaisesRegex(ps.PresentationValidationError,'Fictional project'):
            ps.finalize_plan_design(plan,BRIEF)
        self.assertEqual(ps.finalize_plan_design(plan,{'theme_name':'minimal'}),plan)

    def test_fictional_content_and_structured_blocks_checked(self):
        plan=ps.finalize_plan_design(f.make_plan(),BRIEF)
        for key in ('content','designer'):
            content=f.make_content()
            if key=='content':content['slides'][-1][key]='На сайте можно выбрать даты и забронировать номер.'
            else:content['slides'][-1][key]={'layout':'sections','blocks':[{'heading':'Купить','text':'Купите сейчас на нашем сайте'}]}
            with self.assertRaisesRegex(ps.PresentationValidationError,'Fictional project'):
                ps.validate_content(content,plan)
        content=f.make_content();content['slides'][-1]['content']='Возможный сценарий бронирования пока является прототипом. Работающего сайта нет.'
        self.assertEqual(ps.validate_content(content,plan)['slides'][-1]['content'],content['slides'][-1]['content'])

    def test_concept_booking_disclaimers_are_not_sales_claims(self):
        for text in ('Концепция без реального бронирования.',
                     'Бронирование отсутствует.',
                     'Проект не предлагает бронирование.',
                     'Концепция без бронирования и продажи номеров.'):
            with self.subTest(text=text):ps.validate_concept_copy(text)
        for text in ('Бронирование открыто.', 'Забронируйте номер на нашем сайте.'):
            with self.subTest(text=text), self.assertRaises(ps.PresentationValidationError):
                ps.validate_concept_copy(text)

    def test_sparse_sections_get_photo_but_comparison_kept(self):
        base={'_designer_design':pp.choose_design(BRIEF),'designer':{'layout':'sections','blocks':[{'heading':'Овощи','text':'Сезонное меню'},{'heading':'Травы','text':'Местные продукты'}]}}
        slides=[copy.deepcopy(base) for _ in range(8)]
        slides[3]['designer']['layout']='comparison'
        pd.plan_designer_images(slides,'designer_soft')
        self.assertFalse(slides[1]['_skip_ai_image']);self.assertTrue(slides[3]['_skip_ai_image'])

    def test_editorial_photo_rhythm_and_text_preservation(self):
        with tempfile.TemporaryDirectory() as td:
            path=Path(td)/'photo.png';Image.new('RGB',(1280,720)).save(path)
            asset={'path':str(path),'width':1280,'height':720,'origin':'ai'}
            layouts=set()
            for number in range(2,8):
                body='Гости могут отдыхать на террасе с видом на виноградники. В концепции предусмотрены тихие зоны для прогулок.'
                scene=pd.build_scene({'title':'Отдых среди виноградников','content':body,'_designer_design':pp.choose_design(BRIEF)},asset,number,8)
                texts=[x for x in scene['nodes'] if x.get('role')=='body']
                self.assertEqual(' '.join(' '.join(x['text'] for x in texts).split()),body)
                layouts.add(tuple(tuple(x['box']) for x in scene['nodes'] if x['type']=='image' or x.get('role')=='title'))
            self.assertGreaterEqual(len(layouts),5)
