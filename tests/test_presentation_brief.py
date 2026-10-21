import asyncio
import json
import unittest
import uuid
from types import SimpleNamespace
from unittest.mock import AsyncMock
import test_presentation_support as fixtures
load_functions = fixtures.load_functions
State = fixtures.State
import presentation_support as ps

OPTIONS=['Виды, питание и общение китов','Миграции и жизнь в океане','Угрозы и защита китов']
BRIEF={'needs_audience':True,'needs_focus':True,'options':OPTIONS}


class BriefRulesTests(unittest.TestCase):
    def test_short_topic_needs_brief(self):
        for topic in ['киты','Сделай презентацию про китов','Маркетинг','Make a presentation about space']:
            self.assertTrue(ps.should_clarify_presentation({'mode':'ai','topic':topic}))

    def test_detailed_and_source_requests_skip_brief(self):
        detailed=('Сделай презентацию для учеников 9 класса о солнечной энергетике. '
                  'Объясни устройство панели, инвертора, аккумулятора, движение зарядов, '
                  'преимущества, ограничения, разные варианты применения и влияние погоды. '
                  'В конце сравни автономную и сетевую систему. '
                  'Всего восемь слайдов, светлый фон, зелёные акценты, схемы и понятные иллюстрации. '
                  'Пиши кратко, используй только проверяемые данные.')
        self.assertFalse(ps.should_clarify_presentation({'mode':'ai','topic':detailed}))
        self.assertFalse(ps.should_clarify_presentation({'mode':'user','topic':'Киты','user_text':'Материал пользователя'}))

    def test_choose_for_me_bypasses_questions(self):
        for text in ['Киты, на твоё усмотрение','Сделай без вопросов про китов','киты, выбери за меня',*ps.CHOOSE_PRESENTATION_LABELS.values()]:
            self.assertTrue(ps.wants_presentation_defaults(text),text)
            self.assertFalse(ps.should_clarify_presentation({'topic':text}),text)
        self.assertFalse(ps.wants_presentation_defaults('Сам придумай фото, тема — киты'))

    def test_known_audience_not_asked_twice(self):
        for text in ['Киты для 5 класса','Киты для учеников 9–11 классов','Для детей 10 лет, киты', 'Whales for students']:
            b=ps.validate_presentation_brief(BRIEF,text)
            self.assertFalse(b['needs_audience'],text)
        b=ps.validate_presentation_brief(BRIEF,'Презентация про китов')
        self.assertTrue(b['needs_audience'])

    def test_only_missing_field_is_shown(self):
        b={'needs_audience':False,'needs_focus':True,'options':OPTIONS}
        msg=ps.presentation_brief_message(b,'ru',12)
        self.assertNotIn('Для кого',msg)
        self.assertIn('12 слайдов',msg)
        self.assertIn('Миграции',msg)
        b={'needs_audience':True,'needs_focus':False,'options':[]}
        msg=ps.presentation_brief_message(b,'ru',8)
        self.assertIn('Для кого',msg);self.assertNotIn('раскрыть',msg)

    def test_invalid_classification_rejected(self):
        for value in [None,{}, {'needs_audience':'false','needs_focus':True,'options':OPTIONS},
                      {**BRIEF,'options':['a','a','a']}, {**BRIEF,'options':['x'*81,'b','c']}]:
            with self.assertRaises(ps.PresentationValidationError):ps.validate_presentation_brief(value,'киты')

    def test_number_selection_does_not_confuse_age(self):
        for text in ['2','вариант 2','Для 5 класса, направление 2']:
            self.assertIn(OPTIONS[1],ps.expand_presentation_choice(text,OPTIONS))
        self.assertEqual(ps.expand_presentation_choice('для 2 класса',OPTIONS),'для 2 класса')
        self.assertEqual(ps.expand_presentation_choice('2',[]),'2')
        self.assertEqual(ps.expand_presentation_choice('вариант 23',OPTIONS),'вариант 23')

    def test_supported_languages_have_button_and_copy(self):
        for lang in ['ru','en','de','ar','zh','es','fr']:
            self.assertTrue(ps.CHOOSE_PRESENTATION_LABELS[lang])
            self.assertIn('8',ps.presentation_brief_message(BRIEF,lang,8))
            self.assertTrue(ps.presentation_reply_hint(lang))


class BriefFlowTests(unittest.IsolatedAsyncioTestCase):
    def env(self,**extras):
        env=fixtures.WorkflowTests().env()
        env.update(asyncio=asyncio,uuid=uuid,pick_theme=lambda t:('default',{}),
                   presentation_brief_kb=lambda lang:'brief_keyboard',
                   ask_grok=AsyncMock(return_value=json.dumps(BRIEF)))
        env.update(extras)
        return env
    def message(self,text=''):
        return fixtures.WorkflowTests().message(text)

    async def test_bare_topic_shows_one_question_and_stores_options(self):
        ns=load_functions(['_start_presentation_request','pres_clarify_handler'],self.env())
        state=State({});m=self.message('Сделай презентацию про китов')
        await ns['_start_presentation_request'](m,state,m.text,'ru')
        self.assertEqual(state.current,'clarify')
        self.assertEqual(state.data['pres_brief_options'],OPTIONS)
        self.assertEqual(m.answer.await_count,1)
        ns['process_slides'].assert_not_awaited()
        m.text='Для 5 класса, вариант 2'
        await ns['pres_clarify_handler'](m,state)
        ns['process_slides'].assert_awaited_once()
        self.assertIn('Для 5 класса',state.data['extra'])
        self.assertIn(OPTIONS[1],state.data['extra'])
        self.assertEqual(ns['ask_grok'].await_count,1)

    async def test_default_button_preserves_form_details_and_starts_plan(self):
        ns=load_functions(['_start_presentation_request','pres_clarify_handler'],self.env())
        state=State({});m=self.message()
        await ns['_start_presentation_request'](m,state,'Киты','ru',
            settings={'slides':12,'photo_mode':'own','theme_name':'minimal','content_lang':'en'})
        self.assertIn('12 слайдов',m.answer.await_args.args[0])
        m.text=ps.CHOOSE_PRESENTATION_LABELS['ru']
        await ns['pres_clarify_handler'](m,state)
        self.assertEqual(state.data['slides'],12)
        self.assertEqual(state.data['photo_mode'],'own')
        self.assertEqual(state.data['content_lang'],'en')
        self.assertEqual(state.data['extra'],'')
        ns['process_slides'].assert_awaited_once()
        self.assertEqual(ns['ask_grok'].await_count,1)

    async def test_explicit_defaults_skip_classifier(self):
        ns=load_functions(['_start_presentation_request'],self.env())
        state=State({})
        await ns['_start_presentation_request'](self.message(),state,'Киты, на твоё усмотрение','ru')
        ns['ask_grok'].assert_not_awaited(); ns['process_slides'].assert_awaited_once()

    async def test_complete_short_brief_goes_directly_to_plan(self):
        ns=load_functions(['_start_presentation_request'],self.env(ask_grok=AsyncMock(return_value=json.dumps(
            {'needs_audience':False,'needs_focus':False,'options':[]}))))
        m=self.message();state=State({})
        await ns['_start_presentation_request'](m,state,'Миграции китов для студентов-биологов','ru')
        m.answer.assert_not_awaited();ns['process_slides'].assert_awaited_once()

    async def test_failed_classifier_still_offers_button(self):
        for error in [TimeoutError(),ValueError('provider unavailable')]:
            ns=load_functions(['_start_presentation_request'],self.env(ask_grok=AsyncMock(side_effect=error)))
            state=State({});m=self.message()
            await ns['_start_presentation_request'](m,state,'киты','ru')
            self.assertEqual(state.current,'clarify')
            self.assertEqual(m.answer.await_args.kwargs['reply_markup'],'brief_keyboard')
            self.assertNotIn('provider',m.answer.await_args.args[0])

    async def test_cancel_during_classifier_does_not_resurrect_request(self):
        state=State({})
        async def cancel(*a,**kw):
            await state.clear();return json.dumps(BRIEF)
        ns=load_functions(['_start_presentation_request'],self.env(ask_grok=AsyncMock(side_effect=cancel)))
        m=self.message()
        await ns['_start_presentation_request'](m,state,'киты','ru')
        m.answer.assert_not_awaited();ns['process_slides'].assert_not_awaited()
        self.assertEqual(state.data,{})

    async def test_empty_or_attachment_reply_does_not_start_generation(self):
        ns=load_functions(['pres_clarify_handler'],self.env())
        state=State({'pres_topic':'киты'});m=self.message(None)
        await ns['pres_clarify_handler'](m,state)
        ns['process_slides'].assert_not_awaited()
        self.assertEqual(state.data,{'pres_topic':'киты'})

    async def test_new_request_clears_stale_plan_colors_and_photos(self):
        ns=load_functions(['_start_presentation_request'],self.env())
        state=State({'pres_plan':{'title':'old'},'custom_colors':{'a':'b'},'user_photos':['old.png']})
        await ns['_start_presentation_request'](self.message(),state,'Киты, выбери за меня','ru')
        self.assertIsNone(state.data['pres_plan']);self.assertIsNone(state.data['custom_colors'])
        self.assertEqual(state.data['user_photos'],[])

    async def test_insufficient_balance_skips_model(self):
        ns=load_functions(['_start_presentation_request'],self.env(can_afford=lambda *args:False))
        await ns['_start_presentation_request'](self.message(),State({}),'киты','ru')
        ns['ask_grok'].assert_not_awaited();ns['process_slides'].assert_not_awaited()
        ns['send_no_credits_notice'].assert_awaited_once()

    async def test_topic_menu_preserves_selected_language(self):
        start=AsyncMock()
        ns=load_functions(['process_topic'],self.env(_start_presentation_request=start))
        await ns['process_topic'](self.message('Киты'),State({'content_lang':'de'}))
        self.assertEqual(start.await_args.kwargs['settings'],{'content_lang':'de'})

    async def test_free_chat_offers_confirmation_before_shared_brief(self):
        start=AsyncMock()
        env=self.env(_start_presentation_request=start,offer_document_intake=AsyncMock(),
            get_user=lambda uid:{'lang_chosen':True},chat_rate_limited=lambda uid:False,
            prefers_chat_reply=lambda t:False,looks_like_document_request=lambda t:True,
            detect_requested_format=lambda t:'presentation',PRESENTATION_MIN_SLIDES=3)
        ns=load_functions(['handle_free_text_request'],env)
        m=self.message('Сделай презентацию про китов');state=State({})
        await ns['handle_free_text_request'](m,state,m.text)
        start.assert_not_awaited()
        ns['offer_document_intake'].assert_awaited_once_with(m,state,m.text,'presentation','ru')

    async def test_miniapp_preserves_form_values_in_shared_brief(self):
        start=AsyncMock()
        env=self.env(_start_presentation_request=start,
            ensure_answerable=lambda m:m,get_user=lambda uid:{},
            _clip_text=lambda value,n:(value or '')[:n])
        ns=load_functions(['_handle_miniapp_payload'],env)
        m=self.message();state=State({})
        await ns['_handle_miniapp_payload'](m,state,
            {'topic':'Киты','slides':16,'style':'minimal','photo_mode':'own'},'gen_presentation')
        start.assert_awaited_once()
        settings=start.await_args.kwargs['settings']
        self.assertEqual(settings['slides'],16)
        self.assertEqual(settings['photo_mode'],'own')
        self.assertEqual(settings['theme_name'],'minimal')

    def test_keyboard_contains_choice_and_exit(self):
        ns=load_functions(['presentation_brief_kb'],self.env(
            ReplyKeyboardMarkup=lambda **kw:kw, KeyboardButton=lambda **kw:kw))
        keyboard=ns['presentation_brief_kb']('ru')['keyboard']
        self.assertEqual(keyboard[0][0]['text'],'✨ Выбери за меня')
        self.assertEqual(keyboard[1][0]['text'].strip(),'btn_main_menu')


if __name__=='__main__':unittest.main()
