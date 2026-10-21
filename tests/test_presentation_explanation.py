"""Offline coverage of request continuity, not an evaluation of model prose."""
import copy
import json
from types import SimpleNamespace
import unittest
from unittest.mock import AsyncMock

import test_presentation_support as f
import presentation_support as ps
import presentation_quality as pq
from presentation_explanation import clean_understanding, request_context


CASES = [
    ('Оперативная память', 'Студенты без технических знаний',
     'Объяснить отличие от накопителя', 'Не придумывай результаты тестов'),
    ('Кофейня', 'Начинающие предприниматели',
     'Объяснить путь заказа', 'Не считай отдельного кассира обязательным'),
    ('Городская библиотека', 'Первокурсники',
     'Показать поиск и выдачу книг', 'Не выдумывай единые сроки возврата'),
    ('История письменности', 'Взрослая аудитория без подготовки',
     'Сравнить способы записи', 'Не придумывай даты и источники'),
    ('Семейный альбом', 'Мои родственники',
     'Описать только присланные воспоминания', 'Не придумывай родство и события'),
]


def context_from_author(prompt):
    return json.loads(prompt.split('Исходный запрос и материал: ', 1)[1].split('\nОбложка создаётся отдельно.', 1)[0])


class ExplanationPlanTests(unittest.TestCase):
    def test_optional_metadata_survives_revalidation_and_preview(self):
        plan = f.make_plan(6)
        plan['understanding'] = {'goal': 'Объяснить устройство', 'requirements': ['Не выдумывать цифры'],
                                 'assumptions': ['Рассматриваем типичный вариант']}
        plan['slides'][1]['purpose'] = 'definition'
        plan['slides'][-1]['purpose'] = 'application'
        saved = ps.validate_plan(ps.validate_plan(plan, 6), 6)
        self.assertEqual(saved['understanding'], plan['understanding'])
        self.assertEqual(saved['slides'][-1]['purpose'], 'application')
        preview = ps.format_plan(saved)
        self.assertIn('Цель: Объяснить устройство', preview)
        self.assertIn('Предположения — можно изменить: Рассматриваем типичный вариант', preview)
        self.assertNotIn('purpose', preview)

    def test_legacy_or_malformed_optional_fields_do_not_block_plan(self):
        old = ps.validate_plan(f.make_plan(6), 6)
        self.assertNotIn('understanding', old)
        for value in (None, [], {'goal': False, 'requirements': [None, 7], 'assumptions': 'invented'}):
            plan = f.make_plan(6)
            plan['understanding'] = value
            plan['slides'][1]['purpose'] = ['process']
            cleaned = ps.validate_plan(plan, 6)
            self.assertNotIn('understanding', cleaned)
            self.assertNotIn('purpose', cleaned['slides'][1])
            self.assertEqual(len(cleaned['slides']), 6)

    def test_advisory_metadata_is_bounded_without_truncating_raw_request(self):
        data = {'topic': 'topic', 'user_text': 'begin ' + 'source ' * 2000 + ' END_CONSTRAINT', 'extra': 'RAW_REQUIREMENT'}
        plan = f.make_plan(6)
        plan['understanding'] = {'goal': 'x'*241, 'requirements': ['a']*50+['b']*50,
                                 'assumptions': [str(i) for i in range(8)]}
        context = request_context(data, plan)
        self.assertEqual(context['source']['user_text'], data['user_text'])
        self.assertEqual(context['source']['extra'], 'RAW_REQUIREMENT')
        self.assertEqual(context['plan_interpretation']['requirements'], ['a', 'b'])
        self.assertEqual(len(context['plan_interpretation']['assumptions']), 3)
        self.assertNotIn('goal', context['plan_interpretation'])

    def test_same_context_reaches_author_and_editor_across_subjects(self):
        for topic, audience, goal, constraint in CASES:
            with self.subTest(topic=topic):
                data = {'slides': 6, 'mode': 'user', 'topic': topic,
                        'user_text': 'Материал пользователя. ' + constraint, 'extra': 'Светлый фон'}
                plan = f.make_plan(6)
                plan['audience'] = audience
                plan['understanding'] = {'goal': goal, 'requirements': [constraint], 'assumptions': []}
                plan['slides'][1]['purpose'] = 'definition'
                plan = ps.validate_plan(plan, 6)
                author = context_from_author(ps.content_prompt(data, plan, ''))
                self.assertEqual(author, request_context(data, plan))
                editor = pq.review_prompt(data, plan, f.make_content(6))
                self.assertIn(json.dumps(author, ensure_ascii=False), editor)
                self.assertEqual(author['audience'], audience)
                self.assertEqual(author['plan_interpretation']['goal'], goal)
                self.assertEqual(author['source']['user_text'], data['user_text'])


class ExplanationFlowTests(unittest.IsolatedAsyncioTestCase):
    async def test_actual_planner_preserves_goal_without_extra_model_call(self):
        plan = f.make_plan(6)
        plan['understanding'] = {'goal':'Цель из запроса','requirements':['Сохранить ограничение'],'assumptions':[]}
        plan['slides'][1]['purpose'] = 'comparison'
        model = AsyncMock(return_value=json.dumps(plan))
        data = {'slides':6,'topic':'Тема','user_text':'Ограничение в исходном материале'}
        env = f.load_functions(['_create_presentation_plan'], {
            'ask_grok':model, 'content_gen_lang':lambda *a:'ru',
            'grok_json_lang_instruction':lambda *a:'', 'grok_failed':lambda raw:False,
            'extract_json':json.loads})
        saved = await env['_create_presentation_plan'](data,'ru')
        self.assertEqual(saved['understanding']['goal'], 'Цель из запроса')
        self.assertEqual(saved['slides'][1]['purpose'], 'comparison')
        self.assertEqual(model.await_count, 1)

    async def test_mechanism_review_receives_audience_source_and_task(self):
        data = {'topic':'Устройство','user_text':'Не считай косвенную связь прямой','extra':'Светлый фон'}
        plan = f.make_plan(3)
        plan['audience'] = 'Новички'
        plan['understanding'] = {'goal':'Показать роли частей'}
        plan['slides'][1]['purpose'] = 'structure'
        content = f.make_content(3)
        content['slides'][0].update(visual_kind='technical_diagram', image_prompt='', diagram={
            'kind':'structure','nodes':[{'label':'Система','text':'Целое'},{'label':'Часть','text':'Роль части'}]})
        content['slides'][1].update(visual_kind='technical_diagram', image_prompt='', diagram={
            'kind':'structure','nodes':[{'label':'Система','text':'Целое'},{'label':'Часть','text':'Роль части'}]})
        model = AsyncMock(return_value='{"issues":[]}')
        self.assertEqual(await pq.review_content(data,plan,content,model,json.loads), [])
        self.assertEqual(model.await_count, 2)  # Existing mechanism + editorial passes only.
        mechanism = model.await_args_list[0].args[0]
        self.assertIn(json.dumps(request_context(data,plan),ensure_ascii=False), mechanism)
        self.assertIn('"purpose": "structure"', mechanism)

    async def test_actual_builder_keeps_original_request_on_repair(self):
        env, state, sent = f.BuildTests().build_env(3)
        state.data['user_text'] = 'Исходное ограничение: не выдумывай результаты измерений'
        state.data['pres_plan']['audience'] = 'Первокурсники без подготовки'
        state.data['pres_plan']['understanding'] = {'goal':'Объяснить роли частей','requirements':['Без вымышленных измерений']}
        state.data['pres_plan']['slides'][1]['purpose'] = 'structure'
        state.data['pres_plan_key'] = ps.plan_key(state.data)
        env['review_content'] = AsyncMock(side_effect=[[
            {'slide':2,'problem':'logic','fix':'Уточнить условие связи без нового факта'}], []])
        await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-987654321)),state)
        self.assertEqual(len(sent),2)
        self.assertEqual(env['ask_grok'].await_count,2)
        for call in env['ask_grok'].await_args_list:
            context = context_from_author(call.args[0])
            self.assertEqual(context['audience'],'Первокурсники без подготовки')
            self.assertIn('не выдумывай результаты измерений',context['source']['user_text'])
            self.assertEqual(context['plan_interpretation']['goal'],'Объяснить роли частей')
        env['spend_credits'].assert_called_once()
