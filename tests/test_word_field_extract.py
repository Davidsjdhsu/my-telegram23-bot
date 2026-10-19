import json
import tempfile
import unittest
from pathlib import Path
from docx import Document
import fixed_word_templates as fixed
import car_sale_template as car
import word_field_extract as parser
import test_fixed_word_templates as workflow


REQUEST = '''Исполнитель: ООО «Тестовый исполнитель»
Заказчик: ООО «Тестовый заказчик»
Город: Москва
Дата: 20.10.2026
Услуги: Подготовка инструкции по работе с электронным архивом: изучение текущего порядка хранения файлов, разработка структуры папок, описание правил именования документов и проведение консультации для сотрудников.
Срок оказания услуг: до 30.10.2026
Стоимость: 15000 рублей
Порядок оплаты: после подписания акта оказанных услуг'''
EXPECTED = dict(zip(('provider', 'customer', 'city', 'date', 'services', 'period', 'price', 'payment'),
                    (line.split(':', 1)[1].strip() for line in REQUEST.splitlines())))


class InputTests(unittest.IsolatedAsyncioTestCase):
    async def test_reported_request_reaches_delivered_docx_and_review(self):
        for kind in ('services_ru', 'services'):
            e, m, state = workflow.WorkflowTests().env(kind, REQUEST)
            await e['word_build'](m, state)
            self.assertEqual(state.data['fixed_word']['fields'], EXPECTED)
            review = '\n'.join(c.args[0] for c in m.answer.call_args_list)
            for value in EXPECTED.values():
                self.assertIn(value, review)
            with tempfile.TemporaryDirectory() as td:
                saved = Path(td) / 'filled.docx'
                async def capture(path, **kw):
                    if str(path).endswith('.docx'):
                        saved.write_bytes(Path(path).read_bytes())
                m.answer_document.side_effect = capture
                await e['word_build'](m, state)
                text = '\n'.join(n.text or '' for n in Document(saved).element.iter() if n.tag.endswith('}t'))
                for key, value in EXPECTED.items():
                    self.assertIn(value, text)
                    self.assertNotIn('[' + fixed.CATALOG[kind]['fields'][key] + ']', text)
            e['ask_grok'].assert_not_awaited()
            e['spend_credits'].assert_called_once()

    async def test_short_date_and_city_work_for_every_fixed_form(self):
        for kind, spec in fixed.CATALOG.items():
            values = spec['sample']
            lines = [('Дата' if k == 'date' else 'Город' if k == 'city' else label) + ': ' + values[k]
                     for k, label in spec['fields'].items()]
            e, m, state = workflow.WorkflowTests().env(kind, '\n'.join(lines))
            await e['word_build'](m, state)
            self.assertEqual(state.data['fixed_word']['fields'], values, kind)
            e['ask_grok'].assert_not_awaited()

    async def test_mixed_labels_and_prose_are_both_extracted(self):
        source = 'Исполнитель: ООО Тест\nУслуга — подготовка инструкции. Оплата после подписания акта.'
        e, m, state = workflow.WorkflowTests().env(text=source)
        e['ask_grok'].return_value = json.dumps({'fields': {'services': 'подготовка инструкции', 'payment': 'после подписания акта'}})
        await e['word_build'](m, state)
        self.assertEqual(state.data['fixed_word']['fields'], {'provider': 'ООО Тест', 'services': 'подготовка инструкции', 'payment': 'после подписания акта'})
        e['ask_grok'].assert_awaited_once()

    async def test_mixed_model_failure_never_delivers_a_partially_parsed_order(self):
        for response in (RuntimeError('offline'), 'invalid json'):
            e, m, state = workflow.WorkflowTests().env(text='Исполнитель: ООО Тест\nОплата после подписания акта.')
            if isinstance(response, Exception):
                e['ask_grok'].side_effect = response
            else:
                e['ask_grok'].return_value = response
            await e['word_build'](m, state)
            self.assertEqual(state.current, 'extra')
            m.answer_document.assert_not_awaited()
            e['spend_credits'].assert_not_called()
            self.assertIn('после подписания акта', state.data['topic'])

    async def test_unmatched_label_is_shown_in_review(self):
        e, m, state = workflow.WorkflowTests().env(text='Исполнитель: ООО Тест\nОсобая договорённость: хранить резервную копию год')
        e['ask_grok'].return_value = json.dumps({'fields': {'provider': 'ООО Тест'}})
        await e['word_build'](m, state)
        review = '\n'.join(c.args[0] for c in m.answer.call_args_list)
        self.assertIn('Не удалось сопоставить', review)
        self.assertIn('хранить резервную копию год', review)
        m.answer_document.assert_not_awaited()

    async def test_greeting_does_not_block_an_explicit_request(self):
        e, m, state = workflow.WorkflowTests().env(text='Здравствуйте!\nИсполнитель: ООО Тест')
        e['ask_grok'].return_value = json.dumps({'fields': {}})
        await e['word_build'](m, state)
        self.assertEqual(state.current, 'confirm')
        self.assertEqual(state.data['fixed_word']['fields'], {'provider': 'ООО Тест'})

    async def test_explicit_clear_and_alias_correction_preserve_other_values(self):
        e, m, state = workflow.WorkflowTests().env(text=REQUEST)
        await e['word_build'](m, state)
        await state.update_data(extra='Стоимость: 20000 рублей\nДата: ')
        await fixed.preview(m, state, e)
        fields = state.data['fixed_word']['fields']
        self.assertEqual(fields['price'], '20000 рублей')
        self.assertNotIn('date', fields)
        self.assertEqual(fields['services'], EXPECTED['services'])
        e['ask_grok'].assert_not_awaited()

    async def test_free_text_correction_overrides_base_value_without_repeating_base_model(self):
        e, m, state = workflow.WorkflowTests().env(text=REQUEST)
        await e['word_build'](m, state)
        await state.update_data(extra='Теперь стоимость составляет 18000 рублей.')
        e['ask_grok'].return_value = json.dumps({'fields': {'price': '18000 рублей'}})
        await fixed.preview(m, state, e)
        self.assertEqual(state.data['fixed_word']['fields']['price'], '18000 рублей')
        e['ask_grok'].assert_awaited_once()

    async def test_old_preview_is_reparsed_after_fix(self):
        e, m, state = workflow.WorkflowTests().env(text=REQUEST)
        state.data['fixed_word'] = {'fields': {'provider': EXPECTED['provider']}, 'base_fingerprint': fixed.base_fingerprint(state.data), 'kind': 'services_ru', 'fingerprint': 'old'}
        await e['word_build'](m, state)
        self.assertEqual(state.data['fixed_word']['fields'], EXPECTED)
        m.answer_document.assert_not_awaited()

    async def test_car_mixed_input_and_short_names(self):
        e, m, state = workflow.car_tests.CarWorkflowTests().env({'word_kind': 'dkp_car', 'topic': 'Продавец: Петров\nПокупатель: Иванов\nДата: 20.10.2026\nЦена автомобиля составляет 1200000 рублей.'})
        e['ask_grok'].return_value = json.dumps({'fields': {'price': '1200000 рублей'}})
        await e['word_build'](m, state)
        self.assertEqual(state.data['car_sale']['fields'], {'seller_name': 'Петров', 'buyer_name': 'Иванов', 'date': '20.10.2026', 'price': '1200000 рублей'})
        self.assertNotIn('received_amount', state.data['car_sale']['fields'])

    def test_payment_terms_are_not_payment_facts(self):
        fields, rest = parser.split_labeled('Порядок оплаты: после приёмки', fixed.CATALOG['services_act_ru']['fields'])
        self.assertEqual(fields, {})
        self.assertIn('после приёмки', rest)

    def test_free_text_scope_is_checked_alongside_recognized_fields(self):
        self.assertTrue(fixed.scope_issue({'word_kind': 'lawsuit', 'topic': 'Дата: 20.10.2026\nНужен арбитражный иск'}))

    def test_label_normalization_and_colons_inside_values(self):
        fields, rest = parser.split_labeled('  ДАТА : 20.10.2026\nУслуги: Инструкция: разделы А и Б', fixed.CATALOG['services_ru']['fields'])
        self.assertEqual(fields, {'date': '20.10.2026', 'services': 'Инструкция: разделы А и Б'})
        self.assertEqual(rest, '')
