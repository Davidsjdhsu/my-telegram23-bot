"""User-selected pages exclude the built-in cover and contents, including billing."""
import json
import unittest
from types import SimpleNamespace
from unittest.mock import AsyncMock

import academic_templates as academic
from test_office import WordWorkflowTests
from test_presentation_support import load_functions


class PageContractTests(unittest.IsolatedAsyncioTestCase):
    def test_chat_keyboard_displays_body_pages_without_changing_price_or_parser(self):
        env = load_functions(['word_size_kb'], {
            'KeyboardButton': lambda **kw: SimpleNamespace(**kw),
            'ReplyKeyboardMarkup': lambda **kw: SimpleNamespace(**kw),
            'word_cost': lambda n: n * 5, 'tr': lambda key, lang: key,
        })
        for kind in academic.KINDS:
            kb = env['word_size_kb']('ru', kind, True)
            label = kb.keyboard[2][0].text
            self.assertIn('Титульный +', label)
            self.assertIn('6 стр. · 30 кр.', label)
            self.assertEqual('оглавление' in academic.volume_hint(kind), kind in ('referat', 'coursework'))
            import re
            self.assertEqual(re.findall(r'\d+', label)[0], '6')
        self.assertEqual(env['word_size_kb']('ru', 'doc').keyboard[1][0].text, '6 стр. · 30 кр.')

    async def run_handler(self, kind, actual_body, fail_send=False):
        env, message, state = WordWorkflowTests().env()
        state.get_data.return_value = {'word_kind': kind, 'word_pages': 6, 'topic': 'Учебная тема',
                                      'extra': 'Вуз: Учебный вуз\nДисциплина: Информатика\nФИО: Анна Петрова\nГруппа: ИС-21\nПреподаватель: Ирина Петрова\nГород: Москва\nКурс: 1\nГод: 2026'}
        env['WORD_KIND_LABELS'] = {k: {'ru': k} for k in academic.KINDS}
        env['word_cost'] = lambda n: n * 5
        content = {'title': 'Тема', 'sections': [{'title': 'Проблема', 'content': 'Пояснение.'}]}
        env['ask_grok'].return_value = json.dumps(content)
        async def prepare(value, requested, *args, **kwargs):
            self.assertEqual(requested, 6 + academic.front_matter_pages(kind))
            return value, actual_body + academic.front_matter_pages(kind)
        env['prepare_word_files'] = AsyncMock(side_effect=prepare)
        order = []
        async def send(*args, **kwargs):
            order.append('send')
            if fail_send and len(order) == 2:
                raise RuntimeError('offline fixture')
        message.answer_document.side_effect = send
        env['spend_credits'].side_effect = lambda *a, **kw: order.append('charge')
        await env['word_build'](message, state)
        return env, message, order

    async def test_all_genres_add_front_matter_but_charge_only_selected_body(self):
        for kind in academic.KINDS:
            for actual in (6,):
                with self.subTest(kind=kind, actual=actual):
                    env, message, order = await self.run_handler(kind, actual)
                    self.assertEqual(order, ['send', 'send', 'charge'])
                    env['spend_credits'].assert_called_once_with(-9901, 'word', amount=min(6, actual) * 5)
                    notice = '\n'.join(str(c.args[0]) for c in message.answer.call_args_list)
                    self.assertIn(f'{actual} стр. текста', notice)
                    self.assertIn('без доплаты', notice)
                    prompt = env['ask_grok'].call_args.args[0]
                    self.assertIn('6 страниц основного текста', prompt)
                    self.assertIn('НЕ входят в выбранный объём', prompt)

    async def test_failed_pdf_delivery_never_charges_for_cover_or_text(self):
        env, _, order = await self.run_handler('essay', 6, fail_send=True)
        self.assertEqual(order, ['send', 'send'])
        env['spend_credits'].assert_not_called()

    async def test_cover_only_is_not_delivered_or_charged(self):
        env, message, order = await self.run_handler('essay', 0)
        self.assertEqual(order, [])
        message.answer_document.assert_not_awaited()
        env['spend_credits'].assert_not_called()
