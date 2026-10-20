"""Academic partial delivery requires consent; refinement reuses saved content."""
import copy
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, patch

from docx import Document
import academic_templates as academic
from test_office import WordWorkflowTests
from word_quality import (WORD_ACCEPT, WORD_REFINE, apply_word_section_patch,
                          focused_word_sections, prepare_word_files,
                          word_review_findings, needs_word_revision)


class CompletionTests(unittest.IsolatedAsyncioTestCase):
    def setup_flow(self, kind='notes', actual=4, fill=.9):
        env, message, state = WordWorkflowTests().env()
        data = {'word_kind': kind, 'word_pages': 6, 'topic': 'Оперативная память',
                'extra': 'Вуз: Учебный вуз\nФИО: Анна Петрова\nГруппа: ИС-21\nКурс: 1\n'
                         'Дисциплина: Информатика\nПреподаватель: Ирина Петрова\nГород: Москва\nГод: 2026'}
        state.get_data.side_effect = lambda: dict(data)
        async def update(**kwargs): data.update(kwargs)
        async def clear(): data.clear()
        state.update_data.side_effect = update
        state.clear.side_effect = clear
        env['WORD_KIND_LABELS'] = {k: {'ru': k} for k in academic.KINDS}
        env['word_cost'] = lambda n: n * 5
        value = {'title': 'Память', 'sections': [{'title': 'Работа с данными', 'content': 'Контрольный текст.'}]}
        env['ask_grok'].return_value = json.dumps(value)
        async def prepare(content, requested, *args, **kwargs):
            return content, actual + academic.front_matter_pages(kind)
        env['prepare_word_files'] = AsyncMock(side_effect=prepare)
        return env, message, state, data, {'body_page_fills': [fill] * actual}

    async def test_short_or_sparse_academic_order_is_held_without_delivery_or_charge(self):
        for kind in academic.KINDS:
            for actual, fill in ((4, .9), (6, .2), (8, .9)):
                env, m, state, data, layout = self.setup_flow(kind, actual, fill)
                with patch.object(academic, 'inspect_layout', return_value=layout):
                    await env['word_build'](m, state)
                m.answer_document.assert_not_awaited()
                env['spend_credits'].assert_not_called()
                self.assertIn('content', data['word_quality_pending'])
                state.set_state.assert_awaited_with('Form:waiting_word_quality')
                self.assertIn('кредиты не списаны', m.answer.call_args.args[0])
                buttons = m.answer.call_args.kwargs['reply_markup'].keyboard
                self.assertEqual([row[0].text for row in buttons[:2]], [WORD_REFINE, WORD_ACCEPT])

    async def test_accept_uses_saved_text_without_model_and_charges_once_after_both_files(self):
        env, m, state, data, layout = self.setup_flow()
        with patch.object(academic, 'inspect_layout', return_value=layout):
            await env['word_build'](m, state)
        env['ask_grok'].reset_mock()
        events = []
        async def send(*args, **kwargs): events.append('send')
        m.answer_document.side_effect = send
        env['spend_credits'].side_effect = lambda *a, **k: events.append('charge')
        await env['word_build'](m, state, quality_action='accept')
        self.assertEqual(events, ['send', 'send', 'charge'])
        env['ask_grok'].assert_not_awaited()
        env['spend_credits'].assert_called_once_with(-9901, 'word', amount=20)
        self.assertIsNone(env['prepare_word_files'].call_args.kwargs['expand'])
        await env['word_build'](m, state, quality_action='accept')
        self.assertEqual(events, ['send', 'send', 'charge'])

    async def test_refinement_resumes_saved_text_and_stale_request_cannot_accept_it(self):
        env, m, state, data, layout = self.setup_flow()
        with patch.object(academic, 'inspect_layout', return_value=layout):
            await env['word_build'](m, state)
            env['ask_grok'].reset_mock()
            saved = copy.deepcopy(data['word_quality_pending']['content'])
            await env['word_build'](m, state, quality_action='refine')
        env['ask_grok'].assert_not_awaited()  # mocked prepare, no new initial generation
        self.assertEqual(env['prepare_word_files'].call_args.args[0], saved)
        self.assertIsNotNone(env['prepare_word_files'].call_args.kwargs['expand'])
        self.assertEqual(env['prepare_word_files'].call_args.kwargs['max_revisions'], 2)
        data['topic'] = 'Другая тема'
        await env['word_build'](m, state, quality_action='accept')
        env['spend_credits'].assert_not_called()
        m.answer_document.assert_not_awaited()

    async def test_failed_completion_notice_cannot_repeat_completed_order(self):
        env, m, state, data, layout = self.setup_flow()
        with patch.object(academic, 'inspect_layout', return_value=layout):
            await env['word_build'](m, state)
        async def answer(text, **kwargs):
            if text.startswith('msg_doc_ready'):
                raise RuntimeError('fixture_final_notice')
        m.answer.side_effect = answer
        await env['word_build'](m, state, quality_action='accept')
        self.assertEqual(data, {})
        self.assertEqual(m.answer_document.await_count, 2)
        env['spend_credits'].assert_called_once()
        await env['word_build'](m, state, quality_action='accept')
        self.assertEqual(m.answer_document.await_count, 2)
        env['spend_credits'].assert_called_once()

    async def test_failed_delivery_or_balance_change_retains_pending_without_charge(self):
        for case in ('pdf', 'balance', 'busy'):
            env, m, state, data, layout = self.setup_flow()
            with patch.object(academic, 'inspect_layout', return_value=layout):
                await env['word_build'](m, state)
            if case == 'pdf':
                m.answer_document.side_effect = [None, RuntimeError('fixture')]
            elif case == 'balance':
                env['can_afford'] = lambda *args: False
            else:
                env['start_job'].return_value = (False, 'busy')
            await env['word_build'](m, state, quality_action='accept')
            env['spend_credits'].assert_not_called()
            self.assertIn('word_quality_pending', data)

    async def test_second_revision_is_bounded_and_measured_before_promotion(self):
        initial = {'sections': [{'title': 'Основная часть', 'content': 'first'}]}
        expanded = [{'sections': [{'title': 'Основная часть', 'content': text}]}
                    for text in ('second', 'third')]
        with tempfile.TemporaryDirectory() as td:
            docx, pdf = str(Path(td)/'a.docx'), str(Path(td)/'a.pdf')
            async def build(value, path): Path(path).write_text(value['sections'][0]['content'])
            async def convert(source, target): Path(target).write_text(Path(source).read_text())
            def measure(path): return {'first': 4, 'second': 5, 'third': 7}[Path(path).read_text()]
            revise = AsyncMock(side_effect=expanded)
            value, count = await prepare_word_files(initial, 7, docx, pdf, build=build, convert=convert,
                measure=measure, expand=revise, max_revisions=10)
            self.assertEqual(count, 7)
            self.assertEqual(revise.await_count, 2)  # hard bound, even for invalid caller limit
            self.assertEqual(revise.call_args.args[3]['revision_attempt'], 2)
            self.assertEqual(Path(docx).read_text(), Path(pdf).read_text())

    def test_focused_patch_preserves_untouched_sections_tables_and_numbers(self):
        value = {'meta': {'author': 'Иван'}, 'sections': [
            {'title': 'Введение', 'content': 'Исходная задача.'},
            {'title': 'Механизм', 'content': 'Дано 42 файла для проверки условий.',
             'table': {'headers': ['Число'], 'rows': [['42']]}},
            {'title': 'Пример', 'content': 'Пример с пояснением выбора.'},
            {'title': 'Заключение', 'content': 'Исходный вывод.'}]}
        targets = focused_word_sections(value, 100)
        self.assertEqual([x['index'] for x in targets], [1, 2])
        patch = {'sections': [{**x, 'content': value['sections'][x['index']]['content'] +
                               ' Добавлено новое пояснение.'} for x in targets]}
        merged = apply_word_section_patch(value, patch, targets)
        self.assertEqual(merged['meta'], value['meta'])
        self.assertEqual(merged['sections'][0], value['sections'][0])
        self.assertEqual(merged['sections'][-1], value['sections'][-1])
        self.assertEqual(merged['sections'][1]['table'], value['sections'][1]['table'])
        patch['sections'][0]['content'] = 'Утрачено числовое условие и важные данные.'
        self.assertIs(apply_word_section_patch(value, patch, targets), value)

    def test_repetition_is_review_signal_not_fake_originality_score(self):
        sentence = 'Оперативная память используется для временного хранения данных работающих программ.'
        findings = word_review_findings({'sections': [{'content': sentence + ' ' + sentence}]}, 6, 6)
        self.assertEqual(findings['duplicate_sentences'], 1)
        self.assertTrue(needs_word_revision(findings))
        self.assertNotIn('originality_percent', findings)

    def test_notes_title_does_not_repeat_genre_and_guidance_does_not_force_introduction(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td)/'notes.docx'
            academic.render(path, 'Конспект по теме «Оперативная память»',
                            [{'title': 'Ключевые понятия', 'content': 'Пояснение.'}], 'notes')
            text = '\n'.join(p.text for p in Document(path).paragraphs)
            self.assertIn('КОНСПЕКТ', text)
            self.assertNotIn('Конспект по теме', text)
            self.assertIn('Оперативная память', text)
        self.assertIn('без разделов «Введение» и «Заключение»', academic.guidance('notes'))
