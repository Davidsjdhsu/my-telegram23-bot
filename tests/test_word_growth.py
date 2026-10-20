import copy
import json
import unittest
from unittest.mock import AsyncMock, patch

import academic_templates as academic
import test_academic_completion
from word_quality import (word_growth_targets, apply_word_additions, WORD_EDITORIAL_GUIDANCE,
                          WORD_REVIEW_SYSTEM)


class GrowthTests(unittest.TestCase):
    def setUp(self):
        self.value = {'title': 'Тема', 'meta': {'author': 'Иванов'},
                      'sections': [
                          {'title': 'Введение', 'content': 'Исходный вопрос.'},
                          {'title': 'Основная часть', 'content': 'Первый абзац: 42 файла.\n\nВторой абзац.',
                           'table': {'headers': ['Число'], 'rows': [['42']]}},
                          {'title': 'Заключение', 'content': 'Итоговый вывод.'}]}
        self.targets = word_growth_targets(self.value, 100)
        self.patch = {'additions': [
            {'section_index': 1, 'after_paragraph': 0, 'paragraphs': ['Новое пояснение причин.']},
            {'section_index': 1, 'after_paragraph': 1, 'paragraphs': ['Разбор условного примера.']}]}

    def test_budget_is_only_missing_words_and_excludes_front_and_conclusion(self):
        words = sum(len(s['content'].split()) for s in self.value['sections'])
        self.assertEqual(sum(t['add_words'] for t in self.targets), 100 - words)
        self.assertEqual([t['index'] for t in self.targets], [1])
        self.assertEqual(word_growth_targets(self.value, words), [])

    def test_insertions_preserve_original_data_and_anchors_without_mutation(self):
        before = copy.deepcopy(self.value)
        actual = apply_word_additions(self.value, self.patch, self.targets)
        self.assertEqual(self.value, before)
        self.assertEqual(actual['meta'], before['meta'])
        self.assertEqual(actual['sections'][0], before['sections'][0])
        self.assertEqual(actual['sections'][2], before['sections'][2])
        self.assertEqual(actual['sections'][1]['table'], before['sections'][1]['table'])
        self.assertEqual(actual['sections'][1]['content'],
                         'Первый абзац: 42 файла.\nНовое пояснение причин.\n\nВторой абзац.\nРазбор условного примера.')

    def test_malformed_or_excessive_growth_cannot_replace_saved_work(self):
        for update in ({'section_index': 0}, {'section_index': True}, {'after_paragraph': -1},
                       {'after_paragraph': 99}, {'paragraphs': ['']},
                       {'paragraphs': ['Таблица | данные']}, {'paragraphs': ['a\nb']},
                       {'paragraphs': ['слово ' * 200]}):
            candidate = copy.deepcopy(self.patch)
            candidate['additions'][0].update(update)
            self.assertIs(apply_word_additions(self.value, candidate, self.targets), self.value)

    def test_repeated_or_reworded_sentence_is_not_used_as_padding(self):
        sentence = 'Ответственность за результат требует самостоятельно планировать задачи и учитывать сроки коллег.'
        self.value['sections'][1]['content'] = sentence
        self.targets = word_growth_targets(self.value, 100)
        for repeated in (sentence, sentence.replace('сроки коллег', 'сроки других коллег')):
            patch_value = {'additions': [{'section_index': 1, 'after_paragraph': 0,
                                         'paragraphs': [repeated]}]}
            self.assertIs(apply_word_additions(self.value, patch_value, self.targets), self.value)

    def test_generic_editor_does_not_inject_storage_examples(self):
        for prompt in (WORD_EDITORIAL_GUIDANCE, WORD_REVIEW_SYSTEM):
            self.assertNotIn('облач', prompt)
            self.assertNotIn('синхронизац', prompt)


class GrowthWorkflowTests(unittest.IsolatedAsyncioTestCase):
    async def test_real_bot_callback_grows_essay_and_retains_unresolved_volume(self):
        env, message, state, data, layout = test_academic_completion.CompletionTests().setup_flow('essay', actual=4)
        original = {'title': 'Удалённая работа', 'meta': {'author': 'Иванов'},
                    'sections': [{'title': 'Аргументы', 'content': 'Исходный аргумент и его пояснение.'}]}
        # Start through the actual handler, then exercise its second measured revision.
        env['ask_grok'].side_effect = [
            json.dumps(original),
            json.dumps({'additions': [{'section_index': 0, 'after_paragraph': 0,
                                      'paragraphs': ['Дополнительный условный пример раскрывает новый аспект.']}]})]

        async def prepare(value, requested, *args, **kwargs):
            findings = {'short_pages': True, 'sparse_last_page': False, 'sparse_body_pages': [],
                        'excess_pages': False, 'body_page_fills': [.95] * 4, 'revision_attempt': 2}
            grown = await kwargs['expand'](value, 5, requested, findings)
            self.assertIn(original['sections'][0]['content'], grown['sections'][0]['content'])
            self.assertIn('Дополнительный условный пример', grown['sections'][0]['content'])
            self.assertEqual(kwargs['max_revisions'], 2)
            return grown, 5
        env['prepare_word_files'] = AsyncMock(side_effect=prepare)
        with patch.object(academic, 'inspect_layout', return_value=layout):
            await env['word_build'](message, state)
        self.assertEqual(env['ask_grok'].await_count, 2)
        self.assertIn('add_words', env['ask_grok'].await_args.args[0])
        message.answer_document.assert_not_awaited()
        env['spend_credits'].assert_not_called()
        self.assertIn('Дополнительный условный пример',
                      data['word_quality_pending']['content']['sections'][0]['content'])
