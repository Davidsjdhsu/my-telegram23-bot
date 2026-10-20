import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, patch

from docx import Document
import academic_templates as academic
from academic_reference import reference_guidance, COMPUTER_MEMORY_SOURCES
import test_academic_completion


class NotesPrecisionTests(unittest.TestCase):
    def test_missing_position_has_one_label_and_filled_position_is_preserved(self):
        with tempfile.TemporaryDirectory() as td:
            for position in ("", "Доцент, кандидат технических наук"):
                path = Path(td) / "notes.docx"
                academic.render(path, "Память", [{"title": "RAM", "content": "• Определение: текст."}],
                                "notes", {"teacher": "Ирина Петрова", "teacher_position": position})
                lines = [p.text for p in Document(path).paragraphs]
                field = next(s for s in lines if s.startswith("Должность:"))
                self.assertEqual(field, "Должность: " + (position or "__________"))
                self.assertIn("Преподаватель: Ирина Петрова", lines)
                self.assertNotIn("Должность: Степень", "\n".join(lines))
                self.assertIn("• Определение: текст.", lines)

    def test_computer_brief_is_not_injected_into_unrelated_topics(self):
        for topic in ("Память человека", "Удалённая работа", "История Рима", "Рыбы"):
            self.assertEqual(reference_guidance({"topic": topic}), "")
        for topic in ("Оперативная память", "RAM и SSD", "Файл подкачки", "HDD"):
            hint = reference_guidance({"topic": topic})
            self.assertTrue(all(url in hint for url in COMPUTER_MEMORY_SOURCES))
        self.assertEqual(reference_guidance({"topic": "Литература", "teacher": "RAM"}), "")


class NotesPrecisionWorkflowTests(unittest.IsolatedAsyncioTestCase):
    async def test_actual_author_and_all_revision_routes_receive_same_references(self):
        for attempt, short, excess in ((1, False, False), (2, True, False), (2, False, True)):
            env, message, state, data, layout = test_academic_completion.CompletionTests().setup_flow("notes", actual=6)
            data["topic"] = "Оперативная память и накопители компьютера"
            original = {"title": "Память", "sections": [
                {"title": "Работа RAM", "content": "• Данные: " + "пояснение " * 80},
                {"title": "Накопители", "content": "• Архив: " + "объяснение " * 80},
            ]}
            env["ask_grok"].side_effect = [json.dumps(original), "{}"]

            async def prepare(value, requested, *args, **kwargs):
                await kwargs["expand"](value, 5 if short else 8 if excess else 7, requested, {
                    "short_pages": short, "sparse_last_page": False, "sparse_body_pages": [],
                    "excess_pages": excess, "body_page_fills": [.95] * (4 if short else 7),
                    "revision_attempt": attempt,
                })
                return value, 7
            env["prepare_word_files"] = AsyncMock(side_effect=prepare)
            with patch.object(academic, "inspect_layout", return_value=layout):
                await env["word_build"](message, state)
            self.assertEqual(env["ask_grok"].await_count, 2)
            for call in env["ask_grok"].await_args_list:
                self.assertIn(COMPUTER_MEMORY_SOURCES[0], call.args[0])
                self.assertIn("20–55", call.args[0])
            self.assertEqual(message.answer_document.await_count, 2)
            env["spend_credits"].assert_called_once()
