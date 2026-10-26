"""Failure injection against the actual builder, without paid APIs or Telegram."""
import unittest
from types import SimpleNamespace
from unittest.mock import AsyncMock

import presentation_quality as quality
import test_presentation_support as fixtures

MESSAGE = SimpleNamespace(from_user=SimpleNamespace(id=-987654321))


class FailureAudit(unittest.IsolatedAsyncioTestCase):
    def repeated_deck(self):
        env, state, sent = fixtures.BuildTests().build_env(3)
        # The fixture repeats a paragraph. Use the actual local checker, which
        # the general delivery fixture otherwise replaces with an empty result.
        env['local_issues'] = quality.local_issues
        return env, state, sent

    async def test_local_repetition_exhaustion_delivers_after_factual_review(self):
        env, state, sent = self.repeated_deck()
        await env['_build_presentation'](MESSAGE, state)
        self.assertEqual(env['ask_grok'].await_count, 3)
        self.assertEqual(env['review_content'].await_count, 3)
        self.assertEqual(len(sent), 2)
        env['spend_credits'].assert_called_once()

    async def test_local_repetition_cannot_hide_factual_error(self):
        env, state, sent = self.repeated_deck()
        env['review_content'] = AsyncMock(return_value=[
            {'slide': 2, 'problem': 'unsupported', 'fix': 'Remove the unsupported guarantee.'}])
        await env['_build_presentation'](MESSAGE, state)
        self.assertEqual(env['review_content'].await_count, 3)
        self.assertIn('unsupported', env['ask_grok'].await_args.args[0])
        self.assertEqual(sent, [])
        env['generate_image'].assert_not_awaited()
        env['spend_credits'].assert_not_called()

    async def test_local_repetition_cannot_bypass_unavailable_review(self):
        env, state, sent = self.repeated_deck()
        env['review_content'] = AsyncMock(side_effect=quality.DeckQualityError('unavailable'))
        await env['_build_presentation'](MESSAGE, state)
        self.assertEqual(env['review_content'].await_count, 3)
        self.assertEqual(sent, [])
        env['spend_credits'].assert_not_called()

    async def test_structural_error_remains_blocking(self):
        env, state, sent = self.repeated_deck()
        env['local_issues'] = lambda *a: [
            {'slide': 2, 'problem': 'missing', 'fix': 'Restore the requested diagram.'}]
        await env['_build_presentation'](MESSAGE, state)
        self.assertEqual(sent, [])
        env['review_content'].assert_not_awaited()
        env['generate_image'].assert_not_awaited()
        env['spend_credits'].assert_not_called()

    async def test_pdf_send_failure_then_retry_has_one_charge(self):
        env, state, sent = fixtures.BuildTests().build_env(3)
        original = env['bot'].send_document
        async def fail_pdf(chat_id, file, **kwargs):
            if file.filename.endswith('.pdf'):
                raise ConnectionError('synthetic PDF send outage')
            return await original(chat_id, file, **kwargs)
        env['bot'].send_document = fail_pdf
        await env['_build_presentation'](MESSAGE, state)
        self.assertEqual(len(sent), 1)
        env['spend_credits'].assert_not_called()
        self.assertTrue(state.data)
        env['bot'].send_document = original
        await env['_build_presentation'](MESSAGE, state)
        self.assertEqual(len(sent), 3)
        env['spend_credits'].assert_called_once()
        self.assertEqual(env['ask_grok'].await_count, 1)  # Accepted text survives retry.
        self.assertFalse(state.data)
