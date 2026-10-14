"""Failure-path checks: no external API calls or production balances."""
import io
import asyncio
import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock
import unittest

import test_presentation_support as f
import presentation_support as ps
import presentation_designer as pd
from pypdf import PdfReader
from PIL import Image
import tempfile


class DeliveryTests(unittest.IsolatedAsyncioTestCase):
    def message(self):
        return SimpleNamespace(from_user=SimpleNamespace(id=-987654321),
                               chat=SimpleNamespace(id=-987654321))

    async def test_initial_telegram_failure_releases_job(self):
        env, state, sent = f.BuildTests().build_env(3)
        env['bot'].send_message = AsyncMock(side_effect=ConnectionError('synthetic outage'))
        try:
            await env['_build_presentation'](self.message(), state)
        except ConnectionError:
            pass
        env['finish_job'].assert_called_once_with(-987654321)
        env['ask_grok'].assert_not_awaited()
        env['spend_credits'].assert_not_called()

    async def test_pdf_preserves_full_long_title(self):
        env, state, sent = f.BuildTests().build_env(3)
        title = 'Solar energy and sustainable development: research methods, practical applications and future opportunities'
        state.data['pres_plan']['title'] = title
        await env['_build_presentation'](self.message(), state)
        self.assertEqual(len(sent), 2)
        pdf = PdfReader(io.BytesIO(sent[1][1]))
        text = ' '.join(pdf.pages[0].extract_text().split())
        self.assertIn(title, text)

    async def test_pptx_send_failure_cleans_generated_files_but_keeps_retry_state(self):
        env, state, sent = f.BuildTests().build_env(3)
        env['bot'].send_document = AsyncMock(side_effect=ConnectionError('synthetic outage'))
        await env['_build_presentation'](self.message(), state)
        env['spend_credits'].assert_not_called()
        env['finish_job'].assert_called_once()
        self.assertTrue(state.data)
        attempted_path = env['bot'].send_document.await_args.args[1].path
        self.assertFalse(Path(attempted_path).exists())
        self.assertFalse(Path(attempted_path).parent.exists())

    async def test_cancelled_generation_releases_job_without_charge(self):
        env, state, sent = f.BuildTests().build_env(3)
        env['ask_grok'] = AsyncMock(side_effect=asyncio.CancelledError())
        with self.assertRaises(asyncio.CancelledError):
            await env['_build_presentation'](self.message(), state)
        env['finish_job'].assert_called_once()
        env['spend_credits'].assert_not_called()

    async def test_busy_job_never_starts_model_or_changes_balance(self):
        env, state, sent = f.BuildTests().build_env(3)
        env['start_job'].return_value = (False, 'Already building')
        await env['_build_presentation'](self.message(), state)
        env['ask_grok'].assert_not_awaited()
        env['spend_credits'].assert_not_called()
        env['finish_job'].assert_not_called()

    async def test_missing_balance_stops_before_job(self):
        env, state, sent = f.BuildTests().build_env(3)
        env['can_afford'] = lambda *args: False
        await env['_build_presentation'](self.message(), state)
        env['start_job'].assert_not_called()
        env['ask_grok'].assert_not_awaited()
        env['spend_credits'].assert_not_called()

    async def test_last_notification_failure_does_not_leave_paid_draft_active(self):
        env, state, sent = f.BuildTests().build_env(3)
        env['bot'].send_message = AsyncMock(side_effect=[None, ConnectionError('synthetic outage'), None])
        await env['_build_presentation'](self.message(), state)
        self.assertEqual(len(sent), 2)
        env['spend_credits'].assert_called_once()
        self.assertEqual(state.data, {})
        self.assertEqual(env['bot'].send_message.await_count, 2)


class CoverTests(unittest.TestCase):
    def test_sections_keep_text_inside_panel_with_bottom_padding(self):
        blocks = [{'heading': 'Участник', 'text': 'Исследует участок и описывает образцы.'} for _ in range(4)]
        scene = pd.build_scene({'title':'Распределение задач','content':'Роли участников.',
                                'designer':{'layout':'sections','blocks':blocks}},None,2,8)
        panel = next(n for n in scene['nodes'] if n['type']=='panel')
        bottom = panel['box'][1]+panel['box'][3]
        for node in scene['nodes']:
            if node.get('role') in ('body','subheading'):
                self.assertLessEqual(node['box'][1]+node['box'][3],bottom-20)

    def test_maximum_cjk_plan_title_kept_at_readable_size(self):
        title = ('现代科学技术的应用和未来发展方向' * 8)[:120]
        scene = pd.build_scene({'title': title, 'point': '研究方法与实践'}, None, 1, 8, 'zh')
        node = next(n for n in scene['nodes'] if n.get('role') == 'title')
        self.assertEqual(node['text'].replace('\n', ''), title)
        self.assertGreaterEqual(node['size'], 40)

    def test_long_user_photo_cover_and_subtitle_never_overlap(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / 'photo.png'
            Image.new('RGB', (200, 800), 'green').save(path)
            asset = {'path': str(path), 'width': 200, 'height': 800, 'origin': 'user'}
            title = ('现代科学技术的应用和未来发展方向' * 8)[:120]
            subtitle = ('研究方法与实践' * 30)[:180]
            scene = pd.build_scene({'title': title, 'point': subtitle}, asset, 1, 8, 'zh')
            photo = next(n for n in scene['nodes'] if n['type'] == 'image')
            for node in scene['nodes']:
                if node['type'] != 'text': continue
                x,y,w,h = node['box']; px,py,pw,ph = photo['box']
                self.assertFalse(x < px+pw and x+w > px and y < py+ph and y+h > py)
            texts = [n['text'].replace('\n','') for n in scene['nodes'] if n['type']=='text']
            self.assertIn(title,texts); self.assertIn(subtitle,texts)


if __name__ == '__main__':
    unittest.main()
