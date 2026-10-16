"""Offline check of the real builder and converter; no APIs, Telegram or billing."""
import asyncio
import io
from pathlib import Path
import sys
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT / 'tests')]
from pypdf import PdfReader
from pptx import Presentation
import presentation_pdf
from test_presentation_support import BuildTests


async def main():
    env, state, sent = BuildTests().build_env(8)
    env['ensure_pdf_converter'] = presentation_pdf.ensure_pdf_converter
    env['convert_presentation_pdf'] = presentation_pdf.convert_presentation_pdf
    await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-1)), state)
    assert len(sent) == 2, 'Expected both PPTX and PDF'
    deck = Presentation(io.BytesIO(sent[0][1]))
    pdf = PdfReader(io.BytesIO(sent[1][1]))
    assert len(deck.slides) == len(pdf.pages) == 8
    assert all(page.images for page in pdf.pages)
    env['spend_credits'].assert_called_once()  # Mock: no real user is charged.
    assert state.data == {}
    print('presentation_pdf delivery_check=passed pages=8 images=yes files=2 external_calls=0', flush=True)


if __name__ == '__main__':
    asyncio.run(main())
