"""Bounded check of the bot recovery path. Never sends Telegram or charges users.

--live permits ONE image generation and ONE vision call using existing server
credentials. Text/plan/editor and delivery/billing are fixtures; PPTX/PDF are real.
Without --live every provider is a fixture. Never import the running bot module.
"""
import asyncio
import base64
import io
import json
import os
from pathlib import Path
import sys
import time
from types import SimpleNamespace
from unittest.mock import AsyncMock

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT / 'tests')]
from PIL import Image
from pptx import Presentation
from pypdf import PdfReader
import test_presentation_support as fixtures
import presentation_support as support
import presentation_pdf as pdf


async def main():
    live = '--live' in sys.argv
    out = Path('/tmp/alix-recovery-check')
    out.mkdir(exist_ok=True)
    env, state, sent = fixtures.BuildTests().build_env(3)
    plan = fixtures.make_plan(3)
    plan['title'] = 'Доберман'
    for slide, title in zip(plan['slides'], ['Доберман', 'На прогулке', 'Дома']):
        slide['title'] = title
        slide['point'] = title
    content = fixtures.make_content(3)
    content['title'] = 'Доберман'
    content['slides'][0].update(content='Доберман стоит на траве в парке.', image_prompt='One adult black and tan Doberman standing on grass in a quiet park, full body including all four paws, soft daylight, simple background')
    content['slides'][1].update(content='Собака отдыхает в комнате.', image_prompt='A Doberman resting in a bright room')
    state.data.update(topic='Доберман', pres_plan=plan)
    state.data['pres_plan_key'] = support.plan_key(state.data)
    env['ask_grok'] = AsyncMock(return_value=json.dumps(content))
    env.update(ensure_pdf_converter=pdf.ensure_pdf_converter, convert_presentation_pdf=pdf.convert_presentation_pdf)
    original = env['generate_image']
    async def missing_second(prompt, path, **kw):
        if 'generated-1-' in path:
            return False
        return await original(prompt, path, **kw)
    env['generate_image'] = AsyncMock(side_effect=missing_second)
    message = SimpleNamespace(from_user=SimpleNamespace(id=-987654321))
    await env['_build_presentation'](message, state)
    assert not sent and env['spend_credits'].call_count == 0
    counts = {'image': 0, 'vision': 0}
    client = None
    if live:
        import httpx
        from openai import AsyncOpenAI
        client = AsyncOpenAI(api_key=os.environ['XAI_API_KEY'], base_url='https://api.x.ai/v1', max_retries=0)
        async def heavy(fn, *args):
            return await asyncio.to_thread(fn, *args)
        api = fixtures.load_functions(['generate_image','ask_presentation_vision','_prepare_vision_image'], {
            'httpx':httpx, 'asyncio':asyncio, 'Image':Image, 'run_heavy':heavy, 'base64':base64, 'os':os,
            'REPLICATE_API_TOKEN':os.environ['REPLICATE_API_TOKEN'], 'client':client,
            'VISION_MODEL':os.getenv('XAI_VISION_MODEL','grok-4.6')})
        async def generate_once(*args, **kwargs):
            if counts['image']:
                return False
            counts['image'] += 1
            return await api['generate_image'](*args, **kwargs)
        async def inspect_once(*args, **kwargs):
            if counts['vision']:
                raise RuntimeError('Live check budget exhausted')
            counts['vision'] += 1
            return await api['ask_presentation_vision'](*args, **kwargs)
        env.update(generate_image=generate_once, ask_presentation_vision=inspect_once)
    else:
        env['generate_image'] = original
    started = time.perf_counter()
    try:
        await env['_build_presentation'](message, state)
    finally:
        if client:
            await client.close()
    report = {'live':live, 'calls':counts, 'seconds':round(time.perf_counter()-started,2),
              'files':len(sent), 'author_calls':env['ask_grok'].await_count,
              'simulated_billing_calls':env['spend_credits'].call_count,
              'scope':'Bot recovery function; real PPTX/PDF; fixture text/editor/Telegram/billing'}
    if len(sent) == 2:
        report['slides'] = len(Presentation(io.BytesIO(sent[0][1])).slides)
        report['pages'] = len(PdfReader(io.BytesIO(sent[1][1])).pages)
        for name, payload in sent:
            (out / ('recovered' + Path(name).suffix)).write_bytes(payload)
    (out / 'result.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
    print('RECOVERY_CHECK ' + json.dumps(report), flush=True)
    assert len(sent) == 2, 'Recovery did not complete within the bounded check'
    assert env['ask_grok'].await_count == 1
    assert env['spend_credits'].call_count == 1


if __name__ == '__main__':
    asyncio.run(main())
