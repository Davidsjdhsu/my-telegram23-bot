import asyncio
import copy
import io
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import AsyncMock, patch

import httpx
from PIL import Image
import replicate_images as ri
import presentation_recovery as recovery
import presentation_visuals as pv
import test_presentation_support as f

REAL_CLIENT = httpx.AsyncClient
GOOD = {'relevant': True, 'has_text': False, 'has_chart': False, 'defective': False}


def ready(id='job1'):
    return {'id': id, 'status': 'succeeded', 'output': ['https://replicate.delivery/image.png']}


def png():
    out = io.BytesIO()
    Image.new('RGB', (96, 72), 'blue').save(out, format='PNG')
    return out.getvalue()


class ProviderTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = str(Path(self.temp.name) / 'image.png')
        self.requests = []

    def mock_http(self, responder):
        def handle(request):
            self.requests.append(request)
            return responder(request)
        transport = httpx.MockTransport(handle)
        return patch.object(ri.httpx, 'AsyncClient', side_effect=lambda **kw: REAL_CLIENT(transport=transport, **kw))

    def count_posts(self):
        return sum(r.method == 'POST' for r in self.requests)

    async def test_waits_beyond_old_thirty_poll_limit(self):
        polls = 0
        def response(request):
            nonlocal polls
            if request.method == 'POST':
                return httpx.Response(201, json={'id': 'job1', 'status': 'starting'})
            if request.url.host == 'replicate.delivery':
                self.assertNotIn('authorization', request.headers)
                return httpx.Response(200, content=png())
            polls += 1
            return httpx.Response(200, json=ready() if polls > 35 else {'status': 'processing'})
        job = {}
        with self.mock_http(response), patch.object(ri, 'POLL_SECONDS', 0):
            ok = await ri.generate_prediction('PRIVATE PROMPT', self.path, '4:3', 'SECRET', job)
        self.assertTrue(ok)
        self.assertEqual(polls, 36)
        self.assertEqual(self.count_posts(), 1)
        self.assertEqual(job['id'], 'job1')
        self.assertEqual(job['reason'], 'ready')
        with Image.open(self.path) as image:
            self.assertEqual(image.size, (96, 72))

    async def test_timeout_and_fresh_process_resume_existing_job(self):
        phase = 'pending'
        def response(request):
            if request.method == 'POST':
                return httpx.Response(201, json={'id': 'job1', 'status': 'starting'})
            if request.url.host == 'replicate.delivery':
                return httpx.Response(200, content=png())
            return httpx.Response(200, json=ready() if phase == 'ready' else {'status': 'processing'})
        job = {}
        with self.mock_http(response), patch.object(ri, 'WAIT_SECONDS', .01), patch.object(ri, 'POLL_SECONDS', .02):
            with self.assertLogs(ri.log, level='INFO') as logs:
                self.assertFalse(await ri.generate_prediction('PRIVATE', self.path, '4:3', 'SECRET', job))
            self.assertEqual(job['reason'], 'waiting_timeout')
            self.assertEqual(job['id'], 'job1')
            self.assertNotIn('PRIVATE', str(logs.output))
            self.assertNotIn('SECRET', str(logs.output))
            self.assertNotIn('job1', str(logs.output))
            job = json.loads(json.dumps(job))
            phase = 'ready'
            self.assertTrue(await ri.generate_prediction('PRIVATE', self.path, '4:3', 'SECRET', job))
        self.assertEqual(self.count_posts(), 1)
        self.assertNotIn('replicate.delivery', json.dumps(job))

    async def test_transient_get_and_download_errors_keep_id_without_post_retry(self):
        for fail_download in (False, True):
            self.requests.clear()
            phase = 'fail'
            def response(request):
                if request.method == 'POST':
                    return httpx.Response(201, json={'id': 'job1', 'status': 'starting'})
                if phase == 'fail' and (request.url.host == 'replicate.delivery') == fail_download:
                    raise httpx.ConnectError('SECRET', request=request)
                if request.url.host == 'replicate.delivery':
                    return httpx.Response(200, content=png())
                return httpx.Response(200, json=ready())
            job = {}
            with self.mock_http(response), patch.object(ri.asyncio, 'sleep', AsyncMock()):
                self.assertFalse(await ri.generate_prediction('text', self.path, '4:3', 'SECRET', job))
                self.assertEqual(job['id'], 'job1')
                self.assertEqual(job['reason'], 'generation_unavailable')
                phase = 'ready'
                self.assertTrue(await ri.generate_prediction('text', self.path, '4:3', 'SECRET', job))
            self.assertEqual(self.count_posts(), 1)

    async def test_terminal_failure_allows_fresh_job_on_explicit_retry(self):
        for status in ('failed', 'canceled', 'aborted'):
            self.requests.clear()
            phase = 'fail'
            def response(request):
                if request.method == 'POST':
                    return httpx.Response(201, json={'id': 'job1', 'status': status} if phase == 'fail' else ready('job2'))
                return httpx.Response(200, content=png())
            job = {}
            with self.mock_http(response):
                self.assertFalse(await ri.generate_prediction('text', self.path, '4:3', 'SECRET', job))
                self.assertNotIn('id', job)
                phase = 'ready'
                self.assertTrue(await ri.generate_prediction('text', self.path, '4:3', 'SECRET', job))
            self.assertEqual(self.count_posts(), 2)

    async def test_expired_job_does_not_loop_forever(self):
        phase = 'new'
        def response(request):
            if request.method == 'POST':
                return httpx.Response(201, json={'id': 'job1', 'status': 'starting'} if phase == 'new' else ready('job2'))
            if request.url.host == 'replicate.delivery':
                return httpx.Response(200, content=png())
            return httpx.Response(404)
        job = {}
        with self.mock_http(response), patch.object(ri, 'POLL_SECONDS', 0):
            self.assertFalse(await ri.generate_prediction('text', self.path, '4:3', 'SECRET', job))
            self.assertEqual(job['reason'], 'expired')
            self.assertNotIn('id', job)
            phase = 'ready'
            self.assertTrue(await ri.generate_prediction('text', self.path, '4:3', 'SECRET', job))
        self.assertEqual(self.count_posts(), 2)

    async def test_ambiguous_creation_never_automatically_reposts(self):
        def response(request):
            raise httpx.ReadTimeout('SECRET', request=request)
        with self.mock_http(response), self.assertLogs(ri.log, level='INFO') as logs:
            self.assertFalse(await ri.generate_prediction('PRIVATE', self.path, '4:3', 'SECRET', {}))
        self.assertEqual(self.count_posts(), 1)
        self.assertNotIn('SECRET', str(logs.output))

    async def test_changed_prompt_does_not_reuse_unrelated_job(self):
        def response(request):
            return httpx.Response(201, json=ready()) if request.method == 'POST' else httpx.Response(200, content=png())
        job = {}
        with self.mock_http(response):
            self.assertTrue(await ri.generate_prediction('one', self.path, '4:3', 'SECRET', job))
            self.assertTrue(await ri.generate_prediction('two', self.path, '4:3', 'SECRET', job))
        self.assertEqual(self.count_posts(), 2)

    async def test_outer_cancellation_keeps_prediction(self):
        def response(request):
            return httpx.Response(201, json={'id': 'job1', 'status': 'starting'})
        job = {}
        with self.mock_http(response), patch.object(ri, 'POLL_SECONDS', 10):
            with self.assertRaises(TimeoutError):
                await asyncio.wait_for(ri.generate_prediction('text', self.path, '4:3', 'SECRET', job), .01)
        self.assertEqual(job['id'], 'job1')
        self.assertEqual(job['reason'], 'waiting_timeout')


class PipelineTests(unittest.IsolatedAsyncioTestCase):
    async def test_real_build_checkpoint_resumes_provider_job_and_bills_once(self):
        env, state, sent = f.BuildTests().build_env(3)
        original = env['generate_image']
        phase = 'pending'
        calls = []
        async def generate(prompt, path, *, job, **kwargs):
            if 'generated-1-' in path:
                calls.append(copy.deepcopy(job))
                if phase == 'pending':
                    job.update(id='pending1', reason='waiting_timeout', status='processing')
                    return False
                self.assertEqual(job['id'], 'pending1')
            return await original(prompt, path, **kwargs)
        env['generate_image'] = generate
        message = SimpleNamespace(from_user=SimpleNamespace(id=-987654321))
        await env['_build_presentation'](message, state)
        self.assertEqual(sent, [])
        env['spend_credits'].assert_not_called()
        notice = env['bot'].send_message.await_args.args[1]
        self.assertIn('Не дождались готового изображения: 2', notice)
        self.assertNotIn('не прошло проверку качества', notice)
        phase = 'ready'
        await env['_build_presentation'](message, state)
        self.assertEqual(len(sent), 2)
        self.assertEqual(len(calls), 2)
        self.assertEqual(env['ask_grok'].await_count, 1)
        env['spend_credits'].assert_called_once_with(message.from_user.id, 'presentation', amount=48)
        self.assertNotIn('без фото', env['bot'].send_message.await_args.args[1])

    async def test_rejected_replacement_keeps_prompt_and_attempt_on_resume(self):
        with tempfile.TemporaryDirectory() as td:
            slides = [{'title': 'Dogs'}]
            calls = []
            async def generate(prompt, path, *, job, **kw):
                calls.append((prompt, copy.deepcopy(job)))
                if len(calls) == 2:
                    job.update(id='replacement', reason='waiting_timeout')
                    return False
                Path(path).write_bytes(png())
                return True
            inspector = AsyncMock(side_effect=[json.dumps({**GOOD, 'defective': True,
                'replacement_scene': 'A dog resting on a clean grass field'}), json.dumps(GOOD)])
            assets, warnings = await pv.prepare_images('Dogs', slides, [], {}, td, generate, inspector, json.loads, resumable=True)
            self.assertIsNone(assets[0])
            slides = json.loads(json.dumps(slides))
            assets, warnings = await pv.prepare_images('Dogs', slides, [], {}, td, generate, inspector, json.loads,
                                                       owned=(assets, warnings), resumable=True)
            self.assertEqual(calls[1][0], calls[2][0])
            self.assertEqual(calls[2][1]['id'], 'replacement')
            self.assertIsNotNone(assets[0])
            self.assertEqual(warnings['ai_omitted'], 0)

    def test_notice_classifies_failures_and_legacy_checkpoint_remains_truthful(self):
        slides = [{'_image_issue': 'waiting_timeout'}, {'_image_issue': 'quality_rejected'}, {'_image_issue': 'review_unavailable'}]
        notice = recovery.recovery_notice([1, 2, 3], 'ru', 60, slides)
        self.assertIn('Не дождались готового изображения: 1', notice)
        self.assertIn('не прошло проверку качества: 2', notice)
        self.assertIn('завершить проверку изображения: 3', notice)
        self.assertNotIn('балл', notice)
        legacy = recovery.recovery_notice([5], 'ru', 60)
        self.assertNotIn('проверку качества', legacy)
        self.assertIn('Кредиты не списаны', legacy)
