"""Transient GET failures must not create another billable prediction."""
import ast
import io
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import AsyncMock, patch
import httpx
import replicate_images as ri
from PIL import Image

TransportError = httpx.TransportError

class Response:
    def __init__(self, data=None, status=200, content=b''):
        self.data=data;self.status_code=status;self.content=content
    def json(self):return self.data
    def raise_for_status(self):
        if self.status_code>=400:
            response=httpx.Response(self.status_code,request=httpx.Request('GET','https://api.replicate.com/v1/predictions/one'))
            response.raise_for_status()

class ReplicateDeliveryTests(unittest.IsolatedAsyncioTestCase):
    def setup_generator(self, gets, post_error=None):
        image=io.BytesIO();Image.new('RGB',(80,60),'blue').save(image,'PNG')
        values=[Response(content=image.getvalue()) if value=='image' else value for value in gets]
        http=SimpleNamespace(post=AsyncMock(side_effect=post_error,return_value=Response({'id':'one','status':'starting'},201)),get=AsyncMock(side_effect=values))
        class Client:
            async def __aenter__(self):return http
            async def __aexit__(self,*args):pass
        async def heavy(fn):return fn()
        client_patch = patch.object(ri.httpx, 'AsyncClient', lambda **kw: Client())
        sleep_patch = patch.object(ri.asyncio, 'sleep', AsyncMock())
        client_patch.start(); sleep_patch.start()
        self.addCleanup(client_patch.stop); self.addCleanup(sleep_patch.stop)
        env={'REPLICATE_API_TOKEN':'test-only','Image':Image,'run_heavy':heavy}
        root=Path(__file__).resolve().parents[1]
        fn=next(n for n in ast.parse((root/'main.py').read_text()).body if isinstance(n,ast.AsyncFunctionDef) and n.name=='generate_image')
        exec(compile(ast.Module(body=[fn],type_ignores=[]),'actual-generate-image','exec'),env)
        return env['generate_image'],http

    async def test_poll_read_error_reuses_prediction(self):
        gen,http=self.setup_generator([TransportError('fixture'),Response({'status':'succeeded','output':['https://image/result']}),'image'])
        with tempfile.TemporaryDirectory() as td:
            self.assertTrue(await gen('Public fixture',str(Path(td)/'image.png')))
        self.assertEqual(http.post.await_count,1)
        self.assertEqual(http.get.await_args_list[0].args,http.get.await_args_list[1].args)
        self.assertNotIn('headers',http.get.await_args_list[-1].kwargs)

    async def test_download_503_retries_same_asset(self):
        gen,http=self.setup_generator([Response({'status':'succeeded','output':['https://image/result']}),Response(status=503),'image'])
        with tempfile.TemporaryDirectory() as td:self.assertTrue(await gen('Public fixture',str(Path(td)/'image.png')))
        self.assertEqual(http.post.await_count,1)
        self.assertEqual(http.get.await_args_list[1].args,http.get.await_args_list[2].args)

    async def test_auth_error_is_not_retried(self):
        gen,http=self.setup_generator([Response(status=401)])
        with tempfile.TemporaryDirectory() as td:self.assertFalse(await gen('Public fixture',str(Path(td)/'image.png')))
        self.assertEqual(http.get.await_count,1)

    async def test_ambiguous_creation_is_not_retried(self):
        gen,http=self.setup_generator([],TransportError('fixture'))
        with tempfile.TemporaryDirectory() as td:self.assertFalse(await gen('Public fixture',str(Path(td)/'image.png')))
        self.assertEqual(http.post.await_count,1);http.get.assert_not_awaited()

    async def test_persistent_get_failure_is_bounded(self):
        gen,http=self.setup_generator([TransportError('fixture'),TransportError('fixture'),TransportError('fixture')])
        with tempfile.TemporaryDirectory() as td:self.assertFalse(await gen('Public fixture',str(Path(td)/'image.png')))
        self.assertEqual(http.get.await_count,3);self.assertEqual(http.post.await_count,1)

    async def test_invalid_image_is_rejected_by_actual_wrapper(self):
        gen,http=self.setup_generator([Response({'status':'succeeded','output':['https://image/result']}),Response(content=b'not an image')])
        job={}
        with tempfile.TemporaryDirectory() as td:
            self.assertFalse(await gen('Public fixture',str(Path(td)/'image.png'),job=job))
        self.assertEqual(job,{'reason':'invalid_image'})
