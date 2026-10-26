import tempfile,time,unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock
from aiohttp import web
from aiohttp.test_utils import TestClient,TestServer
from studio_files import FileLibrary

class LibraryTests(unittest.IsolatedAsyncioTestCase):
 async def test_owner_isolation_grouping_retrieval_and_throttling(self):
  with tempfile.TemporaryDirectory() as d:
   lib=FileLibrary(Path(d)/'files');msg=SimpleNamespace(document=SimpleNamespace(file_id='saved'))
   lib.record(1,msg,'Тест','pptx','one');lib.record(1,msg,'Тест','pdf','one')
   self.assertEqual(len(lib.public(1)),1);self.assertEqual(lib.public(2),[])
   self.assertNotIn('saved',str(lib.public(1)))
   studio=SimpleNamespace(authorized=lambda r:{'user':{'id':int(r.headers.get('User','1'))}},check_write=lambda r:None)
   bot=SimpleNamespace(send_document=AsyncMock());app=web.Application();lib.attach(studio,app,bot)
   async with TestClient(TestServer(app)) as client:
    url='/studio/files/'+lib.public(1)[0]['id']+'/open'
    self.assertEqual((await client.post(url,headers={'User':'2'})).status,404)
    self.assertEqual((await client.post(url)).status,200)
    self.assertEqual(bot.send_document.await_count,2)
    self.assertEqual((await client.post(url)).status,429)
    self.assertEqual(bot.send_document.await_count,2)
