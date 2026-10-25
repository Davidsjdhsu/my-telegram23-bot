import time
import unittest
from unittest.mock import AsyncMock
from xml.etree import ElementTree as ET
from aiohttp import web, CookieJar
from aiohttp.test_utils import TestClient, TestServer
from studio_private import Studio, COOKIE
from studio_palette_preview import render_preview, scene_for, svg_for, PALETTES, BASE

class PreviewSceneTests(unittest.TestCase):
 def test_all_templates_light_dark_all_pages(self):
  for style in ('strict','soft'):
   for number in range(1,51):
    for palette in ('tech','school'):
     for page in range(4):
      with self.subTest(style=style,number=number,palette=palette,page=page):
       scene=scene_for(style,f'd{number:02}',palette,page)
       svg=svg_for(scene,'test');root=ET.fromstring(svg)
       self.assertEqual(root[1].get('fill'),scene['background'])
       self.assertNotIn(str(BASE),svg)
       self.assertNotIn('<script',svg)
       images=root.findall('.//{*}image')
       photo=[im for im in images if im.get('href')=='/studio/constructor/coffee.jpg']
       self.assertEqual(len(photo),int(page in (0,2)))
       if photo:self.assertEqual(photo[0].get('opacity'),'1')
       self.assertTrue(any('/preview-pattern/' in im.get('href','') for im in images))
 def test_all_palettes_use_production_background_and_bounded_cache(self):
  render_preview.cache_clear()
  for palette,config in PALETTES.items():
   for page in range(4):
    svg=render_preview('strict','d16',palette,page)
    self.assertEqual(ET.fromstring(svg)[1].get('fill').upper(),config['bg'].upper())
  self.assertLessEqual(render_preview.cache_info().currsize,128)

class PreviewEndpointTests(unittest.IsolatedAsyncioTestCase):
 async def asyncSetUp(self):
  self.send=AsyncMock();self.action=AsyncMock()
  self.studio=Studio(None,self.send,self.action,'http://example.test')
  app=web.Application();self.studio.register(app)
  self.client=TestClient(TestServer(app),cookie_jar=CookieJar(unsafe=True));await self.client.start_server()
 async def asyncTearDown(self):
  await self.client.close()
  for archive in self.studio.archives:archive.close()
 def login(self):
  self.studio.sessions['test-session']={'expires':time.monotonic()+30,'user':{'id':465823470}}
  self.client.session.cookie_jar.update_cookies({COOKIE:'test-session'})
 async def test_requires_session_for_preview_and_pattern(self):
  for path in ('palette-preview?style=strict&key=d16&palette=tech&page=0','preview-pattern/strict-d16.jpg'):
   self.assertEqual((await self.client.get('/studio/templates/'+path)).status,401)
 async def test_preview_and_assets_without_generation_or_charge(self):
  self.login()
  r=await self.client.get('/studio/templates/palette-preview',params={'style':'soft','key':'d16','palette':'school','page':'0'})
  self.assertEqual(r.status,200)
  svg=(await r.json())['svg']
  for im in ET.fromstring(svg).findall('.//{*}image'):
   asset=await self.client.get(im.get('href'));self.assertEqual(asset.status,200)
  self.send.assert_not_awaited();self.action.assert_not_awaited()
 async def test_invalid_choices_rejected(self):
  self.login();valid={'style':'strict','key':'d16','palette':'tech','page':'0'}
  for field,value in [('style','standard'),('key','d00'),('key','d51'),('key','../../main.py'),('palette','bad'),('page','4'),('page','-1')]:
   r=await self.client.get('/studio/templates/palette-preview',params={**valid,field:value});self.assertEqual(r.status,400)
  r=await self.client.get('/studio/templates/preview-pattern/main.py');self.assertEqual(r.status,404)
