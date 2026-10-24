import unittest
from unittest.mock import AsyncMock, Mock
from aiohttp import web
from constructor_pricing import quote
from constructor_billing import ConstructorBilling

class PricingTests(unittest.TestCase):
 def deck(self,p=0,u=0):
  return {'version':1,'slides':[{'elements':[{'type':'photo'}]*p+[{'type':'pattern'}]*u}]}
 def test_every_allowed_combination(self):
  for photos in range(4):
   for patterns in range(4):
    cost=quote(self.deck(photos,patterns))['cost']
    self.assertEqual(cost,16+max(0,photos-1)*5+patterns*2)
    self.assertLessEqual(cost,32)
 def test_example_and_limits(self):
  decks=[self.deck(1,0),self.deck(1,1),self.deck(2,1),self.deck(3,3)]
  self.assertEqual(quote({'version':1,'slides':[d['slides'][0] for d in decks]})['cost'],89)
  for d in [self.deck(4,0),self.deck(0,4),{'version':1,'slides':[]},{'version':1,'slides':[{'elements':[]}]*21},self.deck()|{'slides':[{'elements':[{'type':'text'}]*7}]}]:
   with self.assertRaises(ValueError):quote(d)

class BillingTests(unittest.IsolatedAsyncioTestCase):
 def env(self,balance=100,admins=()):
  self.user={'credits':balance};self.persist=AsyncMock(return_value=True);self.send=AsyncMock();self.finish=Mock()
  return ConstructorBilling(lambda uid:self.user,self.persist,lambda uid:(True,''),self.finish,set(admins))
 async def test_charge_after_delivery_and_idempotence_across_instances(self):
  b=self.env()
  async def send(uid,data):
   self.assertEqual(self.user['credits'],100)
   self.assertEqual(self.user['constructor_orders']['abc']['status'],'pending')
  self.send.side_effect=send
  self.assertEqual((await b.deliver(1,'abc',32,b'file',self.send))['charged'],32)
  self.assertEqual(self.user['credits'],68)
  b2=ConstructorBilling(lambda uid:self.user,self.persist,lambda uid:(True,''),self.finish,set())
  self.assertTrue((await b2.deliver(1,'abc',32,b'file',self.send))['duplicate'])
  self.send.assert_awaited_once();self.assertEqual(self.user['credits'],68)
 async def test_insufficient_balance_and_admin(self):
  b=self.env(15)
  with self.assertRaises(web.HTTPPaymentRequired):await b.deliver(1,'a',16,b'',self.send)
  self.send.assert_not_awaited();self.persist.assert_not_awaited()
  b=self.env(0,{1});self.assertEqual((await b.deliver(1,'a',16,b'',self.send))['charged'],0)
  self.assertEqual(self.user['credits'],0)
 async def test_uncertain_send_does_not_charge_or_resend(self):
  b=self.env();self.send.side_effect=TimeoutError()
  with self.assertRaises(TimeoutError):await b.deliver(1,'a',16,b'',self.send)
  self.assertEqual(self.user['credits'],100)
  with self.assertRaises(web.HTTPConflict):await b.deliver(1,'a',16,b'',self.send)
  with self.assertRaises(web.HTTPConflict):await b.deliver(1,'different',16,b'',self.send)
  self.assertEqual(self.send.await_count,1)
 async def test_save_failure_prevents_send(self):
  b=self.env();self.persist.return_value=False
  with self.assertRaises(web.HTTPServiceUnavailable):await b.deliver(1,'a',16,b'',self.send)
  self.send.assert_not_awaited();self.assertEqual(self.user['credits'],100)
  self.assertEqual(self.user['constructor_orders'],{})
 async def test_save_failure_after_send_never_duplicates(self):
  b=self.env();self.persist.side_effect=[True,False]
  with self.assertRaises(web.HTTPServiceUnavailable):await b.deliver(1,'a',16,b'',self.send)
  self.assertTrue((await b.deliver(1,'a',16,b'',self.send))['duplicate'])
  self.send.assert_awaited_once();self.assertEqual(self.user['credits'],84)
