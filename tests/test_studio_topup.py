import asyncio
import unittest
from types import SimpleNamespace as NS
from unittest.mock import AsyncMock
from aiohttp import web
from aiohttp.test_utils import TestClient,TestServer
from studio_topup import attach
from credit_shortage import details,notice

class TopupTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.users={1:{'credits':40},2:{'credits':9}}
        def authorized(request):
            if not request.headers.get('Tester'):raise web.HTTPUnauthorized()
            return {'user':{'id':int(request.headers['Tester'])}}
        def check(request):
            if request.headers.get('X-Studio-Request')!='1':raise web.HTTPForbidden()
        async def body(request,limit):return await request.json()
        self.studio=NS(authorized=authorized,check_write=check,body=body)
        self.provider=AsyncMock(return_value='https://yookassa.ru/test-only')
        self.save=AsyncMock()
        api=dict(get_user=lambda uid:self.users[uid],billing_saved=self.save,create_yookassa_payment=self.provider,
                 available_credits=lambda uid:self.users[uid]['credits'],PAYMENT_PROVIDER_TOKEN='')
        app=web.Application();attach(self.studio,app,api)
        self.client=TestClient(TestServer(app));await self.client.start_server()
        self.headers={'Tester':'1','X-Studio-Request':'1'}
        self.body={'amount':60,'id':'test-only-key-123456'}
    async def asyncTearDown(self):await self.client.close()
    async def test_auth_validation_and_no_charge_on_invoice_creation(self):
        self.assertEqual((await self.client.post('/studio/topup',json=self.body)).status,401)
        self.assertEqual((await self.client.post('/studio/topup',json=self.body,headers={'Tester':'1'})).status,403)
        for amount in (0,59,10000,60.5,True,'100'):
            self.assertEqual((await self.client.post('/studio/topup',json=self.body|{'amount':amount},headers=self.headers)).status,400)
        self.provider.assert_not_awaited()
        r=await self.client.post('/studio/topup',json=self.body,headers=self.headers)
        self.assertEqual(r.status,200);self.assertFalse((await r.json())['paid'])
        self.assertEqual(self.users[1]['credits'],40)
    async def test_repeat_concurrent_same_invoice_and_other_account_isolation(self):
        a,b=await asyncio.gather(*[self.client.post('/studio/topup',json=self.body,headers=self.headers) for _ in range(2)])
        self.assertEqual(a.status,200);self.assertEqual(b.status,200);self.provider.assert_awaited_once()
        r=await self.client.post('/studio/topup',json=self.body|{'amount':100},headers=self.headers)
        self.assertEqual(r.status,409)
        r=await self.client.get('/studio/topup/'+self.body['id'],headers={'Tester':'2'})
        self.assertEqual(r.status,404)
    async def test_client_paid_flag_or_increased_balance_is_not_payment_proof(self):
        await self.client.post('/studio/topup',json=self.body|{'paid':True,'balance':10000},headers=self.headers)
        self.users[1]['credits']=200
        path='/studio/topup/'+self.body['id']
        self.assertFalse((await (await self.client.get(path,headers=self.headers)).json())['paid'])
        self.users[1]['studio_topups'][self.body['id']]['paid']=True
        self.assertTrue((await (await self.client.get(path,headers=self.headers)).json())['paid'])
    async def test_provider_failure_preserves_balance_and_intent_key(self):
        self.provider.return_value=None
        r=await self.client.post('/studio/topup',json=self.body,headers=self.headers);self.assertEqual(r.status,503)
        self.assertEqual(self.users[1]['credits'],40)
        self.provider.return_value='https://yookassa.ru/test-only'
        r=await self.client.post('/studio/topup',json=self.body,headers=self.headers);self.assertEqual(r.status,200)
        self.assertEqual(self.provider.await_args.kwargs['request_id'],self.body['id'])

class ShortageTests(unittest.TestCase):
    def test_fractional_shortage_and_minimum(self):
        self.assertEqual(details(100,40)['topup'],60)
        self.assertEqual(details(.5,0)['topup'],60)
        self.assertEqual(details(100,9.5)['topup'],91)
        self.assertIn('0,5',notice(.5,0))
        self.assertIn('Минимальное',notice(.5,0))
