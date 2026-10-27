import asyncio
import json
import unittest
from datetime import datetime, timezone
from types import SimpleNamespace as NS
from unittest.mock import AsyncMock
from aiogram.types import Message
import chat_allowance as q
import chat_billing_flow as flow
import wallet
import test_chat_support
import ast
from pathlib import Path

class AllowanceTests(unittest.TestCase):
    def test_shared_ten_any_mix_and_paid_consent(self):
        for kinds in (['chat']*10, ['voice']*10, ['vision']*10, ['chat','voice','vision']*3+['chat']):
            u={'credits':20}
            for i,kind in enumerate(kinds):
                result,ticket=q.admit(u,str(i),kind,affordable=lambda c:False)
                self.assertEqual(result,'ok');q.settle(u,ticket);q.release(u,ticket)
            self.assertEqual(q.status(u)['remaining'],0);self.assertEqual(u['credits'],20)
            self.assertEqual(q.admit(u,'extra','chat',affordable=lambda c:True)[0],'consent')
            q.consent(u,True)
            for kind,price in q.PRICES.items():
                result,ticket=q.admit(u,'paid:'+kind,kind,affordable=lambda c:True)
                self.assertEqual(ticket['cost'],price);q.settle(u,ticket);q.settle(u,ticket);q.release(u,ticket)
            self.assertEqual(u['credits'],13.5);self.assertEqual(len(wallet.public_history(u)['items']),3)

    def test_failure_concurrency_duplicate_restart_and_midnight(self):
        before=datetime(2026,10,27,20,59,59,tzinfo=timezone.utc)
        after=datetime(2026,10,27,21,0,0,tzinfo=timezone.utc)
        u={'credits':1}
        _,t=q.admit(u,'first','vision',affordable=lambda c:False,now=before)
        self.assertEqual(q.admit(u,'parallel','voice',affordable=lambda c:True,now=before)[0],'busy')
        q.release(u,t);self.assertEqual(q.status(u,before)['remaining'],10)
        _,t=q.admit(u,'first','vision',affordable=lambda c:False,now=before)
        q.settle(u,t);q.release(u,t);q.consent(u,True,before)
        restored=json.loads(json.dumps(wallet.snapshot(u)))
        self.assertEqual(q.status(restored,before)['remaining'],9)
        self.assertEqual(q.admit(restored,'first','vision',affordable=lambda c:True,now=before)[0],'duplicate')
        self.assertTrue(q.status(restored,before)['paid_enabled'])
        self.assertEqual(q.status(restored,after)['remaining'],10)
        self.assertFalse(q.status(restored,after)['paid_enabled'])

    def test_no_money_no_consent_and_disable(self):
        u={'credits':0,'chat_allowance':{'day':q.day(),'receipts':{str(i):'chat' for i in range(10)}}}
        self.assertEqual(q.admit(u,'11','chat',affordable=lambda c:False)[0],'consent')
        q.consent(u,True)
        self.assertEqual(q.admit(u,'11','chat',affordable=lambda c:False)[0],'balance')
        q.consent(u,False)
        self.assertEqual(q.admit(u,'11','chat',affordable=lambda c:True)[0],'consent')
        self.assertEqual(u['credits'],0)

class FlowTests(unittest.IsolatedAsyncioTestCase):
    def message(self,i=1):
        return Message.model_validate(dict(message_id=i,date=int(datetime.now().timestamp()),chat={'id':1,'type':'private'},
                                           from_user={'id':1,'is_bot':False,'first_name':'Test'},text='Вопрос'))
    async def test_gate_saves_question_before_consent_and_no_provider_call(self):
        u={'credits':100,'chat_allowance':{'day':q.day(),'receipts':{str(i):'chat' for i in range(10)}}}
        state=NS(get_state=AsyncMock(return_value=None),get_data=AsyncMock(return_value={}))
        bot=NS(send_message=AsyncMock())
        api=dict(get_user=lambda _:u,ADMIN_IDS=set(),can_afford=lambda uid,c:True,bot=bot,billing_saved=AsyncMock())
        self.assertIsNone(await flow.begin(api,self.message(),'chat',text='Вопрос',state=state))
        self.assertEqual(u['credits'],100)
        self.assertEqual(u['billing_resume']['message']['text'],'Вопрос')
        self.assertIn('0,5',bot.send_message.await_args.args[1])
        self.assertIn('10',bot.send_message.await_args.args[1])
        self.assertEqual(q.status(u)['remaining'],0)
        q.consent(u,True)
        ticket=await flow.begin(api,self.message(),'chat',text='Вопрос',state=state)
        q.settle(u,ticket);q.release(u,ticket)
        self.assertEqual(u['credits'],99.5)

    async def test_actual_chat_handlers_free_paid_failed_and_duplicate(self):
        ns,users=test_chat_support.IntegrationTests().chat_ns();u=users[1]
        ns['ask_grok_chat']=AsyncMock(return_value='Ответ')
        u['credits']=0
        for i in range(10):
            await ns['respond_in_chat'](self.message(i+1),'Вопрос','ru')
            # ChatGate protects rapid traffic; reset only its test clock counters.
            ns['_chat_gate']=__import__('chat_support').ChatGate()
        self.assertEqual(q.status(u)['remaining'],0);self.assertEqual(u['credits'],0)
        q.consent(u,True);u['credits']=3
        await ns['respond_in_chat'](self.message(11),'Вопрос','ru')
        self.assertEqual(u['credits'],2.5)
        ns['_chat_gate']=__import__('chat_support').ChatGate()
        await ns['respond_in_chat'](self.message(11),'Вопрос','ru')
        self.assertEqual(ns['ask_grok_chat'].await_count,11)
        ns['_chat_gate']=__import__('chat_support').ChatGate()
        ns['ask_grok_chat'].return_value=None
        await ns['respond_in_chat'](self.message(12),'Вопрос','ru')
        self.assertEqual(u['credits'],2.5)
        self.assertNotIn('_chat_access_active',u)

    async def test_voice_checks_consent_before_transcription_and_order_input_is_free(self):
        tree=ast.parse((Path(__file__).parents[1]/'main.py').read_text())
        cls=next(n for n in tree.body if isinstance(n,ast.ClassDef) and n.name=='VoiceToTextMiddleware')
        class FakeMessage:
            def __init__(self):
                self.from_user=NS(id=1);self.chat=NS(id=1);self.message_id=19
                self.voice=NS(duration=4);self.text=None
                self.answer=AsyncMock(return_value=NS(delete=AsyncMock()))
            def model_copy(self,update):
                self.__dict__.update(update);return self
        u={'credits':0,'chat_allowance':{'day':q.day(),'receipts':{str(i):'chat' for i in range(10)}}}
        state=NS(get_state=AsyncMock(return_value=None),get_data=AsyncMock(return_value={}))
        bot=NS(send_message=AsyncMock())
        env=dict(BaseMiddleware=object,Message=FakeMessage,user_lang=lambda uid:'ru',get_user=lambda uid:u,
                 voice_rate_limited=lambda uid:False,transcribe_voice=AsyncMock(return_value='Мой ответ'),
                 MAX_VOICE_SECONDS=120,tr=lambda *a:a[0],chat_billing_flow=flow,chat_allowance=q,
                 ADMIN_IDS=set(),can_afford=lambda uid,c:False,bot=bot,billing_saved=AsyncMock())
        exec(compile(ast.Module(body=[cls],type_ignores=[]),'main.py','exec'),env)
        middleware=env['VoiceToTextMiddleware']();handler=AsyncMock()
        await middleware(handler,FakeMessage(),{'state':state})
        env['transcribe_voice'].assert_not_awaited();handler.assert_not_awaited()
        state.get_state.return_value='Form:waiting_word_topic'
        await middleware(handler,FakeMessage(),{'state':state})
        env['transcribe_voice'].assert_awaited_once();handler.assert_awaited_once()
        self.assertEqual(u['credits'],0);self.assertEqual(q.status(u)['remaining'],0)
        u['chat_allowance']={'day':q.day(),'receipts':{}};state.get_state.return_value=None
        env['transcribe_voice'].return_value=None
        await middleware(handler,FakeMessage(),{'state':state})
        self.assertEqual(q.status(u)['remaining'],10);self.assertNotIn('_chat_access_active',u)

if __name__=='__main__':unittest.main()
