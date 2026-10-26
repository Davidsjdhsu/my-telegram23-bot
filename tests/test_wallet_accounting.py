"""Real persistence boundary, fractional credits and repeat/failure accounting."""
import asyncio, copy, json, tempfile, unittest
from pathlib import Path
from types import SimpleNamespace as NS
from unittest.mock import Mock, AsyncMock
import wallet
from service_pricing import word_price, excel_price, VERSION
from test_presentation_support import load_functions
import test_chat_support as chat_tests
from order_billing import charge_callback

class LedgerTests(unittest.TestCase):
 def test_fractional_balance_survives_real_serializer_and_restart(self):
  u={'credits':100,'history':[]}
  wallet.apply(u,'chat:1',-.5,'Ответ в чате')
  wallet.apply(u,'topup:1',100,'Пополнение','topup')
  u['presentation_payments']={'order':{'status':'paid','charged':20}}
  u['constructor_orders']={'constructor':{'status':'pending','cost':10}}
  env=load_functions(['_build_users_payload'],dict(users_db={1:u},STARTING_CREDITS=0,USERS_META_KEY='_meta',_processed_yookassa_payment_ids=set(),_processed_tg_charge_ids=set(),_stats_meta={'since':'test'}))
  snapshot=env['_build_users_payload']();restored=json.loads(json.dumps(snapshot))['1']
  self.assertEqual(restored['credits'],199.5)
  self.assertEqual(restored['presentation_payments'],u['presentation_payments'])
  self.assertEqual(restored['constructor_orders'],u['constructor_orders'])
  wallet.apply(restored,'chat:1',-.5,'Duplicate')
  wallet.apply(restored,'topup:1',100,'Duplicate','topup')
  self.assertEqual(restored['credits'],199.5)
  wallet.apply(u,'new',-20,'Word')
  self.assertEqual(snapshot['1']['credits'],199.5)
  self.assertNotIn('new',snapshot['1']['wallet_transactions'])
 def test_balance_and_receipt_never_diverge_or_go_negative(self):
  u={'credits':.5};wallet.apply(u,'one',-.5,'Chat')
  for n in (-.5,-1):
   with self.assertRaises(ValueError):wallet.apply(u,'bad',n,'Denied')
  self.assertEqual(u['credits'],0);self.assertEqual(len(u['wallet_transactions']),1)
  with self.assertRaises(ValueError):wallet.apply(u,'one',-1,'Changed price')
  for n in (.1,float('nan'),float('inf')):
   with self.assertRaises(ValueError):wallet.units(n)
 def test_history_private_ids_deduplicated_and_paginated(self):
  u={'credits':100};wallet.apply(u,'yookassa:private',100,'Пополнение','topup')
  for n in range(55):wallet.apply(u,'chat:'+str(n),-.5,'Ответ')
  old=[dict(id='yookassa:private',ts='2000',delta=100,title='Old')]
  first=wallet.public_history(u,old);second=wallet.public_history(u,old,50)
  self.assertEqual(first['total'],56);self.assertEqual(len(first['items']),50)
  self.assertEqual(len(second['items']),6);self.assertIsNone(second['next_offset'])
  self.assertNotIn('private',json.dumps(first));self.assertNotIn('id',first['items'][0])
 def test_fixed_documents_are_independent_of_layout(self):
  for n in (1,2,8):
   self.assertEqual(word_price(n,'dkp_car'),40);self.assertEqual(word_price(n,'statement'),20)
  self.assertEqual(word_price(2,'essay'),20);self.assertEqual(word_price(6,'essay'),30)
  self.assertEqual(excel_price('price_list'),20);self.assertEqual(excel_price('startup_model'),40)

class PaymentTests(unittest.IsolatedAsyncioTestCase):
 async def test_concurrent_payment_and_restart_charge_once_and_preserve_half_credit(self):
  u={'credits':100.5};disk={}
  async def persist():disk.clear();disk.update(copy.deepcopy(u));await asyncio.sleep(0);return True
  api=dict(get_user=lambda _:u,save_users_now=persist,ADMIN_IDS=set(),can_afford=lambda uid,c:u['credits']>=c,note_success=Mock())
  charge=charge_callback(api,'studio_excel_payments','Excel')
  await asyncio.gather(charge(1,'a',20,'Таблица'),charge(1,'a',20,'Таблица'))
  self.assertEqual(u['credits'],80.5);self.assertEqual(len(u['wallet_transactions']),1)
  u.clear();u.update(copy.deepcopy(disk));await charge_callback(api,'studio_excel_payments','Excel')(1,'a',20,'Таблица')
  self.assertEqual(u['credits'],80.5)
 async def test_failed_receipt_save_retries_without_second_debit(self):
  u={'credits':100.5};save=AsyncMock(side_effect=[True,False,True])
  api=dict(get_user=lambda _:u,save_users_now=save,ADMIN_IDS=set(),can_afford=lambda *a:True,note_success=Mock())
  charge=charge_callback(api,'presentation_payments','Презентация')
  with self.assertRaises(RuntimeError):await charge(1,'a',20,'Тест')
  self.assertEqual(u['credits'],80.5)
  await charge(1,'a',20,'Тест');self.assertEqual(u['credits'],80.5);self.assertEqual(len(u['wallet_transactions']),1)
 async def test_precheckout_rejects_mismatched_amount_and_currency(self):
  bot=NS(answer_pre_checkout_query=AsyncMock())
  ns=load_functions(['pre_checkout_handler'],{'bot':bot})
  for payload,currency,total,ok in [('credits:custom:100','RUB',10000,True),('credits:100','RUB',1,False),('credits:100','USD',10000,False),('bad:100','RUB',10000,False),('credits:-1','RUB',-100,False)]:
   await ns['pre_checkout_handler'](NS(id='test',invoice_payload=payload,currency=currency,total_amount=total))
   self.assertEqual(bot.answer_pre_checkout_query.await_args.kwargs['ok'],ok)
 async def test_verified_telegram_payment_duplicate_keeps_fraction(self):
  u={'credits':99.5};save=AsyncMock(return_value=True)
  ns=load_functions(['successful_payment_handler','billing_saved'],dict(user_lang=lambda _:'ru',get_user=lambda _:u,_processed_tg_charge_ids=set(),save_users_now=save,_notify_admins=AsyncMock(),log_payment=Mock(),PAYMENT_NOTIFY_IDS=set(),tr=lambda *a,**kw:'ok',asyncio=asyncio))
  m=NS(from_user=NS(id=1,username='test'),answer=AsyncMock(),successful_payment=NS(invoice_payload='credits:custom:100',telegram_payment_charge_id='receipt',currency='RUB',total_amount=10000))
  await ns['successful_payment_handler'](m);await ns['successful_payment_handler'](m)
  self.assertEqual(u['credits'],199.5);self.assertEqual(len(u['wallet_transactions']),1)
  m.successful_payment.telegram_payment_charge_id='bad';m.successful_payment.total_amount=1
  await ns['successful_payment_handler'](m);self.assertEqual(u['credits'],199.5)
 async def test_chat_voice_success_failure_and_repeat(self):
  for voice,cost in ((False,.5),(True,3)):
   ns,users=chat_tests.IntegrationTests().chat_ns()
   u=users[1];ns['CREDIT_COSTS']={'chat':.5,'voice':3}
   ns['spend_credits']=lambda uid,action,operation_id:wallet.apply(u,operation_id,-cost,action)
   ns['maybe_send_voice_reply']=AsyncMock()
   m=NS(from_user=NS(id=1),chat=NS(id=1),message_id=987,voice=NS() if voice else None)
   ns['ask_grok_chat']=AsyncMock(return_value='Готовый ответ')
   await ns['respond_in_chat'](m,'Объясни закон','ru')
   self.assertEqual(u['credits'],100-cost)
   await ns['respond_in_chat'](m,'Объясни закон','ru')
   self.assertEqual(u['credits'],100-cost)
   m.message_id=988;ns['ask_grok_chat'].return_value=None
   await ns['respond_in_chat'](m,'Другой вопрос','ru');self.assertEqual(u['credits'],100-cost)
   self.assertNotIn('_chat_reserved',u)
