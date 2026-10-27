"""Telegram admission, consent and explicit resume of saved requests."""
import time
import secrets
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, Message, Update
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.base import StorageKey
import chat_allowance as quota
from wallet import request_key
from credit_shortage import details, notice

def state_for(api, uid):
    return FSMContext(storage=api['dp'].storage, key=StorageKey(bot_id=api['bot'].id, chat_id=uid, user_id=uid))

async def remember(api, m, state=None, *, mode='order', text=None):
    uid = m.from_user.id
    state = state or state_for(api, uid)
    try:
        message = m.model_dump(mode='json', exclude_none=True)
    except AttributeError:
        message = None
    record = dict(token=secrets.token_hex(8), expires=time.time()+86400, mode=mode,
                  message=message, state=await state.get_state() if hasattr(state,'get_state') else None, data=await state.get_data())
    if mode == 'chat' and text is not None:
        record['text'] = text
    mini = getattr(m, '_billing_payload', None)
    if mini:
        record['mini'] = mini
    api['get_user'](uid)['billing_resume'] = record
    await api['billing_saved']()
    return record

def resume_keyboard(record, label='Продолжить создание'):
    return InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=label, callback_data='billing_resume:'+record['token'])]])

async def shortage(api, m, state=None, cost=None):
    u = api['get_user'](m.from_user.id)
    required = cost if cost is not None else u.get('_last_shortage', 0)
    balance = api['available_credits'](m.from_user.id)
    d = details(required, balance)
    # A chat gate already saved its request, including the original voice/photo.
    record = u.get('billing_resume') if cost is not None else None
    if not record:
        record = await remember(api, m, state)
    rows = [[InlineKeyboardButton(text=f'Пополнить на {d["topup"]} ₽', callback_data=f'topup_amt:{d["topup"]}')],
            [InlineKeyboardButton(text='Другая сумма', callback_data='topup_custom')],
            [InlineKeyboardButton(text='Назад к запросу', callback_data='billing_resume:'+record['token'])]]
    await api['bot'].send_message(m.from_user.id, notice(required, balance), reply_markup=InlineKeyboardMarkup(inline_keyboard=rows))

async def begin(api, m, kind, *, text=None, state=None):
    uid = m.from_user.id
    u = api['get_user'](uid)
    result, ticket = quota.admit(u, request_key(m, kind), kind,
                                 affordable=lambda cost: api['can_afford'](uid, cost),
                                 unlimited=uid in api['ADMIN_IDS'])
    if result == 'ok':
        return ticket
    if result in ('consent', 'balance'):
        record = await remember(api, m, state, mode='chat', text=text)
        if result == 'balance':
            await shortage(api, m, state, quota.PRICES[kind])
        else:
            keyboard = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text='✅ Да, продолжить за кредиты', callback_data='chat_paid:yes:'+record['token'])],
                [InlineKeyboardButton(text='⏳ Подожду бесплатные обращения', callback_data='chat_paid:no:'+record['token'])]])
            await api['bot'].send_message(uid, quota.NOTICE, reply_markup=keyboard)
    elif result == 'busy':
        await api['bot'].send_message(uid, 'Дождитесь ответа на предыдущий запрос.')
    else:
        # Retry persistence for a delivered response; never run or bill it again.
        await api['billing_saved']()
        await api['bot'].send_message(uid, 'Ответ на это обращение уже отправлен. Повторного списания нет.')
    return None

async def after_payment(api, uid):
    r = api['get_user'](uid).get('billing_resume')
    if r and r['expires'] > time.time():
        await api['bot'].send_message(uid, 'Ваш запрос сохранён. Можно вернуться и продолжить.',
                                      reply_markup=resume_keyboard(r, 'Продолжить запрос'))

async def resume(api, uid, token):
    u = api['get_user'](uid)
    r = u.get('billing_resume')
    if not r or r['token'] != token or r['expires'] <= time.time():
        await api['bot'].send_message(uid, 'Этот запрос уже продолжен или устарел. Отправьте новый запрос.')
        return
    if u.get('busy') or u.get('_chat_access_active') or u.get('_billing_resuming'):
        await api['bot'].send_message(uid, 'Дождитесь завершения текущей работы.')
        return
    u['_billing_resuming'] = True
    try:
        state = state_for(api, uid)
        await state.set_state(r['state'])
        await state.set_data(r['data'])
        u.pop('billing_resume', None)
        await api['billing_saved']()
        if r.get('mini'):
            action, payload = r['mini']
            await api['_handle_miniapp_payload'](api['_FakeMessageForApi'](uid), state, payload, action)
        elif r.get('message'):
            message = Message.model_validate(r['message'], context={'bot': api['bot']})
            if message.from_user.id != uid or message.chat.id != uid:
                raise ValueError('Saved request owner mismatch')
            if r['mode'] == 'chat' and r.get('text') is not None and not message.voice:
                await api['respond_in_chat'](message, r['text'], api['user_lang'](uid))
            else:
                await api['dp'].feed_update(api['bot'], Update(update_id=0, message=message))
        elif r['state'] == 'Form:waiting_confirm':
            await api['_build_presentation'](api['ensure_answerable'](api['_FakeMessageForApi'](uid)), state)
        else:
            await api['bot'].send_message(uid, 'Настройки восстановлены. Продолжите с последнего шага заказа.')
    finally:
        u.pop('_billing_resuming', None)
