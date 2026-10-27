"""Shared daily allowance. Admission and settlement contain no awaits.

One active interaction per account prevents concurrent voice/photo/text requests
from consuming the same last free turn. Only delivered replies are settled.
"""
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from wallet import apply, request_key

LIMIT = 10
PRICES = {'chat': .5, 'voice': 3, 'vision': 3}
LABELS = {'chat': 'Ответ в чате', 'voice': 'Голосовое сообщение', 'vision': 'Разбор фотографии'}

def day(now=None):
    return (now or datetime.now(timezone.utc)).astimezone(ZoneInfo('Europe/Moscow')).date().isoformat()

def status(user, now=None):
    today = day(now)
    record = user.get('chat_allowance', {})
    used = len(record.get('receipts', {})) if record.get('day') == today else 0
    return dict(day=today, limit=LIMIT, remaining=max(0, LIMIT-used),
                paid_enabled=user.get('chat_paid_day') == today)

def consent(user, enabled, now=None):
    user['chat_paid_day'] = day(now) if enabled else None

def admit(user, key, kind, *, affordable, unlimited=False, now=None):
    current = status(user, now)
    receipt = user.get('chat_allowance', {}).get('receipts', {}).get(key)
    if receipt or key in user.get('wallet_transactions', {}):
        return 'duplicate', None
    active = user.get('_chat_access_active')
    if active:
        return ('ok', active) if active['key'] == key else ('busy', None)
    free = current['remaining'] > 0 or unlimited
    if not free and not current['paid_enabled']:
        return 'consent', None
    cost = 0 if free else PRICES[kind]
    if cost and not affordable(cost):
        return 'balance', None
    ticket = dict(key=key, kind=kind, day=current['day'], cost=cost, unlimited=unlimited)
    user['_chat_access_active'] = ticket
    user['_chat_reserved'] = cost
    return 'ok', ticket

def settle(user, ticket):
    if ticket['unlimited']:
        return
    key, cost = ticket['key'], ticket['cost']
    if cost:
        apply(user, key, -cost, LABELS[ticket['kind']])
    else:
        record = user.get('chat_allowance', {})
        if record.get('day') != ticket['day']:
            record = user['chat_allowance'] = {'day': ticket['day'], 'receipts': {}}
        record['receipts'].setdefault(key, ticket['kind'])

def release(user, ticket):
    if user.get('_chat_access_active') == ticket:
        user.pop('_chat_access_active', None)
        user.pop('_chat_reserved', None)

NOTICE = ('💬 Вы использовали 10 бесплатных обращений на сегодня.\n\n'
          'Продолжить общение за кредиты?\n\n'
          '✍️ Текстовый ответ — 0,5 кредита\n'
          '🎙 Ответ на голосовое сообщение — 3 кредита\n'
          '📷 Разбор фотографии — 3 кредита\n\n'
          'Списание только за успешно полученные ответы после вашего согласия. '
          'Отключить платные ответы: /free.\n\n'
          '🕛 В 00:00 по Москве бесплатный лимит обновится, а разрешение на платные ответы завершится.')
