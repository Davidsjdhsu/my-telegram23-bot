"""Consistent, exact credit shortfall for chat and Mini App."""
import math
import json
from aiohttp import web
from wallet import amount

def details(required, balance):
    required, balance = amount(required), amount(balance)
    missing = max(0, amount(required-balance))
    return dict(code='insufficient_credits', required=required, balance=balance,
                missing=missing, topup=max(60, math.ceil(missing)), minimum=60)

def error(required, balance):
    value = details(required, balance)
    return web.HTTPPaymentRequired(text=json.dumps(dict(ok=False, error='Недостаточно кредитов.', **value)), content_type='application/json')

def number(n):
    return str(amount(n)).replace('.', ',')

def notice(required, balance):
    d = details(required, balance)
    text = ('💳 Не хватает кредитов для продолжения.\n\n'
            f'Стоимость — {number(d["required"])} кредитов\n'
            f'Доступно — {number(d["balance"])} кредитов\n'
            f'Не хватает — {number(d["missing"])} кредитов\n\n'
            'Пополните баланс, чтобы продолжить. 1 ₽ = 1 кредит.')
    if d['topup'] > d['missing']:
        text += '\nМинимальное пополнение — 60 ₽. Остаток сохранится на балансе.'
    return text
