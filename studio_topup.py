"""Authenticated top-up intents; only verified payment handlers mark paid."""
import asyncio
import json
import re
import time
from aiohttp import web

def attach(studio, app, api):
    locks = {}

    async def create(request):
        uid = studio.authorized(request)['user']['id']
        studio.check_write(request)
        body = await studio.body(request, 1024)
        amount, key = body.get('amount'), body.get('id')
        if type(amount) is not int or not 60 <= amount <= 9999 or not isinstance(key, str) or not re.fullmatch(r'[a-zA-Z0-9_-]{16,80}', key):
            raise web.HTTPBadRequest(text='Укажите целую сумму от 60 до 9999 ₽.')
        async with locks.setdefault(uid, asyncio.Lock()):
            u = api['get_user'](uid)
            rows = u.setdefault('studio_topups', {})
            now = time.time()
            old = rows.get(key)
            if old and old['amount'] != amount:
                raise web.HTTPConflict(text='Сумма счёта изменилась. Откройте пополнение заново.')
            if old and old.get('url') and old['expires'] > now:
                return web.json_response(dict(ok=True, id=key, url=old['url'], provider=old['provider'], paid=bool(old.get('paid'))), headers={'Cache-Control':'no-store'})
            if old and (old.get('paid') or old['expires'] <= now):
                raise web.HTTPConflict(text='Этот счёт уже оплачен или истёк. Откройте пополнение заново.')
            if any(now-r['created'] < 5 for k,r in rows.items() if k != key):
                raise web.HTTPTooManyRequests(text='Подождите несколько секунд перед новым счётом.')
            if not old:
                # Bounded history of intents, separate from permanent payment ledger.
                rows = {k:r for k,r in rows.items() if r['expires'] > now}
                if len(rows) >= 20:
                    raise web.HTTPTooManyRequests(text='Слишком много открытых счетов. Используйте уже созданный счёт.')
                u['studio_topups'] = rows
                rows[key] = dict(amount=amount, created=now, expires=now+86400, paid=False)
                await api['billing_saved']()
            url = await api['create_yookassa_payment'](uid, amount, amount, request_id=key)
            provider = 'yookassa'
            if not url and api.get('PAYMENT_PROVIDER_TOKEN'):
                # Native Telegram invoice shares the verified successful_payment handler.
                provider = 'telegram'
                receipt = {'receipt': {'items': [dict(description=f'Пополнение на {amount} кредитов', quantity='1.00',
                    amount={'value':f'{amount:.2f}','currency':'RUB'}, vat_code=1, payment_mode='full_payment', payment_subject='service')]}}
                url = await api['bot'].create_invoice_link(title='Пополнение AlixDoc', description=f'{amount} ₽ = {amount} кредитов',
                    payload=f'credits:studio:{key}:{amount}', provider_token=api['PAYMENT_PROVIDER_TOKEN'], currency='RUB',
                    prices=[api['LabeledPrice'](label=f'{amount} кредитов', amount=amount*100)],
                    need_email=True, send_email_to_provider=True, provider_data=json.dumps(receipt))
            if not url:
                raise web.HTTPServiceUnavailable(text='Не удалось открыть оплату. Баланс не изменился. Попробуйте позже.')
            rows[key].update(url=url, provider=provider)
            await api['billing_saved']()
            return web.json_response(dict(ok=True, id=key, url=url, provider=provider, paid=False), headers={'Cache-Control':'no-store'})

    async def status(request):
        uid = studio.authorized(request)['user']['id']
        u = api['get_user'](uid)
        row = u.get('studio_topups', {}).get(request.match_info['intent'])
        if not row:
            raise web.HTTPNotFound()
        # No client invoice status or balance difference is proof of payment.
        return web.json_response(dict(ok=True, paid=bool(row.get('paid')), balance=u.get('credits', 0),
                                     available=api['available_credits'](uid)), headers={'Cache-Control':'no-store'})

    app.router.add_post('/studio/topup', create)
    app.router.add_get('/studio/topup/{intent}', status)
