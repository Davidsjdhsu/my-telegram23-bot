"""One durable balance mutation per order; pending intent is safe to resume."""
from aiohttp import web
import asyncio

_locks = {}


def charge_callback(api, ledger_name, label):
    async def apply(uid, order_id, cost, title):
        user = api['get_user'](uid)
        ledger = user.setdefault(ledger_name, {})
        old = ledger.get(order_id)
        if old and old['status'] == 'paid':
            # Retry persistence too: previous save may have failed in this process.
            if not await api['save_users_now']():
                raise RuntimeError('Unable to persist receipt')
            return old['charged']
        if old and old.get('version') != 1:
            raise RuntimeError('Legacy payment requires reconciliation')
        charged = 0 if uid in api['ADMIN_IDS'] else cost
        if old and old.get('cost', cost) != cost:
            raise ValueError('Стоимость заказа не совпала.')
        if not old and not api['can_afford'](uid, cost):
            raise web.HTTPPaymentRequired(text='Недостаточно кредитов.')
        if int(user.get('credits', 0)) < charged:
            raise web.HTTPPaymentRequired(text='Недостаточно кредитов.')
        ledger[order_id] = {'status': 'pending', 'charged': charged, 'cost': cost, 'version': 1}
        if not await api['save_users_now']():
            if not old:
                ledger.pop(order_id, None)
            raise RuntimeError('Unable to persist payment intent')
        # No await between balance, receipt and history: every saved snapshot
        # contains either the old balance+pending or new balance+paid, never half.
        user['credits'] = int(user.get('credits', 0)) - charged
        ledger[order_id] = {'status': 'paid', 'charged': charged, 'cost': cost, 'version': 1}
        user['generations'] = user.get('generations', 0) + 1
        user.setdefault('history', []).append(label + ' · ' + title[:180])
        if not await api['save_users_now']():
            raise RuntimeError('Unable to persist payment receipt')
        api['note_success'](uid)
        return charged

    async def charge(uid, order_id, cost, title):
        key=(id(api['get_user']),uid)
        async with _locks.setdefault(key,asyncio.Lock()):
            return await apply(uid,order_id,cost,title)
    return charge
