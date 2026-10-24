"""Persist delivery intent and receipt in the same durable record as the balance.

An uncertain Telegram delivery is never automatically resent or charged. This
avoids duplicate debits after a timeout/restart; support can reconcile pending
deliveries. Admin accounts preserve their existing unlimited entitlement.
"""
from aiohttp import web


class ConstructorBilling:
    def __init__(self, get_user, persist, start, finish, admins):
        self.get_user, self.persist = get_user, persist
        self.start, self.finish, self.admins = start, finish, admins

    async def deliver(self, uid, fingerprint, cost, data, send):
        user = self.get_user(uid)
        orders = user.setdefault('constructor_orders', {})
        old = orders.get(fingerprint)
        if old:
            if old['status'] == 'paid':
                return {'ok': True, 'duplicate': True, 'cost': old['cost'], 'charged': old['charged']}
            raise web.HTTPConflict(text='Статус предыдущей отправки уточняется. Проверьте чат и обратитесь в поддержку; повторного списания нет.')
        ok, reason = self.start(uid)
        if not ok:
            raise web.HTTPConflict(text=reason)
        try:
            if uid not in self.admins and int(user.get('credits') or 0) < cost:
                raise web.HTTPPaymentRequired(text=f'Недостаточно кредитов. Нужно {cost}.')
            # Do not accumulate unresolved attempts or allow spending held funds.
            if any(row.get('status') == 'pending' for row in orders.values()):
                raise web.HTTPConflict(text='Сначала уточните статус предыдущей отправки в поддержке.')
            orders[fingerprint] = {'status': 'pending', 'cost': cost}
            if not await self.persist():
                del orders[fingerprint]
                raise web.HTTPServiceUnavailable(text='Не удалось сохранить заказ. Кредиты не списаны.')
            await send(uid, data)
            charged = 0 if uid in self.admins else cost
            user['credits'] = int(user.get('credits') or 0) - charged
            orders[fingerprint] = {'status': 'paid', 'cost': cost, 'charged': charged}
            user.setdefault('history', []).append(f'Конструктор · {cost} кредитов · PPTX')
            if not await self.persist():
                # Keep the in-process receipt; persisted pending state prevents
                # automatic re-delivery after restart even if this save failed.
                raise web.HTTPServiceUnavailable(text='Файл отправлен. Не удалось подтвердить сохранение оплаты; обратитесь в поддержку, не повторяйте заказ.')
            return {'ok': True, 'cost': cost, 'charged': charged}
        finally:
            self.finish(uid)
