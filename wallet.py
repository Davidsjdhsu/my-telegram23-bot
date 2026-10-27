"""Atomic-in-memory credit ledger; the caller must persist the entire user snapshot.

Amounts use half-credit integer units internally. No rounding, clamping or
inferred historic transactions. The operation id is the idempotency key.
"""
from decimal import Decimal, InvalidOperation
from datetime import datetime, timezone
import copy

PERSISTED = ('wallet_transactions','constructor_orders','studio_word_payments',
             'studio_excel_payments','presentation_payments','fixed_word_completed_orders',
             'car_sale_completed_orders','chat_allowance','chat_paid_day','billing_resume','studio_topups','chat_price_notice','delivery_payments','studio_ui_lang','studio_content_lang')

def units(value):
    try: n=Decimal(str(value if value is not None else 0))*2
    except (InvalidOperation,ValueError): raise ValueError('Invalid credit amount')
    if not n.is_finite() or n != n.to_integral_value(): raise ValueError('Invalid credit amount')
    return int(n)

def amount(value):
    n=units(value)
    return n//2 if n%2==0 else n/2

def apply(user, key, delta, title, kind='debit'):
    ledger=user.setdefault('wallet_transactions',{})
    if key in ledger:
        old=ledger[key]
        if units(old['delta']) != units(delta): raise ValueError('Transaction amount changed')
        return old
    before=units(user.get('credits',0)); change=units(delta)
    if before+change<0: raise ValueError('Недостаточно кредитов.')
    row=dict(id=key,ts=datetime.now(timezone.utc).isoformat(timespec='seconds'),
             delta=amount(delta),balance_after=amount((before+change)/2),title=str(title)[:180],kind=kind)
    # No await between mutation and receipt: persisted snapshots contain both.
    user['credits']=row['balance_after'];ledger[key]=row
    return row

def snapshot(user):
    return {k:copy.deepcopy(user[k]) for k in PERSISTED if k in user}

def public_history(user, legacy=(), offset=0, limit=50):
    rows=list(user.get('wallet_transactions',{}).values())
    keys={r['id'] for r in rows}
    rows += [r for r in legacy if r['id'] not in keys]
    rows=[r for _,r in sorted(enumerate(rows),key=lambda pair:(pair[1].get('ts',''),pair[0]),reverse=True)]
    # Provider IDs, order IDs and personal input never leave the server.
    public=[{k:r.get(k) for k in ('ts','delta','balance_after','title','kind')} for r in rows[offset:offset+limit]]
    return {'items':public,'total':len(rows),'next_offset':offset+limit if len(rows)>offset+limit else None}

def request_key(message, action):
    import uuid
    uid=getattr(getattr(message,'from_user',None),'id',0)
    chat=getattr(getattr(message,'chat',None),'id',uid)
    mid=getattr(message,'message_id',None)
    return action+':'+str(chat)+':'+str(mid if mid is not None else uuid.uuid4().hex)


class BillingPersistenceError(RuntimeError):
    pass


def persistence_notice():
    return 'Результат отправлен. Не удалось подтвердить сохранение оплаты. Не повторяйте заказ: обратитесь в @AlixDocSupport для сверки баланса и истории.'
