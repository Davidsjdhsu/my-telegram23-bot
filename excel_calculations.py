"""A small allowlisted calculation grammar shared by estimate validation/export.

No user formula strings, eval, external references or arbitrary Excel functions.
"""
from decimal import Decimal, InvalidOperation

DEFAULT = {'mode': 'cost'}
LABELS = {'qty': 'Количество', 'price': 'Цена', 'number': 'Число'}
SYMBOLS = {'multiply': '×', 'divide': '÷', 'add': '+', 'subtract': '−'}

def scalar(value, maximum=1000000000):
    if isinstance(value, bool) or not isinstance(value, (str, int, float)):
        raise ValueError('В расчёте нужно указать число.')
    try: n = Decimal(str(value).replace(' ', '').replace('\u00a0', '').replace(',', '.'))
    except InvalidOperation: raise ValueError('В расчёте нужно указать число.')
    if not n.is_finite() or not 0 <= n <= maximum or n.as_tuple().exponent < -3:
        raise ValueError(f'Число в расчёте: от 0 до {maximum}, до трёх знаков после запятой.')
    return float(n)

def normalize_calculation(value=None):
    if value is None: return dict(DEFAULT)
    if not isinstance(value, dict): raise ValueError('Проверьте настройки расчёта.')
    mode = value.get('mode')
    if mode == 'cost': return dict(DEFAULT)
    if mode in ('discount', 'markup'):
        return {'mode': mode, 'percent': scalar(value.get('percent'), 100 if mode == 'discount' else 1000)}
    if mode != 'custom': raise ValueError('Выберите вариант расчёта.')
    left, right, op = value.get('left'), value.get('right'), value.get('op')
    if not all(isinstance(x, str) for x in (left, right, op)) or left not in LABELS or right not in LABELS or op not in SYMBOLS:
        raise ValueError('Выберите столбцы и действие для расчёта.')
    result = {'mode': mode, 'left': left, 'right': right, 'op': op}
    for side in ('left', 'right'):
        if result[side] == 'number': result[side + '_value'] = scalar(value.get(side + '_value'))
    if op == 'divide' and right == 'number' and result['right_value'] == 0:
        raise ValueError('Делить на ноль нельзя. Укажите другой делитель.')
    return result

def calculate(row, spec=None):
    c = normalize_calculation(spec)
    def dec(key):
        v = row[key]
        if v == '': return None
        return Decimal(str(v))
    if c['mode'] != 'custom':
        q, p = dec('qty'), dec('price')
        if q is None or p is None: return None
        result = q * p
        if c['mode'] != 'cost': result *= 1 + Decimal(str(c['percent'])) / 100 * (-1 if c['mode'] == 'discount' else 1)
    else:
        a, b = [Decimal(str(c[s + '_value'])) if c[s] == 'number' else dec(c[s]) for s in ('left', 'right')]
        if a is None or b is None: return None
        if c['op'] == 'divide' and b == 0: raise ValueError('Деление на ноль: измените делитель или расчёт.')
        if c['op'] == 'multiply': result = a * b
        elif c['op'] == 'divide': result = a / b
        elif c['op'] == 'add': result = a + b
        else: result = a - b
    if abs(result) > Decimal('999999999999'):
        raise ValueError('Результат расчёта слишком велик: максимум 999 999 999 999 по модулю.')
    return float(result)

def description(spec=None):
    c = normalize_calculation(spec)
    if c['mode'] == 'cost': return 'Количество × Цена'
    if c['mode'] in ('discount', 'markup'):
        sign = '−' if c['mode'] == 'discount' else '+'
        return f'Количество × Цена × (1 {sign} {c["percent"]:g} / 100)'
    def label(side): return f'{c[side + "_value"]:g}' if c[side] == 'number' else LABELS[c[side]]
    return f'{label("left")} {SYMBOLS[c["op"]]} {label("right")}'

def formula(row, spec=None):
    c = normalize_calculation(spec)
    if c['mode'] == 'cost': return f'=D{row}*E{row}'
    if c['mode'] in ('discount', 'markup'):
        sign = '-' if c['mode'] == 'discount' else '+'
        return f'=D{row}*E{row}*(1{sign}\'Расчёт\'!$B$3/100)'
    refs = {'qty': f'D{row}', 'price': f'E{row}'}
    a, b = [f"'Расчёт'!$B${4 if s == 'left' else 5}" if c[s] == 'number' else refs[c[s]] for s in ('left', 'right')]
    if c['op'] == 'divide': return f'=IF({b}=0,"",{a}/{b})'
    return '=' + a + {'multiply': '*', 'add': '+', 'subtract': '-'}[c['op']] + b
