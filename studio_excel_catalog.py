"""Validated, editable estimate input for the private mini-app pilot."""
from decimal import Decimal, InvalidOperation
from excel_styles import PALETTES, PALETTE_LABELS, STYLE_LABELS

MAX_ROWS = 60

def text(value, limit, label):
    if not isinstance(value, str) or len(value) > limit or any(ord(c) < 32 and c not in '\n\t' for c in value):
        raise ValueError(f'Проверьте поле «{label}» (до {limit} символов).')
    return value.strip()

def number(value, label, maximum):
    if isinstance(value, bool) or not isinstance(value, (int, float, str)):
        raise ValueError(f'{label}: введите число от 0 до {maximum}.')
    try:
        n = Decimal(str(value).replace('\u00a0', '').replace(' ', '').replace(',', '.'))
    except InvalidOperation:
        raise ValueError(f'{label}: введите число.')
    if not n.is_finite() or not 0 <= n <= maximum or n.as_tuple().exponent < -3:
        raise ValueError(f'{label}: от 0 до {maximum}, не более трёх знаков после запятой.')
    return float(n)

def normalize(body, draft=False):
    if not isinstance(body, dict) or body.get('kind', 'expense_estimate') != 'expense_estimate':
        raise ValueError('В новом экране пока доступна смета.')
    title = text(body.get('title', ''), 180, 'Название')
    if not draft and not title: raise ValueError('Укажите название сметы.')
    style, palette = body.get('style', 'standard'), body.get('palette', 'green')
    if not isinstance(style, str) or style not in STYLE_LABELS or not isinstance(palette, str) or palette not in PALETTES:
        raise ValueError('Выберите оформление сметы.')
    rows = body.get('rows', [])
    if not isinstance(rows, list) or len(rows) > MAX_ROWS or (not draft and not rows):
        raise ValueError(f'Добавьте от 1 до {MAX_ROWS} позиций.')
    result = []
    for i, row in enumerate(rows, 1):
        if not isinstance(row, dict): raise ValueError('Проверьте строки сметы.')
        item = {k: text(row.get(k, ''), limit, label) for k, limit, label in
                [('name', 250, 'Позиция'), ('unit', 25, 'Единица'), ('category', 80, 'Раздел')]}
        if not draft and not item['name']: raise ValueError(f'Строка {i}: укажите название позиции.')
        for key, label, maximum in [('qty', 'Количество', 1000000), ('price', 'Цена', 1000000000)]:
            value = row.get(key, '')
            item[key] = '' if draft and value == '' else number(value, f'Строка {i}, {label}', maximum)
        result.append(item)
    total = sum(Decimal(str(r['qty'] or 0)) * Decimal(str(r['price'] or 0)) for r in result)
    if total > Decimal('999999999999'): raise ValueError('Общая сумма слишком велика: максимум 999 999 999 999 ₽.')
    return dict(kind='expense_estimate', title=title, subtitle=text(body.get('subtitle', ''), 300, 'Примечание'),
                style=style, palette=palette if style == 'designer' else 'green', rows=result,
                source=text(body.get('source', ''), 6000, 'Исходные данные'))

def catalog(cost):
    return dict(ok=True, cost=cost, max_rows=MAX_ROWS, styles=STYLE_LABELS, palettes=PALETTE_LABELS,
                colors={key: list(value) for key, value in PALETTES.items()})
