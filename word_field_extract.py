"""Shared parsing for fixed Word forms; explicit fields never hide free text."""
import json
import re

VERSION = 'field_input_v1'

# Only resolve a short label when exactly one field in this form accepts it.
# Do not infer dates, amounts, payment facts or party roles from their values.
ALIASES = {
    'date': ('Дата',), 'city': ('Город', 'Место составления'),
    'number': ('Номер',), 'price': ('Стоимость', 'Цена', 'Общая стоимость'),
    'price_words': ('Сумма прописью', 'Стоимость прописью', 'Цена прописью'),
    'services': ('Услуги', 'Описание услуг', 'Перечень услуг'),
    'provider': ('Исполнитель',), 'customer': ('Заказчик',),
    'owner': ('Наймодатель', 'Арендодатель', 'ФИО арендодателя'),
    'tenant': ('Наниматель', 'Арендатор', 'ФИО арендатора'),
    'seller_name': ('Продавец',), 'buyer_name': ('Покупатель',),
    'rent': ('Арендная плата', 'Плата за аренду'),
    'tax': ('НДС',), 'extra': ('Дополнительно',),
    'author': ('Заявитель',), 'employee': ('Работник',),
    'employer': ('Работодатель',), 'plaintiff': ('Истец',),
    'defendant': ('Ответчик',), 'goods': ('Товары',),
    'work': ('Выполненные работы',), 'amount': ('Сумма',),
}


def normalize_label(value):
    return ' '.join(value.casefold().replace('ё', 'е').split())


def label_names(labels):
    exact = {normalize_label(label): key for key, label in labels.items()}
    candidates = {}
    for key, label in labels.items():
        aliases = list(ALIASES.get(key, ()))
        # Planned payment terms must not populate an act's payment-fact field.
        if key == 'payment' and 'сведения об оплате' not in label.casefold():
            aliases += ['Порядок оплаты', 'Условия оплаты']
        for alias in aliases:
            candidates.setdefault(normalize_label(alias), set()).add(key)
    names = {alias: next(iter(keys)) for alias, keys in candidates.items() if len(keys) == 1}
    names.update(exact)
    return names


def split_labeled(source, labels):
    """Return explicit fields and every unrecognised line, without losing text."""
    names = label_names(labels)
    fields, remaining = {}, []
    for line in source.splitlines():
        label, sep, value = line.partition(':')
        key = names.get(normalize_label(label)) if sep else None
        if key:
            fields[key] = value.strip()
        elif line.strip():
            remaining.append(line)
    return fields, '\n'.join(remaining)


def input_review_lines(source, labels, fields):
    """Make unmatched labelled input visible without treating a greeting as failure."""
    names = label_names(labels)
    unmatched = []
    for line in source.splitlines():
        label, sep, value = line.partition(':')
        value = value.strip()
        if sep and value and normalize_label(label) not in names:
            if not any(value in v or v in value for v in fields.values() if v):
                unmatched.append(line.strip())
    if not unmatched:
        return []
    lines = ['', 'Не удалось сопоставить с полями следующие строки. Если они важны, уточни данные перед сборкой:']
    lines += [line[:300] + ('…' if len(line) > 300 else '') for line in unmatched[:10]]
    if len(unmatched) > 10:
        lines.append('Ещё строк без соответствия: ' + str(len(unmatched) - 10))
    return lines


async def extract_form_fields(data, labels, previous, same_base, api, validate,
                              system_prompt, max_tokens):
    base = str(data.get('user_text') or data.get('topic') or '')
    extra = str(data.get('extra') or '')
    source = base + '\n' + extra

    async def extract_piece(piece):
        explicit, remaining = split_labeled(piece, labels)
        fields = {}
        if remaining:
            raw = await api['ask_grok'](
                'Поля: ' + json.dumps(labels, ensure_ascii=False) +
                '\nИсходные данные:\n' + piece,
                max_tokens=max_tokens, temperature=0, system_prompt=system_prompt)
            obj = api['extract_json'](raw)
            fields = validate(obj.get('fields'), piece)
        clean_explicit = validate(explicit, piece)
        if any(value and key not in clean_explicit for key, value in explicit.items()):
            raise ValueError('invalid_explicit_word_field')
        # Blank explicit values are deliberate clears, even if a model fills them.
        for key in explicit:
            fields.pop(key, None)
        fields.update(clean_explicit)
        return fields, explicit

    if same_base:
        fields = validate(previous.get('fields', {}), source)
    else:
        fields, _ = await extract_piece(base)
    if extra.strip():
        updates, explicit = await extract_piece(extra)
        for key in explicit:
            fields.pop(key, None)
        fields.update(updates)
    return validate(fields, source)
