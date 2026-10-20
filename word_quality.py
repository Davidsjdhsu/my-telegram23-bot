"""Word export and lossless text expansion, shared by bot entry points."""
from collections import Counter
import asyncio
import copy
import os
import re
import tempfile
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pypdf import PdfReader
from presentation_pdf import _convert_office_pdf, PresentationPDFError


INFORMATIONAL_WORD_KINDS = frozenset({'doc', 'notes', 'report', 'essay', 'referat', 'coursework'})
WORD_REFINE = 'Доработать объём'
WORD_ACCEPT = 'Получить текущий вариант'


def word_quality_keyboard(button, markup, main_label):
    return markup(keyboard=[[button(text=WORD_REFINE)], [button(text=WORD_ACCEPT)],
                            [button(text=main_label)]], resize_keyboard=True)


def word_quality_choice_notice(requested, actual, credits, sparse):
    return (f'После доработки получилось {actual} стр. текста вместо выбранных {requested}. '
            if actual != requested else f'Получилось {actual} стр. текста. ') + (
            'Есть слабо заполненные страницы. ' if sparse else '') + (
            f'Можно доработать сохранённый текст или получить текущий вариант за {credits} кредитов. '
            'Титульный и оглавление бесплатно. Пока файлы не выданы, кредиты не списаны. '
            'Повторная доработка не оплачивается отдельно.')
WORD_ORIGINALITY_GUIDANCE = (
    'Строй самостоятельное объяснение вокруг задачи пользователя: тезис, обоснование, '
    'конкретный пример, границы применимости. Не копируй готовые работы и не пересказывай '
    'один источник заменой слов на синонимы. Заимствованные идеи атрибутируй, дословные цитаты '
    'оформляй кавычками и ссылкой на реально предоставленный или проверенный источник. '
    'Если источника нет, не выдумывай ссылку или цитату. Не приписывай автору личный опыт. '
    'Не добавляй ошибки, невидимые символы, подмену букв или искусственные обороты. '
    'Не обещай процент оригинальности или результат детектора ИИ без фактической проверки. '
)
WORD_EDITORIAL_GUIDANCE = (
    'Пиши ясно для указанной аудитории. Если жанру нужно введение, сразу обозначь вопрос; '
    'не начинай с «Данный документ предназначен» и описания границ самого документа. '
    'основное сравнение, объяснения и примеры вынеси в основную часть. '
    'Вывод отвечает на поставленный вопрос и помогает сделать выбор, а не повторяет введение. '
    'Не пересказывай таблицу по строкам: поясни существенные различия и условия применения. '
    'Каждый абзац добавляет новую мысль; отделяй абзацы символами \\n\\n внутри content. '
    'Обычно достаточно 40–90 слов на абзац; не собирай сравнение всех вариантов в один абзац. '
    'Удали общие фразы и смысловые повторы, даже если одна мысль выражена разными словами. '
    'Раскрывай каждый аргумент в одном месте; следующий абзац должен добавлять причину, '
    'следствие, условие, возражение или разобранный пример, а не пересказывать предыдущий. '
    'Сохраняй оговорки и исключения: не превращай «обычно» или «при условии» в «всегда» или «невозможно». '
    'Не выдумывай статистику, исследования, источники, цитаты, личный опыт или реквизиты. '
    'Не заменяй отсутствующий источник словами «по разным оценкам». '
    'Условные примеры явно обозначай как условные. Данные и таблицы пользователя сохраняй точно. '
    'Отдельное требование пользователя об учебном характере документа явно укажи в тексте. '
    'Перед возвратом JSON проверь причинные связи и категоричные утверждения о доступности, '
    'безопасности и невозможности действия. Различай свойства объекта, настройки и условия работы. '
    'Указывай конкретное условие или исключение рядом с утверждением; не добавляй оговорки механически '
    'к достоверным определениям. Объяснения и примеры должны относиться только к теме пользователя. '
    + WORD_ORIGINALITY_GUIDANCE
)

# Word must not inherit the presentation art-director persona or its Russian-only
# instruction. Keep the document's requested depth and language above style tricks.
WORD_SYSTEM = (
    'Ты редактор полноразмерных Word-документов. Выполняй задачу в указанном жанре, '
    'на запрошенном языке и в запрошенном формате ответа. Документ должен раскрывать тему, '
    'сохраняя жанр: конспект содержит тематические тезисы и пояснения, эссе — связное рассуждение. '
    'Если задан объём в словах, распределяй его по разделам и соблюдай ориентир с отклонением '
    'не более 10%, если есть достаточно материала. Приоритеты: точность смысла, сохранность '
    'пользовательских данных, полнота объяснения, затем стиль. Не сокращай содержание ради '
    'лаконичности. Таблицы не заменяют объяснений условий и практического примера. '
    'Перед ответом перечитай весь текст и устрани противоречия и необоснованные обобщения. '
    + WORD_EDITORIAL_GUIDANCE
)

WORD_REVIEW_SYSTEM = (
    'Ты независимый редактор учебных и информационных документов. Проверяй черновик, '
    'а не соглашайся с автором. За один проход исправь фактические обобщения, причинные ошибки, '
    'противоречия и недостаточную глубину объяснений. Признаки из диагностики — только подсказки: '
    'проверь также остальные утверждения. Не считай сгенерированный черновик достоверным источником. '
    'Сохраняй данные и таблицы, которые явно дал пользователь. Если таблица упрощает действительность, '
    'оставь ячейки без изменений, а условия и ограничения объясни в основном тексте. '
    'В остальных фразах черновика исправляй ошибки по смыслу, а не косметически. '
    'Не добавляй выдуманные источники или результаты исследований. Если факт нельзя уверенно '
    'подтвердить, убери неподтверждённую конкретику или обозначь пределы применимости. '
    'Не добавляй в документ отчёт о своей проверке. Верни только JSON с sections (title, content), '
    'с теми же заголовками и порядком. Соблюдай заданные ориентиры слов по разделам. '
    + WORD_EDITORIAL_GUIDANCE
)


def word_section_budgets(content, target, *, allow_shorten=False):
    """Give the optional editor concrete prose budgets rather than one vague total."""
    sections = content.get('sections', [])
    short = re.compile(r'введени|вывод|заключени|introduction|conclusion|summary', re.I)
    nonbody = re.compile(r'содержани|литератур|источник|библиограф|references|bibliograph|contents', re.I)
    weights = [0 if nonbody.search(s.get('title', '')) else .3 if short.search(s.get('title', '')) else 1
               for s in sections]
    # Never ask the editor to shorten a section below the preservation gate.
    # Keep existing prose budgets, then redistribute the remaining total among
    # sections that can grow. References are retained but are not padded.
    floors = [len((s.get('content') or '').split()) for s in sections]
    if allow_shorten:
        import math
        floors = [math.ceil(count * .8) if weights[i] else count for i, count in enumerate(floors)]
    budgets = list(floors)
    remaining = max(int(target), sum(floors))
    active = {i for i, weight in enumerate(weights) if weight > 0}
    remaining -= sum(floors[i] for i in range(len(sections)) if i not in active)
    while active:
        total = sum(weights[i] for i in active)
        locked = {i for i in active if remaining * weights[i] / total < floors[i]}
        if locked:
            remaining -= sum(floors[i] for i in locked)
            active -= locked
            continue
        ordered = sorted(active)
        for i in ordered:
            budgets[i] = int(remaining * weights[i] / total)
        for i in ordered[:remaining - sum(budgets[i] for i in ordered)]:
            budgets[i] += 1
        break
    return [{'title': s.get('title', ''), 'target_words': budget}
            for s, budget in zip(sections, budgets)]


def clean_word_meta(meta, kind):
    """Omit unspecified optional details in generated informational documents only."""
    result = dict(meta or {})
    if kind not in INFORMATIONAL_WORD_KINDS:
        return result
    for key, value in result.items():
        if not isinstance(value, str):
            continue
        value = value.strip()
        # Match known placeholder labels, not every string in square brackets.
        if re.fullmatch(r'\[(?:указать\s+)?(?:учебное заведение|название учебного заведения|'
                        r'наименование организации|организация|организация/учебное заведение|'
                        r'предмет/дисциплина(?:/МДК)?|предмет|дисциплина|номер группы|группа|'
                        r'ФИО(?: студента| преподавателя| научного руководителя)?|'
                        r'должность и/или ФИО преподавателя|город|год|специальность|'
                        r'school|institution|organization|author|name|city|year)\]', value, re.I):
            result[key] = ''
    return result


def word_generation_schema(kind, schema):
    """Template placeholders must not teach the model to invent metadata."""
    if kind not in INFORMATIONAL_WORD_KINDS:
        return schema
    import json
    return json.dumps({key: '' for key in json.loads(schema)}, ensure_ascii=False)


def word_pdf_page_count(pdf_path):
    return len(PdfReader(pdf_path).pages)


def word_pdf_layout(pdf_path, *, front_matter_pages=0):
    """Measure text extent on every body page, excluding covers and footers.

    A cheap signal for text documents, not a visual or semantic quality score.
    The builder uses 2 cm top/bottom margins. Tables count toward the extent.
    """
    fills = []
    for page in PdfReader(pdf_path).pages:
        height = float(page.mediabox.height)
        top, bottom = height - 56.7, 56.7
        baselines = []
        def visit(text, cm, tm, font, size):
            y = tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
            if text.strip() and bottom <= y <= top:
                baselines.append(y - min(float(size), 18) * .25)
        page.extract_text(visitor_text=visit)
        fill = min(1., max(0., (top - min(baselines)) / (top - bottom))) if baselines else 0.
        fills.append(round(fill, 3))
    body = fills[front_matter_pages:]
    return {'last_page_fill': body[-1] if body else None,
            'body_page_fills': body, 'front_matter_pages': front_matter_pages}


def readable_word_paragraphs(text):
    """Split long plain prose at sentence boundaries, changing whitespace only.

    Leave lists, markdown tables, code and short abbreviations/initials alone.
    The model still handles semantic paragraph grouping and long single sentences.
    """
    result = []
    for para in text.split('\n'):
        if len(para.split()) <= 100 or re.match(r'\s*(?:[-*•]|\d+[.)])\s', para) or '|' in para or '`' in para:
            result.append(para)
            continue
        cuts = []
        for match in re.finditer(r'(?<=[.!?])\s+(?=[А-ЯЁA-Z])', para):
            previous = para[:match.start()].split()[-1].strip('«"(').rstrip('.!?')
            if len(previous) > 3 and previous.lower() not in {'prof', 'сокр', 'тов'}:
                cuts.append((match.start(), match.end()))
        start = 0
        for left, right in cuts:
            words = len(para[start:left].split())
            if words >= 55:
                result.append(para[start:left])
                start = right
        result.append(para[start:])
    return '\n'.join(result)


def word_review_findings(content, actual, requested, layout=None):
    paragraphs = [p for s in content.get('sections', []) for p in (s.get('content') or '').split('\n') if p.strip()]
    # These are review candidates, never a verdict that the statement is false.
    absolute = re.compile(r'\b(?:всегда|никогда|невозмож\w*|недоступ\w*|гарантиру\w*|always|never|impossible|guaranteed)\b', re.I)
    conditional = re.compile(r'\b(?:если|обычно|может|могут|зависит|при условии|if|unless|may|depends)\b', re.I)
    claims = sum(bool(absolute.search(sentence)) and not conditional.search(sentence)
                 for p in paragraphs for sentence in re.split(r'(?<=[.!?])\s+', p))
    fill = (layout or {}).get('last_page_fill')
    sentences = [' '.join(s.casefold().split()) for p in paragraphs
                 for s in re.split(r'(?<=[.!?])\s+', p) if len(s.split()) >= 8 and '«' not in s]
    repeats = sum(n - 1 for n in Counter(sentences).values() if n > 1)
    return {
        'short_pages': actual < requested,
        'sparse_last_page': requested > 1 and fill is not None and fill < .5,
        'sparse_body_pages': [i + 1 + (layout or {}).get('front_matter_pages', 0)
                              for i, value in enumerate((layout or {}).get('body_page_fills', []))
                              if value < .5],
        'excess_pages': actual > requested,
        'long_paragraphs': sum(len(p.split()) > 100 and not p.lstrip().startswith('|') for p in paragraphs),
        'categorical_claims': claims,
        'duplicate_sentences': repeats,
        'last_page_fill': fill,
        'body_page_fills': (layout or {}).get('body_page_fills', []),
    }


def needs_word_revision(findings):
    return any(findings[k] for k in ('short_pages', 'sparse_last_page', 'sparse_body_pages',
                                     'excess_pages', 'long_paragraphs', 'categorical_claims')) or bool(
                                         findings.get('duplicate_sentences'))


def word_volume_needs_choice(actual, requested, layout=None):
    """Do not silently treat an unresolved academic volume mismatch as complete."""
    fills = (layout or {}).get('body_page_fills', [])
    return actual != requested or any(fill < .5 for fill in fills)


def focused_word_sections(content, target, *, limit=3, allow_shorten=False):
    """Spend the last bounded model call on the largest substantive deficits."""
    budgets = word_section_budgets(content, target, allow_shorten=allow_shorten)
    sections = content.get('sections', [])
    eligible = [i for i, s in enumerate(sections) if not re.search(
        r'введени|вывод|заключени|содержани|литератур|источник|библиограф|introduction|conclusion|references',
        s.get('title', ''), re.I)]
    if not eligible:
        return []
    delta = lambda i: budgets[i]['target_words'] - len((sections[i].get('content') or '').split())
    chosen = sorted(eligible, key=lambda i: -abs(delta(i)))[:limit]
    # Allocate the requested net change to selected sections only. The other
    # sections remain byte-for-byte unchanged in the patch merge.
    current = sum(len((s.get('content') or '').split()) for s in sections)
    change = target - current
    each = int(change / len(chosen))
    return [{'index': i, 'title': sections[i].get('title', ''),
             'target_words': max(int(len((sections[i].get('content') or '').split()) * .8 + .999),
                                 len((sections[i].get('content') or '').split()) + each +
                                 (change - each * len(chosen) if pos == 0 else 0))}
            for pos, i in enumerate(sorted(chosen))]


def apply_word_section_patch(content, patch, targets):
    """Keep unselected prose, user metadata and all tables outside model output."""
    if not isinstance(patch, dict) or not isinstance(patch.get('sections'), list):
        return content
    rows = patch['sections']
    if len(rows) != len(targets):
        return content
    result = copy.deepcopy(content)
    for row, spec in zip(rows, targets):
        if (not isinstance(row, dict) or row.get('index') != spec['index'] or
                row.get('title') != spec['title'] or not isinstance(row.get('content'), str)):
            return content
        result['sections'][spec['index']]['content'] = row['content']
    return merge_word_expansion(content, result)


def word_growth_targets(content, target, *, limit=3):
    """Budget only new prose. Existing paragraphs do not consume response tokens."""
    current = sum(len((s.get('content') or '').split()) for s in content.get('sections', []))
    if target <= current:
        return []
    targets = focused_word_sections(content, target, limit=limit)
    return [{'index': t['index'], 'title': t['title'],
             'add_words': t['target_words'] - len((content['sections'][t['index']].get('content') or '').split()),
             'paragraphs': [p for p in (content['sections'][t['index']].get('content') or '').splitlines() if p.strip()]}
            for t in targets if t['target_words'] > len((content['sections'][t['index']].get('content') or '').split())]


def apply_word_additions(content, patch, targets):
    """Insert prose at explicit anchors; never let growth rewrite existing data."""
    if not isinstance(patch, dict) or not isinstance(patch.get('additions'), list):
        return content
    specs = {t['index']: t for t in targets}
    additions = {}
    seen = []
    def sentences(text):
        return [set(re.findall(r'\w+', s.casefold())) for s in re.split(r'(?<=[.!?])\s+', text)
                if len(s.split()) >= 8]
    for section in content.get('sections', []):
        seen.extend(sentences(section.get('content') or ''))
    added_words = 0
    for row in patch['additions']:
        if not isinstance(row, dict):
            return content
        index, anchor, paragraphs = row.get('section_index'), row.get('after_paragraph'), row.get('paragraphs')
        if (type(index) is not int or index not in specs or type(anchor) is not int
                or not 0 <= anchor < len(specs[index]['paragraphs'])
                or not isinstance(paragraphs, list) or not paragraphs):
            return content
        for para in paragraphs:
            if not isinstance(para, str) or not para.strip() or '\n' in para or '|' in para:
                return content
            for words in sentences(para):
                # A conservative lexical guard, not a semantic/plagiarism verdict.
                if any(len(words & old) / max(1, len(words | old)) >= .85 for old in seen):
                    print('Word growth_rejected=repeated_sentence', flush=True)
                    return content
                seen.append(words)
            added_words += len(para.split())
        additions.setdefault((index, anchor), []).extend(p.strip() for p in paragraphs)
    budget = sum(t['add_words'] for t in targets)
    if not added_words or added_words > max(budget * 1.35, budget + 60):
        return content
    result = copy.deepcopy(content)
    for index, spec in specs.items():
        # Preserve every original line, including empty ones, and insert after
        # the numbered nonempty paragraph. Tables and metadata remain untouched.
        lines, paragraph_index = [], -1
        for line in (content['sections'][index].get('content') or '').splitlines():
            lines.append(line)
            if line.strip():
                paragraph_index += 1
                lines.extend(additions.get((index, paragraph_index), []))
        result['sections'][index]['content'] = '\n'.join(lines)
    print(f'Word growth requested_words={budget} added_words={added_words}', flush=True)
    return result


def word_revision_notice(lang):
    return {
        'ru': 'Уточняю текст и проверяю заполненность страниц.',
        'en': 'Refining the text and checking page coverage.',
        'de': 'Ich überarbeite den Text und prüfe den Seitenumfang.',
        'es': 'Estoy revisando el texto y el contenido de las páginas.',
        'fr': 'Je révise le texte et vérifie le contenu des pages.',
        'zh': '正在完善文字并检查页面内容。',
        'ar': 'أراجع النص وأتحقق من امتلاء الصفحات.',
    }.get(lang, 'Refining the text and checking page coverage.')


def expansion_word_target(content, actual_pages, requested_pages, *, front_matter_pages=0, layout=None):
    words = sum(len((s.get('content') or '').split()) for s in content.get('sections', []))
    actual_body = max(1, actual_pages - front_matter_pages)
    requested_body = max(1, requested_pages - front_matter_pages)
    fills = (layout or {}).get('body_page_fills', [])
    if words and fills and sum(fills) > .1:
        # Use the observed typography/table footprint, aiming for a 75%-full ending.
        # Compression remains limited by the lossless metadata/table and prose gates.
        target = round(words * max(.75, requested_body - .25) / sum(fills))
        floor = (words * .8 + .999) if actual_pages > requested_pages else words
        return min(16000, max(int(floor), target))
    return min(16000, max(requested_body * 280, round(words * requested_body / actual_body * 1.1)))


async def prepare_word_files(content, requested_pages, docx_path, pdf_path, *, build,
                             convert=None, expand=None, measure=word_pdf_page_count, inspect=None,
                             max_revisions=1):
    """Measure actual pages; bound revisions without risking the valid original pair.

    All expensive work is bounded. Editorial/expansion failures are non-blocking;
    initial build or integrity failures still propagate before delivery and billing.
    """
    convert = convert or convert_word_pdf
    await build(content, docx_path)
    await convert(docx_path, pdf_path)
    actual = await asyncio.to_thread(measure, pdf_path)
    async def layout_at(path):
        if inspect is None:
            return None
        try:
            return await asyncio.to_thread(inspect, path)
        except Exception as exc:
            print('Word optional_layout_failure=' + type(exc).__name__, flush=True)
            return None
    layout = await layout_at(pdf_path)
    findings = word_review_findings(content, actual, requested_pages, layout)
    attempts = 0
    accepted = False
    while needs_word_revision(findings) and expand is not None and attempts < min(2, max_revisions):
        attempts += 1
        promoting = False
        try:
            expanded = await expand(copy.deepcopy(content), actual, requested_pages,
                                    dict(findings, revision_attempt=attempts))
            candidate = merge_word_expansion(content, expanded)
            if candidate != content:
                with tempfile.TemporaryDirectory(prefix='word-revision-', dir=Path(docx_path).parent) as td:
                    revised_docx, revised_pdf = Path(td) / 'revised.docx', Path(td) / 'revised.pdf'
                    await build(candidate, str(revised_docx))
                    await convert(str(revised_docx), str(revised_pdf))
                    revised_pages = await asyncio.to_thread(measure, revised_pdf)
                    revised_layout = await layout_at(revised_pdf)
                    revised_findings = word_review_findings(candidate, revised_pages, requested_pages, revised_layout)
                    print(f'Word revision_candidate pages={revised_pages} findings={revised_findings}', flush=True)
                    # Same page count can improve substance or readability. Do not trade
                    # those for a new page-count error, emptier ending or lost prose.
                    closer = abs(revised_pages - requested_pages) < abs(actual - requested_pages)
                    same_pages = revised_pages == actual
                    old_fill, new_fill = findings['last_page_fill'], revised_findings['last_page_fill']
                    fill_ok = old_fill is None or (new_fill is not None and new_fill >= old_fill - .03)
                    improved = (
                        (old_fill is not None and new_fill is not None and new_fill > old_fill + .05)
                        or revised_findings['long_paragraphs'] < findings['long_paragraphs']
                        or revised_findings['categorical_claims'] < findings['categorical_claims']
                        or revised_findings['duplicate_sentences'] < findings['duplicate_sentences']
                        or len(revised_findings['sparse_body_pages']) < len(findings['sparse_body_pages'])
                    )
                    if closer or (same_pages and fill_ok and improved):
                        promoting = True
                        os.replace(revised_docx, docx_path)
                        os.replace(revised_pdf, pdf_path)
                        content, actual, findings = candidate, revised_pages, revised_findings
                        accepted = True
                    else:
                        print('Word revision_rejected=layout_not_improved', flush=True)
        except Exception as exc:
            if promoting:
                raise  # Never deliver a possibly mismatched DOCX/PDF pair.
            print('Word optional_revision_failure=' + type(exc).__name__, flush=True)
            break  # Provider/conversion failures are not an excuse for more paid calls.
        if not word_volume_needs_choice(actual, requested_pages, findings):
            break  # The second call is reserved for unresolved physical page volume.
    print(f'Word pagination requested={requested_pages} actual={actual} revisions={attempts} accepted={accepted}', flush=True)
    print('Word quality ' + ' '.join(f'{k}={v}' for k, v in findings.items()), flush=True)
    return content, actual


def word_page_notice(lang, requested, actual, credits):
    messages = {
        'ru': 'В готовом документе {actual} стр. при выбранных {requested}. Стоимость — {credits} кредитов; дополнительные страницы не увеличивают цену.',
        'en': 'The document has {actual} pages; {requested} were requested. Cost: {credits} credits. Extra pages do not increase the price.',
        'de': 'Das Dokument hat {actual} Seiten; ausgewählt waren {requested}. Kosten: {credits} Credits. Zusätzliche Seiten erhöhen den Preis nicht.',
        'es': 'El documento tiene {actual} páginas; se solicitaron {requested}. Coste: {credits} créditos. Las páginas adicionales no aumentan el precio.',
        'fr': 'Le document contient {actual} pages, pour {requested} demandées. Coût : {credits} crédits. Les pages supplémentaires ne majorent pas le prix.',
        'zh': '文档共 {actual} 页，所选为 {requested} 页。费用为 {credits} 积分；额外页面不加价。',
        'ar': 'يحتوي المستند على {actual} صفحة، والمطلوب {requested}. التكلفة {credits} رصيدًا؛ الصفحات الإضافية لا تزيد السعر.',
    }
    return messages.get(lang, messages['en']).format(actual=actual, requested=requested, credits=credits)


def merge_word_expansion(original, expanded):
    """Only accept expanded paragraph text; never replace tables or user metadata."""
    def reject(reason):
        print('Word revision_rejected=' + reason, flush=True)
        return original
    if not isinstance(expanded, dict):
        return reject('invalid_object')
    old, new = original.get('sections', []), expanded.get('sections', [])
    if not isinstance(new, list):
        return reject('invalid_sections')
    if len(old) != len(new) or not all(isinstance(b, dict) for b in new):
        return reject('section_count')
    if any(a.get('title') != b.get('title') or not isinstance(b.get('content'), str)
           for a, b in zip(old, new)):
        return reject('section_heading_or_content')
    for a, b in zip(old, new):
        before, after = a.get('content') or '', b['content']
        # A revision may polish prose, but must not erase a section or supplied numbers.
        if before.strip() and not after.strip():
            return reject('empty_section')
        if len(after.split()) < len(before.split()) * .8:
            return reject('section_shortened')
        numbers = lambda text: Counter(re.findall(r'\d+(?:[.,]\d+)*', text))
        if numbers(before) - numbers(after):
            return reject('numbers_changed')
        table_lines = lambda text: [p.strip() for p in text.splitlines() if p.strip().startswith('|')]
        if table_lines(before) != table_lines(after):
            return reject('table_changed')
    return {**original, 'sections': [{**a, 'content': b['content']} for a, b in zip(old, new)]}


def validate_word_pdf(pdf_path, docx_path, _unused=None):
    pdf_path = Path(pdf_path)
    if not pdf_path.is_file() or not 1000 <= pdf_path.stat().st_size <= 49_000_000:
        raise PresentationPDFError('invalid_word_pdf_size')
    try:
        result = PdfReader(pdf_path)
        if result.is_encrypted or not result.pages:
            raise PresentationPDFError('invalid_word_pdf')
        with ZipFile(docx_path) as archive:
            root = ET.fromstring(archive.read('word/document.xml'))
        text = ' '.join(node.text or '' for node in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
        # Include table cells as well as paragraphs. Normalize whitespace and hyphenation.
        normalize = lambda s: re.sub(r'\s+|[\u00ad\-]', '', s).casefold()
        source_chars = Counter(normalize(text))
        exported = normalize(' '.join(page.extract_text() or '' for page in result.pages))
        present = Counter(exported)
        missing = sum((source_chars - present).values())
        if not exported or missing > max(3, sum(source_chars.values()) * .03):
            raise PresentationPDFError('word_pdf_text_loss')
        # Character coverage is only a gross loss guard, not proof of correct layout.
        for page in result.pages:
            if min(float(page.mediabox.width), float(page.mediabox.height)) < 100:
                raise PresentationPDFError('invalid_word_page_geometry')
    except PresentationPDFError:
        raise
    except Exception as exc:
        raise PresentationPDFError('unreadable_word_pdf') from exc


async def convert_word_pdf(docx_path, pdf_path, *, timeout=120, queue_timeout=240):
    result = await _convert_office_pdf(docx_path, pdf_path, None, source_suffix='.docx',
        export_filter='pdf:writer_pdf_Export', validator=validate_word_pdf,
        timeout=timeout, queue_timeout=queue_timeout)

    from academic_templates import cache_page_references
    await asyncio.to_thread(cache_page_references, docx_path, pdf_path)
    return result
