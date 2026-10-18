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
WORD_EDITORIAL_GUIDANCE = (
    'Пиши ясно для указанной аудитории. Введение кратко объясняет цель и границы темы; '
    'основное сравнение, объяснения и примеры вынеси в основную часть. '
    'Вывод отвечает на поставленный вопрос и помогает сделать выбор, а не повторяет введение. '
    'Не пересказывай таблицу по строкам: поясни существенные различия и условия применения. '
    'Каждый абзац добавляет новую мысль; отделяй абзацы символами \\n\\n внутри content. '
    'Обычно достаточно 40–90 слов на абзац; не собирай сравнение всех вариантов в один абзац. '
    'Удали общие фразы и повторы. '
    'Сохраняй оговорки и исключения: не превращай «обычно» или «при условии» в «всегда» или «невозможно». '
    'Не выдумывай статистику, исследования, источники, цитаты, личный опыт или реквизиты. '
    'Не заменяй отсутствующий источник словами «по разным оценкам». '
    'Условные примеры явно обозначай как условные. Данные и таблицы пользователя сохраняй точно. '
    'Отдельное требование пользователя об учебном характере документа явно укажи в тексте. '
    'Перед возвратом JSON проверь причинные связи и категоричные утверждения о доступности, '
    'безопасности и невозможности действия. Различай свойства объекта, настройки и условия работы. '
    'Указывай конкретное условие или исключение рядом с утверждением; не добавляй оговорки механически '
    'к достоверным определениям. Требование связи для обмена изменениями не означает, что нельзя '
    'работать с уже сохранённой локальной копией. Не выдавай синхронизацию за независимую резервную копию.'
)

# Word must not inherit the presentation art-director persona or its Russian-only
# instruction. Keep the document's requested depth and language above style tricks.
WORD_SYSTEM = (
    'Ты редактор полноразмерных Word-документов. Выполняй задачу в указанном жанре, '
    'на запрошенном языке и в запрошенном формате ответа. Документ должен раскрывать тему, '
    'а не быть тезисами для слайда или кратким пересказом. '
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
    'Например, облачные файлы могут быть доступны без интернета, если локальная копия заранее '
    'загружена и сервис поддерживает нужный офлайн-режим. Локальное хранение само по себе '
    'не устраняет сетевые угрозы, а облако не устраняет все физические риски. '
    'Не переносись к этим примерам, если тема документа иная. '
    'Не добавляй выдуманные источники или результаты исследований. Если факт нельзя уверенно '
    'подтвердить, убери неподтверждённую конкретику или обозначь пределы применимости. '
    'Не добавляй в документ отчёт о своей проверке. Верни только JSON с sections (title, content), '
    'с теми же заголовками и порядком. Соблюдай заданные ориентиры слов по разделам. '
    + WORD_EDITORIAL_GUIDANCE
)


def word_section_budgets(content, target):
    """Give the optional editor concrete prose budgets rather than one vague total."""
    sections = content.get('sections', [])
    short = re.compile(r'введени|вывод|заключени|introduction|conclusion|summary', re.I)
    nonbody = re.compile(r'содержани|литератур|источник|библиограф|references|bibliograph|contents', re.I)
    weights = [0 if nonbody.search(s.get('title', '')) else .3 if short.search(s.get('title', '')) else 1
               for s in sections]
    total = sum(weights)
    return [{'title': s.get('title', ''), 'target_words': round(target * w / total) if total else 0}
            for s, w in zip(sections, weights)]


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


def word_pdf_layout(pdf_path):
    """Approximate body extent on the last page, excluding the footer area.

    A cheap signal for text documents, not a visual or semantic quality score.
    The builder uses 2 cm top/bottom margins. Tables count toward the extent.
    """
    page = PdfReader(pdf_path).pages[-1]
    height = float(page.mediabox.height)
    top, bottom = height - 56.7, 56.7
    baselines = []
    def visit(text, cm, tm, font, size):
        y = tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
        if text.strip() and bottom <= y <= top:
            baselines.append(y - min(float(size), 18) * .25)
    page.extract_text(visitor_text=visit)
    fill = min(1., max(0., (top - min(baselines)) / (top - bottom))) if baselines else 0.
    return {'last_page_fill': round(fill, 3)}


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
    return {
        'short_pages': actual < requested,
        'sparse_last_page': actual == requested and requested > 1 and fill is not None and fill < .5,
        'long_paragraphs': sum(len(p.split()) > 100 and not p.lstrip().startswith('|') for p in paragraphs),
        'categorical_claims': claims,
        'last_page_fill': fill,
    }


def needs_word_revision(findings):
    return any(findings[k] for k in ('short_pages', 'sparse_last_page', 'long_paragraphs', 'categorical_claims'))


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


def expansion_word_target(content, actual_pages, requested_pages):
    words = sum(len((s.get('content') or '').split()) for s in content.get('sections', []))
    return min(16000, max(requested_pages * 280, round(words * requested_pages / max(1, actual_pages) * 1.1)))


async def prepare_word_files(content, requested_pages, docx_path, pdf_path, *, build,
                             convert=None, expand=None, measure=word_pdf_page_count, inspect=None):
    """Measure actual pages; try one revision without risking the valid original pair.

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
    if needs_word_revision(findings) and expand is not None:
        attempts = 1
        promoting = False
        try:
            expanded = await expand(copy.deepcopy(content), actual, requested_pages, dict(findings))
            candidate = merge_word_expansion(content, expanded)
            if candidate != content:
                with tempfile.TemporaryDirectory(prefix='word-revision-', dir=Path(docx_path).parent) as td:
                    revised_docx, revised_pdf = Path(td) / 'revised.docx', Path(td) / 'revised.pdf'
                    await build(candidate, str(revised_docx))
                    await convert(str(revised_docx), str(revised_pdf))
                    revised_pages = await asyncio.to_thread(measure, revised_pdf)
                    revised_layout = await layout_at(revised_pdf)
                    revised_findings = word_review_findings(candidate, revised_pages, requested_pages, revised_layout)
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
                    )
                    if closer or (same_pages and fill_ok and improved):
                        promoting = True
                        os.replace(revised_docx, docx_path)
                        os.replace(revised_pdf, pdf_path)
                        content, actual, findings = candidate, revised_pages, revised_findings
                        accepted = True
        except Exception as exc:
            if promoting:
                raise  # Never deliver a possibly mismatched DOCX/PDF pair.
            print('Word optional_revision_failure=' + type(exc).__name__, flush=True)
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
    if not isinstance(expanded, dict):
        return original
    old, new = original.get('sections', []), expanded.get('sections', [])
    if not isinstance(new, list):
        return original
    if len(old) != len(new) or not all(isinstance(b, dict) for b in new):
        return original
    if any(a.get('title') != b.get('title') or not isinstance(b.get('content'), str)
           for a, b in zip(old, new)):
        return original
    for a, b in zip(old, new):
        before, after = a.get('content') or '', b['content']
        # A revision may polish prose, but must not erase a section or supplied numbers.
        if before.strip() and not after.strip():
            return original
        if len(after.split()) < len(before.split()) * .8:
            return original
        numbers = lambda text: Counter(re.findall(r'\d+(?:[.,]\d+)*', text))
        if numbers(before) - numbers(after):
            return original
        table_lines = lambda text: [p.strip() for p in text.splitlines() if p.strip().startswith('|')]
        if table_lines(before) != table_lines(after):
            return original
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
    return await _convert_office_pdf(docx_path, pdf_path, None, source_suffix='.docx',
        export_filter='pdf:writer_pdf_Export', validator=validate_word_pdf,
        timeout=timeout, queue_timeout=queue_timeout)
