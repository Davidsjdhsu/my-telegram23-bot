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
    'Каждый абзац добавляет новую мысль. Удали общие фразы и повторы. '
    'Сохраняй оговорки и исключения: не превращай «обычно» или «при условии» в «всегда» или «невозможно». '
    'Не выдумывай статистику, исследования, источники, цитаты, личный опыт или реквизиты. '
    'Не заменяй отсутствующий источник словами «по разным оценкам». '
    'Условные примеры явно обозначай как условные. Данные и таблицы пользователя сохраняй точно. '
    'Отдельное требование пользователя об учебном характере документа явно укажи в тексте.'
)


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


def expansion_word_target(content, actual_pages, requested_pages):
    words = sum(len((s.get('content') or '').split()) for s in content.get('sections', []))
    return min(16000, max(requested_pages * 280, round(words * requested_pages / max(1, actual_pages) * 1.1)))


async def prepare_word_files(content, requested_pages, docx_path, pdf_path, *, build,
                             convert=None, expand=None, measure=word_pdf_page_count):
    """Measure actual pages; try one revision without risking the valid original pair.

    All expensive work is bounded. Editorial/expansion failures are non-blocking;
    initial build or integrity failures still propagate before delivery and billing.
    """
    convert = convert or convert_word_pdf
    await build(content, docx_path)
    await convert(docx_path, pdf_path)
    actual = await asyncio.to_thread(measure, pdf_path)
    attempts = 0
    if actual < requested_pages and expand is not None:
        attempts = 1
        promoting = False
        try:
            expanded = await expand(copy.deepcopy(content), actual, requested_pages)
            candidate = merge_word_expansion(content, expanded)
            if candidate != content:
                with tempfile.TemporaryDirectory(prefix='word-revision-', dir=Path(docx_path).parent) as td:
                    revised_docx, revised_pdf = Path(td) / 'revised.docx', Path(td) / 'revised.pdf'
                    await build(candidate, str(revised_docx))
                    await convert(str(revised_docx), str(revised_pdf))
                    revised_pages = await asyncio.to_thread(measure, revised_pdf)
                    # Never accept a revision that shortens the file or greatly overshoots.
                    if revised_pages > actual and abs(revised_pages - requested_pages) <= abs(actual - requested_pages):
                        promoting = True
                        os.replace(revised_docx, docx_path)
                        os.replace(revised_pdf, pdf_path)
                        content, actual = candidate, revised_pages
        except Exception as exc:
            if promoting:
                raise  # Never deliver a possibly mismatched DOCX/PDF pair.
            print('Word optional_revision_failure=' + type(exc).__name__, flush=True)
    print(f'Word pagination requested={requested_pages} actual={actual} revisions={attempts}', flush=True)
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
