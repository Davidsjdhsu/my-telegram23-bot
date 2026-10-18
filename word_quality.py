"""Word export and lossless text expansion, shared by bot entry points."""
from collections import Counter
import re
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pypdf import PdfReader
from presentation_pdf import _convert_office_pdf, PresentationPDFError


def merge_word_expansion(original, expanded):
    """Only accept expanded paragraph text; never replace tables or user metadata."""
    old, new = original.get('sections', []), expanded.get('sections', [])
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
