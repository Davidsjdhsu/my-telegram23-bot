"""Validated private-menu preferences and palettes. No model calls."""
import json
from pathlib import Path
CONTENT_LANGUAGES = frozenset(('ru', 'en', 'de', 'ar', 'zh', 'es', 'fr'))
INTERFACE_LANGUAGES = frozenset(('ru', 'en'))
PALETTES = {p['key']: p for p in json.loads((Path(__file__).parent / 'studio_ui/templates/palettes.json').read_text())}

def content_language(value, fallback='ru'):
    return value if isinstance(value, str) and value in CONTENT_LANGUAGES else fallback

def preferences(body):
    if (not isinstance(body, dict) or set(body) != {'ui_lang', 'content_lang'}
            or not isinstance(body['ui_lang'], str) or body['ui_lang'] not in INTERFACE_LANGUAGES
            or not isinstance(body['content_lang'], str) or body['content_lang'] not in CONTENT_LANGUAGES):
        raise ValueError('Unsupported language')
    return dict(body)
