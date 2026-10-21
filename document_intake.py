"""Conservative, cost-free signals for a guided document request."""
import re

FORMATS = {'excel': '📊 Excel', 'word': '📄 Word', 'presentation': '🎞 Презентация'}
YES = '✅ Да, всё верно'
CHAT = '💬 Ответить в чате'
OWN = '📎 Только мои данные / текст'
AI = '✨ Подготовить содержание по теме'
STANDARD = 'Стандартный'
DESIGNER = 'Дизайнерский'

STRICT_SOURCE = re.compile(
    r'ничего\s+не\s+(?:добавля|придумыва|досочиня)|не\s+добавля\w*\s+(?:нов\w*\s+)?(?:строк|позиц)|'
    r'(?:используй|использовать|возьми|сохрани)\s+только|строго\s+по\s+(?:моим|этим|данным)|'
    r'only\s+(?:these|my|provided)|do\s+not\s+(?:add|invent)', re.I)
EXAMPLE = re.compile(r'придумай|условн\w*\s+(?:пример|данн)|примерн\w*\s+(?:цифр|данн)|sample\s+data|invent', re.I)
FORMATS_RE = {
    'excel': r'\b(?:excel|xlsx|эксель|эксел\w*|иксель|смет\w*|таблиц\w*|бюджет\w*|прайс\w*|финмодел\w*)\b',
    'presentation': r'\b(?:презентац\w*|през[ау]|слайд\w*|pptx|powerpoint|presentation)\b',
    'word': r'\b(?:word|docx|ворд|документ\w*|эссе|реферат\w*|конспект\w*|доклад\w*|курсов\w*|договор\w*|заявлени\w*|резюме|доверенност\w*|расписк\w*|претензи\w*|акт|согласи\w*|соглашени\w*|гарантийн\w*\s+письм\w*|nda|коммерческ\w*\s+предложени\w*)\b',
}
REQUEST = re.compile(r'\b(?:сдела\w*|созда\w*|собер\w*|собра\w*|напиш\w*|состав\w*|оформ\w*|подготов\w*|нуж\w*|хочу|make|create|write|build)\b', re.I)
DECLINE = re.compile(r'\bне\s+(?:хочу|нужен|нужна|нужно|надо|создавай|делай)\s+(?:мне\s+)?(?:файл|документ|таблиц\w*|презентац\w*|word|excel)', re.I)
QUESTION = re.compile(r'^\s*(?:почему|зачем|что\s+такое|как\s+|объясни|поясни|расскажи|проверь|реши\s+задачу|explain|what\s+is|how\s+)',re.I)


def structured_numbers(text):
    """Require multiple numeric records or a quantity/price expression, not a page count."""
    lines = re.split(r'[\n;]+', text or '')
    numeric = [line for line in lines if re.search(r'\d', line) and re.search(r'[а-яa-z]',line,re.I)
               and not re.search(r'\d+\s*(?:стр\w*|слайд\w*|pages?|slides?)\b',line,re.I)]
    return len(numeric) >= 2 or bool(re.search(r'\d[\d .,]*\s*(?:шт\w*|банк\w*|канистр\w*|штук\w*|кг|м²)\s+по\s+\d',text,re.I))


def pasted_prose(text):
    return len(text) >= 300 and (len([s for s in text.splitlines() if s.strip()]) >= 3 or len(text.split()) >= 70) and not REQUEST.search(text[:120])


def source_mode(text, fmt):
    if STRICT_SOURCE.search(text): return 'user'
    if EXAMPLE.search(text): return 'ai'
    if fmt == 'excel': return 'user' if structured_numbers(text) else None
    if re.search(r'готовый\s+текст|мой\s+текст|не\s+переписывай|оформи\s+(?:этот|следующий|текст)',text,re.I) or pasted_prose(text): return 'user'
    return 'ai'


def suggest_format(text):
    if DECLINE.search(text) or QUESTION.search(text): return None
    matches = [(m.start(), fmt) for fmt, pattern in FORMATS_RE.items() if (m := re.search(pattern,text,re.I))]
    if matches:
        pos,fmt=min(matches)
        if REQUEST.search(text) or pos < 12 or structured_numbers(text): return fmt
    if structured_numbers(text): return 'excel'
    if pasted_prose(text): return 'word'
    return None


def needs_topic(text):
    # Only bare format commands lack a topic. Never discard a concrete short topic.
    t=text.lower().strip(' .!')
    for pattern in FORMATS_RE.values(): t=re.sub(pattern,'',t)
    t=REQUEST.sub('',t)
    t=re.sub(r'\b(?:мне|пожалуйста|нужна|нужен|на|в|файл)\b','',t)
    return not re.search(r'[а-яa-z0-9]',t)


def excel_appearance(text):
    t=text.lower()
    if re.search(r'дизайнер\w*|designer',t): return 'designer'
    if re.search(r'стандартн\w*|обычн\w*|standard',t): return 'standard'
    return None


def excel_palette(text):
    for fragment,key in [('графит','copper'),('медн','copper'),('изумруд','emerald'),('песоч','emerald'),('черниль','lilac'),('сирен','lilac'),('бордов','burgundy'),('кремов','burgundy'),('зелён','green'),('зелен','green')]:
        if fragment in text.lower(): return key
    return None
