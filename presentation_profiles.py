"""Deterministic deck art direction. No model, network or user storage."""
import copy
import re

PROFILES = {
    'editorial': dict(label='Журнальный', font='Georgia', body_font='Calibri', bold=False,
        light='#F5EBDD', dark='#241B17', ink='#39271D', body='#564238', accent='#9B522E', bright='#E2AB7D', secondary='#EAD5B8'),
    'technology': dict(label='Технологический', font='Trebuchet MS', body_font='Arial', bold=True,
        light='#EDF4FA', dark='#0C1C2C', ink='#132C43', body='#344F62', accent='#006C82', bright='#72DDDF', secondary='#CEDFEA'),
    'business': dict(label='Деловой', font='Calibri', body_font='Calibri', bold=True,
        light='#F8FAF9', dark='#152B29', ink='#193F3A', body='#39524D', accent='#176C59', bright='#93D5B7', secondary='#E1EAE5'),
    'education': dict(label='Учебный', font='Verdana', body_font='Calibri', bold=True,
        light='#F2F4E7', dark='#20352C', ink='#234231', body='#3D5346', accent='#426B2D', bright='#C9DE91', secondary='#DFE7C9'),
}
LABELS = {
    'en': {'editorial':'Editorial', 'technology':'Technology', 'business':'Business', 'education':'Educational'},
    'de': {'editorial':'Magazin', 'technology':'Technologie', 'business':'Business', 'education':'Lernen'},
    'fr': {'editorial':'Éditorial', 'technology':'Technologique', 'business':'Professionnel', 'education':'Pédagogique'},
    'es': {'editorial':'Editorial', 'technology':'Tecnológico', 'business':'Empresarial', 'education':'Educativo'},
}


def choose_profile(data):
    """Explicit style wins, then audience, then topic. Last explicit request wins."""
    parts = [str(data.get(k) or '') for k in ('topic', 'user_text', 'extra')]
    text = '\n'.join(parts).lower().replace('ё', 'е')
    explicit = []
    for key, pattern in {
        'editorial': r'журнальн\w*|редакционн\w*|editorial',
        'technology': r'технологическ\w*\s+(?:стил\w*|дизайн\w*|оформлен\w*)|(?:tech|technology)\s+style',
        'business': r'делов\w*\s+(?:стил\w*|дизайн\w*|оформлен\w*)|business\s+style',
        'education': r'учебн\w*\s+(?:стил\w*|дизайн\w*|оформлен\w*)|educational\s+style',
    }.items():
        explicit.extend((m.start(), key) for m in re.finditer(pattern, text))
    if explicit:
        return max(explicit)[1]
    audience = text + ' ' + str(data.get('audience') or '').lower()
    if re.search(r'ученик|школьник|класс[а-я]*\b|для детей|school|students|biology|биолог|истори[яию]|природ|океан|кит[а-я]*\b', audience):
        return 'education'
    topic = str(data.get('topic') or data.get('user_text') or '').lower().replace('ё', 'е')
    if re.search(r'коф|ресторан|кухн|путешеств|отел|интерьер|мод[аыу]\b|coffee|travel|hotel|food|fashion', topic):
        return 'editorial'
    if re.search(r'технолог|робот|искусственн|нейросет|\bии\b|\bai\b|technolog|robot|software|космос|space', topic):
        return 'technology'
    return 'business'


def choose_design(data):
    family = choose_profile(data); profile = PROFILES[family]
    text = '\n'.join(str(data.get(k) or '') for k in ('topic','user_text','extra')).lower().replace('ё','е')
    background = 'dark' if family == 'technology' else 'light'
    matches = list(re.finditer(r'(темн\w*|светл\w*|бел\w*|черн\w*|dark|light|white|black)\s+(?:фон\w*|background)', text))
    if matches:
        background = 'dark' if re.match(r'темн|черн|dark|black',matches[-1][1]) else 'light'
    colors = []
    color_words = {'медн': '#C48356', 'кремов': '#EAD5B8', 'бирюзов': '#42BBC1',
        'син': '#3B64CF', 'зелен': '#43875F', 'оранжев': '#D77C39', 'фиолетов': '#9670CF',
        'copper':'#C48356', 'cream':'#EAD5B8', 'teal':'#42BBC1', 'blue':'#3B64CF', 'green':'#43875F'}
    # Read colors only next to design vocabulary, not from subjects like blue whales.
    for m in re.finditer(r'[^.\n]*(?:акцент|палитр|цвет|accent|palette|color)[^.\n]*',text):
        clause=m[0]
        colors.extend((m.start()+c.start(), c[0]) for c in re.finditer(r'#[0-9a-f]{6}\b',clause))
        for stem, value in color_words.items():
            colors.extend((m.start()+c.start(), value) for c in re.finditer(r'\b'+stem+r'\w*',clause))
    colors = [v for _,v in sorted(colors)]
    accent = colors[-1] if colors else profile['bright' if background=='dark' else 'accent']
    secondary = colors[-2] if len(colors)>1 else profile['secondary']
    return {'family':family,'background':background,'accent':accent,'secondary':secondary}


def profile_summary(design, lang='ru'):
    family=design['family'];p=PROFILES[family]
    name=LABELS.get(lang,{}).get(family,p['label'] if lang=='ru' else LABELS['en'][family])
    if lang=='ru':
        return f"{name}: {'тёмный' if design['background']=='dark' else 'светлый'} фон, заголовки {p['font']}, разные композиции слайдов"
    return f"{name}: {design['background']} background, {p['font']} headings, varied slide compositions"


def relative_luminance(color):
    rgb=[int(color[i:i+2],16)/255 for i in (1,3,5)]
    return sum((v/12.92 if v<=.04045 else ((v+.055)/1.055)**2.4)*w for v,w in zip(rgb,(.2126,.7152,.0722)))


def readable_color(color, background, target=4.5):
    """Preserve hue while moving lightness to accessible text contrast."""
    lum=relative_luminance(background)
    rgb=[int(color[i:i+2],16) for i in (1,3,5)]
    for step in range(21):
        goal=255 if lum<.3 else 0
        candidate='#'+''.join(f'{round(v+(goal-v)*step/20):02X}' for v in rgb)
        cl=relative_luminance(candidate)
        if (max(lum,cl)+.05)/(min(lum,cl)+.05)>=target:return candidate
    return '#FFFFFF' if lum<.3 else '#000000'


def resolved_profile(design):
    p=copy.deepcopy(PROFILES[design['family']]);dark=design['background']=='dark'
    p['bg']=p['dark'] if dark else p['light']
    p['ink']=p['secondary'] if dark and design['family']=='editorial' else '#F3F6F8' if dark else p['ink']
    p['body']='#DCE4E5' if dark else p['body']
    p['accent']=readable_color(design['accent'],p['bg'])
    p['secondary']=design['secondary'];p['muted']=readable_color('#7D827D',p['bg'])
    return p
