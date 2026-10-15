"""Versioned catalogue of 50 design kits, selected locally without paid calls.

A kit combines an art direction, typography, geometry and twelve semantic compositions.
Compositions share tested drawing primitives; they are not 600 hand-authored templates.
"""
import hashlib
import re

VERSION = 1
COMPOSITIONS = ('cover_split', 'cover_band', 'photo_left', 'photo_right',
                'photo_top', 'photo_bottom', 'process', 'sections',
                'comparison', 'conclusion', 'statement', 'chart')
# Each row is an authored direction, not a random colour permutation.
# name | family | layout system | title font | body font | paper | ink | accent | topic signals
_ROWS = '''
Медовый атлас|education|field|Georgia|Calibri|FFF6DC|392A12|956000|пчел,улей,мед,bee,honey
Ботанический дневник|education|margin|Georgia|Calibri|F1F4E7|283D29|526B35|растен,ботан,цвет,botan,plant
Океаническая экспедиция|education|panorama|Trebuchet MS|Arial|EAF4F3|153F4B|087A89|океан,кит,морск,риф,ocean,whale
Лесной маршрут|education|split|Verdana|Calibri|EEF1E6|293B24|4E6A2F|лес,дерев,forest,woodland
Мир животных|education|gallery|Trebuchet MS|Calibri|FAF0E2|45331F|A25C25|животн,собак,кошк,animal,dog,cat
Космическая обсерватория|technology|cinema|Trebuchet MS|Arial|111828|F1F3FF|A9BAFF|космос,звезд,планет,space,astronom
Лабораторный журнал|education|report|Verdana|Calibri|F0F5F8|26394F|32679A|хими,лаборатор,chemistr,laborator
Физика света|education|poster|Arial|Calibri|FDF3DF|332B28|A95125|физик,свет,оптик,physics,optics
Исторический архив|education|margin|Times New Roman|Calibri|F4EDDF|3E3025|825037|истори,архив,history,archive
Географический атлас|education|field|Trebuchet MS|Calibri|EEF2EA|27413D|386F65|географ,геолог,материк,geograph,geolog
Тосканский журнал|editorial|editorial|Georgia|Calibri|F5EBDD|39271D|9B522E|тоскан,виноград,итал,tuscan,vineyard,italy
Тихая роскошь|editorial|gallery|Georgia|Arial|F6F1E8|38322C|776044|бутик,отел,курорт,boutique,hotel,resort
Скандинавский дом|editorial|split|Trebuchet MS|Calibri|F1F2ED|303D36|5B7164|скандинав,дом,интерьер,scandinav,interior,home
Городская архитектура|editorial|poster|Arial|Arial|F0EFEB|292D33|7F533C|архитект,здан,urban,architect
Путевой альбом|editorial|panorama|Georgia|Calibri|F8EFE4|40342D|A25038|путешеств,туризм,маршрут,travel,tourism
Кофейная редакция|editorial|margin|Georgia|Calibri|F3E7D7|3A251C|8B4F2D|коф,бариста,coffee,barista
Ресторанная история|editorial|editorial|Times New Roman|Arial|F8EFE6|492C2B|994B43|ресторан,кухн,блюд,restaurant,cuisine
Пекарня утром|editorial|gallery|Georgia|Calibri|FFF3E3|533A26|A5682F|пекар,хлеб,выпеч,bakery,bread
Фермерский рынок|editorial|field|Trebuchet MS|Calibri|F2F0DB|38432B|62732C|ферм,овощ,урожай,farm,harvest
Чайный ритуал|editorial|split|Georgia|Calibri|EBF0E7|2C443A|4A7460|чай,матча,tea,matcha
Инженерная система|technology|report|Trebuchet MS|Arial|EDF4FA|132C43|006C82|робот,склад,автоматиза,robot,warehouse
Цифровой горизонт|technology|cinema|Arial|Arial|101D2B|EDF8FF|60D3E2|ии,нейросет,искусственн,artificial,ai
Продуктовый интерфейс|technology|split|Verdana|Calibri|EFF2FC|253354|5267A0|приложен,софт,программ,software,application
Энергия будущего|technology|poster|Trebuchet MS|Arial|F3F5E5|283C32|577329|энерг,солнеч,электр,energy,solar
Кибербезопасность|technology|margin|Arial|Arial|111F21|E7FAF5|6DD3AE|кибер,защит,безопасност,cyber,secur
Стратегический отчёт|business|report|Calibri|Calibri|F8FAF9|193F3A|176C59|стратег,компан,проект,strateg,company
Финансовый обзор|business|margin|Georgia|Arial|EEF2F6|253A56|496B98|финанс,инвест,банк,financ,invest,bank
Стартап на сцене|business|poster|Trebuchet MS|Calibri|F9EEE7|4C2C27|B14D36|стартап,питч,startup,pitch
Командная работа|business|split|Verdana|Calibri|F2F0F9|3A3457|72629B|команд,сотрудник,персонал,team,hr
Портрет бренда|business|editorial|Georgia|Arial|F9F0E6|492F2D|A35D45|бренд,маркетинг,brand,marketing
Клинический обзор|education|report|Arial|Calibri|EFF7F7|24434B|267985|медицин,клиник,диагност,medical,clinic
Движение и спорт|education|poster|Verdana|Arial|F7F2E6|343A31|7B7029|спорт,фитнес,трениров,sport,fitness
Психология общения|education|gallery|Georgia|Calibri|F3EDF4|433649|86638D|психолог,эмоци,psycholog,emotion
Баланс и здоровье|editorial|split|Trebuchet MS|Calibri|EBF3ED|32453C|597B62|здоров,питан,wellness,health
Нейронаука|technology|cinema|Trebuchet MS|Arial|211D32|F5EFFC|B9A1E1|мозг,нейрон,neuro,brain
Галерея искусства|editorial|gallery|Georgia|Arial|F8F5EF|302F2C|8D503A|искусств,живопис,худож,art,museum
Модная редакция|editorial|editorial|Times New Roman|Arial|F9EFEF|40292F|9B5365|мода,одежд,стил,fashion,clothes
Музыкальная сцена|editorial|cinema|Trebuchet MS|Arial|221C2D|F9EFF7|E6A3C7|музык,концерт,music,concert
Кинопанорама|editorial|panorama|Georgia|Calibri|181D26|F4EFE4|DAB786|кино,фильм,cinema,film
Книжная полка|editorial|margin|Georgia|Calibri|F5EEDC|3E342B|86643F|книг,литератур,чтен,book,literatur
Мастерская предметов|business|gallery|Trebuchet MS|Arial|F3EEE6|3B3730|7F643E|предмет,издел,дизайн продукт,object,craft
Автомобильный обзор|technology|panorama|Arial|Arial|1A222A|EFF4F6|ABC8DA|автомоб,транспорт,car,vehicle
Городская мобильность|technology|field|Verdana|Calibri|F0F4EB|304335|5A763B|велосипед,мобильност,велодорож,bike,mobility
Мир игр|technology|poster|Trebuchet MS|Arial|211E36|F6F1FF|B8A0EF|игр,гейм,game,gaming
Вещи и материалы|business|split|Georgia|Calibri|F4F0E9|3F3B34|807452|материал,мебел,текстил,material,furniture
Экологический обзор|education|report|Trebuchet MS|Calibri|EDF3E6|304830|557F39|эколог,климат,environment,climate
Социальная инициатива|business|field|Verdana|Calibri|FCF0E6|493A31|9C6742|социальн,волонтер,благотвор,charity,social
Образовательный курс|education|split|Verdana|Calibri|F1F1F8|363B58|5E6EA5|курс,обучен,образован,course,educat
Данные и исследования|business|report|Arial|Arial|EDF2F6|2F4153|4B6E91|данн,исследован,опрос,research,data
Событие и фестиваль|editorial|poster|Trebuchet MS|Calibri|FBEEE7|4B3331|B05A45|фестивал,событи,праздник,festival,event
'''.strip().splitlines()

TEMPLATES = {}
for i, row in enumerate(_ROWS, 1):
    name, family, system, font, body_font, paper, ink, accent, tags = row.split('|')
    key = f'd{i:02}'
    # Per-kit proportions/scale are stable; a change of theme never scrambles a confirmed plan.
    TEMPLATES[key] = dict(id=key, number=i, name=name, family=family, system=system,
        font=font, body_font=body_font, paper='#'+paper, ink='#'+ink, accent='#'+accent,
        tags=tuple(tags.split(',')), margin=(48,56,64,72,80)[(i-1)%5],
        split=(.42,.46,.50,.54,.58)[(i-1)%5],
        title_size=(76,70,84,72,80)[(i-1)%5], bold=font not in ('Georgia','Times New Roman'),
        compositions=COMPOSITIONS, sequence=(i-1)%4)

# Topic-specific composition decisions where two directions would otherwise share a grid.
for key, changes in {
    'd29': {'system':'report','split':.52,'title_size':70},
    'd37': {'system':'gallery','split':.40,'title_size':86},
    'd33': {'system':'margin','split':.52,'title_size':68},
    'd45': {'system':'field','split':.56,'title_size':72},
    'd48': {'system':'report','split':.44,'title_size':72},
    'd44': {'system':'cinema','split':.54,'title_size':88},
}.items():
    TEMPLATES[key].update(changes)


def fingerprint(text):
    return int.from_bytes(hashlib.sha256(text.encode('utf-8')).digest()[:4], 'big')


def choose_template(data, family=None):
    topic = str(data.get('topic') or data.get('user_text') or '').lower().replace('ё','е')
    text = '\n'.join(str(data.get(k) or '') for k in ('topic','user_text','extra')).lower().replace('ё','е')
    explicit = list(re.finditer(r'(?:дизайн|шаблон|design|template)\s*(?:№|#)?\s*(?:d)?(0?[1-9]|[1-4][0-9]|50)\b',text))
    if explicit:
        return TEMPLATES[f'd{int(explicit[-1][1]):02}']
    named = [(text.rfind(t['name'].lower()),t) for t in TEMPLATES.values() if t['name'].lower() in text]
    if named:return max(named,key=lambda x:x[0])[1]
    candidates=[t for t in TEMPLATES.values() if not family or t['family']==family]
    def match(tag, haystack):
        # Short tags must be words, never accidental substrings (AI in "training").
        return bool(re.search(r'\b'+re.escape(tag)+r'\b',haystack)) if tag=='ии' or (tag.isascii() and len(tag)<=3) else tag in haystack
    scores=[]
    for t in candidates:
        subject=[len(tag) for tag in t['tags'] if match(tag,topic)]
        context=[len(tag) for tag in t['tags'] if match(tag,text)]
        scores.append(((100*max(subject)+len(subject)) if subject else max(context,default=0),t))
    best=max(score for score,_ in scores)
    if best:return next(t for score,t in scores if score==best)
    # Unknown themes use a neutral kit, chosen consistently from the actual subject, not audience.
    pool=candidates if family else [TEMPLATES[k] for k in ('d26','d29','d41','d49')]
    return pool[fingerprint(topic)%len(pool)]


def template_palette(template):
    dark=int(template['paper'][1:3],16)<80
    return dict(label=template['name'],font=template['font'],body_font=template['body_font'],
        bold=template['bold'],light='#F5F3EE' if dark else template['paper'],
        dark=template['paper'] if dark else '#1F2529',ink='#29343A' if dark else template['ink'],
        body='#3D4B50' if dark else template['ink'],accent=template['accent'],
        bright=template['accent'] if dark else '#D3E0D6',secondary='#C8D2CF')


def composition_for(item, number, total):
    d=item.get('_designer_design',{});t=TEMPLATES.get(d.get('template_id'))
    if not t:return None
    seed=d.get('variant',0)
    if number==1:return 'cover_band' if (seed+t['sequence'])%2 else 'cover_split'
    layout=item.get('designer',{}).get('layout','auto')
    if item.get('chart'):return 'chart'
    if layout in ('process','sections','comparison','conclusion','statement'):return layout
    return COMPOSITIONS[2+(number-2+t['sequence']+seed)%4]
