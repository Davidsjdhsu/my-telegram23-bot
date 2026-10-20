"""Neutral university-derived covers and deterministic academic Word layout.

Provenance and adaptation boundaries are recorded in assets/academic/sources.json.
The profiles are not a claim of approval by every university.
"""
from pathlib import Path
import hashlib
import textwrap
import re
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ASSETS = Path(__file__).resolve().parent / 'assets' / 'academic'
KINDS = frozenset(('referat', 'coursework', 'report', 'essay', 'notes'))
COVER_LABELS = {'referat': 'РЕФЕРАТ', 'coursework': 'КУРСОВАЯ РАБОТА',
                'report': 'ДОКЛАД', 'essay': 'ЭССЕ', 'notes': 'КОНСПЕКТ'}
BLANKS = {
    'referat': 'Реферат — титульный лист',
    'coursework': 'Курсовая — титульный лист',
    'essay': 'Эссе — титульный лист',
    'report': 'Доклад — титульный лист',
    'notes': 'Конспект — титульный лист',
    'course_project': 'Курсовой проект — титульный лист',
    'lab': 'Лабораторная — титульный лист',
    'practice': 'Практика — титульный лист',
    'bachelor': 'ВКР бакалавра — титульный лист',
    'diploma': 'Дипломная — титульный лист',
    'master': 'Магистерская — титульный лист',
    'assignment': 'Задание на ВКР',
}
NOTICE = ('Титульные адаптированы по отдельным образцам ТГУ, Московского международного '
          'университета, Самарского университета и Международного банковского института, без чужих логотипов. '
          'Это базовый вариант: требования вашей кафедры могут отличаться.')
META_LABELS = {'org':'Ведомство', 'school':'Учебное заведение', 'department':'Факультет или кафедра',
               'discipline':'Дисциплина', 'author':'ФИО студента', 'group':'Группа',
               'teacher':'Руководитель', 'teacher_position':'Должность руководителя',
               'specialty':'Направление подготовки', 'profile':'Профиль', 'city':'Город', 'year':'Год',
               'course':'Курс', 'record_book':'Номер зачётной книжки', 'institute':'Институт'}
ALIASES = {'вуз':'school','университет':'school','учебное заведение':'school','фио':'author',
           'автор':'author','студент':'author','фио студента':'author','преподаватель':'teacher',
           'фио преподавателя':'teacher','фио научного руководителя':'teacher','научный руководитель':'teacher',
           'факультет':'department','кафедра':'department','специальность':'specialty',
           'номер зачетной книжки':'record_book','№ зачётной книжки':'record_book',
           '№ зачетной книжки':'record_book','зачётная книжка':'record_book',
           'должность преподавателя':'teacher_position','ученая степень, звание':'teacher_position',
           'учёная степень, звание':'teacher_position'}
TITLE_FIELDS = ('school', 'discipline', 'author', 'group', 'teacher', 'city', 'year')
TITLE_FIELDS_BY_KIND = {
    'referat': TITLE_FIELDS,
    'coursework': ('school', 'author', 'group', 'teacher', 'city', 'year'),
    'essay': ('school', 'author', 'group', 'teacher', 'city', 'year'),
    'report': ('school', 'course', 'author', 'group', 'teacher', 'city', 'year'),
    'notes': ('school', 'discipline', 'course', 'author', 'group', 'teacher', 'city', 'year'),
}
LEAVE_TITLE_BLANKS = 'Оставить пустые поля'


def request_meta(data):
    return explicit_meta({}, '\n'.join(str(data.get(k) or '') for k in ('topic', 'user_text', 'extra')))


def title_request_key(data):
    # A choice to leave gaps belongs only to this request, never to the next one.
    value = '\0'.join(str(data.get(k) or '') for k in
                      ('word_kind', 'topic', 'user_text', 'extra', 'word_pages'))
    return hashlib.sha256(value.encode()).hexdigest()


def missing_title_fields(data):
    meta = request_meta(data)
    return [key for key in TITLE_FIELDS_BY_KIND.get(data.get('word_kind'), TITLE_FIELDS)
            if not meta.get(key)]


def title_question(data):
    fields = '\n'.join(META_LABELS[key] + ': ' for key in missing_title_fields(data))
    optional = (' Для конспекта можно также указать «Номер зачётной книжки: …» '
                'и «Должность преподавателя: …». Оценка, дата проверки и подписи остаются пустыми.'
                if data.get('word_kind') == 'notes' else '')
    return ('Для титульного листа по учебному образцу не хватает данных. '
            'Пришли их одним сообщением, сохранив названия полей:\n\n' + fields
            + '\n\nКафедру можно добавить отдельной строкой «Кафедра: …».' + optional + ' '
              'Если часть данных неизвестна, нажми «' + LEAVE_TITLE_BLANKS
            + '»: в документе останутся строки для ручного заполнения. '
              'Генерация ещё не началась, кредиты не списаны.')

def blank_path(kind):
    if kind not in BLANKS:
        raise ValueError('unknown_academic_blank')
    return ASSETS / (kind + '.docx')


def explicit_meta(meta, source):
    """Explicit labels win over generated metadata, including deliberate clearing."""
    result = dict(meta or {})
    names = {v.casefold(): k for k, v in META_LABELS.items()}
    names.update(ALIASES)
    for line in source.splitlines():
        label, sep, value = line.partition(':')
        key = names.get(' '.join(label.casefold().strip().split())) if sep else None
        if key:
            result[key] = value.strip()
    return result


def guidance(kind):
    if kind not in KINDS:
        return ''
    common = ('\nУчебная структура: титульный лист, оглавление и номера страниц создаёт программа. '
              'НЕ включай титульный лист и оглавление в sections. Заголовки разделов должны описывать их содержание. '
              'Данные в meta бери только из запроса. Не придумывай вуз, ФИО, группу, город, год, '
              'подписи, результаты экспериментов, источники или цитаты. Не требуй определённого числа источников: '
              'в список включай только источники, предоставленные пользователем или проверенные по доступному тексту. '
              'При отсутствии источников не создавай фиктивную библиографию. ')
    return common + {
        'referat':'Разделы: Введение (цель и задачи), 2–3 тематических раздела, Заключение, источники при наличии. Главы обозначай 1, 2, подразделы 1.1, 1.2. Введение и заключение не нумеруй.',
        'coursework':'Разделы: Введение (проблема, цель, задачи, объект и предмет), теоретическая глава, практическая/аналитическая глава, Заключение, источники при наличии. Главы 1, 2; подразделы 1.1, 1.2. Никаких вымышленных результатов исследования; если данных нет, опиши методику предполагаемого исследования.',
        'report':'Доклад: введение с вопросом, основная часть с объяснением и примерами, выводы, источники при наличии. Не превращай его в курсовую.',
        'essay':'Эссе: постановка проблемы и тезис, аргументы и примеры, рассмотрение возражения, собственный вывод. Титульный лист с надписью ЭССЕ добавляет программа. Без оглавления; основной текст связный, с абзацами, без видимых заголовков разделов. sections служит только для внутреннего плана. Не приписывай автору личный опыт, которого нет в запросе. Не начинай с фразы «Этот текст — учебное эссе»: сразу переходи к вопросу темы.',
        'notes':'Конспект: начни сразу с ключевых понятий, без разделов «Введение» и «Заключение», если пользователь их не просил. Сделай 5–8 тематических разделов, каждый раскрывает отдельный вопрос. Внутри — определения, короткие тезисы с пояснениями, причинные связи, конкретные примеры и условия применения. Используй маркированные пункты там, где перечисляются самостоятельные положения; обычный пункт — 25–60 слов, одна мысль. Не повторяй одни и те же определения в разных разделах. Заверши 3–5 вопросами для самопроверки с краткими ответами. Сохрани содержание исходного материала. Титульный лист с надписью КОНСПЕКТ добавляет программа. Без оглавления. В title верни только название темы без «Конспект по теме».',
    }[kind]


def _text(p, text):
    # Preserve the paragraph and first run formatting from the reference.
    nodes = list(p._p.iter(qn('w:t')))
    if nodes:
        nodes[0].text = str(text)
        nodes[0].set(qn('xml:space'), 'preserve')
        for n in nodes[1:]: n.text = ''
    else:
        p.add_run(str(text))


def _cover(kind, title, meta):
    if kind in ('essay', 'report', 'notes'):
        return _independent_cover(kind, title, meta)
    d = Document(blank_path(kind))
    _ensure_styles(d)
    # Body Normal uses a first-line indent and 14 pt; the source cover uses
    # 12 pt without that indent. Keep its own style so body styling cannot leak in.
    cover_style = d.styles.add_style('Academic Cover', WD_STYLE_TYPE.PARAGRAPH)
    cover_style.font.name = 'Times New Roman'
    cover_style.font.size = Pt(12)
    cover_style.font.color.rgb = RGBColor(0, 0, 0)
    cover_style.paragraph_format.first_line_indent = Cm(0)
    cover_style.paragraph_format.line_spacing = 1
    cover_style.paragraph_format.space_before = Pt(0)
    cover_style.paragraph_format.space_after = Pt(0)
    spare_lines = []
    extra_lines = 0
    teacher_seen = False
    for i,p in enumerate(d.paragraphs):
        p.style = cover_style
        t = p.text.strip()
        # Only original spacer paragraphs may absorb longer user values.
        if not t and (4 <= i <= 13 or 23 <= i <= 27 or 41 <= i <= 46):
            spare_lines.append(p)
        v = None
        if i == 0: v = meta.get('org','')
        elif i == 2: v = meta.get('school') or 'Учебное заведение: ____________________'
        elif i == 3: v = meta.get('department','')
        elif t in ('ТЕМА РАБОТЫ','ТЕМА РЕФЕРАТА'): v = title
        elif t == 'РЕФЕРАТ': v = COVER_LABELS[kind]
        elif t.startswith('по дисциплине'): v = 'по дисциплине «'+(meta.get('discipline') or '____________________')+'»'
        elif t == 'Фамилия Имя Отчество обучающегося': v = meta.get('author') or 'ФИО студента: ____________________'
        elif t.startswith('Направление подготовки'): v = ('Направление подготовки: '+meta['specialty']) if meta.get('specialty') else ''
        elif t.startswith('Направленность'): v = ('Профиль: '+meta['profile']) if meta.get('profile') else ''
        elif t == 'ученая степень, звание': v = meta.get('teacher_position','')
        elif t == 'Руководитель работы':
            teacher_seen = True
            v = t
        elif t == 'Автор работы':
            teacher_seen = False
            v = t
        elif 'И.О. Фамилия' in t: v = meta.get('teacher' if teacher_seen else 'author') or 'ФИО: ____________________'
        elif t.startswith('студент группы'): v = 'Группа '+(meta.get('group') or '____________')
        elif 'подпись' in t: v = 'Подпись: ________________'
        elif t == '[Город] — [Год]': v = (meta.get('city') or 'Город __________')+' — '+(meta.get('year') or 'Год ____')
        if v is not None: _text(p,v)
        if v:
            # Conservative 12 pt Cyrillic line capacity; keep the reference font,
            # and spend whitespace for wrapping rather than shrink user text.
            width = 32 if teacher_seen or 35 <= i <= 40 else 78
            extra_lines += max(0, len(textwrap.wrap(str(v), width=width)) - 1)
        # Preserve the template's spacer paragraphs, even for omitted optional fields.
        # Collapsing all empty lines moves the title and city far above their positions.
        if p.text.strip():
            p.paragraph_format.keep_together = True
        if v == title:
            p.style = d.styles['Title']
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                r.font.name = 'Times New Roman'; r.font.size = Pt(12); r.font.bold = False
    # Distribute the space reduction between the three original gaps. Short
    # fields leave the reference untouched. Long names/title keep city on page 1.
    groups = [spare_lines[:10], spare_lines[10:15], spare_lines[15:]]
    while extra_lines and any(groups):
        for group in groups:
            if group and extra_lines:
                p = group.pop()
                p._p.getparent().remove(p._p)
                extra_lines -= 1
    return d


def _independent_cover(kind, title, meta):
    """Fill each source-derived asset by stable slots, never by a shared cover."""
    d = Document(blank_path(kind))
    _ensure_styles(d)
    value = lambda key, label: meta.get(key) or label + ' __________'
    fields = {
        'school': value('school', 'Учебное заведение:'),
        'department': meta.get('department', ''),
        'org': meta.get('org', ''), 'institute': meta.get('institute', ''),
        'title': '«' + title.upper() + '»' if kind == 'report' else title,
        'author': ('Выполнил: ' if kind == 'notes' else '') + value('author', 'ФИО студента:'),
        'group': ('Выполнил: студент группы ' if kind == 'essay' else 'Группа ') + value('group', ''),
        'course': ('Выполнил студент ' if kind == 'report' else 'Студент ') + value('course', '') + ' курса',
        'teacher': ('Преподаватель: ' if kind == 'notes' else '') + value('teacher', 'ФИО руководителя:'),
        'teacher_position': ('Должность: ' if kind == 'notes' else '') + value('teacher_position', 'Степень, звание:'),
        'discipline': 'по дисциплине «' + value('discipline', '') + '»',
        'record_book': '№ зачётной книжки: ' + value('record_book', ''),
        'city': value('city', 'Город'),
        'year': value('year', 'Год'),
        'city_year': value('city', 'Город') + (' — ' if kind == 'essay' else ' ') + value('year', 'Год'),
    }
    extra_height = 0
    gaps = []
    for p in d.paragraphs:
        if p.style.name == 'Academic Gap':
            gaps.append(p)
            continue
        slots = [e.get(qn('w:name'), '').removeprefix('academic_slot_')
                 for e in p._p.findall(qn('w:bookmarkStart'))
                 if e.get(qn('w:name'), '').startswith('academic_slot_')]
        if slots:
            _text(p, fields[slots[0]])
        # Conservative line capacity at 14 pt; use spare spacing for long fields.
        width = 29 if p.paragraph_format.left_indent else 62
        lines = sum(max(1, len(textwrap.wrap(line, width=width)))
                    for line in p.text.split('\n'))
        extra_height += max(0, lines - 1) * 18
    remaining = extra_height
    # Spread reduction across the existing gaps so the visual hierarchy survives.
    while remaining > 0.5 and gaps:
        active = [p for p in gaps if (p.paragraph_format.space_after or Pt(0)).pt > 0]
        if not active:
            break
        share = remaining / len(active)
        for p in active:
            before = p.paragraph_format.space_after.pt
            reduction = min(before, share)
            p.paragraph_format.space_after = Pt(before - reduction)
            remaining -= reduction
    return d


def _field(p, instruction, cached=''):
    for typ, text in [('begin',None),('instr',instruction),('separate',None),('text',cached),('end',None)]:
        r = p.add_run()
        if typ in ('begin','separate','end'):
            e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'),typ)
        else:
            e = OxmlElement('w:instrText' if typ == 'instr' else 'w:t'); e.text=text
            e.set(qn('xml:space'),'preserve')
        r._r.append(e)


def _ensure_styles(d):
    for name in ('Normal', 'Title', 'Heading 1', 'Heading 2', 'Heading 3', 'TOC 1', 'TOC 2'):
        if name not in d.styles:
            st = d.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            if name != 'Normal': st.base_style = d.styles['Normal']


def _styles(d, covered):
    _ensure_styles(d)
    s=d.sections[0]
    s.page_width,s.page_height=Cm(21),Cm(29.7)
    s.left_margin,s.right_margin=Cm(3),Cm(1.5)
    s.top_margin=s.bottom_margin=Cm(2)
    for name in ('Normal','Title','Heading 1','Heading 2','Heading 3','TOC 1','TOC 2'):
        if name not in d.styles: continue
        st=d.styles[name]; st.font.name='Times New Roman';st.font.size=Pt(14);st.font.color.rgb=RGBColor(0,0,0)
        fonts=st.element.get_or_add_rPr().find(qn('w:rFonts'))
        for attr in list(fonts.attrib):
            if attr.rsplit('}',1)[-1].lower().endswith('theme'): del fonts.attrib[attr]
        for attr in ('ascii','hAnsi','eastAsia','cs'): fonts.set(qn('w:'+attr),'Times New Roman')
        st.font.italic=False
        for e in list(st.element.iter(qn('w:pBdr'))): e.getparent().remove(e)
        pf=st.paragraph_format;pf.space_before=Pt(0);pf.space_after=Pt(0);pf.line_spacing=1.5
        pf.first_line_indent=Cm(0 if name == 'Title' else 1.25);pf.widow_control=True
        pf.left_indent=Cm(0);pf.right_indent=Cm(0)
    for name in ('Heading 1','Heading 2','Heading 3'):
        st=d.styles[name];st.font.bold=True;st.paragraph_format.keep_with_next=True
        props=st.element.get_or_add_pPr()
        for node in list(props.findall(qn('w:outlineLvl'))):props.remove(node)
        level=OxmlElement('w:outlineLvl');level.set(qn('w:val'),str(int(name[-1])-1));props.append(level)
        st.paragraph_format.space_after=Pt(10)
    s.different_first_page_header_footer=covered
    fp=s.first_page_footer.paragraphs[0];fp.clear()
    p=s.footer.paragraphs[0];p.clear();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent=Cm(0)
    _field(p,' PAGE ')
    flag=OxmlElement('w:updateFields');flag.set(qn('w:val'),'true');d.settings.element.append(flag)


def _level(title):
    return 2 if re.match(r'^\d+\.\d+',title) else 1


def front_matter_pages(kind):
    return 2 if kind in ('referat', 'coursework') else 1 if kind in KINDS else 0


def volume_label(kind, pages):
    """Selected pages always mean body pages; front matter is additional and free."""
    prefix = 'Титульный + оглавление + ' if kind in ('referat', 'coursework') else 'Титульный + '
    return prefix + f'{pages} стр. текста'


def volume_hint(kind):
    extra = ' и оглавление' if kind in ('referat', 'coursework') else ''
    return (f'Выбери число страниц основного текста. Титульный лист{extra} — отдельно, '
            'без доплаты. Данные для титульного можно указать вместе с темой: '
            'учебное заведение, дисциплина, ФИО, группа, преподаватель, город и год.')


def delivery_notice(kind, requested, actual, total, credits):
    return (f'В документе: {volume_label(kind, actual)}. Всего в файле {total} стр.\n'
            + (f'Было выбрано {requested} стр. текста; фактически получилось {actual}.\n'
               if actual != requested else '')
            + f'Стоимость — {credits} кредитов. Титульный лист и оглавление без доплаты; '
              'дополнительные страницы текста не увеличивают цену.')


def page_guidance(kind, pages):
    front = front_matter_pages(kind)
    body = max(1, pages)
    words_per_page = 250 if kind == 'essay' else 220 if kind == 'notes' else 200
    compact = kind == 'referat' and pages <= 10
    return (
        f'\nВыбранный объём — {body} страниц основного текста. Дополнительно программа создаст '
        f'{front} служебных стр. (титульный лист и оглавление по жанру); они НЕ входят в выбранный объём. '
        f'Ожидаемый общий объём — {body + front} стр. На основной текст — {body} страниц, '
        f'ориентир {body * words_per_page} слов, включая таблицы, без реквизитов и оглавления. '
        'До ответа посчитай слова по разделам. Распредели объём между содержательными вопросами; '
        'не заменяй полноценное раскрытие темы короткой аннотацией. '
        'Введение и выводы краткие; около 80% текста отведи объяснениям, условиям и примерам. '
        'Не наращивай объём повторениями, общими фразами и описанием границ задания. '
        + ('Это короткий реферат: разделы размещаются последовательно, без обязательного '
           'нового листа для каждого. Сохрани все заданные пользователем темы. ' if compact else '')
    )


def inspect_layout(path, kind):
    from word_quality import word_pdf_layout
    return word_pdf_layout(path, front_matter_pages=front_matter_pages(kind))


def render(path, title, sections, kind, meta=None, add_table=None, requested_pages=None):
    if kind not in KINDS: raise ValueError('unknown_academic_kind')
    meta={k:str(v) for k,v in (meta or {}).items() if isinstance(v,(str,int)) and v}
    covered=kind in KINDS
    # The genre is already printed by the cover. Strip only the redundant
    # generated prefix, preserving the substantive topic and all user fields.
    title = re.sub(r'^\s*(?:Конспект|Реферат|Эссе|Доклад|Курсовая работа)\s+(?:по теме|на тему)\s*[:—-]?\s*',
                   '', title, flags=re.I).strip()
    if title.startswith('«') and title.endswith('»'):
        title = title[1:-1]
    d=_cover(kind,title,meta) if covered else Document()
    _styles(d,covered)
    blocks=[dict(b) for b in sections or [] if str(b.get('title','')).strip().casefold() not in
            ('содержание','оглавление','титульный лист','contents','table of contents')]
    if not covered:
        p=d.add_paragraph(title,style='Title');p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent=Cm(0);p.paragraph_format.space_after=Pt(12)
        for k in ('school','discipline','author','group'):
            if meta.get(k):
                p=d.add_paragraph(meta[k]);p.paragraph_format.first_line_indent=Cm(0)
    toc=kind in ('referat','coursework') and bool(blocks)
    compact = kind == 'referat' and requested_pages is not None and requested_pages <= 10
    if toc:
        p=d.add_paragraph('ОГЛАВЛЕНИЕ');p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.page_break_before = True
        p.paragraph_format.first_line_indent=Cm(0);p.paragraph_format.space_after=Pt(14)
        p.runs[0].bold=True
        for i,b in enumerate(blocks):
            head=str(b.get('title') or '').strip()
            if not head: continue
            p=d.add_paragraph(head+'\t');p.paragraph_format.first_line_indent=Cm(0)
            p.paragraph_format.left_indent=Cm(.6 if _level(head)==2 else 0)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(16.2),WD_TAB_ALIGNMENT.RIGHT,WD_TAB_LEADER.DOTS)
            _field(p,' PAGEREF study_'+str(i)+' \\h ')
    body_start_index = len(d.paragraphs)
    for i,b in enumerate(blocks):
        head=str(b.get('title') or '').strip()
        if head and kind != 'essay':
            level=_level(head)
            p=d.add_paragraph(head,style='Heading '+str(level))
            p.paragraph_format.page_break_before=(toc and level==1 and not compact) or (toc and i==0)
            if compact and i:
                p.paragraph_format.space_before=Pt(12)
            p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            start=OxmlElement('w:bookmarkStart');start.set(qn('w:id'),str(1000+i));start.set(qn('w:name'),'study_'+str(i))
            end=OxmlElement('w:bookmarkEnd');end.set(qn('w:id'),str(1000+i))
            p._p.insert(0,start);p._p.append(end)
        for line in str(b.get('content') or '').splitlines():
            if not line.strip():continue
            p=d.add_paragraph(line.strip());p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            if re.match(r'^(?:[-•*]|\d+[.)])\s',line):
                p.paragraph_format.first_line_indent=Cm(0)
        table=b.get('table')
        if table and add_table:
            first_table = len(d.tables)
            add_table(d,table)
            for t in d.tables[first_table:]:
                keep_small_table = len(t.rows) <= 6 and sum(len(c.text) for row in t.rows for c in row.cells) <= 800
                for row_index, row in enumerate(t.rows):
                    if keep_small_table:
                        trpr = row._tr.get_or_add_trPr()
                        if trpr.find(qn('w:cantSplit')) is None:
                            trpr.append(OxmlElement('w:cantSplit'))
                    for cell in row.cells:
                        tcpr=cell._tc.get_or_add_tcPr()
                        margins=tcpr.find(qn('w:tcMar'))
                        if margins is None:
                            margins=OxmlElement('w:tcMar');tcpr.append(margins)
                        for edge in ('top','left','bottom','right'):
                            node=margins.find(qn('w:'+edge))
                            if node is None:
                                node=OxmlElement('w:'+edge);margins.append(node)
                            node.set(qn('w:w'),'80');node.set(qn('w:type'),'dxa')
                        for p in cell.paragraphs:
                            p.paragraph_format.line_spacing=1
                            p.paragraph_format.first_line_indent=Cm(0)
                            if keep_small_table:
                                p.paragraph_format.keep_with_next = row_index < len(t.rows) - 1
                            for r in p.runs:r.font.name='Times New Roman';r.font.size=Pt(12);r.font.color.rgb=RGBColor(0,0,0)
    if covered and not toc and len(d.paragraphs) > body_start_index:
        d.paragraphs[body_start_index].paragraph_format.page_break_before = True
    d.core_properties.author='';d.core_properties.last_modified_by=''
    d.save(path)


def cache_page_references(docx_path, pdf_path):
    """Cache only academic PAGEREF results from the real export's heading outline.

    Keep every other DOCX package part byte-for-byte; do not save through Writer.
    Word still updates these fields when the document is edited later.
    """
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree
    from pypdf import PdfReader
    import os
    import tempfile
    with ZipFile(docx_path) as z:
        xml=z.read('word/document.xml')
        if b'PAGEREF study_' not in xml:return
        parts={n:z.read(n) for n in z.namelist()}
    root=etree.fromstring(xml)
    reader=PdfReader(pdf_path)
    entries=[]
    def walk(items):
        for item in items:
            if isinstance(item,list):walk(item)
            elif isinstance(item,dict) and item.get('/Title'):
                entries.append((str(item['/Title']),reader.get_destination_page_number(item)+1))
    walk(reader.outline)
    pages={}; cursor=0
    normalize=lambda value:' '.join(value.split())
    for mark in root.iter(qn('w:bookmarkStart')):
        name=mark.get(qn('w:name'),'')
        if not name.startswith('study_'):continue
        head=''.join(mark.getparent().itertext())
        head=''.join(n.text or '' for n in mark.getparent().iter(qn('w:t')))
        for index in range(cursor,len(entries)):
            if normalize(entries[index][0])==normalize(head):
                pages[name]=entries[index][1];cursor=index+1;break
        else:raise ValueError('academic_heading_missing_from_pdf_outline')
    for instruction in root.iter(qn('w:instrText')):
        m=re.fullmatch(r'\s*PAGEREF (study_\d+)\s+\\h\s*',instruction.text or '')
        if not m:continue
        if m[1] not in pages:raise ValueError('academic_page_reference_missing')
        run=instruction.getparent().getnext();separated=False
        while run is not None:
            field=run.find(qn('w:fldChar'))
            if field is not None:
                if field.get(qn('w:fldCharType'))=='separate':separated=True
                elif field.get(qn('w:fldCharType'))=='end':break
            elif separated:
                node=run.find(qn('w:t'))
                if node is not None:
                    node.text=str(pages[m[1]]);break
            run=run.getnext()
        else:raise ValueError('academic_field_result_missing')
    parts['word/document.xml']=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
    fd,temp=tempfile.mkstemp(suffix='.docx',dir=Path(docx_path).resolve().parent)
    os.close(fd)
    try:
        with ZipFile(temp,'w',ZIP_DEFLATED)as z:
            for name,content in parts.items():z.writestr(name,content)
        os.replace(temp,docx_path)
    finally:
        if os.path.exists(temp):os.remove(temp)
