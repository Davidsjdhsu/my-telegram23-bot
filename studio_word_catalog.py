"""Word form definitions derived from the actual runtime templates."""
import re
from academic_templates import KINDS, META_LABELS, render as academic_render
from fixed_word_templates import CATALOG, render as fixed_render, scope_issue
from car_sale_template import FIELDS, render as car_render

STUDY = {'referat':'Реферат','coursework':'Курсовая работа','report':'Доклад','essay':'Эссе','notes':'Конспект'}
# Only slots present in each approved cover are exposed.
COVER_FIELDS = {
 'referat': ['org','school','department','discipline','author','specialty','profile','group','teacher','teacher_position','city','year'],
 'coursework': ['org','school','department','author','specialty','profile','group','teacher','teacher_position','city','year'],
 'essay': ['school','department','author','group','teacher','city','year'],
 'report': ['org','school','institute','department','course','author','group','teacher','teacher_position','city','year'],
 'notes': ['school','department','discipline','course','author','group','record_book','teacher','teacher_position','city','year'],
}
PHYSICAL = {'dkp_car','rent_flat_ru','claim_replace_ru','dkp','rent','proxy','statement','act','loan','claim','consent','marriage_contract','gift','lawsuit','alimony'}

def catalog():
    out=[]
    for kind,title in STUDY.items():
        out.append(dict(kind=kind,title=title,category='study',cover=True,fixed=False,
            fields={k:META_LABELS[k] for k in COVER_FIELDS[kind]},
            notice='Титульный лист — по выбранному русскому учебному образцу. Язык основного текста можно выбрать. Титульный и оглавление не оплачиваются.'))
    for kind,spec in CATALOG.items():
        out.append(dict(kind=kind,title='Аренда квартиры' if kind=='rent_flat_ru' else spec['title'],
            category='physical' if kind in PHYSICAL else 'entity',cover=False,fixed=True,
            fields=spec['fields'],notice=spec['scope']+' '+spec.get('notice','')))
    out.append(dict(kind='dkp_car',title='Купля-продажа автомобиля',category='physical',cover=False,fixed=True,
        fields=FIELDS,notice='Договор РФ между физическими лицами. Неизвестные поля остаются незаполненными.'))
    out.append(dict(kind='doc',title='Обычный документ',category='study',cover=False,fixed=False,
        fields={'org':'Организация','sign':'Подпись / автор'},notice='Первая страница показывает оформление. Основной текст появится после сборки.'))
    out.append(dict(kind='resume',title='Резюме',category='physical',cover=False,fixed=False,
        fields={'name':'ФИО','position':'Желаемая должность','phone':'Телефон','email':'Электронная почта','city':'Город'},
        notice='Предпросмотр первой страницы. Опыт и навыки укажите в исходных данных.'))
    return out

SPECS={r['kind']:r for r in catalog()}

def normalize(body, *, preview=False):
    if not isinstance(body,dict):raise ValueError('Некорректные данные документа.')
    kind=body.get('kind')
    if not isinstance(kind,str) or kind not in SPECS:raise ValueError('Выберите тип документа.')
    spec=SPECS[kind]
    fields=body.get('fields',{})
    if not isinstance(fields,dict) or set(fields)-set(spec['fields']):raise ValueError('Поля не соответствуют типу документа.')
    clean={}
    for k,v in fields.items():
        limit=200 if spec['cover'] else 1500
        if not isinstance(v,str) or len(v)>limit or re.search(r'[\x00-\x08\x0b\x0c\x0e-\x1f]',v):raise ValueError('Слишком длинное или некорректное поле.')
        if spec['fixed'] and ('[' in v or ']' in v):raise ValueError('Уберите квадратные скобки из заполненных полей.')
        clean[k]=v.strip()
    topic=body.get('topic','')
    content=body.get('content','')
    if not isinstance(topic,str) or len(topic)>300 or not isinstance(content,str) or len(content)>10000:raise ValueError('Сократите тему или исходный текст.')
    if not preview and not spec['fixed'] and not topic.strip():raise ValueError('Укажите тему документа.')
    from studio_settings import CONTENT_LANGUAGES
    lang=body.get('content_lang','ru')
    if not isinstance(lang,str) or lang not in CONTENT_LANGUAGES:raise ValueError('Выберите язык.')
    if spec['fixed'] and lang!='ru':raise ValueError('Выбранный российский шаблон заполняется на русском языке.')
    pages=body.get('pages',6)
    if not spec['fixed'] and kind!='resume' and (type(pages) is not int or pages not in (2,6,10,16,20,30,40)):raise ValueError('Выберите объём документа.')
    if spec['fixed'] or kind=='resume':pages=2
    value=dict(kind=kind,topic=topic.strip(),content=content.strip(),fields=clean,content_lang=lang,pages=pages)
    if spec['fixed']:
        # Structured fields are authoritative; do not silently ignore extra prose.
        if content.strip():raise ValueError('Внесите сведения в отдельные поля выбранного документа.')
        if kind in CATALOG:
            issue=scope_issue({'word_kind':kind,'topic':'','extra': '\n'.join(spec['fields'][k]+': '+v for k,v in clean.items())})
            if issue:raise ValueError(issue)
        if kind=='dkp_car':
            from car_sale_template import scope_issue as car_scope_issue
            issue=car_scope_issue({'topic':'','extra':'\n'.join(clean.values())})
            if issue:raise ValueError(issue)
    return value

def price(payload):return payload['pages']*5

def render_preview(path,payload,generic_render=None):
    kind=payload['kind'];fields=payload['fields'];title=payload['topic'] or 'Тема работы'
    if kind in KINDS:academic_render(path,title,[],kind,fields,requested_pages=payload['pages'])
    elif kind in CATALOG:fixed_render(path,kind,fields)
    elif kind=='dkp_car':car_render(path,fields)
    elif generic_render:
        sections=[{'title':'Основной текст','content':'Здесь будет текст по вашей теме.'}]
        generic_render(str(path),title,sections,kind,fields)
    else:raise ValueError('Предпросмотр недоступен.')
