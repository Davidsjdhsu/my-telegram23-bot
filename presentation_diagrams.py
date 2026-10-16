"""Shared, editable explanatory diagrams. No image generation or external services."""
import re

DIAGRAM_INSTRUCTION = '''
Объяснение механизма должно содержать все необходимые переходы: что входит, что меняется,
почему меняется и куда попадает результат. Для цикла объясни также возврат к исходному состоянию.
Не выбрасывай ключевой этап ради краткости. Объединяй соседние этапы в один блок, сохраняя причины.
Пиши естественно: называй конкретное действие предмета. Избегай выдуманных выражений,
например «компрессор гонит круг». Не копируй этот пример в другие темы.
Для каждого слайда добавь visual_kind: photo или technical_diagram.
photo — внешний вид, интерьер, люди, наблюдаемое действие; это иллюстрация, не техническое доказательство.
technical_diagram — внутреннее устройство, отдельные функциональные детали, пути вещества,
энергии, сигнала, анатомия или научный механизм. Не генерируй для них фотореалистичные псевдодетали.
Крупный план отдельного узла, трубок, схемы соединений, разреза или рабочего органа,
который должен объяснить функцию, — technical_diagram, даже если этот узел виден снаружи.
Фото целого прибора в комнате допустимо; фото его детали вместо объяснения её работы — нет.
Вместо них обязательно заполни diagram. Это смысловая схема, а не точный инженерный чертёж.
Для новичка или школьника замени профессиональный жаргон обычными словами либо кратко объясни
термин при первом употреблении. Не придумывай названия производственных линий для бытового процесса.
Если явно запрошена схема, хотя бы один соответствующий слайд обязан содержать diagram.
Формат diagram: {"kind":"sequence|cycle|structure", "nodes":[{"label":"название до 32 символов","text":"объяснение до 100 символов"}]}.
От 2 до 5 узлов; для cycle минимум 3. Общий текст diagram до 440 символов.
sequence: узлы в порядке этапов, стрелки соединяют соседей.
cycle: то же и стрелка от последнего к первому; обязательно объясни обратный переход.
structure: первый узел — целое, остальные — его части, соединения без стрелок, не порядок действий.
Не превращай альтернативы в последовательность, а незамкнутый процесс — в цикл.
Определение вещества или термина не является этапом: включи короткое пояснение в текст
первого настоящего этапа, не создавай для определения отдельный узел со стрелкой.
Подписи и объяснения полностью раскрывают смысл, content передаёт те же факты.
На слайде с diagram chart=null, image_prompt="", designer={"layout":"auto","blocks":[]}.
diagram имеет приоритет над декоративным макетом в обычной и дизайнерской презентации.
На остальных слайдах diagram=null. Не заменяй схемой личные фотографии и фотоальбомы.
'''


def validate_diagram(raw):
    if raw is None:
        return None
    if not isinstance(raw, dict) or raw.get('kind') not in ('sequence', 'cycle', 'structure'):
        raise ValueError('Invalid diagram kind')
    nodes = raw.get('nodes')
    if not isinstance(nodes, list) or not (3 if raw['kind']=='cycle' else 2) <= len(nodes) <= 5:
        raise ValueError('Diagram requires 2–5 nodes; cycle at least 3')
    result=[]
    for node in nodes:
        if not isinstance(node, dict): raise ValueError('Invalid diagram node')
        clean={}
        for key, limit in (('label',32),('text',100)):
            value=node.get(key)
            if not isinstance(value,str) or not value.strip() or len(value.strip())>limit:
                raise ValueError('Shorten diagram '+key)
            clean[key]=value.strip()
        result.append(clean)
    if sum(len(n['label'])+len(n['text']) for n in result)>440:
        raise ValueError('Shorten diagram without omitting essential stages')
    if len({n['label'].casefold() for n in result})!=len(result):
        raise ValueError('Diagram labels must be distinct')
    return {'kind':raw['kind'],'nodes':result}


def effective_diagram(item):
    if item.get('diagram') is not None:
        return validate_diagram(item['diagram'])
    # Preserve old drafts: their real process blocks now get actual arrows, too.
    design=item.get('designer') or {}
    if design.get('layout')=='process':
        return {'kind':'sequence','nodes':[{'label':b['heading'],'text':b['text']} for b in design['blocks']]}
    return None


def diagram_text(item):
    d=effective_diagram(item)
    return '\n'.join(n['label']+': '+n['text'] for n in d['nodes']) if d else ''


def diagram_relations(item):
    d=effective_diagram(item)
    if not d:return []
    pairs=([(0,i) for i in range(1,len(d['nodes']))] if d['kind']=='structure'
           else [(i,i+1) for i in range(len(d['nodes'])-1)])
    if d['kind']=='cycle':pairs.append((len(d['nodes'])-1,0))
    return [{'from':a+1,'to':b+1,'meaning':'part_of' if d['kind']=='structure' else 'next_stage'} for a,b in pairs]


def diagram_requested(data):
    text=' '.join(str(data.get(k) or '') for k in ('topic','extra')).lower()
    text=re.sub(r'(?:без|no|without)\s+(?:схем\w*|diagrams?)','',text)
    return bool(re.search(r'\b(?:схем\w*|diagram\w*|flowchart\w*)\b',text))


def build_diagram_scene(item,number,total,lang='ru',style=None,colors=None,asset=None):
    from presentation_designer import text_node, validate_scene, palette
    from presentation_profiles import resolved_profile, readable_color
    from presentation_catalog import TEMPLATES
    from presentation_style_direction import decorate_fallback, has_custom_palette
    d=effective_diagram(item)
    if not d:raise ValueError('Missing diagram')
    design=item.get('_designer_design',{})
    c=resolved_profile(design) if design.get('family') else palette(item.get('_designer_background')=='dark')
    if colors:
        hx=lambda v:'#'+''.join(f'{x:02X}' for x in v)
        c={**c,'bg':hx(colors['bg']),'ink':hx(colors['ink']),'body':hx(colors['mid']),'accent':hx(colors['ink'])}
    bg=c['bg'];ink=readable_color(c['ink'],bg);accent=readable_color(c['accent'],bg)
    scene={'background':bg,'layout':'diagram' if item.get('diagram') else 'process','composition':d['kind'] if item.get('diagram') else 'process','nodes':[],
           'notes':'Editable explanatory diagram. Connections represent '+d['kind']+'.\n'+diagram_text(item)}
    nodes=scene['nodes']
    def text(value,box,role='body',size=28,minimum=22):
        n=text_node(value,box,ink,size,minimum,role!='body',role,'Arial')
        if lang=='ar':n['rtl']=True
        nodes.append(n)
    text(item['title'],(64,36,830 if asset else 1152,135),'title',56,32)
    if asset:
        from presentation_visuals import fit_rect
        nodes.append({'type':'image','path':asset['path'],'box':list(fit_rect(asset['width'],asset['height'],(930,30,286,140))), 'origin':asset['origin'],'radius':0,'crop':False})
        scene['notes']+='\nSource: user-supplied photograph; full frame preserved.'
    count=len(d['nodes']);gap=20;rh=min(112,(460-gap*(count-1))/count);start=190
    for i,node in enumerate(d['nodes']):
        y=start+i*(rh+gap)
        # Native outline denotes one semantic node; all text remains editable.
        nodes.append({'type':'photo_frame','box':[150,y,1066,rh],'color':accent,'width':1.2,'radius':14 if style=='designer_soft' else 0})
        # A five-node diagram has 76px rows. Two lines at the readable minimum
        # need 58.68px; the old 58px inner box rejected valid two-line labels.
        text(node['label'],(170,y+4,285,rh-8),'subheading',30,22)
        text(node['text'],(480,y+4,715,rh-8),'body',28,22)
    def edge(x1,y1,x2,y2,arrow=True):
        nodes.append({'type':'diagram_edge','box':[min(x1,x2),min(y1,y2),max(1,abs(x2-x1)),max(1,abs(y2-y1))],
                      'points':[x1,y1,x2,y2],'color':accent,'arrow':arrow})
    if d['kind']=='structure':
        mid=start+rh/2
        edge(150,mid,110,mid,False)
        edge(110,mid,110,start+(count-1)*(rh+gap)+rh/2,False)
        for i in range(1,count):
            y=start+i*(rh+gap)+rh/2;edge(110,y,150,y,False)
    else:
        for i in range(count-1):
            y=start+i*(rh+gap)+rh;edge(680,y+2,680,y+gap-2)
        if d['kind']=='cycle':
            a=start+rh/2;b=start+(count-1)*(rh+gap)+rh/2
            edge(150,b,100,b,False);edge(100,b,100,a,False);edge(100,a,150,a)
    text(f'{number:02} / {total:02}',(64,682,1152,28),'footer',16,14)
    if style and design.get('template_id') in TEMPLATES:
        scene=decorate_fallback(scene,item,asset,number,style,has_custom_palette(design))
        actual_ink=next(n['color'] for n in scene['nodes'] if n.get('role')=='title')
        for n in scene['nodes']:
            if n['type'] in ('photo_frame','diagram_edge'):n['color']=actual_ink
    validate_scene(scene)
    return scene
