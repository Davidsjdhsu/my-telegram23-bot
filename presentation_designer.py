"""Opt-in technology layouts. Pure scene planning plus an adapter for the existing bot.

No network, credentials, billing or Telegram operations. Dimensions are CSS pixels.
The same scene can be rendered by the bot or the review/export tool.
"""
import copy
import math
import re
from functools import lru_cache
from pathlib import Path

from PIL import ImageFont
from presentation_visuals import AI_LABEL, fit_rect, set_slide_notes

STYLES = {'designer_soft', 'designer_strict'}
FONT = 'Arial'
LAYOUTS = {'auto', 'photo', 'statement', 'process', 'comparison', 'sections'}


class DesignerValidationError(ValueError):
    pass


def clean(value, limit, field, optional=False):
    if optional and value in (None, ''):
        return ''
    if not isinstance(value, str) or not value.strip() or len(value) > limit:
        raise DesignerValidationError(f'Invalid {field}')
    return value.strip()


def content_instruction():
    return '''Дизайнерское оформление: крупные заголовки, короткие смысловые блоки.
Для каждого содержательного слайда добавь designer:
{"layout":"auto|photo|statement|process|comparison|sections",
"blocks":[{"heading":"короткий подзаголовок","text":"текст"}]}.
Для process нужно 3–4 последовательных шага, comparison ровно 2 стороны,
sections 2–4 смысловых блока. Заголовок блока до 36 символов, текст до 120.
Вся сумма текстов и заголовков блоков до 440 символов. Для auto/photo/statement
blocks=[]; тогда используется content. content обязателен и содержит тот же смысл,
что блоки, для текстовой копии. Не добавляй фактов, процентов или чисел ради дизайна.
Для statement content до 220 символов. Для photo content до 320 символов.
Не менее половины слайдов планируй с предметными фото (photo/auto), включая обложку.
statement используй не более одного раза. process — только реальная последовательность
действий или этапов; для перечня болезней, свойств или мер профилактики используй sections.
Чередуй фото, схемы и сравнения по смыслу; не делай все слайды одинаковыми.
Если план содержит сопоставление, используй comparison, для последовательности действий — process.
Для презентации от 6 слайдов выбери хотя бы один подходящий содержанию структурный макет
(comparison/sections/process), сохраняя фото минимум на половине слайдов. Не выдумывай
этапы или противопоставления ради оформления. Данные пользователя сохраняй. При длинном материале выбирай auto. Все поля и
количество слайдов из основного задания остаются обязательными.'''


def plan_designer_images(slides, style):
    """Keep true diagrams; promote supporting-photo layouts in a text-heavy plan."""
    for i, item in enumerate(slides):
        item['_designer_style'] = style
        item['_designer_cover'] = i == 0
        item['_skip_ai_image'] = bool(item.get('chart')) or (i > 0 and item.get('designer', {}).get('layout') in (
            'process', 'comparison', 'sections', 'statement'))
    target = min(6, math.ceil(len(slides)/2))
    current = sum(not s['_skip_ai_image'] for s in slides)
    # Choose alternate positions first; preserve all structured text when rendered with a photo.
    candidates = list(range(2, len(slides), 2)) + list(range(1, len(slides), 2))
    for i in candidates:
        s = slides[i]
        if current >= target:
            break
        if s['_skip_ai_image'] and not s.get('chart') and s.get('designer', {}).get('layout') in ('statement', 'sections'):
            s['_skip_ai_image'] = False
            current += 1


def validate_designer(item):
    """Unknown/malformed layouts fail early, rather than silently losing text."""
    raw = item.get('designer')
    if raw is None:
        return {'layout': 'auto', 'blocks': []}
    if not isinstance(raw, dict) or raw.get('layout') not in LAYOUTS:
        raise DesignerValidationError('Unknown designer layout')
    layout = raw['layout']
    blocks = raw.get('blocks', [])
    if not isinstance(blocks, list):
        raise DesignerValidationError('Blocks must be a list')
    count = {'process': (3, 4), 'comparison': (2, 2), 'sections': (2, 4)}.get(layout, (0, 0))
    if not count[0] <= len(blocks) <= count[1]:
        raise DesignerValidationError('Layout and block count disagree')
    result = []
    for block in blocks:
        if not isinstance(block, dict):
            raise DesignerValidationError('Invalid block')
        result.append({'heading': clean(block.get('heading'), 36, 'block heading'),
                       'text': clean(block.get('text'), 120, 'block text')})
    if sum(len(b['heading']) + len(b['text']) for b in result) > 440:
        raise DesignerValidationError('Shorten blocks before rendering')
    return {'layout': layout, 'blocks': result}


@lru_cache(maxsize=16)
def metric_font(size, bold=False):
    # DejaVu is already installed on Render; its wider metrics reserve breathing
    # room for the widely available Arial used in the editable PowerPoint file.
    name = 'DejaVuSans-Bold.ttf' if bold else 'DejaVuSans.ttf'
    locations = [Path('/usr/share/fonts/truetype/dejavu') / name,
                 Path('/Library/Fonts') / name,
                 Path(__file__).parent / 'fonts' / name]
    # Bundled review runtime. On Render, the first system path is used.
    runtime = Path.home() / '.cache/codex-runtimes/codex-primary-runtime/dependencies/native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/Resources/fonts/truetype'
    locations.append(runtime / name)
    for path in locations:
        if path.is_file():
            return ImageFont.truetype(str(path), size)
    try:
        return ImageFont.truetype(name, size)
    except OSError:
        return ConservativeMetrics(size)


class ConservativeMetrics:
    """Reserve space when the host lacks DejaVu, without downloading fonts at runtime."""
    def __init__(self, size): self.size = size
    def getlength(self, text):
        return self.size * sum(1.1 if ord(c) > 0x2e7f or c in 'WMЖШЩЮ' else
                               .38 if c.isspace() else .76 for c in text)


def measured(font, text, size):
    if any(ord(c) > 0x2e7f for c in text):
        return sum(size * 1.05 if ord(c) > 0x2e7f else font.getlength(c) for c in text)
    return font.getlength(text)


def wrap(text, width, size, bold=False):
    font = metric_font(size, bold)
    lines = []
    for paragraph in text.split('\n'):
        current = ''
        for word in paragraph.split():
            candidate = f'{current} {word}'.strip()
            if measured(font, candidate, size) <= width:
                current = candidate
                continue
            if current:
                lines.append(current)
            current = ''
            # Long URLs / unspaced scripts still must fit without truncation.
            for letter in word:
                if current and measured(font, current + letter, size) > width:
                    lines.append(current)
                    current = ''
                current += letter
        lines.append(current)
    return '\n'.join(lines)


def text_node(text, box, color, size=30, minimum=28, bold=False, role='body'):
    x, y, width, height = box
    for chosen in range(size, minimum - 1, -1):
        # Reduce the font before breaking ordinary words across lines.
        font = metric_font(chosen, bold)
        words = str(text).split()
        if any(1 < len(w) <= 40 and not any(ord(c) > 0x2e7f for c in w)
               and measured(font, w, chosen) > width - 4 for w in words):
            continue
        wrapped = wrap(str(text), width - 4, chosen, bold)
        lines = len(wrapped.split('\n'))
        if lines * chosen * 1.22 + 5 <= height:
            return {'type': 'text', 'text': wrapped, 'box': list(box), 'color': color,
                    'size': chosen, 'font': FONT, 'bold': bold, 'role': role}
    raise DesignerValidationError(f'Shorten {role}: text does not fit at a readable size')


def is_designer(data):
    return data.get('theme_name') in STYLES and not data.get('custom_colors')


def palette(dark):
    return {'bg': '#080B10' if dark else '#FFFFFF',
            'ink': '#FFFFFF' if dark else '#080B10',
            'body': '#DCE5F6' if dark else '#28323F',
            'accent': '#80F0E1' if dark else '#244DFF',
            'muted': '#B8CBCF' if dark else '#58616F'}


GRADIENT = {'angle': 0, 'stops': [(0, '#080B10', 1), (.45, '#14225B', 1),
                                 (.8, '#234DEC', 1), (1, '#32B8CC', 1)]}


def build_scene(item, asset, number, total, lang='ru', style='designer_soft'):
    """Plan the complete slide before mutating a PPTX; never truncate content."""
    title = clean(item.get('title'), 160, 'title')
    body = clean(item.get('content'), 480, 'body', optional=number == 1)
    design = validate_designer(item)
    layout, blocks = design['layout'], design['blocks']
    if not 1 <= number <= total: raise DesignerValidationError('Invalid slide number')
    if asset and (not Path(asset['path']).is_file() or asset['width'] <= 0 or asset['height'] <= 0):
        raise DesignerValidationError('Invalid image asset')
    if asset and blocks:
        # User photos take priority; preserve all structured text when using a photo layout.
        body = '\n\n'.join(b['heading'] + '\n' + b['text'] for b in blocks)
        blocks, layout = [], 'photo'
    layout = 'cover' if number == 1 else 'chart' if item.get('chart') else 'photo' if asset else layout
    if layout == 'photo' and not asset: layout = 'auto'
    radius = 36 if style == 'designer_soft' else 0
    dark = layout in ('cover', 'statement') or item.get('_designer_background') == 'dark'
    c = palette(dark)
    accent = item.get('_designer_accent')
    if isinstance(accent, str) and re.fullmatch(r'#[0-9a-fA-F]{6}', accent):
        # Keep semantic accent text readable on the chosen background.
        rgb = [int(accent[i:i+2],16)/255 for i in (1,3,5)]
        luminance = sum(v*w for v,w in zip(rgb,(.2126,.7152,.0722)))
        if (dark and luminance >= .45) or (not dark and luminance <= .65):
            c['accent'] = accent
    
    scene = {'background': '#F4F6FA' if layout in ('cover','statement') else c['bg'], 'layout': layout,
             'notes': 'AlixDocBot. ' + style, 'nodes': []}
    nodes = scene['nodes']

    def txt(value, box, role='body', size=30, minimum=26, color=None, bold=False):
        if value: nodes.append(text_node(value, box, color or c['body'], size, minimum, bold, role))

    def panel(box, gradient=None, fill=None):
        nodes.append({'type':'panel','box':list(box),'radius':radius,
                      'gradient':gradient,'fill':fill or c['bg']})

    def photo(box, cover=False):
        # Preserve complete source photographs, including evidence at the edges.
        if asset.get('origin') == 'user':
            panel(box, fill='#EAF0F7')
            x,y,w,h = box
            fitted=fit_rect(asset['width'],asset['height'],(x+16,y+16,w-32,h-32))
            nodes.append({'type':'image','path':str(asset['path']),'box':list(fitted),
                          'origin':'user','radius':0,'crop':False})
        else:
            source_ratio = asset['width'] / asset['height']
            target_ratio = box[2] / box[3]
            retained = min(source_ratio / target_ratio, target_ratio / source_ratio)
            crop = cover or retained >= .72
            image_box = box if crop else fit_rect(asset['width'], asset['height'], box)
            nodes.append({'type':'image','path':str(asset['path']),'box':list(image_box),
                          'origin':'ai','radius':radius,'crop':crop})

    def heading(box=(64,72,1152,214), size=76, minimum=42):
        txt(title,box,'title',size,minimum,c['ink'])

    if layout == 'cover':
        panel((20,20,1240,680),GRADIENT)
        if asset and asset.get('origin') != 'user':
            photo((20,20,1240,680),True)
            # Native overlay guarantees text contrast even if an AI image is too bright.
            panel((20,20,1240,680),{'angle':0,'stops':[(0,'#080B10',.97),
                  (.58,'#080B10',.84),(1,'#080B10',.15)]})
            heading((64,124,1056,310),112,40)
        elif asset:
            try:
                node = text_node(title,(64,112,600,300),c['ink'],80,42,False,'title')
            except DesignerValidationError:
                # Preserve long confirmed titles alongside a complete user photograph.
                node = text_node(title,(64,64,860,310),c['ink'],64,40,False,'title')
                photo((992,64,224,166))
            else:
                photo((710,100,500,510))
            nodes.append(node)
        else: heading((64,124,1080,310),112,40)
        subtitle = item.get('subtitle') or item.get('point') or body
        if subtitle:
            subtitle = clean(subtitle,180,'subtitle')
            # Long descriptions get a full-width band below the photo/title area.
            try:
                subtitle_node = text_node(subtitle,(68,490,600 if asset else 1050,145),c['body'],32,26)
            except DesignerValidationError:
                if asset and asset.get('origin') == 'user':
                    # Rebuild the user cover so the band cannot overlap its photograph.
                    nodes[:] = nodes[:1]
                    heading((64,64,860,310),64,40)
                    photo((992,64,224,166))
                subtitle_node = text_node(subtitle,(68,444,1136,204),c['body'],32,26)
            nodes.append(subtitle_node)
    elif layout == 'photo':
        portrait = asset['width'] / asset['height'] < .9
        try:
            if portrait and len(body) <= 320 and len(title) <= 85:
                title_node = text_node(title,(608,84,608,202),c['ink'],62,40,False,'title')
                body_node = text_node(body,(608,322,608,318),c['body'],30,26)
                nodes.extend([title_node,body_node])
                photo((64,64,480,584))
            else:
                title_node = text_node(title,(64,54,1152,164),c['ink'],68,40,False,'title')
                body_node = text_node(body,(64,268,520,370),c['body'],30,26)
                nodes.extend([title_node,body_node])
                photo((640,242,576,406))
        except DesignerValidationError:
            # A wide copy column still reserves a substantial photograph, never a tiny thumbnail.
            heading((64,54,1152,164),60,40)
            txt(body,(64,250,680,398),size=30,minimum=24)
            photo((800,250,416,398))
    elif layout == 'process':
        heading((64,60,1152,210),72,42)
        panel((32,306,1216,350),GRADIENT)
        count=len(blocks); step_w=(1152-(count-1)*32)/count
        for i,b in enumerate(blocks):
            x=64+i*(step_w+32)
            txt(f'{i+1:02}',(x,332,step_w,65),'step',48,48,'#80F0E1')
            txt(b['heading'],(x,407,step_w,87),'subheading',32,26,'#FFFFFF')
            txt(b['text'],(x,505,step_w,132),size=27,minimum=24,color='#FFFFFF')
            if i < count-1:
                nodes.append({'type':'connector','box':[x+82,370,step_w-67,2],'color':'#80F0E1'})
    elif layout in ('comparison','sections'):
        heading((64,60,1152,190),72,42)
        # Three equal cards, rather than a 2x2 grid with an empty fourth cell.
        cols=3 if len(blocks)==3 else 2
        rows=math.ceil(len(blocks)/cols); row_h=376/rows
        col_w=1152/cols; text_w=col_w-48
        panel((32,250,1216,408),fill='#15202E' if dark else '#F0F4FC')
        for i,b in enumerate(blocks):
            x=64+(i%cols)*col_w; y=270+(i//cols)*row_h
            head_h=110 if cols==3 else 70
            txt(b['heading'],(x,y,text_w,head_h),'subheading',36,26,c['accent'])
            txt(b['text'],(x,y+head_h+4,text_w,row_h-head_h-22),size=29,minimum=24)
    elif layout == 'statement':
        panel((20,20,1240,680),GRADIENT)
        heading((64,72,1110,212),80,44)
        txt(body,(180,330,1000,300),size=44,minimum=28)
    elif layout == 'chart':
        heading((64,58,870 if asset else 1152,158),60,42)
        txt(body,(64,260,400,370),size=30,minimum=24)
        nodes.append({'type':'chart','box':[510,250,706,380],'data':copy.deepcopy(item['chart']),'dark':False})
        if asset: photo((980,56,236,152))
    else:
        heading((64,72,1120,208),76,42)
        txt(body,(180,330,1020,310),size=38,minimum=26)

    label=f'{number:02} / {total:02}'
    if asset and asset.get('origin') == 'ai':
        label += '    ' + AI_LABEL.get(lang,AI_LABEL['en'])
        scene['notes'] += '\nImage: AI-generated illustration, not documentary evidence.'
    elif asset: scene['notes'] += '\nImage: supplied by user. Full frame preserved.'
    # Keep the original body available verbatim as speaker notes, too.
    scene['notes'] += '\n' + (item.get('content') or item.get('point') or '')
    txt(label,(64,672,1120,25),'footer',14,14,c['muted'])
    for node in nodes:
        if node.get('role') == 'title':
            lines=node['text'].split('\n')
            if layout == 'cover' and len(lines) == 2 and not (asset and asset.get('origin') == 'user'):
                if measured(metric_font(node['size']),lines[1],node['size']) + 116 < node['box'][2]:
                    node['stagger'] = 116
            elif not dark and len(lines) > 1:
                node['last_line_color'] = c['accent']
        if node['type'] == 'text' and lang == 'ar': node['rtl'] = True
    validate_scene(scene)
    return scene


def validate_scene(scene):
    for node in scene['nodes']:
        x,y,w,h=node['box']
        if not all(math.isfinite(v) for v in (x,y,w,h)) or min(x,y)<0 or min(w,h)<=0 or x+w>1280.1 or y+h>720.1:
            raise DesignerValidationError('Element outside slide')
    return scene


def render_scene(slide, scene, chart_renderer=None):
    """Existing bot runtime adapter; no new runtime dependency."""
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_ANCHOR
    from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.oxml.ns import qn
    validate_scene(scene)
    bg=slide.background.fill;bg.solid();bg.fore_color.rgb=RGBColor.from_string(scene['background'][1:])

    def rounded(shape,radius,w,h):
        geom=shape._element.spPr.find(qn('a:prstGeom'))
        if geom is None:
            geom=OxmlElement('a:prstGeom');shape._element.spPr.insert(1,geom)
        geom.set('prst','roundRect' if radius else 'rect')
        for child in list(geom): geom.remove(child)
        av=OxmlElement('a:avLst');geom.append(av)
        if radius:
            gd=OxmlElement('a:gd');gd.set('name','adj');gd.set('fmla','val '+str(round(min(.5,radius/min(w,h))*100000)));av.append(gd)

    def fill(shape,node):
        f=shape.fill
        if not node.get('gradient'):
            f.solid();f.fore_color.rgb=RGBColor.from_string(node['fill'][1:]);return
        g=node['gradient'];f.gradient()
        gs_list=f._xPr.find(qn('a:gradFill')).find(qn('a:gsLst'))
        for child in list(gs_list):gs_list.remove(child)
        for offset,color,opacity in g['stops']:
            gs=OxmlElement('a:gs');gs.set('pos',str(round(offset*100000)))
            rgb=OxmlElement('a:srgbClr');rgb.set('val',color[1:])
            alpha=OxmlElement('a:alpha');alpha.set('val',str(round(opacity*100000)));rgb.append(alpha)
            gs.append(rgb);gs_list.append(gs)
        f.gradient_angle=g['angle']

    for node in scene['nodes']:
        px,py,pw,ph=node['box'];x,y,w,h=[v/96 for v in node['box']]
        if node['type']=='panel':
            shape=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if node['radius'] else MSO_SHAPE.RECTANGLE,
                                        Inches(x),Inches(y),Inches(w),Inches(h))
            shape.name='Designer background';shape.line.fill.background()
            rounded(shape,node['radius'],pw,ph);fill(shape,node)
            shape._element.spPr.append(OxmlElement('a:effectLst'))
        elif node['type']=='text':
            shape=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));shape.name='Designer '+node['role']
            tf=shape.text_frame;tf.word_wrap=False
            tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0;tf.vertical_anchor=MSO_ANCHOR.TOP
            for i,line in enumerate(node['text'].split('\n')):
                p=tf.paragraphs[0] if not i else tf.add_paragraph();p.text=line
                p.font.name=node['font'];p.font.size=Pt(node['size']*.75);p.font.bold=node['bold']
                p.font.color.rgb=RGBColor.from_string(node['color'][1:]);p.line_spacing=Pt(node['size']*.75*1.22)
                p.space_after=p.space_before=Pt(0)
                if i == len(node['text'].split('\n'))-1 and node.get('last_line_color'):
                    p.font.color.rgb=RGBColor.from_string(node['last_line_color'][1:])
                if i == 1 and node.get('stagger'):
                    p._p.get_or_add_pPr().set('marL',str(Inches(node['stagger']/96)))
                if node.get('rtl'):
                    from pptx.enum.text import PP_ALIGN
                    p.alignment=PP_ALIGN.RIGHT;p._p.get_or_add_pPr().set('rtl','1')
        elif node['type']=='image':
            shape=slide.shapes.add_picture(node['path'],Inches(x),Inches(y),width=Inches(w),height=Inches(h))
            shape.name='User photograph' if node['origin']=='user' else 'AI illustration'
            if node.get('crop'):
                iw,ih=shape.image.size;source_ratio=iw/ih;target_ratio=w/h
                if source_ratio>target_ratio:
                    shape.crop_left=shape.crop_right=(1-target_ratio/source_ratio)/2
                else:shape.crop_top=shape.crop_bottom=(1-source_ratio/target_ratio)/2
            rounded(shape,node.get('radius',0),pw,ph)
        elif node['type']=='connector':
            shape=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x),Inches(y),Inches(x+w),Inches(y))
            shape.line.color.rgb=RGBColor.from_string(node['color'][1:]);shape.line.width=Pt(1.5)
            tail=OxmlElement('a:tailEnd');tail.set('type','triangle');shape._element.spPr.get_or_add_ln().append(tail)
        elif node['type']=='chart':
            if chart_renderer is None:raise DesignerValidationError('A native chart renderer is required')
            c=palette(node['dark']);conv=lambda color:tuple(bytes.fromhex(color[1:]))
            chart_renderer(slide,x,y,w,h,node['data'],{'ink':conv(c['ink']),'mid':conv(c['body']),
              'mute':conv(c['muted']),'line':conv(c['accent']),'bg':conv(c['bg'])})
    set_slide_notes(slide, scene['notes'])


def render_designer_slide(slide,item,asset,colors,number,total,lang,background,chart_renderer):
    scene=build_scene(item,asset,number,total,lang,item.get('_designer_style','designer_soft'))
    render_scene(slide,scene,chart_renderer)
    return scene
