"""Private, bounded, model-free previews using the production designer scenes."""
import asyncio
import copy
import math
import re
from functools import lru_cache
from pathlib import Path
from xml.etree.ElementTree import Element, SubElement, tostring
from aiohttp import web
from studio_settings import PALETTES

BASE = Path(__file__).resolve().parent
ITEMS = [
 {'title':'Работа кофейни','content':'От выбора напитка до выдачи заказа','designer':{'layout':'auto','blocks':[]}},
 {'title':'Что делает бариста','content':'Бариста уточняет заказ, готовит напиток и проверяет его перед подачей. Вкус зависит от зерна, помола и приготовления.','designer':{'layout':'statement','blocks':[]}},
 {'title':'Пространство кофейни','content':'За стойкой готовят напитки, а в гостевой зоне посетители отдыхают. Удобное расположение оборудования помогает команде работать.','designer':{'layout':'photo','blocks':[]}},
 {'title':'Путь заказа','content':'Выбор напитка, приготовление и выдача гостю.','designer':{'layout':'auto','blocks':[]},'diagram':{'kind':'sequence','nodes':[{'label':'Выбор','text':'Гость выбирает напиток.'},{'label':'Приготовление','text':'Бариста выполняет заказ.'},{'label':'Выдача','text':'Гость получает напиток.'}]}}
]

def scene_for(style, key, palette, page):
    from presentation_profiles import choose_design
    from presentation_designer import build_scene
    design = choose_design({'topic':'Работа кофейни','extra':f'Шаблон №{int(key[1:])}', 'studio_palette':palette})
    item = copy.deepcopy(ITEMS[page])
    item.update(_designer_style='designer_'+style, _designer_design=design)
    asset = {'path':str(BASE/'studio_ui/constructor/coffee.jpg'),'width':1344,'height':768,'origin':'user'} if page in (0,2) else None
    return build_scene(item, asset, page+1, 4, 'ru', 'designer_'+style)

def el(parent, tag, **attrs):
    return SubElement(parent, tag, {k.replace('_','-'):str(v) for k,v in attrs.items()})

def geometry(parent, n, **attrs):
    x,y,w,h=n['box']; mask=n.get('mask','rect')
    if mask=='circle': return el(parent,'ellipse',cx=x+w/2,cy=y+h/2,rx=w/2,ry=h/2,**attrs)
    if mask in ('triangle','hexagon','chamfer'):
        pts={'triangle':[(.5,0),(1,1),(0,1)],'hexagon':[(.25,0),(.75,0),(1,.5),(.75,1),(.25,1),(0,.5)],'chamfer':[(.08,0),(1,0),(1,.92),(.92,1),(0,1),(0,.08)]}[mask]
        return el(parent,'polygon',points=' '.join(f'{x+a*w},{y+b*h}' for a,b in pts),**attrs)
    return el(parent,'rect',x=x,y=y,width=w,height=h,rx=n.get('radius',0),**attrs)

def svg_for(scene, prefix):
    svg=Element('svg',{'xmlns':'http://www.w3.org/2000/svg','viewBox':'0 0 1280 720','role':'img','aria-label':'Пример выбранной палитры'})
    defs=el(svg,'defs')
    el(svg,'rect',width=1280,height=720,fill=scene['background'])
    for i,n in enumerate(scene['nodes']):
        x,y,w,h=n['box']; kind=n['type']; ident=f'{prefix}-{i}'
        if kind=='panel':
            fill=n.get('fill','none')
            if n.get('gradient'):
                g=n['gradient']; angle=math.radians(g['angle'])
                grad=el(defs,'linearGradient',id=ident,x1=.5-.5*math.cos(angle),y1=.5+.5*math.sin(angle),x2=.5+.5*math.cos(angle),y2=.5-.5*math.sin(angle))
                for offset,color,opacity in g['stops']: el(grad,'stop',offset=offset,stop_color=color,stop_opacity=opacity)
                fill=f'url(#{ident})'
            geometry(svg,n,fill=fill)
        elif kind=='text':
            size=n['size']; lines=n['text'].split('\n')
            t=el(svg,'text',font_family=n['font'],font_size=size,font_weight=700 if n['bold'] else 400,fill=n['color'])
            for j,line in enumerate(lines):
                span=el(t,'tspan',x=x+(n.get('stagger',0) if j==1 else 0),y=y+size*.96+j*size*1.22,fill=n.get('last_line_color',n['color']) if j==len(lines)-1 else n['color'])
                span.text=line
        elif kind in ('image','background_image'):
            p=Path(n['path'])
            if p.parent == BASE/'assets/patterns' and re.fullmatch(r'(strict|soft)-d\d{2}\.jpg',p.name):
                url='/studio/templates/preview-pattern/'+p.name
            elif p == BASE/'studio_ui/constructor/coffee.jpg': url='/studio/constructor/coffee.jpg'
            else: raise ValueError('Unexpected preview asset')
            clip=el(defs,'clipPath',id=ident);geometry(clip,n)
            group=el(svg,'g',clip_path=f'url(#{ident})')
            transforms=[]
            if n.get('flip_h'): transforms.append(f'translate({2*x+w},0) scale(-1,1)')
            if n.get('flip_v'): transforms.append(f'translate(0,{2*y+h}) scale(1,-1)')
            im=el(group,'image',href=url,x=x,y=y,width=w,height=h,opacity=n.get('opacity',1),transform=' '.join(transforms))
            im.set('preserveAspectRatio','xMidYMid slice' if n.get('crop') else 'none')
            if n.get('border_color'):geometry(svg,n,fill='none',stroke=n['border_color'],stroke_width=n.get('border_width',1)*4/3)
        elif kind=='photo_frame': geometry(svg,n,fill='none',stroke=n['color'],stroke_width=n['width']*4/3)
        elif kind in ('connector','diagram_edge'):
            marker=el(defs,'marker',id=ident,markerWidth=8,markerHeight=8,refX=7,refY=4,orient='auto')
            el(marker,'path',d='M 0 0 L 8 4 L 0 8 Z',fill=n['color'])
            if kind=='diagram_edge': x1,y1,x2,y2=n['points']
            elif n.get('vertical'): x1,y1,x2,y2=x+w/2,y,x+w/2,y+h
            else: x1,y1,x2,y2=x,y+h/2,x+w,y+h/2
            line=el(svg,'line',x1=x1,y1=y1,x2=x2,y2=y2,stroke=n['color'],stroke_width=2 if kind=='connector' else 8/3)
            if n.get('arrow',True):line.set('marker-end',f'url(#{ident})')
        else: raise ValueError('Unsupported preview node: '+kind)
    return tostring(svg,encoding='unicode')

@lru_cache(maxsize=128)
def render_preview(style,key,palette,page):
    return svg_for(scene_for(style,key,palette,page),f'p-{style}-{key}-{palette}-{page}')

def register(studio,app):
    lock=asyncio.Lock()
    async def preview(request):
        studio.authorized(request)
        q=request.query
        style,key,palette,page=q.get('style'),q.get('key'),q.get('palette'),q.get('page')
        if style not in ('strict','soft') or not isinstance(key,str) or not re.fullmatch(r'd(?:0[1-9]|[1-4][0-9]|50)',key) or palette not in PALETTES or page not in ('0','1','2','3'):
            raise web.HTTPBadRequest(text='Неизвестный вариант оформления')
        async with lock:
            svg=await asyncio.to_thread(render_preview,style,key,palette,int(page))
        return web.json_response({'svg':svg},headers={'Cache-Control':'private, max-age=3600','X-Content-Type-Options':'nosniff'})
    async def pattern(request):
        studio.authorized(request)
        name=request.match_info['name']
        if not re.fullmatch(r'(strict|soft)-d(?:0[1-9]|[1-4][0-9]|50)\.jpg',name):raise web.HTTPNotFound()
        path=BASE/'assets/patterns'/name
        if not path.is_file():raise web.HTTPNotFound()
        return web.FileResponse(path,headers={'Cache-Control':'private, max-age=86400','X-Content-Type-Options':'nosniff'})
    app.router.add_get('/studio/templates/palette-preview',preview)
    app.router.add_get('/studio/templates/preview-pattern/{name}',pattern)
