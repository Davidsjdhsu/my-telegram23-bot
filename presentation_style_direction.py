"""Candidate soft/strict art direction, exercised through the bot's scene adapter.

No topic-specific layouts, providers, credentials, or separate slide exporter.
"""
from pathlib import Path
from presentation_profiles import resolved_profile, readable_color
from presentation_visuals import fit_rect
from presentation_catalog import TEMPLATES, composition_for, template_palette
from presentation_patterns import add_pattern


def photo_mask(asset, mode, number):
    """Only use a crop approved for this image, with a full-frame failure path."""
    user = asset.get('origin') == 'user'
    allowed = asset.get('safe_frames', []) if not user or asset.get('frame_reviewed') is True else []
    sequences = {'graphite': ('rect','circle','triangle','square'),
                 'classic': ('rect','square','circle','triangle'),
                 'technical': ('chamfer','hexagon','square','circle'),
                 'natural': ('rect','triangle','circle','square'),
                 None: ('rect','circle','square','triangle','hexagon','chamfer')}
    sequence = sequences[mode]
    offset = (number-1) % len(sequence)
    requested = sequence[offset]
    if requested in allowed:
        return requested
    if user and requested != 'rect':
        # A circle-safe photograph should not remain rectangular just because this
        # slide's first choice was a triangle. Never infer other safe shapes.
        for shape in sequence[offset:]+sequence[:offset]:
            if shape in allowed:
                return shape
    return 'rounded' if user and mode is None and 'rounded' in allowed else 'rect'


def user_frame_note(scene, asset):
    if not asset or asset.get('origin') != 'user':
        return
    shaped = any(n['type']=='image' and (n.get('crop') or n.get('radius') or n.get('mask','rect')!='rect')
                 for n in scene['nodes'])
    scene['notes'] = scene.get('notes','').replace('Image: supplied by user. Full frame preserved.',
        'Image: supplied by user. Decorative crop approved by image review.' if shaped else
        'Image: supplied by user. Full frame preserved.')


def has_custom_palette(design):
    kit=TEMPLATES[design['template_id']];p=template_palette(kit)
    background='dark' if int(kit['paper'][1:3],16)<80 else 'light'
    return (design['background']!=background or
            design['accent'].upper()!=p['bright' if background=='dark' else 'accent'].upper() or
            design['secondary'].upper()!=p['secondary'].upper())


def decorate_fallback(scene,item,asset,number,style,custom=False):
    """Retain validated legacy evidence/dense-copy geometry with the new finish."""
    soft=style=='designer_soft';d=item['_designer_design'];kit=TEMPLATES[d['template_id']]
    direction=strict_direction(kit)
    c=resolved_profile(d)
    bg=scene['background'] if custom else ('#FAF7F0' if soft else direction['bg'])
    ink=readable_color(c['ink'],bg) if soft or custom else direction['ink']
    accent=readable_color(c['accent'],bg) if soft or custom else direction['accent']
    scene['background']=bg
    for node in scene['nodes']:
        if node['type']=='text' and not custom:
            node['color']=readable_color(accent if node['role']=='subheading' else ink,bg)
        elif node['type']=='table' and not custom:
            node['background']=bg;node['accent']=accent
            for row in node['cells']:
                for cell in row:cell['color']=ink
        elif node['type']=='image' and asset and asset.get('origin')=='user':node['radius']=0
    if not add_pattern(scene,d,style,number,custom=custom) and not custom:
        texture=Path(__file__).resolve().parent/'assets'/('soft-pastel-contours.jpg' if soft else f"strict-{direction['key']}-pattern.jpg")
        if texture.is_file():scene['nodes'].insert(0,{'type':'background_image','path':str(texture),'box':[0,0,1280,720],'origin':'decoration','crop':False,'radius':0})
    # These evidence layouts may touch margins: use an inset stroke, never enlarge them.
    for node in scene['nodes']:
        if node['type']=='image':node['border_color']=accent;node['border_width']=1.5
    if asset and asset.get('origin')=='user':
        # Dense/custom layouts keep their existing image bounds and text geometry.
        mode=None if soft else direction['key']
        mask=photo_mask(asset,mode,number)
        for node in scene['nodes']:
            if node['type']!='image':continue
            x,y,w,h=node['box']
            if mask in ('circle','square','triangle','hexagon'):
                side=min(w,h);node['box']=[x+(w-side)/2,y+(h-side)/2,side,side]
            node['mask']='rect' if mask=='rounded' else mask
            node['radius']=min(42,min(node['box'][2:])*.08) if mask=='rounded' else 0
            node['crop']=mask in ('circle','square','triangle','hexagon')
        user_frame_note(scene,asset)
    return scene


def blend(a,b,t):
    return '#'+''.join(f'{round(int(a[i:i+2],16)*(1-t)+int(b[i:i+2],16)*t):02X}' for i in (1,3,5))


def strict_direction(kit):
    """The existing topic selector chooses a kit; its family selects strict art direction."""
    return {
        'editorial': dict(key='graphite',name='Графитовый',bg='#12171B',ink='#FFFFFF',accent='#D8DEE3',font='Arial',caps=True),
        'business': dict(key='classic',name='Светлый деловой',bg='#F7F8FA',ink='#172C48',accent='#315987',font='Georgia',caps=False),
        'technology': dict(key='technical',name='Технологический',bg='#091B30',ink='#F1F7FF',accent='#73DCEB',font='Trebuchet MS',caps=True),
        'education': dict(key='natural',name='Светлый зелёный',bg='#F4F6EE',ink='#203B2C',accent='#466B4B',font='Verdana',caps=False),
    }[kit['family']]


def build_style_scene(item,asset,number,total,lang,style):
    from presentation_designer import clean, text_node, validate_designer, validate_scene, DesignerValidationError
    design=item['_designer_design'];c=resolved_profile(design);kit=TEMPLATES[design['template_id']]
    soft=style=='designer_soft'
    title=clean(item.get('title'),160,'title')
    body=clean(item.get('content'),480,'body',optional=number==1)
    structured=validate_designer(item);layout=structured['layout'];blocks=structured['blocks']
    if asset and (not Path(asset['path']).is_file() or min(asset['width'],asset['height'])<=0):
        raise DesignerValidationError('Invalid image asset')
    if asset and blocks and layout not in ('comparison','conclusion'):
        body='\n\n'.join(b['heading']+'\n'+b['text'] for b in blocks);blocks=[];layout='photo'
    comp=composition_for(item,number,total)
    cover=number==1
    direction=None if soft else strict_direction(kit)
    mode=direction['key'] if direction else None
    dark=mode in ('graphite','technical')
    bg='#FAF7F0' if soft else (blend(c['dark'],'#080B0F',.55) if dark else blend(c['bg'],'#FFFFFF',.64))
    ink='#FFFFFF' if dark else readable_color(c['ink'],bg)
    accent=readable_color(c['accent'],bg) if soft else '#D8DEE3'
    body_color=blend(ink,bg,.14)
    head_font=(c['font'] if design['family']=='editorial' else 'Trebuchet MS') if soft else 'Arial'
    if direction:
        bg=direction['bg'];ink=direction['ink'];accent=direction['accent']
        body_color=blend(ink,bg,.14);head_font=direction['font']
    scene={'strict_direction':mode,'background':bg,'layout':layout,'composition':comp,'family':design['family'],
           'template_id':kit['id'],'style_direction':2,'nodes':[],
           'notes':f"AlixDocBot. {style}. {kit['id']} {kit['name']}. {comp}\n"+body}
    nodes=scene['nodes']
    def panel(box,fill,radius=0,gradient=None):
        nodes.append({'type':'panel','box':list(box),'fill':fill,'radius':radius,'gradient':gradient})
    def text(value,box,role='body',size=32,minimum=26,color=None):
        if not value:return
        if direction and direction['caps'] and role in ('title','subheading'):value=value.upper()
        node=text_node(value,box,color or body_color,size,minimum,
                       role in ('title','subheading') and not soft and mode!='classic',role,
                       head_font if role in ('title','subheading') else ('Calibri' if soft else 'Arial'))
        if lang=='ar':node['rtl']=True;node['font']='Arial'
        nodes.append(node)
    def heading(box=(64,48,1152,164),size=68):
        text(title,box,'title',size,40,ink)
    def photo(box):
        if not asset:return
        selected=photo_mask(asset,mode,number)
        shape='rect' if selected=='rounded' else selected
        bx,by,bw,bh=box;inner=(bx+16,by+16,bw-32,bh-32)
        if shape in ('circle','square','triangle','hexagon'):
            side=min(inner[2:]);fitted=(bx+(bw-side)/2,by+(bh-side)/2,side,side)
        else:fitted=fit_rect(asset['width'],asset['height'],inner)
        fx,fy,fw,fh=fitted
        frame_color=blend(accent,bg,.24) if soft else {'graphite':'#A8B3BD','classic':'#829AB3',
                                       'technical':'#63BFD4','natural':'#829C71'}[mode]
        radius=42 if soft and shape=='rect' and asset.get('origin')!='user' else 0
        if selected=='rounded':radius=min(42,min(fw,fh)*.08)
        # Native PowerPoint outlines follow the picture mask and remain editable.
        for gap,width in (((12,1.0),(5,1.5)) if soft else ((12,1.2),(5,2.0))):
            nodes.append({'type':'photo_frame','box':[fx-gap,fy-gap,fw+2*gap,fh+2*gap],
                          'mask':shape,'radius':radius+gap if radius else 0,
                          'color':blend(frame_color,bg,.32) if soft and gap==12 else frame_color,'width':width})
        nodes.append({'type':'image','path':str(asset['path']),'box':list(fitted),
                      'origin':asset.get('origin','ai'),'radius':radius,'mask':shape,
                      'crop':shape in ('circle','square','triangle','hexagon')})
    def rule(x,y,w,color=None):panel((x,y,w,2 if soft else 3),color or blend(accent,bg,.40))
    patterned=add_pattern(scene,design,style,number)
    if not soft and not patterned:
        texture=Path(__file__).resolve().parent/'assets'/f'strict-{mode}-pattern.jpg'
        if texture.is_file():
            nodes.append({'type':'background_image','path':str(texture),'box':[0,0,1280,720],
                          'origin':'decoration','radius':0,'crop':False})
    if soft and not patterned:
        texture=Path(__file__).resolve().parent/'assets'/'soft-pastel-contours.jpg'
        if texture.is_file():
            nodes.append({'type':'background_image','path':str(texture),'box':[0,0,1280,720],
                          'origin':'decoration','radius':0,'crop':False,
                          'flip_h':number%2==0,'flip_v':number%3==0})
        else:
            panel((0,0,1280,720),bg,gradient={'angle':24,'stops':[(0,bg,1),(1,blend(bg,c['secondary'],.35),1)]})
    if cover:
        sub=clean(item.get('subtitle') or item.get('point') or body,180,'subtitle',optional=True)
        if soft:
            heading((64,40,1152,193),72)
            if asset:
                if comp=='cover_band' and len(title)<33:
                    photo((64,248,778,420));text(sub,(900,333,316,262),size=32,minimum=26)
                else:
                    photo((438,248,778,420));text(sub,(64,333,316,262),size=32,minimum=26)
            else:text(sub,(64,360,1030,238),size=38)
        elif comp=='cover_band' and len(title)<33:
            heading((64,46,1152,178),72)
            if mode=='classic':
                photo((64,260,750,400));text(sub,(870,330,346,280),size=32)
            else:
                text(sub,(64,330,346,280),size=32);photo((466,260,750,400))
            rule(64,232,1152)
        elif mode=='technical':
            heading((64,62,560,395),72);text(sub,(64,500,550,155),size=30)
            photo((670,160,546,430));rule(670,620,546,accent)
        elif mode=='classic':
            heading((64,52,1152,190),72)
            text(sub,(64,340,400,276),size=32)
            photo((532,254,684,400));rule(64,248,1152)
        elif mode=='natural':
            heading((64,48,1152,170),68);rule(64,231,1152)
            photo((64,272,760,390));text(sub,(876,318,340,312),size=31)
        else:
            heading((64,42,1152,192),76);rule(64,247,1152,accent)
            if asset:
                photo((64,281,710,380));text(sub,(832,340,384,280),size=32,minimum=26)
            else:text(sub,(64,350,1050,260),size=38)
    elif asset and layout not in ('comparison','conclusion') and not item.get('chart'):
        # Pick a deep slot for portraits, a wide slot for landscapes. Never crop.
        portrait=asset['width']/asset['height']<1.15
        if soft:
            heading((64,42,1152,165),68)
            if portrait:
                photo((64,224,534,432));text(body,(654,254,562,390),size=34,minimum=26)
            elif comp=='photo_left':
                photo((64,238,758,416));text(body,(876,231,340,420),size=31,minimum=25)
            elif comp=='photo_right':
                text(body,(64,241,340,412),size=31,minimum=25);photo((458,238,758,416))
            elif comp=='photo_top':
                photo((64,238,584,416));text(body,(700,244,516,410),size=32,minimum=25)
            else:
                text(body,(64,244,516,410),size=32,minimum=25);photo((632,238,584,416))
        else:
            if not portrait and comp in ('photo_top','photo_bottom'):
                heading((64,42,1152,165),70);rule(64,218,1152)
                if comp=='photo_top':
                    photo((64,245,584,410));text(body,(700,250,516,400),size=31,minimum=25)
                else:
                    text(body,(64,250,516,400),size=31,minimum=25);photo((632,245,584,410))
            elif mode=='technical' and comp=='photo_right' and not portrait and len(title)<28:
                heading((64,54,360,262),62);text(body,(64,330,360,330),size=29,minimum=25)
                photo((484,144,732,486));rule(484,655,732,accent)
            elif mode=='natural' and comp=='photo_left' and not portrait and len(title)<28:
                heading((850,62,366,222),58);text(body,(850,320,366,330),size=29,minimum=25)
                photo((64,142,720,460));rule(64,636,720)
            elif mode=='classic' and comp=='photo_right' and not portrait:
                heading((64,48,1152,160),72);rule(64,218,1152)
                text(body,(64,266,380,388),size=30,minimum=25);photo((504,244,712,410))
            elif portrait:
                heading((64,70,590,230),66);text(body,(64,354,590,290),size=31,minimum=25)
                photo((710,48,506,608));rule(676,48,2,accent)
            elif comp=='photo_left':
                heading((64,40,1152,165),82);rule(64,214,1152)
                photo((64,244,734,410))
                text(body,(860,254,356,398),size=30,minimum=25)
            else:
                heading((64,40,1152,170),76);rule(64,218,1152)
                text(body,(64,274,356,378),size=30,minimum=25)
                photo((482,244,734,410))
    elif layout=='process':
        heading()
        n=len(blocks)
        if soft:
            panel((40,235,1200,415),blend(bg,c['secondary'],.22),48)
            width=(1104-44*(n-1))/n
            for i,b in enumerate(blocks):
                x=88+i*(width+44)
                text(f'{i+1:02}',(x,266,width,100),'title',70,44,accent)
                text(b['heading'],(x,385,width,94),'subheading',38,26,ink)
                text(b['text'],(x,495,width,132),size=29,minimum=23)
        elif mode in ('classic','technical'):
            width=(1152-48*(n-1))/n
            for i,b in enumerate(blocks):
                x=64+i*(width+48);rule(x,240,width,accent)
                text(f'{i+1:02}',(x,268,width,100),'title',72,42,accent)
                text(b['heading'],(x,389,width,94),'subheading',38,26,ink)
                text(b['text'],(x,500,width,142),size=29,minimum=23)
        else:
            rh=430/n
            for i,b in enumerate(blocks):
                y=226+i*rh;rule(64,y,1152)
                text(f'{i+1:02}',(64,y+22,112,rh-28),'title',72,42,accent)
                text(b['heading'],(218,y+24,300,rh-28),'subheading',36,26,ink)
                text(b['text'],(594,y+24,622,rh-28),size=30,minimum=24)
    elif layout=='comparison' and structured.get('rows'):
        heading((64,48,1152,158),66)
        widths=[280,436,436];heights=[96]+[354/len(structured['rows'])]*len(structured['rows'])
        values=[['Критерий' if lang=='ru' else 'Criterion']+[b['heading'] for b in blocks]]+[[r[k] for k in ('criterion','left','right')] for r in structured['rows']]
        backgrounds=([blend(bg,c['secondary'],.35)]+[bg,blend(bg,c['secondary'],.15),bg]) if soft else [('#E1E5E8' if mode=='graphite' else accent),bg,bg,bg]
        cells=[]
        for row,values_row in enumerate(values):
            cells.append([text_node(value,(0,0,widths[col]-40,heights[row]-28),
                           ('#111820' if mode=='graphite' else readable_color(ink,backgrounds[row])) if row==0 and not soft else readable_color(accent if row==0 else ink,backgrounds[row]),
                           30 if row==0 else 28,23,row==0 or col==0,'table','Calibri') for col,value in enumerate(values_row)])
        nodes.append({'type':'table','box':[64,220,1152,sum(heights)],'widths':widths,'heights':heights,'cells':cells,
                      'background':bg,'row_backgrounds':backgrounds,'accent':blend(accent,bg,.65)})
    elif layout in ('sections','conclusion','comparison'):
        if layout=='sections' and soft:
            heading((64,48,1152,165),70)
            rh=430/len(blocks)
            for i,b in enumerate(blocks):
                y=235+i*rh
                text(b['heading'],(64,y,340,rh-18),'subheading',36,26,accent)
                text(b['text'],(460,y,756,rh-18),size=32,minimum=24)
        elif layout=='conclusion' and not soft:
            heading((64,80,410,450),70)
            rh=570/len(blocks)
            for i,b in enumerate(blocks):
                y=62+i*rh;rule(570,y,646,blend(accent,bg,.2))
                text(b['heading'],(570,y+26,646,72),'subheading',38,28,accent)
                text(b['text'],(570,y+107,646,rh-114),size=30,minimum=24)
        else:
            heading((64,48,1152,165),70)
            cols=3 if len(blocks)==3 else 2;rows=(len(blocks)+cols-1)//cols
            width=(1152-64*(cols-1))/cols;rh=430/rows
            for i,b in enumerate(blocks):
                x=64+(i%cols)*(width+64);y=242+(i//cols)*rh
                if not soft:rule(x,y,width)
                text(b['heading'],(x,y+12,width,90),'subheading',42,26,accent)
                text(b['text'],(x,y+128,width,rh-142),size=32,minimum=24)
    elif item.get('chart'):
        heading();text(body,(64,254,360,386),size=31)
        nodes.append({'type':'chart','box':[484,237,732,410],'data':item['chart'],'dark':dark,'palette':c})
    else:
        heading((64,70,soft and 1152 or 480,soft and 200 or 530),72)
        if soft:text(body,(240,355,920,270),size=38,minimum=28)
        else:text(body,(632,150,584,480),size=34,minimum=27);rule(570,80,2)
    if asset:scene['notes']+='\nImage: '+('supplied by user. Full frame preserved.' if asset.get('origin')=='user' else 'AI-generated illustration, not documentary evidence. Geometric framing only where approved by image review.')
    user_frame_note(scene,asset)
    text(f'{number:02} / {total:02}',(64,680,1152,24),'footer',13,13,blend(ink,bg,.35))
    return validate_scene(scene)
