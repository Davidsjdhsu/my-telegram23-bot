"""Twelve editable compositions per catalogue kit, used by the live bot renderer."""
from pathlib import Path
from presentation_catalog import TEMPLATES, composition_for
from presentation_profiles import resolved_profile, readable_color
from presentation_visuals import fit_rect


def build_catalog_scene(item, asset, number, total, lang, style):
    from presentation_style_direction import build_style_scene, decorate_fallback, has_custom_palette
    from presentation_designer import DesignerValidationError
    # Keep the proven spacious compositions for dense copy and supplied evidence.
    custom=has_custom_palette(item['_designer_design'])
    evidence=bool(asset and item.get('designer',{}).get('layout') in ('comparison','conclusion'))
    if not custom and not evidence:
        try:
            return build_style_scene(item, asset, number, total, lang, style)
        except DesignerValidationError as exc:
            if not str(exc).startswith('Shorten '):raise
    scene=build_catalog_scene_previous(item, asset, number, total, lang, style)
    return decorate_fallback(scene,item,asset,number,style,custom)


def build_catalog_scene_previous(item, asset, number, total, lang, style):
    # Lazy import avoids a cycle while retaining exactly the same text fitter and adapter.
    from presentation_designer import clean, text_node, validate_designer, validate_scene, DesignerValidationError
    d=item['_designer_design'];t=TEMPLATES[d['template_id']];c=resolved_profile(d)
    title=clean(item.get('title'),160,'title')
    body=clean(item.get('content'),480,'body',optional=number==1)
    structured=validate_designer(item);blocks=structured['blocks']
    comp=composition_for(item,number,total)
    if not 1<=number<=total:raise DesignerValidationError('Invalid slide number')
    if asset and (not Path(asset['path']).is_file() or min(asset['width'],asset['height'])<=0):
        raise DesignerValidationError('Invalid image asset')
    if asset and blocks and comp not in ('comparison','conclusion'):
        # Preserve structured copy even if the image planner promoted a short sections slide.
        body='\n\n'.join(b['heading']+'\n'+b['text'] for b in blocks)
        comp=('photo_left','photo_right','photo_top','photo_bottom')[(number+t['sequence'])%4]
        blocks=[]
    if comp.startswith('photo_') and not asset:comp='statement'
    scene={'background':c['bg'],'layout':structured['layout'],'composition':comp,
           'family':d['family'],'template_id':t['id'],'nodes':[],
           'notes':f"AlixDocBot. {style}. {t['id']} {t['name']}. {comp}\n"+body}
    nodes=scene['nodes'];m=t['margin'];w=1280-2*m;system=t['system']
    radius=0 if style=='designer_strict' else (20,24,28,32,36)[(t['number']-1)%5]
    font=t['font'] if lang not in ('ar','zh') else 'Arial'
    def txt(value,box,role='body',size=32,minimum=24,color=None):
        if not value:return
        node=text_node(value,box,color or c['body'],size,minimum,
            t['bold'] if role=='title' else role=='subheading',role,font if role=='title' else t['body_font'])
        if lang=='ar':node['rtl']=True
        nodes.append(node)
    def head(box=None):
        txt(title,box or (m,48,w,176),'title',t['title_size'],36,
            (c['accent'] if number==1 else c['secondary_ink']) if d['family']=='editorial' else c['accent'] if system in ('field','poster') else c['ink'])
    def photo(box,full=False):
        if not asset:return
        # Image review sees the whole frame, so placement must preserve it too.
        crop=False
        rect=fit_rect(asset['width'],asset['height'],box)
        nodes.append({'type':'image','path':str(asset['path']),'box':list(rect),
                      'origin':asset.get('origin','ai'),'radius':radius,'crop':crop})
    def panel(box,fill):
        nodes.append({'type':'panel','box':list(box),'fill':fill,'radius':0})
    def arrow(box,vertical=False):
        nodes.append({'type':'connector','box':list(box),'color':c['accent'],'vertical':vertical})
    def safe_photo():
        # A wider alternative for unusually long copy; never silently remove text.
        nodes.clear();head((m,40,w,190))
        # Keep a large complete image when a shallow band would cut the subject.
        left=comp!='photo_bottom'
        pw=w*.43;tw=w-pw-40
        photo((m if left else m+tw+40,258,pw,380))
        txt(body,(m+pw+40 if left else m,250,tw,405),size=30,minimum=24)
    if comp.startswith('cover'):
        sub=clean(item.get('subtitle') or item.get('point') or body,180,'subtitle',optional=True)
        try:
            if asset and len(title)>44:
                head((m,40,w,258));txt(sub,(m,352,w*.54,280),size=31)
                photo((m+w*.59,340,w*.41,300))
            elif asset and system=='cinema' and comp=='cover_split':
                photo((0,0,1280,720),True)
                nodes.append({'type':'panel','box':[0,0,1280,720],'radius':0,
                    'gradient':{'angle':0,'stops':[(0,c['bg'],1),(.68,c['bg'],.97),(1,c['bg'],.2)]}})
                head((m,100,800,350));txt(sub,(m,490,760,155),size=33)
            elif asset and comp=='cover_band':
                if system in ('gallery','panorama'):
                    photo((m,48,w,320),True);head((m,403,w*.61,230))
                    txt(sub,(m+w*.65,414,w*.35,220),size=30)
                else:
                    head((m,45,w,210));photo((m,285,w*.62,354),True)
                    txt(sub,(m+w*.66,300,w*.34,330),size=31)
            elif asset:
                left=round(w*t['split']);gap=48
                if system in ('margin','field'):
                    photo((m,62,w-left-gap,560));head((m+w-left,70,left,365))
                    txt(sub,(m+w-left,464,left,184),size=30)
                else:
                    head((m,80,left,355));txt(sub,(m,470,left,180),size=32)
                    photo((m+left+gap,62,w-left-gap,570))
            else:
                head((m,100,w,350));txt(sub,(m,490,w*.78,150),size=34)
        except DesignerValidationError:
            nodes.clear();head((m,40,w,258))
            txt(sub,(m,352,w*.54,280),size=31)
            if asset:photo((m+w*.59,340,w*.41,300))
    elif comp.startswith('photo'):
        try:
            if comp in ('photo_top','photo_bottom') and asset['width']/asset['height'] < w/420:
                raise DesignerValidationError('Use a deeper frame for this image')
            if comp in ('photo_left','photo_right'):
                # The marginal systems reserve a title column; others lead with a full-width title.
                if system in ('margin','editorial') and comp=='photo_right' and len(title)<65 and len(body)<260:
                    title_w=w*.29;head((m,64,title_w,510))
                    image_x=m+title_w+36;image_w=w-title_w-36
                    photo((image_x,55,image_w,345))
                    txt(body,(image_x,438,image_w,205),size=31)
                else:
                    head();pw=round(w*(.46 if system=='report' else t['split']))
                    gap=44;tw=w-pw-gap;photo_left=comp=='photo_left'
                    photo((m if photo_left else m+tw+gap,255,pw,387))
                    txt(body,(m+pw+gap if photo_left else m,260,tw,387),size=32)
            elif comp=='photo_top':
                if len(body)>270:raise DesignerValidationError('Use wider body')
                photo((m,44,w,305 if system!='panorama' else 340),True)
                y=390 if system!='panorama' else 422
                head((m,y,w*.48,650-y));txt(body,(m+w*.53,y,w*.47,650-y),size=30)
            else:
                if len(body)>270:raise DesignerValidationError('Use wider body')
                head((m,50,w*.48,235));txt(body,(m+w*.53,62,w*.47,220),size=31)
                photo((m,330,w,315),True)
        except DesignerValidationError:safe_photo()
    elif comp=='process':
        head();n=len(blocks)
        vertical=system in ('margin','field','report')
        if vertical:
            rh=420/n
            for i,b in enumerate(blocks):
                y=238+i*rh
                txt(f'{i+1:02}',(m,y,66,rh-12),'subheading',34,26,c['accent'])
                txt(b['heading'],(m+100,y,300,rh-10),'subheading',31,22,c['ink'])
                txt(b['text'],(m+438,y,w-438,rh-10),size=28,minimum=22)
                if i<n-1:arrow((m+32,y+46,2,rh-57),True)
        else:
            bw=(w-(n-1)*32)/n
            for i,b in enumerate(blocks):
                x=m+i*(bw+32)
                txt(f'{i+1:02}',(x,263,bw,70),'subheading',48,36,c['accent'])
                txt(b['heading'],(x,360,bw,110),'subheading',32,24,c['ink'])
                txt(b['text'],(x,489,bw,170),size=28,minimum=22)
                if i<n-1:arrow((x+78,299,bw-57,2))
    elif comp=='comparison' and structured.get('rows'):
        head((m,45,w-260 if asset else w,162))
        if asset:photo((1280-m-230,48,230,146))
        widths=[round(w*.24),round(w*.38),w-round(w*.24)-round(w*.38)]
        heights=[78,118,118,118]
        headers=['Критерий' if lang=='ru' else 'Criterion']+[b['heading'] for b in blocks]
        values=[headers]+[[r[k] for k in ('criterion','left','right')] for r in structured['rows']]
        cells=[[text_node(v,(0,0,widths[col]-32,heights[row]-20),c['accent'] if row==0 else c['ink'],
                 30 if row==0 else 28,22,row==0 or col==0,'table',t['body_font'])
                for col,v in enumerate(vals)] for row,vals in enumerate(values)]
        nodes.append({'type':'table','box':[m,224,w,sum(heights)],'widths':widths,'heights':heights,
                      'cells':cells,'background':c['bg'],'accent':c['accent']})
    elif comp=='conclusion':
        try:
            head((m,80,w*.35,340 if asset else 530))
            if asset:photo((m,450,w*.35,185))
            rh=550/len(blocks);x=m+w*.43;bw=w*.57
            for i,b in enumerate(blocks):
                y=80+i*rh;hh=64 if len(blocks)==3 else 85
                txt(b['heading'],(x,y,bw,hh),'subheading',34,24,c['accent'])
                txt(b['text'],(x,y+hh+10,bw,rh-hh-30),size=31,minimum=24)
        except DesignerValidationError:
            nodes.clear();head((m,45,w-260 if asset else w,166))
            if asset:photo((1280-m-230,48,230,146))
            rh=426/len(blocks)
            for i,b in enumerate(blocks):
                y=231+i*rh
                txt(b['heading'],(m,y,w*.30,rh-22),'subheading',36,24,c['accent'])
                txt(b['text'],(m+w*.35,y,w*.65,rh-22),size=32,minimum=24)
    elif comp in ('sections','comparison'):
        head((m,45,w-260 if asset else w,166))
        if asset:photo((1280-m-230,48,230,146))
        n=len(blocks)
        # Dense conclusions use rows. Other kits use a vertical reading order or broad columns.
        rows=comp=='conclusion' or system in ('margin','report','field')
        if rows:
            rh=426/n
            for i,b in enumerate(blocks):
                y=231+i*rh
                txt(b['heading'],(m,y,w*.30,rh-22),'subheading',36,24,c['accent'])
                txt(b['text'],(m+w*.35,y,w*.65,rh-22),size=32,minimum=24)
        else:
            cols=3 if n==3 else 2;nr=(n+cols-1)//cols;bw=(w-40*(cols-1))/cols;bh=420/nr
            for i,b in enumerate(blocks):
                x=m+(i%cols)*(bw+40);y=235+(i//cols)*bh
                hh=85 if nr==1 else 65
                txt(b['heading'],(x,y,bw,hh),'subheading',36,24,c['accent'])
                txt(b['text'],(x,y+hh+12,bw,bh-hh-26),size=31,minimum=22)
    elif comp=='chart':
        head();txt(body,(m,257,390,383),size=30)
        nodes.append({'type':'chart','box':[m+440,249,w-440,400],'data':item['chart'],
                      'dark':d['background']=='dark','palette':c})
    else:
        if system in ('editorial','margin','gallery') and len(title)<65:
            head((m,100,w*.43,510));txt(body,(m+w*.49,160,w*.51,470),size=36,minimum=24)
        else:
            head((m,72,w,240));txt(body,(m+w*.10,355,w*.80,295),size=38,minimum=24,color=c['accent'])
    label=f'{number:02} / {total:02}'
    if asset:
        if asset.get('origin')=='user':scene['notes']+='\nImage: supplied by user. Full frame preserved.'
        else:
            scene['notes']+='\nImage: AI-generated illustration, not documentary evidence.'
    txt(label,(m,675,w,24),'footer',13,13,c['muted'])
    return validate_scene(scene)
