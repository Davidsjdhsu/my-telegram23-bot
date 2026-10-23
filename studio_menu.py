"""Adapters for the private menu; reuse current forms and billing unchanged."""
import json
from pathlib import Path

def recent_payments(path, uid):
    """Bounded tail read; never returns provider IDs or another user's payments."""
    result=[]
    try:
        with open(path,'rb') as f:
            f.seek(0,2);start=max(0,f.tell()-512_000);f.seek(start)
            if start:f.readline()
            lines=f.read().splitlines()
        for line in reversed(lines):
            try:
                row=json.loads(line)
                if str(row.get('uid'))!=str(uid):continue
                result.append({'ts':str(row.get('ts',''))[:30],'credits':int(row['credits'])})
                if len(result)>=30:break
            except (ValueError,TypeError,KeyError):continue
    except OSError:pass
    return result

def forms_html(base):
    source=(Path(base)/'index.html').read_text()
    source=source.replace('Меню — Neon Glassmorphism','AlixDoc · Документы').replace('<head>','<head><base href="/">',1)
    source=source.replace("const miniSession = createMiniAppSession({ getTg, apiBase });",'''const miniSession = (()=>{
      let pending=false,retry=null;
      async function request(path,body){const r=await fetch('/studio/'+path,{method:body?'POST':'GET',headers:body?{'Content-Type':'application/json','X-Studio-Request':'1'}:{},body:body?JSON.stringify(body):undefined,cache:'no-store'});const b=await r.json();if(!r.ok||!b.ok)throw Error(b.error||'request_failed');return b}
      return {loadProfile:()=>request('profile'),async send(action,payload={}){if(pending)return false;pending=true;const key=JSON.stringify({action,payload});if(!retry||retry.key!==key)retry={key,id:crypto.randomUUID()};try{await request('action',{action,payload,requestId:retry.id});retry=null;return true}finally{pending=false}}};
    })();''')
    source=source.replace('new URL(relativePath, window.location.href)', "new URL(relativePath, window.location.origin + '/')")
    source=source.replace('if (tg.initData) {','if (true) {').replace("if (!getTg()?.initData) return;",'')
    # Match existing theme while preserving every form, price and validation rule.
    theme='''<link rel="stylesheet" href="/studio/templates/fonts.css"><style>
    html,body{background:#191b20!important;color:#e7e2d9;font-family:Manrope,sans-serif!important}.display{font-family:'Cormorant Garamond',Georgia,serif!important}.orb,.glow,.glow-1,.glow-2,.bg-glow{display:none!important}input,textarea,select{background:#252830!important;color:#e7e2d9!important;border-color:#514b45!important}.btn-primary,.form-submit,.field-chip.active,.credits-topup,.gauge-topup{background:linear-gradient(110deg,#b59a6b,#d1bb93)!important;color:#211e19!important;box-shadow:none!important}.tabs,.tabbar-float{display:none!important}.form-label,.form-hint,.form-cost{color:#b7b3ad!important}.wrap{max-width:600px;padding-bottom:30px}.studio-return{display:block;padding:18px 22px;color:#c4aa7d;text-decoration:none;font:14px Manrope,sans-serif}body{padding-bottom:24px!important}</style>'''
    source=source.replace('</head>',theme+'</head>').replace('<body>','<body><a class="studio-return" href="/studio/home/">← Меню AlixDoc</a>')
    initialization='''
    const studioScreen=new URLSearchParams(location.search).get('screen');
    if(['form-word','form-excel','form-resume','form-dkp-car','form-dkp-realty','topup','profile','create'].includes(studioScreen))showScreen(studioScreen);
    document.addEventListener('click',e=>{const b=e.target.closest('[data-back-to],.tab-item');if(b){e.preventDefault();e.stopImmediatePropagation();location.href='/studio/home/'}},true);
    '''
    at=source.rfind('</script>');source=source[:at]+initialization+source[at:]
    return source
