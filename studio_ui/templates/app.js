(async()=>{
const root=document.getElementById('alix-real-catalog');
const response=await fetch('catalog.json',{cache:'no-cache'});
if(!response.ok)throw new Error('Catalog unavailable');
const data=await response.json();
const $=s=>root.querySelector(s), $$=s=>Array.from(root.querySelectorAll(s));
const palettes=await (await fetch('palettes.json')).json();
const initial={version:3,palette:'',background_mode:'solid',diagram_layout:'classic',kind:'standard',style:'strict',ordinary:'coffee',designer:'d16',page:0,slides:6,confirmed:false};
function normalize(s){
 s=s&&s.version===3?s:{};
 const n={...initial,...s};
 n.background_mode=n.background_mode==='gradient'?'gradient':'solid';
 n.diagram_layout=['classic','cards','steps'].includes(n.diagram_layout)?n.diagram_layout:'classic';
 n.palette=palettes.some(p=>p.key===n.palette)?n.palette:'';
 n.kind=n.kind==='designer'?'designer':'standard';n.style=n.style==='soft'?'soft':'strict';
 if(!data.some(r=>r.group==='standard'&&r.key===n.ordinary))n.ordinary='coffee';
 if(!data.some(r=>r.group==='strict'&&r.key===n.designer))n.designer='d16';
 n.page=[0,1,2,3].includes(n.page)?n.page:0;n.slides=[4,6,8,10,12,16].includes(+n.slides)?+n.slides:6;n.confirmed=n.confirmed===true;
 return Object.fromEntries(Object.keys(initial).map(k=>[k,n[k]]));
}
let restored;try{restored=JSON.parse(localStorage.getItem('alix-hd-preview'));}catch{}
let state=normalize(window.openai?.widgetState?.privateContent||restored);
const requestedKind=new URLSearchParams(location.search).get('kind');
if(['standard','designer'].includes(requestedKind))state.kind=requestedKind;
let groupRendered='',sent=new Set(),lastSaved='',touchStart=null;
const constrained=navigator.connection?.saveData||/2g/.test(navigator.connection?.effectiveType||'')||navigator.deviceMemory<=2;
const previews=new window.AlixPreviewLoader({limit:constrained?2:4,concurrency:2});
const slideNames=['Обложка','Текст','Фото','Схема'];
let requestID=0,shownURL='',idleTimer=null,thumbObserver=null,outlineModule,paletteModule;
let paletteImport;
function paletteAPI(){return paletteImport??=import('./palette-preview.js').then(m=>paletteModule=m).catch(e=>{paletteImport=null;throw e;});}
function liveMode(on){$('#rc-slide').hidden=on;$('#rc-zoom-image').hidden=on;$('#rc-live-slide').hidden=!on;$('#rc-live-zoom').hidden=!on;}
function paint(url){$('#rc-slide').src=url;$('#rc-zoom-image').src=url;}
function showSlide(r){
 if(state.kind==='designer'&&(state.palette||state.background_mode==='gradient'||state.diagram_layout!=='classic')){showPalette(r);return;}
 const url=r.images[state.page];if(shownURL===url)return;
 paletteModule?.cancel();liveMode(false);
 shownURL=url;const ticket=++requestID;clearTimeout(idleTimer);
 paint(r.thumbs[state.page]);window.AlixMobile?.enter($('.rc-stage'));$('#rc-quality').textContent='Загружаю чёткий слайд…';
 $('#rc-quality').hidden=false;$('#rc-slide').setAttribute('aria-busy','true');
 previews.focus(url).then(objectURL=>{
  if(ticket!==requestID)return;
  paint(objectURL);$('#rc-quality').hidden=true;$('#rc-slide').setAttribute('aria-busy','false');
  if(!constrained)idleTimer=setTimeout(()=>{
   if(ticket!==requestID||document.visibilityState==='hidden')return;
   previews.warm([r.images[(state.page+1)%4],r.images[(state.page+3)%4]]);
  },600);
 }).catch(error=>{
  if(ticket!==requestID||error.name==='AbortError')return;
  $('#rc-slide').setAttribute('aria-busy','false');$('#rc-quality').textContent='Превью. Нажмите, чтобы загрузить чётче';shownURL='';
 });
}
async function showPalette(r){
 const {style,palette,page,background_mode,diagram_layout}=state,key=r.key;
 const id=[style,key,palette,page,background_mode,diagram_layout].join(':');if(shownURL===id)return;
 shownURL=id;const ticket=++requestID;clearTimeout(idleTimer);paletteModule?.cancel();
 // Never display the old palette as though it were the newly selected one.
 liveMode(true);for(const host of [$('#rc-live-slide'),$('#rc-live-zoom')]){host.replaceChildren();host.style.background=(palettes.find(p=>p.key===palette)?.bg||'#1a1b20');host.setAttribute('aria-busy','true');}
 $('#rc-quality').hidden=false;$('#rc-quality').textContent='Обновляю пример…';
 try{
  const api=await paletteAPI();if(ticket!==requestID)return;
  const svg=await api.load(style,key,palette,page,background_mode,diagram_layout);if(ticket!==requestID)return;
  api.paint(svg,[$('#rc-live-slide'),$('#rc-live-zoom')]);
  for(const host of [$('#rc-live-slide'),$('#rc-live-zoom')])host.setAttribute('aria-busy','false');
  $('#rc-quality').hidden=true;
 }catch(error){
  if(ticket!==requestID||error.name==='AbortError')return;
  shownURL='';$('#rc-quality').textContent='Не удалось обновить пример. Нажмите, чтобы повторить';
  for(const host of [$('#rc-live-slide'),$('#rc-live-zoom')])host.setAttribute('aria-busy','false');
 }
}
function observeThumbnails(){
 thumbObserver?.disconnect();
 if('IntersectionObserver' in window){
  thumbObserver=new IntersectionObserver(entries=>{for(const e of entries)if(e.isIntersecting){e.target.src=e.target.dataset.src;thumbObserver.unobserve(e.target);}},{root:$('.rc-rail'),rootMargin:'0px 130px',threshold:0});
  $$('.rc-thumb img').forEach(im=>thumbObserver.observe(im));
 }else $$('.rc-thumb img').forEach(im=>{im.src=im.dataset.src;});
}
const group=()=>state.kind==='standard'?'standard':state.style;
const entries=()=>data.filter(r=>r.group===group());
const selected=()=>entries().find(r=>r.key===(state.kind==='standard'?state.ordinary:state.designer));
function save(){
 const snapshot=JSON.stringify(state);if(snapshot===lastSaved)return;
 lastSaved=snapshot;sent.add(snapshot);if(sent.size>32)sent.delete(sent.values().next().value);
 try{localStorage.setItem('alix-hd-preview',snapshot);}catch{}
 const r=selected();
 window.openai?.setWidgetState?.({modelContent:{preview:'Реальные шаблоны AlixDoc',template:r.name,template_id:r.id,slide_count:state.slides,price:state.slides*(state.kind==='standard'?16:32),confirmed:state.confirmed},privateContent:JSON.parse(snapshot)})?.catch(()=>{});
}
function render(scrollToChoice=false){
 const list=entries(),r=selected(),idx=list.indexOf(r);
 let rebuilt=false;
 $$('.rc-kinds button').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.kind===state.kind)));
 $('.rc-style').hidden=state.kind==='standard';
 $('#rc-palette-field').hidden=state.kind==='standard';
 $('#rc-palette').value=state.palette;
 $('#rc-background-mode').value=state.background_mode;$('#rc-diagram-layout').value=state.diagram_layout;
 const chosen=palettes.find(p=>p.key===state.palette),sample=$('#rc-palette-sample');sample.hidden=!chosen;if(chosen){sample.style.background=chosen.bg;sample.style.color=chosen.ink;sample.style.border='1px solid '+chosen.line;sample.textContent=chosen.name+' · Aa Бб 123';}
 $('#rc-diagrams').hidden=state.kind==='standard';
 $$('.rc-style button').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.style===state.style)));
 if(groupRendered!==group()){
  const select=$('#rc-templates'),rail=$('.rc-rail');select.replaceChildren();rail.replaceChildren();
  for(const entry of list){
   const option=document.createElement('option');option.value=entry.key;option.textContent=entry.name;select.append(option);
   const b=document.createElement('button');b.type='button';b.className='rc-thumb cursor-interaction';b.dataset.key=entry.key;b.setAttribute('aria-label',entry.name);
   const im=document.createElement('img');im.dataset.src=entry.thumbs[0];im.alt='';im.loading='lazy';im.decoding='async';im.draggable=false;
   const label=document.createElement('span');label.textContent=entry.name;b.append(im,label);rail.append(b);
  }
  groupRendered=group();scrollToChoice=true;rebuilt=true;
 }
 $('#rc-templates').value=r.key;$('#rc-count').textContent=list.length+' вариантов';
 const img=$('#rc-slide');img.alt=r.name+' — '+slideNames[state.page];
 $('#rc-zoom-image').alt=img.alt;$('#rc-zoom-title').textContent=r.name+' · '+slideNames[state.page];showSlide(r);
 $$('.rc-slide-tabs button').forEach(b=>b.setAttribute('aria-pressed',String(+b.dataset.page===state.page)));
 $('#rc-name').textContent=r.name;$('#rc-position').textContent=(idx+1)+' / '+list.length;
 $$('.rc-thumb').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.key===r.key)));
 if(scrollToChoice){const rail=$('.rc-rail'),b=$$('.rc-thumb')[idx];rail.scrollLeft=Math.max(0,b.offsetLeft-rail.offsetLeft-(rail.clientWidth-b.clientWidth)/2);}
 if(rebuilt)observeThumbnails();
 $('#rc-pattern').textContent=state.kind==='standard'?'Единый фон на всех слайдах · без узоров':'Узор этого оформления · '+(state.style==='strict'?'строгий':'мягкий');
 $('#rc-slides').value=String(state.slides);$('#rc-price').textContent=state.slides*(state.kind==='standard'?16:32)+' кредитов';
 $('#rc-confirm').textContent='Посмотреть план';
 $('#rc-plan-hint').textContent='Заголовки и примерное расположение текста — в меню. Полный текст и картинки создадим после подтверждения.';
 $('#rc-status').textContent=state.confirmed?'Выбрано: '+r.name+'. Это просмотр — создание файла не запускается.':'';
}
function change(edit,scroll=false){const before=JSON.stringify(state);edit();state=normalize(state);if(before!==JSON.stringify(state)){render(scroll);save();}}
function choose(key){change(()=>{state[state.kind==='standard'?'ordinary':'designer']=key;state.page=0;state.confirmed=false;},true);}
root.addEventListener('click',async e=>{
 const b=e.target.closest('button');if(!b||!root.contains(b))return;
 if(b.dataset.kind){change(()=>{state.kind=b.dataset.kind;state.page=0;state.confirmed=false;});const url=new URL(location.href);url.searchParams.set('kind',state.kind);history.replaceState(null,'',url);}
 else if(b.dataset.style)change(()=>{state.style=b.dataset.style;state.page=0;state.confirmed=false;});
 else if(b.dataset.page!==undefined)change(()=>{state.page=+b.dataset.page;});
 else if(b.dataset.pageMove)change(()=>{state.page=(state.page+Number(b.dataset.pageMove)+4)%4;});
 else if(b.dataset.key)choose(b.dataset.key);
 else if(b.dataset.move){const list=entries();choose(list[(list.indexOf(selected())+Number(b.dataset.move)+list.length)%list.length].key);}
 else if(b.id==='rc-confirm'){
 const topic=$('#rc-topic').value.trim();if(!topic){$('#rc-status').textContent='Напишите тему презентации.';return;}
 if(b.disabled)return;
 b.disabled=true;try{outlineModule??=await import('./outline.js');await outlineModule.open({topic,kind:state.kind,style:state.style,key:selected().key,slides:state.slides,content_lang:window.AlixPreferences.content(),...(state.kind==='designer'?{background_mode:state.background_mode,diagram_layout:state.diagram_layout,...(state.palette?{palette:state.palette}:{})}:{})},b);}catch(e){$('#rc-status').textContent='Не удалось открыть план. Повторите попытку.';}finally{b.disabled=false;}
 }
 else if(b.id==='rc-zoom-open')$('#rc-zoom').showModal();
 else if(b.id==='rc-zoom-close')$('#rc-zoom').close();
});
for(const p of palettes){const o=document.createElement('option');o.value=p.key;o.textContent=p.name;$('#rc-palette').append(o)}
$('#rc-background-mode').addEventListener('change',e=>change(()=>{state.background_mode=e.target.value;state.confirmed=false}));
$('#rc-diagram-layout').addEventListener('change',e=>change(()=>{state.diagram_layout=e.target.value;state.page=3;state.confirmed=false}));
$('#rc-palette').addEventListener('change',e=>change(()=>{state.palette=e.target.value;state.confirmed=false}));
$('#rc-templates').addEventListener('change',e=>choose(e.target.value));
$('#rc-slides').addEventListener('change',e=>change(()=>{state.slides=+e.target.value;state.confirmed=false;}));
$('.rc-stage').addEventListener('pointerdown',e=>{touchStart={x:e.clientX,y:e.clientY};});
$('.rc-stage').addEventListener('pointercancel',()=>{touchStart=null;});
$('.rc-stage').addEventListener('pointerup',e=>{if(!touchStart)return;const dx=e.clientX-touchStart.x,dy=e.clientY-touchStart.y;touchStart=null;if(Math.abs(dx)>35&&Math.abs(dx)>Math.abs(dy))change(()=>{state.page=(state.page+(dx<0?1:3))%4;});});
$('#rc-quality').addEventListener('click',()=>showSlide(selected()));
$('#rc-zoom').addEventListener('keydown',e=>{if(e.key==='ArrowRight'||e.key==='ArrowLeft'){e.preventDefault();change(()=>{state.page=(state.page+(e.key==='ArrowRight'?1:3))%4;});}});
window.addEventListener('pagehide',()=>{clearTimeout(idleTimer);paletteModule?.cancel();previews.destroy();thumbObserver?.disconnect();});
window.addEventListener('pageshow',e=>{if(e.persisted)window.location.reload();});
window.addEventListener('openai:set_globals',e=>{
 const incoming=e.detail?.globals?.widgetState?.privateContent;
 if(!incoming||incoming.version!==3||sent.has(JSON.stringify(incoming)))return;
 const next=normalize(incoming);if(JSON.stringify(next)!==JSON.stringify(state)){state=next;render();}
});
render();save();
 outlineModule=await import('./outline.js');outlineModule.resume($('#rc-confirm'));
})().catch(()=>{document.getElementById('rc-status').textContent='Не удалось открыть каталог. Обновите страницу.';});
