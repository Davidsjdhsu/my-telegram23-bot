(async()=>{
const root=document.getElementById('alix-real-catalog');
const response=await fetch('catalog.json',{cache:'no-cache'});
if(!response.ok)throw new Error('Catalog unavailable');
const data=await response.json();
const $=s=>root.querySelector(s), $$=s=>Array.from(root.querySelectorAll(s));
const initial={version:3,kind:'standard',style:'strict',ordinary:'coffee',designer:'d16',page:0,slides:6,confirmed:false};
function normalize(s){
 s=s&&s.version===3?s:{};
 const n={...initial,...s};
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
let requestID=0,shownURL='',idleTimer=null,thumbObserver=null;
function paint(url){$('#rc-slide').src=url;$('#rc-zoom-image').src=url;}
function showSlide(r){
 const url=r.images[state.page];if(shownURL===url)return;
 shownURL=url;const ticket=++requestID;clearTimeout(idleTimer);
 paint(r.thumbs[state.page]);$('#rc-quality').textContent='Загружаю чёткий слайд…';
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
 window.openai?.setWidgetState?.({modelContent:{preview:'Реальные шаблоны AlixDoc',template:r.name,template_id:r.id,slide_count:state.slides,price:state.slides*(state.kind==='standard'?5:10),confirmed:state.confirmed},privateContent:JSON.parse(snapshot)})?.catch(()=>{});
}
function render(scrollToChoice=false){
 const list=entries(),r=selected(),idx=list.indexOf(r);
 let rebuilt=false;
 $$('.rc-kinds button').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.kind===state.kind)));
 $('.rc-style').hidden=state.kind==='standard';
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
 $('#rc-slides').value=String(state.slides);$('#rc-price').textContent=state.slides*(state.kind==='standard'?5:10)+' кредитов';
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
 if(b.disabled)return;b.disabled=true;$('#rc-status').textContent='Передаю запрос боту…';
 const payload={topic,kind:state.kind,style:state.style,key:selected().key,slides:state.slides};
 const fingerprint=JSON.stringify(payload);if(window.studioLastRequest?.fingerprint!==fingerprint)window.studioLastRequest={fingerprint,id:crypto.randomUUID()};
 try{const r=await fetch('/studio/generate',{method:'POST',headers:{'Content-Type':'application/json','X-Studio-Request':'1'},body:JSON.stringify({...payload,requestId:window.studioLastRequest.id})});if(!r.ok)throw Error(r.status===401?'Сессия закончилась. Откройте «Меню» заново.':r.status===409?'Бот уже обрабатывает запрос. Дождитесь результата.':'Не удалось передать запрос. Повторите попытку.');$('#rc-status').textContent='Запрос принят. Вернитесь в чат бота — там появится план.';}catch(e){$('#rc-status').textContent=e.message;}finally{b.disabled=false;}
 }
 else if(b.id==='rc-zoom-open')$('#rc-zoom').showModal();
 else if(b.id==='rc-zoom-close')$('#rc-zoom').close();
});
$('#rc-templates').addEventListener('change',e=>choose(e.target.value));
$('#rc-slides').addEventListener('change',e=>change(()=>{state.slides=+e.target.value;state.confirmed=false;}));
$('.rc-stage').addEventListener('pointerdown',e=>{touchStart={x:e.clientX,y:e.clientY};});
$('.rc-stage').addEventListener('pointercancel',()=>{touchStart=null;});
$('.rc-stage').addEventListener('pointerup',e=>{if(!touchStart)return;const dx=e.clientX-touchStart.x,dy=e.clientY-touchStart.y;touchStart=null;if(Math.abs(dx)>35&&Math.abs(dx)>Math.abs(dy))change(()=>{state.page=(state.page+(dx<0?1:3))%4;});});
$('#rc-quality').addEventListener('click',()=>showSlide(selected()));
$('#rc-zoom').addEventListener('keydown',e=>{if(e.key==='ArrowRight'||e.key==='ArrowLeft'){e.preventDefault();change(()=>{state.page=(state.page+(e.key==='ArrowRight'?1:3))%4;});}});
window.addEventListener('pagehide',()=>{clearTimeout(idleTimer);previews.destroy();thumbObserver?.disconnect();});
window.addEventListener('pageshow',e=>{if(e.persisted)window.location.reload();});
window.addEventListener('openai:set_globals',e=>{
 const incoming=e.detail?.globals?.widgetState?.privateContent;
 if(!incoming||incoming.version!==3||sent.has(JSON.stringify(incoming)))return;
 const next=normalize(incoming);if(JSON.stringify(next)!==JSON.stringify(state)){state=next;render();}
});
render();save();
})().catch(()=>{document.getElementById('rc-status').textContent='Не удалось открыть каталог. Обновите страницу.';});
