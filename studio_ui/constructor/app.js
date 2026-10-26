(()=>{
'use strict';const S=window.Scene,D=window.Diagrams,$=s=>document.querySelector(s),$$=s=>Array.from(document.querySelectorAll(s));
const storageKey=new URLSearchParams(location.search).has('qa')?'alix-constructor-qa-v1':'alix-constructor-v1';
let pendingImport=null,importing=false;
let deck=S.initial(),slideIndex=0,selected=null,undo=[],redo=[],drag=null,saveTimer,thumbTimer,exporting=false;
try{const saved=localStorage.getItem(storageKey);if(saved)deck=S.validate(JSON.parse(saved));}catch{}
const canvas=$('#canvas'),ctx=document.createElement('canvas').getContext('2d');
const current=()=>deck.slides[slideIndex],element=()=>current().elements.find(e=>e.id===selected);
const measure=(t,e)=>{ctx.font=`${e.italic?'italic ':''}${e.bold?'bold ':''}${e.size}px "${e.font}"`;return ctx.measureText(t).width;};
const status=t=>$('#status').textContent=t;
function snapshot(){return JSON.stringify({deck,slideIndex,selected});}
function remember(){undo.push(snapshot());if(undo.length>45)undo.shift();redo=[];}
function save(){clearTimeout(saveTimer);const previousStatus=$('#status').textContent;saveTimer=setTimeout(()=>{try{localStorage.setItem(storageKey,JSON.stringify(deck));if(!exporting&&$('#status').textContent===previousStatus)status('Изменения сохранены в этом браузере');}catch{status('Память устройства заполнена. Скачайте копию через меню «Проект».');}},450);}
function historyButtons(){$('#undo').disabled=!undo.length;$('#redo').disabled=!redo.length;}
const diagramMeasure=(t,size)=>measure(t,{size,font:'Arial'});
function hasOverflow(e){return e.type==='diagram'?D.overflow(e,diagramMeasure):e.type==='text'&&S.wrap(e,measure).length*e.size*1.22>e.h+5;}
function update(inspector=true){
 $('#import-report').hidden=!deck.importWarnings?.length;
 canvas.innerHTML=S.inner(current(),{prefix:'main',selected,measure});
 let bill;try{bill=S.quote(deck);}catch(error){bill={slides:deck.slides.map(()=> '—'),cost:'—'};status(error.message);}
 $('#slide-label').textContent=`Слайд ${slideIndex+1} из ${deck.slides.length} · ${bill.slides[slideIndex]} кредитов`;
 $('#total-price').textContent=`Всего: ${bill.cost} кредитов`;
 $('#export').textContent=`Получить · ${bill.cost} кредитов`;
 const overflow=current().elements.filter(hasOverflow);
 $('#warning').textContent=overflow.length?'Текст не помещается в рамку. Увеличьте рамку или уменьшите шрифт.':'';
 $('#delete-slide').disabled=deck.slides.length===1;$('#slide-up').disabled=slideIndex===0;$('#slide-down').disabled=slideIndex===deck.slides.length-1;$('#add-slide').disabled=deck.slides.length>=20;
 $('#edit-text').hidden=element()?.type!=='text';
 if(inspector)properties();layers();historyButtons();clearTimeout(thumbTimer);thumbTimer=setTimeout(thumbnails,120);
 if($('#preview-dialog').open)renderPreview();
}
const phone=matchMedia('(max-width:740px)');
function setInspector(open){
 if(open&&!document.body.classList.contains('inspector-open'))$('.inspector-scroll').scrollTop=0;
 document.body.classList.toggle('inspector-open',open&&phone.matches);
 $('#edit-element').setAttribute('aria-expanded',String(open&&phone.matches));
 $('#inspector').inert=phone.matches&&!open;
 if(!open&&$('#inspector').contains(document.activeElement))$('#edit-element').focus({preventScroll:true});
}
phone.addEventListener('change',()=>setInspector(false));
setInspector(false);
const canvasArea=$('.canvas-area');
function fitCanvas(){
 if(!phone.matches){$('#canvas-wrap').style.width='';return;}
 const area=canvasArea.getBoundingClientRect();
 $('#canvas-wrap').style.width=Math.max(0,Math.min(area.width,area.height*16/9))+'px';
}
const canvasObserver=new ResizeObserver(fitCanvas);canvasObserver.observe(canvasArea);
phone.addEventListener('change',fitCanvas);
function thumbnails(){
 $('#slide-list').innerHTML=deck.slides.map((s,i)=>`<button class="slide-thumb ${i===slideIndex?'active':''}" data-slide="${i}" aria-label="Слайд ${i+1}" aria-pressed="${i===slideIndex}">${S.svg(s,{prefix:'thumb'+i,measure})}<span>${String(i+1).padStart(2,'0')}</span></button>`).join('');$('#slide-count').textContent=deck.slides.length;
}
function mutate(fn,inspector=true){remember();fn();update(inspector);save();}
function select(id){selected=id;update();}
const color=(label,key,value)=>`<label>${label}<div class="color-pair"><input type="color" data-prop="${key}" value="${value}" aria-label="${label}"><input type="text" data-hex="${key}" value="${value}" maxlength="7" aria-label="${label} HEX"></div></label>`;
const number=(label,key,value,min,max,step=1)=>`<label>${label}<input type="number" data-prop="${key}" value="${value}" min="${min}" max="${max}" step="${step}"></label>`;
const range=(label,key,value,min,max,step=1)=>`<label>${label}<output data-value="${key}">${value}</output><input type="range" aria-label="${label}" data-prop="${key}" value="${value}" min="${min}" max="${max}" step="${step}"></label>`;
function properties(){const e=element(),b=current().background;
 if(!e){$('#properties').innerHTML=`<h2>Фон слайда</h2><label>Заливка<select data-bg="type"><option value="solid" ${b.type==='solid'?'selected':''}>Один цвет</option><option value="gradient" ${b.type==='gradient'?'selected':''}>Градиент</option></select></label>${color('Основной цвет','bg-color',b.color)}${b.type==='gradient'?color('Второй цвет','bg-second',b.second)+range('Направление градиента','bg-angle',b.angle,0,360):''}<button id="apply-background" class="wide">Применить фон ко всем слайдам</button><div class="separator"></div><p class="empty">Нажмите на текст, фото или узор на слайде, чтобы настроить этот элемент.</p>`;return;}
 let html=`<h2>${{text:'Текст',photo:'Фотография',pattern:'Узор',diagram:'Схема'}[e.type]}</h2>`;
 if(e.type==='diagram')html+=`<label>Вид схемы<select data-prop="kind">${Object.entries(D.NAMES).map(([k,v])=>`<option value="${k}" ${e.kind===k?'selected':''}>${v}</option>`).join('')}</select></label><label>Подписи: один блок на строку<textarea data-prop="labels" maxlength="1100">${S.esc(e.labels.join('\n'))}</textarea></label><p class="file-note">От 2 до 7 блоков. Первая строка в дереве — главный блок.</p>${number('Размер подписей','size',e.size,16,42)}${color('Цвет блоков','fill',e.fill)}${color('Цвет связей','accent',e.accent)}${color('Цвет подписей','color',e.color)}`;
 if(e.type==='text')html+=`<label>Содержание<textarea data-prop="text" maxlength="3000">${S.esc(e.text)}</textarea></label><label>Шрифт<select data-prop="font">${S.FONTS.map(f=>`<option ${f===e.font?'selected':''}>${f}</option>`).join('')}</select></label><div class="grid2">${number('Размер шрифта','size',e.size,12,120)}<label>Выравнивание<select data-prop="align"><option value="left" ${e.align==='left'?'selected':''}>Слева</option><option value="center" ${e.align==='center'?'selected':''}>По центру</option><option value="right" ${e.align==='right'?'selected':''}>Справа</option></select></label></div><div class="row"><button data-toggle="bold" aria-pressed="${e.bold}"><b>Жирный</b></button><button data-toggle="italic" aria-pressed="${e.italic}"><i>Курсив</i></button></div>${color('Цвет текста','color',e.color)}`;
 if(e.type==='photo')html+=`<button id="replace-photo" class="wide">Заменить фотографию</button><p class="file-note">Файл остаётся на вашем компьютере.</p><label>Форма<select data-prop="shape"><option value="rect" ${e.shape==='rect'?'selected':''}>Прямоугольник</option><option value="round" ${e.shape==='round'?'selected':''}>Скруглённая</option><option value="circle" ${e.shape==='circle'?'selected':''}>Круг / овал</option></select></label>${range('Приближение','zoom',e.zoom,1,3,.05)}${range('Кадр по горизонтали','fx',e.fx,0,100)}${range('Кадр по вертикали','fy',e.fy,0,100)}${color('Цвет рамки','border',e.border)}${number('Толщина рамки','borderWidth',e.borderWidth,0,12)}`;
 if(e.type==='pattern')html+=`<div class="choice-patterns">${[['rings','Кольца'],['lines','Линии'],['dots','Точки'],['frame','Рамки']].map(([k,l])=>`<button data-pattern="${k}" aria-pressed="${e.kind===k}">${l}</button>`).join('')}</div>${color('Цвет узора','color',e.color)}${number('Толщина линий','weight',e.weight,1,12)}`;
 html+=`<div class="separator"></div><div class="grid2">${number('Слева','x',Math.round(e.x),0,1200-e.w)}${number('Сверху','y',Math.round(e.y),0,675-e.h)}${number('Ширина','w',Math.round(e.w),40,1200)}${number('Высота','h',Math.round(e.h),30,675)}</div>${e.type==='diagram'?'':range('Поворот','rotation',Math.round(e.rotation),-180,180)}${range('Непрозрачность','opacity',Math.round(e.opacity*100),0,100)}<div class="row"><button id="layer-down">На слой ниже</button><button id="layer-up">На слой выше</button></div><div class="row"><button id="duplicate-element">Копия</button><button id="delete-element" class="danger">Удалить</button></div>`;
 $('#properties').innerHTML=html;
}
function layers(){
 $('#layers').innerHTML=[...current().elements].reverse().map(e=>`<button class="layer ${e.id===selected?'active':''}" data-layer="${e.id}"><span>${{text:'Т',photo:'▧',pattern:'◉',diagram:'▦'}[e.type]}</span><span>${S.esc(e.type==='text'?e.text.slice(0,45):e.type==='photo'?'Фотография':e.type==='diagram'?'Схема · '+D.NAMES[e.kind]:'Узор · '+({rings:'кольца',lines:'линии',dots:'точки',frame:'рамки'}[e.kind]))}</span></button>`).join('');
}
function applyProperty(input){
 const e=element(),k=input.dataset.prop||input.dataset.hex,v=input.value;
 if(input.dataset.hex&&!/^#[\da-f]{6}$/i.test(v))return;
 if(k==='labels'){const labels=v.split('\n').map(x=>x.trim()).filter(Boolean);if(labels.length<2||labels.length>7){input.setCustomValidity('Нужно от 2 до 7 строк.');status('В схеме нужно от 2 до 7 непустых строк.');return;}input.setCustomValidity('');}
 remember();
 if(k.startsWith('bg-')){current().background[k.slice(3)]=k==='bg-angle'?+v:v;}
 else if(e){
  if(k==='labels')e.labels=v.split('\n').map(x=>x.trim()).filter(Boolean).map(x=>x.slice(0,160));
  else if(['text','font','color','border','align','shape','kind','fill','accent'].includes(k))e[k]=v;
  else if(k==='opacity')e[k]=+v/100;
  else e[k]=S.clamp(v,+input.min,+input.max);
  e.x=S.clamp(e.x,0,1200-e.w);e.y=S.clamp(e.y,0,675-e.h);
 }
 const out=$(`[data-value="${k}"]`);if(out)out.textContent=v;
 if(input.type==='color'){const hex=$(`[data-hex="${k}"]`);if(hex)hex.value=v;}
 if(input.dataset.hex){const picker=$(`[data-prop="${k}"]`);if(picker)picker.value=v;}
 update(false);save();
}
$('#properties').addEventListener('input',e=>{if(e.target.dataset.prop||e.target.dataset.hex)applyProperty(e.target);});
$('#properties').addEventListener('change',e=>{if(e.target.dataset.bg)mutate(()=>current().background[e.target.dataset.bg]=e.target.value);});
function add(el){try{S.quote({version:1,slides:[{elements:[...current().elements,el]}]});}catch(e){status(e.message);return;}mutate(()=>{current().elements.push(el);selected=el.id;});}
function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);}
function undoRedo(from,to){if(!from.length)return;to.push(snapshot());const old=JSON.parse(from.pop());deck=old.deck;slideIndex=old.slideIndex;selected=old.selected;$('#deck-name').value=deck.name;update();save();}
document.querySelector('.mode-switch').addEventListener('click',()=>{try{localStorage.setItem(storageKey,JSON.stringify(deck));}catch{}});
document.addEventListener('click',ev=>{
 const b=ev.target.closest('button');if(!b)return;
 if(b.closest('.project-menu'))b.closest('.project-menu').open=false;
 if(b.dataset.diagram){insertDiagram(b.dataset.diagram);$('#diagram-dialog').close();return;}
 if(b.dataset.slide!==undefined){window.AlixMobile?.enter($('#canvas-wrap'));slideIndex=+b.dataset.slide;selected=null;$('#layout').value='free';update();return;}
 if(b.dataset.layer){select(b.dataset.layer);return;}
 const e=element();
 if(b.dataset.toggle&&e){mutate(()=>e[b.dataset.toggle]=!e[b.dataset.toggle]);return;}
 if(b.dataset.pattern&&e){mutate(()=>e.kind=b.dataset.pattern);return;}
 switch(b.id){
 case 'undo':undoRedo(undo,redo);break;case 'redo':undoRedo(redo,undo);break;
 case 'deselect':select(null);setInspector(true);break;
 case 'edit-element':setInspector(!document.body.classList.contains('inspector-open'));break;
 case 'close-inspector':setInspector(false);break;
 case 'add-diagram':$('#diagram-options').innerHTML=Object.entries(D.NAMES).map(([k,v])=>`<button data-diagram="${k}"><svg viewBox="0 0 1060 400">${D.markup(D.create(k))}</svg><strong>${v}</strong></button>`).join('');$('#diagram-dialog').showModal();break;
 case 'close-diagram':$('#diagram-dialog').close();break;
 case 'add-text':editText(true);break;
 case 'edit-text':editText(false);break;
 case 'cancel-text':$('#text-dialog').close();break;
 case 'close-saved':$('#saved-dialog').close();break;
 case 'add-pattern':add(S.pattern(800,310,310,280));setInspector(true);break;
 case 'add-photo':photoMode='add';$('#photo-file').click();break;
 case 'replace-photo':photoMode='replace';$('#photo-file').click();break;
 case 'apply-background':mutate(()=>{const bg=structuredClone(current().background);deck.slides.forEach(s=>s.background=structuredClone(bg));});break;
 case 'duplicate-element':if(e)add({...structuredClone(e),id:S.uid(),x:Math.min(e.x+20,1200-e.w),y:Math.min(e.y+20,675-e.h)});break;
 case 'delete-element':if(e)mutate(()=>{current().elements=current().elements.filter(x=>x!==e);selected=null;});break;
 case 'layer-up':case 'layer-down':if(e)mutate(()=>{const a=current().elements,i=a.indexOf(e),j=S.clamp(i+(b.id==='layer-up'?1:-1),0,a.length-1);[a[i],a[j]]=[a[j],a[i]];});break;
 case 'add-slide':if(deck.slides.length<20)mutate(()=>{deck.slides.push({id:S.uid(),background:structuredClone(current().background),elements:[S.text('Новый слайд',80,90,1000,150,60)]});slideIndex=deck.slides.length-1;selected=null;});break;
 case 'duplicate-slide':if(deck.slides.length<20)mutate(()=>{const s=structuredClone(current());s.id=S.uid();s.elements.forEach(e=>e.id=S.uid());deck.slides.splice(slideIndex+1,0,s);slideIndex++;selected=null;});break;
 case 'delete-slide':if(deck.slides.length>1)mutate(()=>{deck.slides.splice(slideIndex,1);slideIndex=Math.min(slideIndex,deck.slides.length-1);selected=null;});break;
 case 'slide-up':case 'slide-down':{const j=slideIndex+(b.id==='slide-up'?-1:1);if(j>=0&&j<deck.slides.length)mutate(()=>{[deck.slides[j],deck.slides[slideIndex]]=[deck.slides[slideIndex],deck.slides[j]];slideIndex=j;});break;}
 case 'save-project':try{deck.name=$('#deck-name').value.slice(0,80);clearTimeout(saveTimer);localStorage.setItem(storageKey,JSON.stringify(deck));status('Проект сохранён на этом устройстве. Можно закрыть меню и продолжить позже.');$('#saved-dialog').showModal();}catch{status('Не удалось сохранить. Скачайте копию через меню «Проект».');}break;
 case 'backup-project':download(new Blob([JSON.stringify(deck)],{type:'application/octet-stream'}),'AlixDoc-project.alixdoc');status('Копия проекта скачана. Для редактирования выберите «Открыть копию проекта».');break;
 case 'open-ready':window.AlixMobile.openFiles();break;
 case 'import-pptx':if(!importing)$('#pptx-file').click();break;
 case 'import-report':window.alert((deck.importWarnings||[]).join('\n\n'));break;
 case 'cancel-import':pendingImport=null;$('#import-dialog').close();break;
 case 'accept-import':if(pendingImport){const next=pendingImport;mutate(()=>{deck=next;slideIndex=0;selected=null;$('#deck-name').value=deck.name;});pendingImport=null;$('#import-dialog').close();status('PPTX открыт как копия. Проверьте все слайды. Текст и фотографии можно менять.');}break;
 case 'open-project':$('#project-file').click();break;
 case 'mobile-preview':case 'preview':renderPreview();$('#preview-dialog').showModal();break;
 case 'close-preview':$('#preview-dialog').close();break;
 case 'prev-preview':case 'next-preview':slideIndex=(slideIndex+(b.id==='next-preview'?1:deck.slides.length-1))%deck.slides.length;selected=null;update();break;
 case 'export':exportDeck();break;
 }
});
let textTarget=null;
function editText(isNew){
 const e=element();if(!isNew&&e?.type!=='text')return;
 textTarget=isNew?null:e.id;
 $('#text-input').value=isNew?'':e.text;$('#text-counter').textContent=$('#text-input').value.length+' / 3000';
 $('#text-dialog').showModal();$('#text-input').focus();
}
$('#text-input').addEventListener('input',()=>$('#text-counter').textContent=$('#text-input').value.length+' / 3000');
$('#text-form').addEventListener('submit',event=>{
 event.preventDefault();const value=$('#text-input').value;if(!value.trim()){ $('#text-input').focus();return; }
 if(textTarget){const target=current().elements.find(e=>e.id===textTarget&&e.type==='text');if(target)mutate(()=>target.text=value);}
 else{const t=S.text(value,100,180,900,350,32);while(hasOverflow(t)&&t.size>16)t.size-=2;add(t);}
 $('#text-dialog').close();status('Текст добавлен. Цвет, шрифт и положение можно изменить в настройках.');
});
canvas.addEventListener('dblclick',()=>editText(false));
function insertDiagram(kind,palette='strict'){if(deck.slides.length>=20){status('В примере можно создать до 20 слайдов.');return;}mutate(()=>{const p=D.PALETTES[palette]||D.PALETTES.strict,d=D.create(kind,palette);deck.slides.push({id:S.uid(),background:{type:'solid',color:p.bg,second:p.bg,angle:0},elements:[S.text('Схема: '+D.NAMES[kind],70,60,1060,90,48,p.color,'Georgia'),d]});slideIndex=deck.slides.length-1;selected=d.id;});}
function renderPreview(){$('#preview-canvas').innerHTML=S.svg(current(),{prefix:'present',measure});$('#preview-name').textContent=`${deck.name} · ${slideIndex+1} / ${deck.slides.length}`;}
$('#deck-name').value=deck.name;$('#deck-name').addEventListener('input',e=>mutate(()=>deck.name=e.target.value.slice(0,80),false));
$('#layout').addEventListener('change',ev=>{
 const mode=ev.target.value;if(mode==='free')return;
 mutate(()=>{const p=current().elements.find(e=>e.type==='photo'),texts=current().elements.filter(e=>e.type==='text');
  if(mode==='cover'){if(p){Object.assign(p,{x:0,y:0,w:1200,h:675,shape:'rect',borderWidth:0,opacity:.4});current().elements=current().elements.filter(e=>e!==p);current().elements.unshift(p);}texts.forEach((e,i)=>Object.assign(e,{x:80,y:i?365:160,w:1030,h:i?200:170,color:'#ffffff'}));}
  else{if(p)Object.assign(p,{x:mode==='right'?650:75,y:135,w:475,h:405,opacity:1});texts.forEach((e,i)=>Object.assign(e,{x:mode==='right'?75:600,y:i?365:145,w:500,h:i?205:180}));}
 });
});
function point(ev){const r=canvas.getBoundingClientRect();return {x:(ev.clientX-r.left)*1200/r.width,y:(ev.clientY-r.top)*675/r.height};}
canvas.addEventListener('pointerdown',ev=>{
 if(ev.button!==0)return;const resize=ev.target.closest('[data-resize]'),g=ev.target.closest('[data-id]'),id=resize?.dataset.resize||g?.dataset.id;if(!id){select(null);return;}
 select(id);const p=point(ev),e=element();drag={id,start:p,original:structuredClone(e),before:snapshot(),resize:!!resize,moved:false};canvas.setPointerCapture(ev.pointerId);ev.preventDefault();
});
canvas.addEventListener('pointermove',ev=>{
 if(!drag)return;const p=point(ev),e=element();if(!e)return;const dx=p.x-drag.start.x,dy=p.y-drag.start.y;if(Math.abs(dx)+Math.abs(dy)<2&&!drag.moved)return;drag.moved=true;
 if(drag.resize){const a=-drag.original.rotation*Math.PI/180,rx=dx*Math.cos(a)-dy*Math.sin(a),ry=dx*Math.sin(a)+dy*Math.cos(a);e.w=S.clamp(drag.original.w+rx,40,1200-e.x);e.h=S.clamp(drag.original.h+ry,30,675-e.y);}
 else{e.x=S.clamp(drag.original.x+dx,0,1200-e.w);e.y=S.clamp(drag.original.y+dy,0,675-e.h);if(Math.abs(e.x+e.w/2-600)<7)e.x=600-e.w/2;if(Math.abs(e.y+e.h/2-337.5)<7)e.y=337.5-e.h/2;}
 canvas.innerHTML=S.inner(current(),{prefix:'main',selected,measure});
});
function endDrag(){if(!drag)return;if(drag.moved){undo.push(drag.before);if(undo.length>45)undo.shift();redo=[];save();}drag=null;update();}
canvas.addEventListener('pointerup',endDrag);canvas.addEventListener('pointercancel',endDrag);
document.addEventListener('keydown',ev=>{
 if(/INPUT|TEXTAREA|SELECT/.test(ev.target.tagName))return;
 if((ev.ctrlKey||ev.metaKey)&&ev.key.toLowerCase()==='z'){ev.preventDefault();undoRedo(ev.shiftKey?redo:undo,ev.shiftKey?undo:redo);return;}
 const e=element();if(!e)return;
 if(ev.key==='Delete'||ev.key==='Backspace'){ev.preventDefault();mutate(()=>{current().elements=current().elements.filter(x=>x!==e);selected=null;});}
 if(['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(ev.key)){ev.preventDefault();mutate(()=>{const d=ev.shiftKey?10:1;e.x=S.clamp(e.x+(ev.key==='ArrowLeft'?-d:ev.key==='ArrowRight'?d:0),0,1200-e.w);e.y=S.clamp(e.y+(ev.key==='ArrowUp'?-d:ev.key==='ArrowDown'?d:0),0,675-e.h);});}
});
let photoMode='add';const readFile=f=>new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result);r.onerror=reject;r.readAsDataURL(f);});
const loadImage=src=>new Promise((resolve,reject)=>{const im=new Image();im.onload=()=>resolve(im);im.onerror=()=>reject(Error('Не удалось открыть изображение'));im.src=src;});
$('#photo-file').addEventListener('change',async ev=>{const f=ev.target.files[0];ev.target.value='';if(!f)return;if(f.size>15000000){status('Выберите фото до 15 МБ.');return;}try{const im=await loadImage(await readFile(f)),ratio=Math.min(1,1800/Math.max(im.width,im.height)),c=document.createElement('canvas');c.width=Math.round(im.width*ratio);c.height=Math.round(im.height*ratio);c.getContext('2d').drawImage(im,0,0,c.width,c.height);const src=c.toDataURL('image/jpeg',.9);
 const target=photoMode==='replace'&&element()?.type==='photo'?element():null;if(target)mutate(()=>Object.assign(target,{src,iw:c.width,ih:c.height,fx:50,fy:50,zoom:1}));else add({...S.photo(630,140,480,390),src,iw:c.width,ih:c.height});}catch(e){status(e.message);}});
$('#pptx-file').addEventListener('change',async ev=>{
 const file=ev.target.files[0];ev.target.value='';if(!file||importing)return;
 if(!/\.pptx$/i.test(file.name)||file.size>10000000){status('Выберите файл .pptx до 10 МБ.');return;}
 importing=true;$('#import-pptx').disabled=true;status('Открываю PPTX. Текущий проект пока сохранён…');
 const controller=new AbortController(),timeout=setTimeout(()=>controller.abort(),45000);
 try{
  const response=await fetch('/studio/import-pptx',{method:'POST',headers:{'Content-Type':'application/octet-stream','X-Studio-Request':'1','X-Presentation-Name':encodeURIComponent(file.name)},body:file,signal:controller.signal});
  if(response.status===401)throw Error('Откройте «Меню» заново в Telegram.');
  const result=await response.json();if(!response.ok)throw Error(result.error||'Не удалось открыть PPTX.');
  const next=S.validate(result.deck);S.quote(next);pendingImport=next;
  $('#import-notice').textContent=result.notice;
  $('#import-summary').textContent=`Слайдов: ${next.slides.length}. Сборка после редактирования: ${result.price.cost} кредитов, с отдельным подтверждением.`;
  $('#import-warnings').replaceChildren(...result.warnings.map(text=>{const li=document.createElement('li');li.textContent=text;return li;}));
  $('#import-dialog').showModal();status('Файл проверен. Подтвердите открытие копии.');
 }catch(error){status(error.name==='AbortError'?'Не удалось открыть файл за 45 секунд. Проект не изменён.':error.message);}
 finally{clearTimeout(timeout);importing=false;$('#import-pptx').disabled=false;}
});
$('#project-file').addEventListener('change',async ev=>{const f=ev.target.files[0];ev.target.value='';if(!f)return;try{if(f.size>15000000)throw Error('Проект слишком большой.');const next=S.validate(JSON.parse(await f.text()));mutate(()=>{deck=next;slideIndex=0;selected=null;$('#deck-name').value=deck.name;});}catch(e){status(e.message||'Не удалось открыть проект.');}});
async function rasterElement(e){
 const c=document.createElement('canvas'),scale=2;c.width=Math.round(e.w*scale);c.height=Math.round(e.h*scale);const g=c.getContext('2d');g.scale(scale,scale);g.globalAlpha=e.opacity;
 if(e.type==='photo'){const im=await loadImage(e.src),box=S.crop(e);g.beginPath();if(e.shape==='circle')g.ellipse(e.w/2,e.h/2,e.w/2,e.h/2,0,0,Math.PI*2);else if(e.shape==='round')g.roundRect(0,0,e.w,e.h,e.radius);else g.rect(0,0,e.w,e.h);g.save();g.clip();g.drawImage(im,box.x,box.y,box.w,box.h);g.restore();if(e.borderWidth){g.strokeStyle=e.border;g.lineWidth=e.borderWidth;g.stroke();}}
 else{const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${e.w}" height="${e.h}" viewBox="0 0 ${e.w} ${e.h}">${S.patternMarkup(e)}</svg>`;const im=await loadImage('data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg));g.drawImage(im,0,0,e.w,e.h);}
 return c.toDataURL('image/png');
}
async function exportDeck(){
 if(deck.importWarnings?.length&&!window.confirm('При импорте были отличия ('+deck.importWarnings.length+'). Проверьте слайды перед сборкой. Продолжить?'))return;
 if(exporting)return;if($('#properties textarea:invalid')){status('Сначала исправьте количество блоков схемы: от 2 до 7.');return;}const bad=deck.slides.findIndex(s=>s.elements.some(hasOverflow));if(bad>=0){slideIndex=bad;selected=current().elements.find(hasOverflow).id;update();status('Сначала поправьте текст, который не помещается в рамку.');return;}
 let bill;try{bill=S.quote(deck);}catch(e){status(e.message);return;}
 if(!window.confirm(`Стоимость каждого слайда: ${bill.slides.join(' + ')} кредитов.\nВсего: ${bill.cost} кредитов.\nОснова — 10, фото сверх первого — +3, узор — +2.\nСписание после отправки PPTX. Для администратора действует безлимит.\nСобрать презентацию?`))return;
 exporting=true;$('#export').disabled=true;$('#open-ready').hidden=true;status('Готовлю PPTX с вашими настройками…');
 try{const payload=structuredClone(deck);payload.confirmedCost=bill.cost;payload.pricingVersion=bill.version;for(const s of payload.slides)for(const e of s.elements){if(e.type==='text')e.lines=S.wrap(e,measure);else if(e.type==='diagram')e.parts=D.parts(e,diagramMeasure);else e.raster=await rasterElement(e);}
 const response=await fetch('/studio/export',{method:'POST',headers:{'Content-Type':'application/json','X-Studio-Request':'1'},body:JSON.stringify(payload)});if(!response.ok)throw Error((await response.json()).error||'Не удалось собрать PPTX');const receipt=await response.json();clearTimeout(saveTimer);status(receipt.duplicate?'Этот файл уже отправлен в чат. Повторного списания нет.':`Готово. Презентация в чате. Списано ${receipt.charged} кредитов.`);$('#open-ready').hidden=false;}catch(e){status('Не удалось получить презентацию: '+e.message);}finally{exporting=false;$('#export').disabled=false;}
}
window.addEventListener('beforeunload',()=>{try{localStorage.setItem(storageKey,JSON.stringify(deck));}catch{}});
const params=new URLSearchParams(location.search),incoming=params.get('diagram');if(Object.keys(D.NAMES).includes(incoming)){insertDiagram(incoming,params.get('palette')||'strict');params.delete('diagram');params.delete('palette');history.replaceState(null,'',location.pathname+(params.size?'?'+params:'')+location.hash);}
update();thumbnails();document.fonts.ready.then(()=>update(false));
})();
