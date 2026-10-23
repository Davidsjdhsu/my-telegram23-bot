// A lightweight plan preview: no body copy, image generation or canvas rendering.
let dialog, current, timer, generation=0, returnFocus, currentPayload;
const memoryKey='alix-outline-preview';
const el=(tag,cls,text)=>{const n=document.createElement(tag);if(cls)n.className=cls;if(text)n.textContent=text;return n};
async function request(url,body){
 const controller=new AbortController(),timeout=setTimeout(()=>controller.abort(),15000);
 try{
  const r=await fetch(url,{method:body?'POST':'GET',headers:body?{'Content-Type':'application/json','X-Studio-Request':'1'}:{},body:body?JSON.stringify(body):undefined,cache:'no-store',signal:controller.signal});
  if(!r.ok){const raw=await r.text();let message;try{message=JSON.parse(raw).error}catch{}throw Error(message||(r.status===401?'Откройте «Меню» заново в Telegram.':r.status===404?'План устарел. Составьте его заново.':r.status===402?'Недостаточно кредитов. Пополните баланс в профиле.':r.status===409?'Предыдущая работа ещё выполняется. Попробуйте позже.':'Не удалось связаться с сервером. Повторите попытку.'));}
  return await r.json();
 }finally{clearTimeout(timeout)}
}
function remember(id){try{sessionStorage.setItem(memoryKey,id)}catch{}}
function forget(){try{sessionStorage.removeItem(memoryKey)}catch{}}
function close(){clearTimeout(timer);generation++;if(['completed','build_failed'].includes(current?.status))forget();dialog.close();returnFocus?.focus();}
function shell(){
 if(dialog)return;
 dialog=el('dialog','op-dialog');dialog.setAttribute('aria-label','План вашей презентации');
 const header=el('header','op-header'),caption=el('div');caption.append(el('small','','Предпросмотр'),el('h2','','Ваша презентация'));
 const back=el('button','','Закрыть ×');back.type='button';back.onclick=close;header.append(caption,back);
 dialog.append(header,el('div','op-body'),el('footer','op-footer'));document.body.append(dialog);
 dialog.addEventListener('cancel',e=>{e.preventDefault();close()});
}
function footer(message){const f=dialog.querySelector('footer');f.replaceChildren(el('p','op-notice',message));return f;}
function failure(message,retry,label='Повторить'){
 const f=footer(message),b=el('button','op-primary',label);b.onclick=()=>{b.disabled=true;retry()};f.append(b);
}
function busy(){current=null;dialog.querySelector('.op-body').replaceChildren(el('p','op-loading','Составляем названия и порядок слайдов…'));footer('Сейчас создаётся только короткий план. Полный текст и изображения — после подтверждения.');}
function wave(width){
 const ns='http://www.w3.org/2000/svg',svg=document.createElementNS(ns,'svg');svg.setAttribute('viewBox','0 0 300 10');svg.setAttribute('preserveAspectRatio','none');svg.setAttribute('aria-hidden','true');svg.style.width=width+'%';
 const path=document.createElementNS(ns,'path');path.setAttribute('d','M0 5 Q5 1 10 5 T20 5 T30 5 T40 5 T50 5 T60 5 T70 5 T80 5 T90 5 T100 5 T110 5 T120 5 T130 5 T140 5 T150 5 T160 5 T170 5 T180 5 T190 5 T200 5 T210 5 T220 5 T230 5 T240 5 T250 5 T260 5 T270 5 T280 5 T290 5 T300 5');svg.append(path);return svg;
}
function paint(row){
 const unchanged=current?.id===row.id&&dialog.querySelector('.op-deck');current=row;
 if(!unchanged){const body=dialog.querySelector('.op-body');body.replaceChildren();
 body.append(el('h3','op-title',row.title),el('p','op-notice','Заголовки — ваш план. Волнистые линии показывают примерное место и объём текста. Расположение может уточниться при сборке.'));
 const list=el('ol','op-deck');
 row.titles.forEach((title,i)=>{
  const item=el('li','op-item'),card=el('div','op-slide'+(i===0?' op-cover':''));
  for(const k of ['bg','ink'])if(/^#[0-9a-f]{6}$/i.test(row.palette?.[k]||''))card.style.setProperty('--op-'+k,row.palette[k]);
  if(['Calibri','Georgia','Trebuchet MS'].includes(row.palette?.font))card.style.setProperty('--op-font','"'+row.palette.font+'", '+(row.palette.font==='Georgia'?'serif':'sans-serif'));
  card.append(el('h4','',title));
  const content=el('div','op-composition'),lines=el('div','op-lines');lines.setAttribute('aria-label','Место для основного текста');
  (i===0?[90,65]:[100,92,98,75,96,60]).forEach(w=>lines.append(wave(w)));
  content.append(lines,el('div','op-photo','Место для изображения'));card.append(content);
  item.append(el('span','op-number','Слайд '+(i+1)),card);list.append(item);
 });body.append(list);}
 clearTimeout(timer);
 if(['building','completed','build_failed'].includes(row.status)){
  const messages={building:['Собираем презентацию…','Создаём текст и изображения, затем проверим и отправим PPTX и PDF. Можно оставаться здесь — статус обновится автоматически.'],completed:['Готово ✓','Презентация собрана. PPTX и PDF уже отправлены в чат бота.'],build_failed:['Не удалось завершить сборку',row.error||'Проверьте сообщения бота перед повторной попыткой.']};
  const [title,note]=messages[row.status],f=footer(note),heading=el('div','op-build-state '+row.status,title);heading.setAttribute('role','status');heading.setAttribute('aria-live','polite');f.prepend(heading);
  if(row.status==='building'){
   remember(row.id);timer=setTimeout(()=>poll(row.id,generation),2500);
  }else{
   const button=el('button','op-primary',row.status==='completed'?'Открыть файл':'Вернуться в чат');button.type='button';
   button.onclick=()=>window.AlixMobile.openFiles();f.append(button);
   // Keep the result available after a reload until the user closes it.
   remember(row.id);
  }
  return;
 }
 const f=footer('Стоимость — '+row.cost+' кредитов. Списание после успешной выдачи PPTX и PDF.');
 const confirm=el('button','op-primary','Собрать · '+row.cost+' кредитов');confirm.type='button';
 confirm.onclick=async()=>{confirm.disabled=true;confirm.textContent='Запускаем сборку…';const ticket=generation;
  try{const next=await request('/studio/outline/'+encodeURIComponent(row.id)+'/confirm',{});if(ticket===generation)paint(next)}
  catch(e){if(ticket===generation){paint(row);dialog.querySelector('footer').prepend(el('p','op-error',e.message));}}
 };f.append(confirm);
 const edit=el('button','op-secondary','Изменить запрос');edit.type='button';edit.onclick=()=>{forget();close();document.getElementById('rc-topic')?.focus()};f.append(edit);
}
async function poll(id,ticket){
 try{const row=await request('/studio/outline/'+encodeURIComponent(id));if(ticket!==generation)return;
  if(row.status==='planning'||row.status==='submitting'){timer=setTimeout(()=>poll(id,ticket),1800);return;}
  if(row.status==='failed'){forget();failure(row.error,()=>{close();if(currentPayload)open(currentPayload,returnFocus);else document.getElementById('rc-topic')?.focus()},currentPayload?'Повторить':'Изменить запрос');return;}
  paint(row);
 }catch(e){if(ticket===generation)failure(e.message,()=>poll(id,ticket));}
}
export async function open(payload,trigger){
 shell();clearTimeout(timer);const ticket=++generation;returnFocus=trigger;currentPayload=payload;dialog.showModal();busy();
 try{const row=await request('/studio/outline',payload);remember(row.id);if(ticket!==generation)return;
  if(['ready','building','completed','build_failed'].includes(row.status))paint(row);else await poll(row.id,ticket);
 }catch(e){if(ticket===generation)failure(e.message,()=>{close();open(payload,trigger)});}
}
export function resume(trigger){let id;try{id=sessionStorage.getItem(memoryKey)}catch{}if(!id)return;
 shell();returnFocus=trigger;dialog.showModal();busy();poll(id,++generation);
}
window.addEventListener('pagehide',()=>{clearTimeout(timer);generation++});
