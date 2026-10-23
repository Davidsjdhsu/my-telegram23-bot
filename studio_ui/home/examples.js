// Preview only: no generation, file delivery, downloads or billing requests.
let catalogPromise, active;
const safe=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
async function catalog(){
 if(!catalogPromise)catalogPromise=(async()=>{
  const control=new AbortController(),timer=setTimeout(()=>control.abort(),12000);
  try{const r=await fetch('/studio/home/examples/catalog.json',{signal:control.signal});if(!r.ok)throw Error('catalog');return await r.json();}
  finally{clearTimeout(timer)}
 })().catch(e=>{catalogPromise=null;throw e});
 return catalogPromise;
}
export async function open(kind,opener){
 if(active)return;
 // Reserve the viewer while its small manifest loads, preventing double opens.
 active={pending:true};
 let item;
 try{item=(await catalog())[kind];if(!item?.pages?.length)throw Error('sample');}
 catch(e){active=null;throw e;}
 const dialog=document.createElement('dialog');dialog.className='as-viewer';dialog.setAttribute('aria-labelledby','as-example-title');
 dialog.innerHTML='<header class="as-viewer-head"><div><small>Пример · только просмотр</small><h2 id="as-example-title">'+safe(item.title)+'</h2></div><button type="button" data-view="close" aria-label="Закрыть пример">✕</button></header><div class="as-viewer-tools"><button type="button" data-view="zoom" aria-label="Увеличить страницу">＋ Увеличить</button><button type="button" data-view="fit">Вписать</button><span class="as-viewer-scale" aria-live="polite"></span></div><div class="as-viewer-stage" tabindex="0" aria-label="Страница примера. Используйте увеличение для чтения."><img class="as-viewer-image" draggable="false" hidden><div class="as-viewer-status" role="status"><span></span><button type="button" data-view="retry" hidden>Повторить</button></div></div><footer class="as-viewer-footer"><div class="as-viewer-pages"><button type="button" data-view="prev" aria-label="Предыдущая страница">←</button><span aria-live="polite"></span><button type="button" data-view="next" aria-label="Следующая страница">→</button></div><p>'+safe(item.note)+'</p><small class="as-viewer-help">'+(item.pages.length>1?'Листайте свайпом или стрелками. ':'')+'Увеличенную страницу можно двигать.</small></footer>';
 document.getElementById('alix-studio-mixed').append(dialog);
 const stage=dialog.querySelector('.as-viewer-stage'),img=dialog.querySelector('img'),status=dialog.querySelector('.as-viewer-status'),counter=dialog.querySelector('.as-viewer-pages span');
 let index=0,zoom=1,touch=null,timer,disposed=false,attempt=0;
 const tg=window.Telegram?.WebApp;
 const tgBack=tg?.isVersionAtLeast?.('6.1')?tg.BackButton:null,wasBackVisible=!!tgBack?.isVisible;
 if(history.state?.alixExample)history.replaceState(null,'',location.pathname+location.search);
 const beforeURL=location.href;
 history.pushState({alixExample:kind},'',location.pathname+location.search+'#example='+kind);
 function fit(){
  const p=item.pages[index],base=Math.max(1,Math.min(stage.clientWidth-24,(stage.clientHeight-24)*p.width/p.height));
  img.style.width=Math.round(base*zoom)+'px';img.style.height='auto';
  stage.classList.toggle('is-zoomed',zoom>1);stage.style.touchAction=zoom>1?'auto':'pan-y pinch-zoom';
  dialog.querySelector('.as-viewer-scale').textContent=zoom===1?'Целиком':Math.round(zoom*100)+'%';
  dialog.querySelector('[data-view="zoom"]').disabled=zoom>=3;
 }
 function message(text,retry=false){status.hidden=!text;status.querySelector('span').textContent=text;status.querySelector('button').hidden=!retry;}
 function loadPage(retry=false){
  clearTimeout(timer);const version=++attempt;img.hidden=true;zoom=1;stage.scrollTo(0,0);
  counter.textContent=(kind==='pres'?'Слайд ':kind==='excel'?'Лист ':'Страница ')+(index+1)+' из '+item.pages.length;
  dialog.querySelector('[data-view="prev"]').disabled=index===0;
  dialog.querySelector('[data-view="next"]').disabled=index===item.pages.length-1;
  message('Загружаем страницу…');fit();
  img.alt=item.title+' — '+counter.textContent.toLowerCase();
  img.onload=()=>{if(disposed||version!==attempt)return;clearTimeout(timer);message('');img.hidden=false;fit();window.AlixMobile?.enter(img);};
  img.onerror=()=>{if(disposed||version!==attempt)return;clearTimeout(timer);message('Не удалось загрузить страницу. Проверьте соединение.',true);};
  timer=setTimeout(()=>{if(!disposed&&version===attempt)message('Страница загружается дольше обычного. Можно повторить.',true)},12000);
  img.src=item.pages[index].src+(retry?'?retry='+Date.now():'');
 }
 function change(delta){const next=index+delta;if(next<0||next>=item.pages.length)return;index=next;loadPage();}
 function cleanup(){
  if(disposed)return;disposed=true;clearTimeout(timer);img.onload=img.onerror=null;observer.disconnect();
  window.removeEventListener('popstate',onBack);tgBack?.offClick(close);
  if(!wasBackVisible)tgBack?.hide();active=null;dialog.remove();opener?.focus({preventScroll:true});
 }
 function onBack(){dialog.close();cleanup();}
 function close(){if(disposed)return;dialog.close();cleanup();if(history.state?.alixExample===kind)history.back();else history.replaceState(null,'',beforeURL);}
 const observer=new ResizeObserver(fit);observer.observe(stage);
 window.addEventListener('popstate',onBack);tgBack?.onClick(close);tgBack?.show();
 dialog.addEventListener('cancel',e=>{e.preventDefault();close()});
 dialog.addEventListener('click',e=>{
  const action=e.target.closest('[data-view]')?.dataset.view;
  if(action==='close')close();if(action==='next')change(1);if(action==='prev')change(-1);if(action==='retry')loadPage(true);
  if(action==='zoom'){zoom=Math.min(3,zoom+1);fit();}if(action==='fit'){zoom=1;fit();stage.scrollTo(0,0);}
 });
 dialog.addEventListener('keydown',e=>{if(zoom!==1)return;if(e.key==='ArrowRight'){e.preventDefault();change(1)}if(e.key==='ArrowLeft'){e.preventDefault();change(-1)}});
 stage.addEventListener('touchstart',e=>{touch=zoom===1&&e.touches.length===1?{x:e.touches[0].clientX,y:e.touches[0].clientY}:null},{passive:true});
 stage.addEventListener('touchmove',e=>{if(e.touches.length>1)touch=null},{passive:true});
 stage.addEventListener('touchcancel',()=>{touch=null},{passive:true});
 stage.addEventListener('touchend',e=>{if(!touch||zoom!==1)return;const t=e.changedTouches[0],dx=t.clientX-touch.x,dy=t.clientY-touch.y;touch=null;if(Math.abs(dx)>55&&Math.abs(dx)>Math.abs(dy)*1.5)change(dx<0?1:-1)},{passive:true});
 active={dialog};dialog.showModal();loadPage();dialog.querySelector('[data-view="close"]').focus({preventScroll:true});
}
