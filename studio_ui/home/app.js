(()=>{
'use strict';
const root=document.getElementById('alix-studio-mixed'),main=root.querySelector('main');
const escape=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const icon=n=>'<svg class="as-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true">'+({house:'<path d="m3 11 9-8 9 8v10h-6v-7H9v7H3Z"/>',wallet:'<rect x="3" y="5" width="18" height="15" rx="3"/><path d="M15 10h6v5h-6Z"/>','arrow-up-right':'<path d="M6 18 18 6M6 6h12v12"/>','file-text':'<path d="M5 3h10l4 4v14H5ZM8 11h8M8 15h8M8 18h5"/>','table-2':'<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/>',sparkles:'<path d="m12 2 3 7 7 3-7 3-3 7-3-7-7-3 7-3Z"/>'}[n]||'<path d="M5 12h14m-5-5 5 5-5 5"/>')+'</svg>';
const samples={pres:{format:'Презентация',ext:'Дизайнерская · 4 слайда',thumb:'/studio/templates/assets/soft-d16-0-small-3c4fea139c.webp'},word:{format:'Документ Word',ext:'Эссе · 3 страницы',thumb:'/studio/home/examples/essay-thumb.webp'},excel:{format:'Таблица Excel',ext:'Бюджет за месяц',thumb:'/studio/home/examples/excel-thumb.webp'}};
let sampleViewer,legalViewer;
async function openLegal(kind,button){
 if(button.disabled)return;button.disabled=true;
 try{legalViewer??=await import('/studio/home/legal.js');await legalViewer.open(kind,button);}
 catch(e){error='Не удалось открыть документ. Попробуйте ещё раз.';render();}
 finally{button.disabled=false;}
}
function legalLinks(){return '<section class="as-legal-links" aria-label="Документы сервиса"><h2>Документы сервиса</h2><button data-legal="offer">Публичная оферта<span aria-hidden="true">↗</span></button><button data-legal="terms">Пользовательское соглашение<span aria-hidden="true">↗</span></button></section>'}
async function openSample(kind,button){
 if(button.disabled)return;button.disabled=true;
 try{sampleViewer??=await import('/studio/home/examples.js');await sampleViewer.open(kind,button);}
 catch(e){error='Не удалось открыть пример. Проверьте соединение и нажмите ещё раз.';render();}
 finally{button.disabled=false;}
}
let page=['profile','files'].includes(new URLSearchParams(location.search).get('screen'))?new URLSearchParams(location.search).get('screen'):'home',data=null,error='',busy=false,request=null,loading=false;
const connection=root.querySelector('.as-connection');
const transition=()=>window.AlixMobile?.enter(main);
function connectionState(message,retry=false){connection.hidden=!message;connection.querySelector('span').textContent=message;connection.querySelector('button').hidden=!retry;}
const credit=()=>data?(data.unlimited?'∞':escape(data.credits)):'…';
function cover(k){return '<div class="as-cover as-real-cover '+k+'">'+(data?'<img src="'+samples[k].thumb+'" alt="" loading="lazy" decoding="async">':'')+'</div>'}
function formats(){return '<section aria-label="Создание файлов"><button class="as-hero cursor-interaction" data-format="pres" aria-label="Настроить презентацию"><span class="as-kicker">'+icon('sparkles')+' От идеи к слайдам</span><span class="as-card-title">Ваша презентация.</span><span class="as-card-desc">С вашим настроением и оформлением.</span><span class="as-hero-bottom"><small>Обычная — 10 кредитов/слайд<br>Дизайнерская — 20</small><span class="as-card-arrow">'+icon('arrow-up-right')+'</span></span></button><div class="as-pair"><button class="as-format cursor-interaction" data-format="word">'+icon('file-text')+'<strong>Документ Word</strong><small>Учёба, работа, личное</small></button><button class="as-format excel cursor-interaction" data-format="excel">'+icon('table-2')+'<strong>Таблица Excel</strong><small>Цифры под контролем</small></button></div></section>'}
function back(){return '<button class="as-back cursor-interaction" data-nav="profile">'+icon('arrow-left')+' Назад</button>'}function gallery(){return '<section><div class="as-section-head"><h2>Посмотрите, что получится</h2></div><div class="as-gallery">'+Object.entries(samples).map(([k,v])=>'<button class="as-example" data-sample="'+k+'">'+cover(k)+'<strong>'+v.format+'</strong><small>'+v.ext+'</small></button>').join('')+'</div></section>'}
function transactions(){
 if(!data)return '<p>Загружаем…</p>';
 const h=data.transactions||{items:[]};
 const html=h.items.map(t=>{const plus=t.delta>0;const amount=new Intl.NumberFormat('ru-RU',{maximumFractionDigits:1}).format(Math.abs(t.delta));const date=new Date(t.ts);const ts=Number.isNaN(date.getTime())?t.ts:date.toLocaleString('ru-RU');return '<div class="as-transaction"><span class="as-tx-icon">'+icon(plus?'wallet':'file-text')+'</span><div><span class="as-tx-title" data-no-translate>'+escape(t.title)+'</span><small>'+escape(ts)+'</small></div><div class="as-tx-amount '+(plus?'positive':'')+'"><strong>'+(plus?'+':t.delta<0?'−':'')+escape(amount)+'</strong><small>кредитов</small><small>Остаток: '+(t.balance_after==null?'не сохранён':escape(t.balance_after))+'</small></div></div>'}).join('');
 return (html||'<p class="as-notice">Операций пока нет.</p>')+(h.next_offset!=null?'<button data-action="more-transactions" class="as-quiet">Показать ещё</button>':'');
}
async function moreTransactions(button){
 button.disabled=true;
 try{const r=await fetch('/studio/transactions?offset='+data.transactions.next_offset,{cache:'no-store'});if(!r.ok)throw Error('Не удалось загрузить операции.');const b=await r.json();data.transactions={...b,items:[...data.transactions.items,...b.items]};render()}
 catch(e){error=e.message;render()}
}
function profile(){return '<div class="as-profile-head"><div class="as-avatar">'+escape((data?.name||'…')[0])+'</div><div><h1>'+escape(data?.name||'Профиль')+'</h1><small>Ваше пространство в AlixDoc</small></div></div><section class="as-wallet" aria-label="Карта баланса AlixDoc"><div class="as-wallet-top"><div class="as-wallet-brand" aria-label="AlixDoc">Alix<span>Doc</span></div><span class="as-wallet-caption">Личный баланс</span></div><div class="as-wallet-label">'+icon('wallet')+' Доступный баланс</div><div class="as-wallet-value"><strong>'+credit()+'</strong><span>кредитов</span></div><div class="as-wallet-owner"><span>Владелец</span><strong>'+escape(data?.name||'…')+'</strong></div><button class="as-primary" data-form="topup">Пополнить баланс</button></section><div class="as-section-head"><h2>История операций</h2></div><p class="as-notice">Пополнения и списания с остатком после операции. Старые списания без сохранённой записи не восстанавливаются приблизительно.</p>'+transactions()+'<div class="as-profile-bottom"><button data-nav="rates">Стоимость услуг</button><button data-nav="support">Поддержка</button></div><a class="as-link" href="/studio/languages/">Языки</a><button class="as-quiet" data-action="refresh">Обновить баланс</button>'+legalLinks()}
function render(){
root.querySelectorAll('.as-nav button').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.nav===page)));
if(page==='home')main.innerHTML='<div class="as-welcome"><div><small>Рады видеть вас,</small><p>'+escape(data?.name||'…')+'</p></div><button class="as-credit" data-nav="profile">'+credit()+' кредитов</button></div><h1 class="as-headline">Ваши идеи.<br><span style="color:var(--as-gold)">Красивый результат.</span></h1>'+formats()+gallery();
if(page==='profile')main.innerHTML=profile();
if(page==='files')main.innerHTML='<h1 class="as-headline">Мои файлы</h1><label>Поиск по названию<input id="file-search" type="search" placeholder="Название работы"></label><div id="file-results"></div><p class="as-notice">Повторное получение сохранённого файла бесплатно. Старые работы без сохранённой ссылки можно найти в чате.</p><button class="as-primary" data-send="history">Открыть историю в Telegram</button>';
if(page==='rates')main.innerHTML=back()+'<h1 class="as-headline">Стоимость услуг</h1><div class="as-rates"><div class="as-rate"><strong>Презентации</strong><p>Обычная — 10 кредитов за слайд<br>Дизайнерская — 20 кредитов за слайд<br>Обложка учитывается. PPTX и PDF включены.</p></div><div class="as-rate"><strong>Конструктор</strong><p>10–20 кредитов за слайд. Основа с одним фото — 10; второе и третье фото — по +3; каждый узор — +2 (до трёх). Шрифт, цвет и расположение включены. Не больше 20 кредитов за слайд. До 20 слайдов. Списание после отправки PPTX.</p></div><div class="as-rate"><strong>Word и Excel</strong><p>Word: простые документы — 20, подробные договоры — 40 за файл. Учебные работы и обычный текст — 5 за страницу текста, минимум 20; титульный и оглавление бесплатно. Excel — 20 за таблицу; бюджет проекта и финансовая модель — 40. Оформление и пароль формул включены.</p></div><div class="as-rate"><strong>Чат, фото и голос</strong><p>Ответ в чате — 0,5 кредита. Разбор фото — 3. Голосовой запрос с ответом — 3, без дополнительной платы за текстовый ответ. Навигация и уточнение заказа бесплатны. Ошибки без результата не оплачиваются.</p></div></div>';
if(page==='support')main.innerHTML=back()+'<h1 class="as-headline">Помощь</h1><p class="as-subtitle">Откройте справку бота в Telegram.</p><button class="as-primary" data-send="help">Открыть справку</button>'+legalLinks();
if(page==='files')drawFiles();
main.insertAdjacentHTML('beforeend','<p class="as-notice" role="status">'+escape(error)+'</p>');
}
function drawFiles(query=''){
 const box=document.getElementById('file-results');if(!box)return;
 const q=query.trim().toLocaleLowerCase(),files=data?.files||[];
 const rows=files.filter(f=>f.title.toLocaleLowerCase().includes(q));
 const older=(data?.history||[]).slice().reverse().filter(h=>h.toLocaleLowerCase().includes(q)&&!files.some(f=>h.includes(f.title)));
 box.innerHTML=rows.map(f=>'<div class="as-row"><div data-no-translate><strong>'+escape(f.title)+'</strong><small>'+escape(f.formats.join(' · '))+'</small></div><button data-file="'+escape(f.id)+'">Получить в чате</button></div>').join('')+older.map(h=>'<div class="as-row"><div data-no-translate>'+escape(h)+'</div></div>').join('')||'<p>Ничего не найдено.</p>';
}
root.addEventListener('input',e=>{if(e.target.id==='file-search')drawFiles(e.target.value)});
async function openFile(id,button){
 button.disabled=true;
 try{const r=await fetch('/studio/files/'+encodeURIComponent(id)+'/open',{method:'POST',headers:{'X-Studio-Request':'1'}});if(!r.ok)throw Error('Не удалось отправить файл. Повторите позже.');window.AlixMobile.openFiles();}
 catch(e){error=e.message;render()}finally{button.disabled=false}
}
async function load(){
 if(loading)return;loading=true;connectionState('');
 main.setAttribute('aria-busy','true');
 if(!data){main.inert=true;root.querySelector('.as-nav').inert=true;}
 const controller=new AbortController(),deadline=setTimeout(()=>controller.abort(),15000);
 const notice=setTimeout(()=>connectionState('Загружаем меню…'),1800);
 const options={cache:'no-store',signal:controller.signal};
 try{
  let r=await fetch('/studio/profile',options);
  if(r.status===401){
   const initData=window.Telegram?.WebApp?.initData;
   if(!initData)throw Error('Откройте «Меню» в чате AlixDocBot в Telegram.');
   const auth=await fetch('/studio/session',{...options,method:'POST',headers:{'Content-Type':'application/json','X-Studio-Request':'1'},body:JSON.stringify({initData})});
   if(!auth.ok)throw Error(auth.status===403?'Это меню пока доступно только владельцу. Откройте его заново в Telegram.':'Не удалось войти. Попробуйте ещё раз.');
   r=await fetch('/studio/profile',options);
  }
  if(!r.ok)throw Error('Не удалось загрузить меню. Попробуйте ещё раз.');
  const b=await r.json();if(!b.ok||!b.profile)throw Error('Не удалось загрузить профиль.');
  data=b.profile;window.AlixPreferences?.apply(data);error='';connectionState('');render();main.inert=false;root.querySelector('.as-nav').inert=false;
 }catch(e){connectionState(e.name==='AbortError'?'Загрузка заняла слишком много времени. Проверьте соединение и повторите.':e.message,true);}
 finally{clearTimeout(deadline);clearTimeout(notice);loading=false;main.setAttribute('aria-busy','false');}
}
async function send(action){if(busy)return;busy=true;error='Отправляем…';render();if(!request||request.action!==action)request={action,requestId:crypto.randomUUID(),payload:{}};try{const r=await fetch('/studio/action',{method:'POST',headers:{'Content-Type':'application/json','X-Studio-Request':'1'},body:JSON.stringify(request)});const b=await r.json();if(!r.ok||!b.ok)throw Error('Не удалось отправить запрос. Повторите попытку.');request=null;error='Ответ отправлен в чат бота.';window.Telegram?.WebApp?.close();}catch(e){error=e.message}finally{busy=false;render()}}
root.addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;if(b.dataset.action==='more-transactions'){moreTransactions(b)}else if(b.dataset.file){openFile(b.dataset.file,b)}else if(b.dataset.legal){openLegal(b.dataset.legal,b)}else if(b.dataset.nav){page=b.dataset.nav;error='';render();main.scrollTop=0;transition()}else if(b.dataset.format){location.href=b.dataset.format==='pres'?'/studio/templates/?kind=standard':b.dataset.format==='word'?'/studio/word/':'/studio/excel/'}else if(b.dataset.form){location.href='/studio/forms/?screen='+b.dataset.form}else if(b.dataset.sample){openSample(b.dataset.sample,b)}else if(b.dataset.send)send(b.dataset.send);else if(['refresh','retry'].includes(b.dataset.action))load()});
render();root.querySelector('.as-nav').inert=true;load();
})();
