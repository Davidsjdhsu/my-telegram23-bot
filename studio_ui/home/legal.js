// Read the existing published documents without copying or rewriting their terms.
const documents={offer:{path:'/offer.html',title:'Публичная оферта'},terms:{path:'/terms.html',title:'Пользовательское соглашение'}};
let active;
export async function open(kind,opener){
 const spec=documents[kind];if(!spec||active)return;
 const dialog=document.createElement('dialog');dialog.className='as-viewer as-legal-dialog';dialog.setAttribute('aria-labelledby','as-legal-title');
 dialog.innerHTML='<header class="as-viewer-head"><div><small>AlixDoc</small><h2 id="as-legal-title">Документы сервиса</h2></div><button type="button" data-legal-close aria-label="Закрыть документ">✕</button></header><article class="as-legal-content" tabindex="0" aria-busy="true"></article>';
 document.getElementById('alix-studio-mixed').append(dialog);active=dialog;
 const article=dialog.querySelector('article'),tg=window.Telegram?.WebApp;
 const back=tg?.isVersionAtLeast?.('6.1')?tg.BackButton:null,wasBackVisible=!!back?.isVisible;
 let disposed=false,control,timer;
 if(history.state?.alixLegal)history.replaceState(null,'',location.pathname+location.search);
 const previousURL=location.href;history.pushState({alixLegal:kind},'',location.pathname+location.search+'#document='+kind);
 function cleanup(){
  if(disposed)return;disposed=true;control?.abort();clearTimeout(timer);back?.offClick(close);if(!wasBackVisible)back?.hide();
  window.removeEventListener('popstate',onBack);dialog.remove();active=null;opener?.focus({preventScroll:true});
 }
 function onBack(){dialog.close();cleanup();}
 function close(){dialog.close();cleanup();if(history.state?.alixLegal===kind)history.back();else history.replaceState(null,'',previousURL);}
 function clean(source){
  // Only document text/structure and email links. Never execute fetched markup.
  if(source.nodeType===Node.TEXT_NODE)return document.createTextNode(source.textContent);
  if(source.nodeType!==Node.ELEMENT_NODE||!['DIV','H1','H2','H3','P','UL','OL','LI','STRONG','B','EM','I','A','BR','SPAN'].includes(source.tagName))return document.createTextNode('');
  const el=document.createElement(source.tagName.toLowerCase());
  if(source.classList.contains('meta'))el.className='meta';
  if(source.tagName==='A'&&/^mailto:[^\s<>]+$/i.test(source.getAttribute('href')||''))el.setAttribute('href',source.getAttribute('href'));
  for(const child of source.childNodes)el.append(clean(child));return el;
 }
 async function load(){
  control?.abort();clearTimeout(timer);const current=new AbortController();control=current;
  article.setAttribute('aria-busy','true');article.innerHTML='<p role="status">Загружаем документ…</p>';
  timer=setTimeout(()=>current.abort(),12000);
  try{
   const response=await fetch(spec.path,{cache:'no-cache',signal:current.signal});if(!response.ok)throw Error('load');
   const parsed=new DOMParser().parseFromString(await response.text(),'text/html'),wrap=parsed.querySelector('.wrap');
   if(!wrap||wrap.querySelector('h1')?.textContent.trim()!==spec.title)throw Error('document');
   if(disposed||current!==control)return;article.replaceChildren(clean(wrap));article.scrollTop=0;
  }catch(e){if(!disposed&&current===control)article.innerHTML='<p role="alert">Не удалось загрузить документ. Проверьте соединение и повторите.</p><button type="button" data-legal-retry>Повторить</button>';}
  finally{if(current===control){clearTimeout(timer);article.setAttribute('aria-busy','false');}}
 }
 dialog.addEventListener('cancel',e=>{e.preventDefault();close()});
 dialog.addEventListener('click',e=>{if(e.target.closest('[data-legal-close]'))close();else if(e.target.closest('[data-legal-retry]'))load()});
 window.addEventListener('popstate',onBack);back?.onClick(close);back?.show();
 dialog.showModal();dialog.querySelector('button').focus({preventScroll:true});void load();
}
