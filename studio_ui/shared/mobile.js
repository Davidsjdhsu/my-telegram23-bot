(()=>{
'use strict';
const doc=document.documentElement,tg=window.Telegram?.WebApp;
const reduced=matchMedia('(prefers-reduced-motion: reduce)');
function enter(node){if(!node||reduced.matches||!node.animate)return;node.getAnimations().forEach(a=>a.cancel());node.animate([{opacity:.45,transform:'translateY(6px)'},{opacity:1,transform:'translateY(0)'}],{duration:170,easing:'ease-out'});}
let frame=0;
function size(){cancelAnimationFrame(frame);frame=requestAnimationFrame(()=>{
 const visual=window.visualViewport;
 // Telegram's viewportHeight can lag a keyboard resize. Use the visible minimum.
 const heights=[innerHeight,visual?.height,tg?.viewportHeight].filter(n=>Number.isFinite(n)&&n>0);
 doc.style.setProperty('--app-height',Math.floor(Math.min(...heights))+'px');
 for(const edge of ['top','bottom','left','right']){
  const inset=Math.max(tg?.safeAreaInset?.[edge]||0,tg?.contentSafeAreaInset?.[edge]||0);
  doc.style.setProperty('--app-safe-'+edge,inset+'px');
 }
});}
function openFiles(){
 const app=window.Telegram?.WebApp;
 if(app?.close&&(app.initData||(app.platform&&app.platform!=='unknown'))){app.close();return;}
 location.href='https://t.me/AlixDocBot';
}
const paths={home:'<path d="m3 11 9-8 9 8v10h-6v-7H9v7H3Z"/>',files:'<path d="M3 6h7l2 3h9v12H3Z"/>',profile:'<circle cx="12" cy="8" r="4"/><path d="M4 22v-3a8 8 0 0 1 16 0v3"/>'};
function icon(name){return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true">'+paths[name]+'</svg>'}
function navigation(){
 const existing=document.querySelector('.as-nav');
 if(existing){for(const button of existing.querySelectorAll('[data-nav]')){button.querySelector('i')?.remove();button.insertAdjacentHTML('afterbegin',icon(button.dataset.nav));if(button.dataset.nav==='home')button.lastChild.textContent='Главное меню';}return;}
 if(!document.querySelector('.studio-editor,.rc-app,.studio-return,.language-page'))return;
 const nav=document.createElement('nav');nav.className='studio-bottom-nav';nav.setAttribute('aria-label','Главная навигация');
 nav.innerHTML=Object.entries({home:'Главное меню',files:'Мои файлы',profile:'Профиль'}).map(([key,label])=>'<a href="/studio/home/?screen='+key+'">'+icon(key)+'<span>'+label+'</span></a>').join('');document.body.append(nav);document.body.classList.add('with-studio-nav');
 const back=document.querySelector('.studio-return');if(back&&document.body.classList.contains('studio-editor')){try{const previous=new URL(document.referrer);if(previous.origin===location.origin&&previous.pathname==='/studio/templates/')back.href=previous.href;}catch{}}
 if(back){const fn=()=>{location.href=back.href};tg?.BackButton?.onClick(fn);tg?.BackButton?.show();addEventListener('pagehide',()=>{tg?.BackButton?.offClick(fn);tg?.BackButton?.hide()});}
}
window.AlixMobile={enter,openFiles};
navigation();
addEventListener('resize',size,{passive:true});window.visualViewport?.addEventListener('resize',size,{passive:true});
for(const event of ['viewportChanged','safeAreaChanged','contentSafeAreaChanged'])tg?.onEvent?.(event,size);
tg?.ready();tg?.expand();size();
enter(document.querySelector('.as-content,.rc-app,.workspace,.bg-root'));
// Cross-document transitions are progressive enhancement; older WebViews use
// the same short entrance animation without intercepting links or adding waits.
addEventListener('pageshow',()=>size());
})();
