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
window.AlixMobile={enter,openFiles};
addEventListener('resize',size,{passive:true});window.visualViewport?.addEventListener('resize',size,{passive:true});
for(const event of ['viewportChanged','safeAreaChanged','contentSafeAreaChanged'])tg?.onEvent?.(event,size);
tg?.ready();tg?.expand();size();
enter(document.querySelector('.as-content,.rc-app,.workspace,.bg-root'));
// Cross-document transitions are progressive enhancement; older WebViews use
// the same short entrance animation without intercepting links or adding waits.
addEventListener('pageshow',()=>size());
})();
