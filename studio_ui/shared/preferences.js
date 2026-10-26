(()=>{'use strict';
const languages={"ru": "Русский", "en": "English", "de": "Deutsch", "ar": "العربية", "zh": "中文", "es": "Español", "fr": "Français"};
let prefs={ui_lang:'ru',content_lang:'ru'};try{prefs={...prefs,...JSON.parse(localStorage.getItem('alix-languages')||'{}')}}catch{}
let dictionary={},observer;const originals=new WeakMap(),attributes=new WeakMap();
function translate(){
 if(!dictionary||!document.body)return;
 document.documentElement.lang=prefs.ui_lang;
 observer?.disconnect();
 const walk=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT);
 while(walk.nextNode()){
  const node=walk.currentNode,el=node.parentElement;
  if(!el||el.closest('script,style,textarea,svg,canvas,[data-no-translate],#slide-list,#layers,#rc-topic,.as-row,.as-avatar,.as-wallet-owner,.as-profile-head h1,.as-welcome p,.op-title,.op-slide h4'))continue;
  const now=node.nodeValue,old=originals.get(node);
  const original=old&&(now===old.rendered||now===old.original)?old.original:now;
  const term=original.trim();let translated=prefs.ui_lang==='en'?(dictionary[term]||term):term;if(prefs.ui_lang==='en'&&!dictionary[term])translated=term.replace(/^(\d+|∞) кредитов$/, '$1 credits').replace(/^(\d+) кредитов за слайд$/, '$1 credits per slide').replace(/^(\d+) кредитов \/ слайд$/, '$1 credits / slide').replace(/^(До |до )?(\d+) кредитов$/,(_,up,n)=>(up?'Up to ':'')+n+' credits').replace(/^(Титульный \+ )?(\d+) стр\.( текста)?$/,(_,cover,n)=> (cover?'Title page + ':'')+n+' pages'+(cover?' of text':'')).replace(/^Слайд (\d+) из (\d+)$/, 'Slide $1 of $2');
  const rendered=original.replace(term,translated);originals.set(node,{original,rendered});if(now!==rendered)node.nodeValue=rendered;
 }
 for(const el of document.querySelectorAll('[placeholder],[aria-label]')){
  if(el.closest('[data-no-translate]'))continue;
  let saved=attributes.get(el)||{};
  for(const attr of ['placeholder','aria-label']){
   if(!el.hasAttribute(attr))continue;
   const now=el.getAttribute(attr),old=saved[attr];const original=old&&(now===old.rendered||now===old.original)?old.original:now;
   const rendered=prefs.ui_lang==='en'?(dictionary[original]||original):original;
   saved[attr]={original,rendered};if(now!==rendered)el.setAttribute(attr,rendered);
  }attributes.set(el,saved);
 }
 observer?.observe(document.body,{childList:true,subtree:true,characterData:true});
}
function fill(select){select.innerHTML=Object.entries(languages).map(([k,v])=>'<option value="'+k+'">'+v+'</option>').join('');select.value=prefs.content_lang;select.addEventListener('change',()=>select.dataset.touched='1')}
function apply(value){prefs={ui_lang:value.ui_lang==='en'?'en':'ru',content_lang:languages[value.content_lang]?value.content_lang:'ru'};try{localStorage.setItem('alix-languages',JSON.stringify(prefs))}catch{};for(const s of document.querySelectorAll('[data-content-language]'))if(!s.dataset.touched)s.value=prefs.content_lang;translate();document.dispatchEvent(new CustomEvent('alix-preferences',{detail:prefs}));}
window.AlixPreferences={apply,fill,get:()=>({...prefs}),content:()=>document.querySelector('[data-content-language]')?.value||prefs.content_lang};
for(const select of document.querySelectorAll('[data-content-language]'))fill(select);
let timer;observer=new MutationObserver(()=>{clearTimeout(timer);timer=setTimeout(translate,40)});
fetch('/studio/shared/en.json').then(r=>r.json()).then(d=>{dictionary=d;translate()}).catch(()=>{});
if(!document.getElementById('alix-studio-mixed'))fetch('/studio/profile',{cache:'no-store'}).then(r=>r.ok?r.json():Promise.reject()).then(b=>apply(b.profile)).catch(()=>{});
})();
