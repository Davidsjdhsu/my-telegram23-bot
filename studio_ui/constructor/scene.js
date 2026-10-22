(function(scope){
const D=typeof module!=='undefined'?require('./diagrams.js'):scope.Diagrams;
const W=1200,H=675,FONTS=['Arial','Georgia','Times New Roman','Trebuchet MS','Verdana'];
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const clamp=(v,a,b)=>Math.max(a,Math.min(b,Number(v)||0));
const uid=()=> 'e'+Math.random().toString(36).slice(2,10);
const text=(value,x,y,w,h,size=32,color='#eee7dc',font='Arial')=>({id:uid(),type:'text',text:value,x,y,w,h,size,color,font,bold:false,italic:false,align:'left',opacity:1,rotation:0});
const photo=(x,y,w,h)=>({id:uid(),type:'photo',src:'coffee.jpg',iw:1344,ih:768,x,y,w,h,shape:'round',radius:30,fx:50,fy:50,zoom:1,opacity:1,rotation:0,border:'#c4aa7d',borderWidth:2});
const pattern=(x,y,w,h,kind='rings')=>({id:uid(),type:'pattern',x,y,w,h,kind,color:'#c4aa7d',opacity:.35,rotation:0,weight:3});
function initial(){return {version:1,name:'Моя кофейня',slides:[
 {id:uid(),background:{type:'solid',color:'#14293b',second:'#314253',angle:135},elements:[pattern(880,28,270,270),photo(650,133,475,405),text('Моя\nкофейня',76,146,540,190,76,'#eee7dc','Georgia'),text('Место, в которое хочется вернуться',80,382,480,120,28,'#c4aa7d')]},
 {id:uid(),background:{type:'solid',color:'#f4eee3',second:'#ded2be',angle:135},elements:[pattern(850,360,290,250,'lines'),{...photo(75,125,400,400),shape:'circle'},text('Атмосфера\nи детали',550,130,565,160,60,'#382f28','Georgia'),text('Хороший кофе, спокойная музыка и удобное пространство.\n\nЗдесь можно встретиться с друзьями или сделать паузу в течение дня.',550,327,510,230,27,'#51483e')]},
 {id:uid(),background:{type:'gradient',color:'#47293a',second:'#1e2736',angle:135},elements:[pattern(740,115,350,400,'dots'),text('Ваша идея',85,170,880,120,80,'#e8dac2','Georgia'),text('Добавьте свой текст, фотографию и детали оформления.',90,338,750,120,32,'#e7e2d9')]}
]};}
function wrap(e,measure){
 const lines=[];
 for(const para of String(e.text).split('\n')){
  if(!para){lines.push('');continue;}let line='';
  for(const word of para.split(/\s+/)){
   const proposed=line?line+' '+word:word;
   if(measure(proposed,e)<=e.w-4){line=proposed;continue;}
   if(line){lines.push(line);line='';}
   if(measure(word,e)<=e.w-4){line=word;continue;}
   for(const ch of word){if(line&&measure(line+ch,e)>e.w-4){lines.push(line);line=ch;}else line+=ch;}
  }
  if(line)lines.push(line);
 }
 return lines;
}
function patternMarkup(e){
 const w=e.w,h=e.h,c=esc(e.color),sw=e.weight;let out='';
 if(e.kind==='rings')for(let i=0;i<5;i++)out+=`<ellipse cx="${w/2}" cy="${h/2}" rx="${w*(.46-i*.075)}" ry="${h*(.46-i*.075)}" fill="none" stroke="${c}" stroke-width="${sw}"/>`;
 if(e.kind==='lines')for(let i=0;i<8;i++)out+=`<line x1="${w*.08}" y1="${h*(.1+i*.095)}" x2="${w*.92}" y2="${h*(.23+i*.095)}" stroke="${c}" stroke-width="${sw}"/>`;
 if(e.kind==='dots')for(let row=0;row<6;row++)for(let col=0;col<6;col++)out+=`<circle cx="${w*(.08+col*.168)}" cy="${h*(.08+row*.168)}" r="${Math.min(w,h)*.014+sw/2}" fill="${c}"/>`;
 if(e.kind==='frame')for(let i=0;i<3;i++)out+=`<rect x="${w*(.04+i*.055)}" y="${h*(.04+i*.055)}" width="${w*(.92-i*.11)}" height="${h*(.92-i*.11)}" rx="${Math.min(w,h)*.08}" fill="none" stroke="${c}" stroke-width="${sw}"/>`;
 return out;
}
function crop(e){const s=Math.max(e.w/e.iw,e.h/e.ih)*e.zoom;return {x:(e.w-e.iw*s)*e.fx/100,y:(e.h-e.ih*s)*e.fy/100,w:e.iw*s,h:e.ih*s};}
function element(e,prefix,measure){
 const id=prefix+'-'+e.id;let body='';
 if(e.type==='text'){
  const lines=wrap(e,measure),anchor=e.align==='center'?'middle':e.align==='right'?'end':'start',x=e.align==='center'?e.w/2:e.align==='right'?e.w:0;
  body=`<defs><clipPath id="${id}"><rect width="${e.w}" height="${e.h}"/></clipPath></defs><rect width="${e.w}" height="${e.h}" fill="transparent"/><text clip-path="url(#${id})" font-family="${esc(e.font)}" font-size="${e.size}" font-weight="${e.bold?700:400}" font-style="${e.italic?'italic':'normal'}" fill="${e.color}" text-anchor="${anchor}">${lines.map((l,i)=>`<tspan x="${x}" y="${e.size*.95+i*e.size*1.22}">${esc(l)}</tspan>`).join('')}</text>`;
 }else if(e.type==='photo'){
  const shape=e.shape==='circle'?`<ellipse cx="${e.w/2}" cy="${e.h/2}" rx="${e.w/2}" ry="${e.h/2}"/>`:`<rect width="${e.w}" height="${e.h}" rx="${e.shape==='round'?e.radius:0}"/>`,c=crop(e);
  body=`<defs><clipPath id="${id}">${shape}</clipPath></defs><image href="${esc(e.src)}" x="${c.x}" y="${c.y}" width="${c.w}" height="${c.h}" clip-path="url(#${id})"/><g fill="none" stroke="${e.border}" stroke-width="${e.borderWidth}">${shape}</g>`;
 }else if(e.type==='diagram')body=`<rect width="${e.w}" height="${e.h}" fill="transparent"/>`+D.markup(e,(t,size)=>measure(t,{size,font:'Arial'}));
 else body=`<rect width="${e.w}" height="${e.h}" fill="transparent"/>`+patternMarkup(e);
 return `<g data-id="${e.id}" transform="translate(${e.x} ${e.y}) rotate(${e.rotation} ${e.w/2} ${e.h/2})" opacity="${e.opacity}">${body}</g>`;
}
function inner(slide,{prefix='s',selected=null,measure=(t,e)=>t.length*e.size*.55}={}){
 const b=slide.background,a=b.angle*Math.PI/180,dx=Math.cos(a)*50,dy=Math.sin(a)*50;let out=`<defs><linearGradient id="${prefix}-bg" x1="${50-dx}%" y1="${50-dy}%" x2="${50+dx}%" y2="${50+dy}%"><stop stop-color="${b.color}"/><stop offset="1" stop-color="${b.second}"/></linearGradient></defs><rect width="1200" height="675" fill="${b.type==='gradient'?`url(#${prefix}-bg)`:b.color}"/>`;
 for(const e of slide.elements)out+=element(e,prefix,measure);
 const e=slide.elements.find(e=>e.id===selected);
 if(e)out+=`<g pointer-events="none" transform="translate(${e.x} ${e.y}) rotate(${e.rotation} ${e.w/2} ${e.h/2})"><rect x="-3" y="-3" width="${e.w+6}" height="${e.h+6}" fill="none" stroke="#dfba72" stroke-width="3" stroke-dasharray="8 5"/><rect data-resize="${e.id}" pointer-events="all" x="${e.w-10}" y="${e.h-10}" width="20" height="20" fill="#dfba72" stroke="#1a1b20" stroke-width="3"/></g>`;
 return out;
}
function svg(slide,options){return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1200 675">${inner(slide,options)}</svg>`;}
function validate(raw){
 if(!raw||raw.version!==1||!Array.isArray(raw.slides)||!raw.slides.length||raw.slides.length>20)throw Error('Неверный проект: нужно от 1 до 20 слайдов.');
 const clean=JSON.parse(JSON.stringify(raw));clean.name=String(raw.name||'Моя презентация').slice(0,80);
 for(const s of clean.slides){s.id=uid();const b=s.background;if(!b||!/^#[\da-f]{6}$/i.test(b.color)||!/^#[\da-f]{6}$/i.test(b.second))throw Error('Некорректный цвет фона.');b.type=b.type==='gradient'?'gradient':'solid';b.angle=clamp(b.angle,0,360);
  if(!Array.isArray(s.elements)||s.elements.length>50)throw Error('Слишком много элементов.');
  for(const e of s.elements){e.id=uid();if(!['text','photo','pattern','diagram'].includes(e.type))throw Error('Неизвестный элемент.');e.w=clamp(e.w,40,W);e.h=clamp(e.h,30,H);e.x=clamp(e.x,0,W-e.w);e.y=clamp(e.y,0,H-e.h);e.rotation=clamp(e.rotation,-180,180);e.opacity=clamp(e.opacity,0,1);
   if(e.type==='text'){e.text=String(e.text||'').slice(0,3000);e.font=FONTS.includes(e.font)?e.font:'Arial';e.size=clamp(e.size,12,120);e.bold=!!e.bold;e.italic=!!e.italic;e.align=['left','center','right'].includes(e.align)?e.align:'left';}
   if(e.type==='diagram'){if(!Array.isArray(e.labels)||e.labels.length<2||e.labels.length>7)throw Error('В схеме нужно от 2 до 7 блоков.');e.labels=e.labels.map(x=>String(x).slice(0,160));e.kind=Object.keys(D.NAMES).includes(e.kind)?e.kind:'cards';e.size=clamp(e.size,16,42);e.rotation=0;for(const key of ['fill','accent'])if(!/^#[\da-f]{6}$/i.test(e[key]))throw Error('Некорректный цвет схемы.');}
   if(e.type==='pattern'){e.kind=['rings','lines','dots','frame'].includes(e.kind)?e.kind:'rings';e.weight=clamp(e.weight,1,12);}
   if(e.type!=='photo'&&!/^#[\da-f]{6}$/i.test(e.color))throw Error('Некорректный цвет элемента.');
   if(e.type==='photo'){if(e.src!=='coffee.jpg'&&!/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+/=]+$/.test(e.src))throw Error('Разрешены только локально добавленные фотографии.');if(e.src.length>4000000)throw Error('Фотография слишком большая.');e.iw=clamp(e.iw,1,10000);e.ih=clamp(e.ih,1,10000);e.fx=clamp(e.fx,0,100);e.fy=clamp(e.fy,0,100);e.zoom=clamp(e.zoom,1,3);e.radius=clamp(e.radius,0,200);e.shape=['rect','round','circle'].includes(e.shape)?e.shape:'rect';if(!/^#[\da-f]{6}$/i.test(e.border))e.border='#c4aa7d';e.borderWidth=clamp(e.borderWidth,0,12);}
  }
 }
 return clean;
}
const api={W,H,FONTS,esc,clamp,uid,text,photo,pattern,initial,wrap,crop,patternMarkup,inner,svg,validate};
if(typeof module!=='undefined')module.exports=api;else scope.Scene=api;
})(typeof window==='undefined'?globalThis:window);
