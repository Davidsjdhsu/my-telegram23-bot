(function(scope){
const NAMES={cards:'Этапы',stairs:'Ступени',timeline:'Временная шкала',cycle:'Цикл',tree:'Дерево',compare:'Сравнение'};
const PALETTES={strict:{bg:'#14293b',fill:'#233e51',accent:'#c4aa7d',color:'#f1e9dc'},soft:{bg:'#f4eee3',fill:'#e2e9df',accent:'#467460',color:'#243d32'},burgundy:{bg:'#34232f',fill:'#513346',accent:'#d1b38a',color:'#f4ece2'}};
const labels=['Заказ','Оплата','Приготовление','Проверка','Выдача'];
function create(kind='cards',palette='strict'){const p=PALETTES[palette]||PALETTES.strict;return {type:'diagram',id:'e'+Math.random().toString(36).slice(2,10),kind,x:70,y:180,w:1060,h:400,labels:kind==='tree'?['Кофейня','Бар','Кухня','Зал','Доставка']:kind==='compare'?['В зале','С собой']:kind==='cycle'?['План','Работа','Проверка','Улучшение']:labels.slice(),fill:p.fill,accent:p.accent,color:p.color,size:24,opacity:1,rotation:0};}
function wrapLabel(text,width,size,measure){let out=[],line='';for(const word of text.split(/\s+/)){const next=line?line+' '+word:word;if(line&&measure(next,size)>width){out.push(line);line=word;}else line=next;}if(line)out.push(line);return out;}
function parts(e,measure=(t,s)=>t.length*s*.55){
 const out=[],sx=e.w/1060,sy=e.h/400,fs=e.size*Math.min(sx,sy),n=e.labels.length;
 const shape=(geometry,x,y,w,h,fill,stroke='none',rotation=0)=>out.push({geometry,x:x*sx,y:y*sy,w:w*sx,h:h*sy,fill,stroke,rotation});
 const txt=(value,x,y,w,h,size=fs,color=e.color,bold=false)=>out.push({geometry:'textbox',x:x*sx,y:y*sy,w:w*sx,h:h*sy,fill:'none',stroke:'none',rotation:0,text:value,lines:wrapLabel(value,w*sx-12,size,measure),size,color,bold});
 function box(i,x,y,w=180,h=110){shape('roundRect',x,y,w,h,e.fill,e.accent);txt(String(i+1).padStart(2,'0'),x+10,y+9,w-20,25,Math.min(fs,22*Math.min(sx,sy)),e.accent,true);txt(e.labels[i],x+4,y+43,w-8,h-48);}
 function arrow(x1,y1,x2,y2){const dx=(x2-x1)*sx,dy=(y2-y1)*sy,len=Math.hypot(dx,dy);out.push({geometry:'rightArrow',x:(x1+x2)*sx/2-len/2,y:(y1+y2)*sy/2-5,w:len,h:10,fill:e.accent,stroke:'none',rotation:Math.atan2(dy,dx)*180/Math.PI});}
 if(['cards','stairs','timeline'].includes(e.kind)){
  const gap=24,w=(1060-gap*(n-1))/n;
  for(let i=0;i<n;i++){const x=i*(w+gap),y=e.kind==='stairs'?260-i*220/Math.max(1,n-1):e.kind==='timeline'?(i%2?232:30):130;
   if(e.kind==='timeline'){shape('ellipse',x+w/2-9,191,18,18,e.accent);if(i<n-1)arrow(x+w/2+12,200,x+w+gap+w/2-12,200);shape('rect',x+w/2-1,i%2?209:140,2,i%2?23:51,e.accent);}
   else if(i<n-1){const nextY=e.kind==='stairs'?260-(i+1)*220/Math.max(1,n-1):y;arrow(x+w+4,y+55,x+w+gap-4,nextY+55);}
   box(i,x,y,w,110);
  }
 }else if(e.kind==='cycle'){
  const w=168,h=90,cx=530,cy=200,rx=400,ry=145;
  const nodes=e.labels.map((_,i)=>{const a=-Math.PI/2+i*Math.PI*2/n;return{x:cx+rx*Math.cos(a),y:cy+ry*Math.sin(a)};});
  for(let i=0;i<n;i++){const a=nodes[i],b=nodes[(i+1)%n],dx=b.x-a.x,dy=b.y-a.y;const t=Math.min(.47,Math.min(Math.abs(dx)>0?(w/2+8)/Math.abs(dx):Infinity,Math.abs(dy)>0?(h/2+8)/Math.abs(dy):Infinity));arrow(a.x+dx*t,a.y+dy*t,b.x-dx*t,b.y-dy*t);}
  nodes.forEach((p,i)=>box(i,p.x-w/2,p.y-h/2,w,h));
 }else if(e.kind==='tree'){
  const w=Math.min(210,(1060-20*(n-2))/(n-1)),gap=(1060-w*(n-1))/Math.max(1,n-2);
  for(let i=1;i<n;i++){const x=(i-1)*(w+gap);arrow(530,128,x+w/2,240);box(i,x,245,w,120);}box(0,400,10,260,115);
 }else{
  const cols=n<=3?n:Math.ceil(n/2),rows=Math.ceil(n/cols),gap=24,w=(1060-gap*(cols-1))/cols,h=rows===1?240:170;
  e.labels.forEach((_,i)=>box(i,(i%cols)*(w+gap),rows===1?80:Math.floor(i/cols)*205,w,h));
 }
 return out;
}
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function markup(e,measure){return parts(e,measure).map(p=>{let body='';if(p.geometry==='textbox'){const lh=p.size*1.22,yy=Math.max(p.size,(p.h-p.lines.length*lh)/2+p.size);body=`<text font-family="Arial" font-size="${p.size}" font-weight="${p.bold?700:400}" fill="${p.color}" text-anchor="middle">${p.lines.map((l,i)=>`<tspan x="${p.w/2}" y="${yy+i*lh}">${esc(l)}</tspan>`).join('')}</text>`;}else if(p.geometry==='ellipse')body=`<ellipse cx="${p.w/2}" cy="${p.h/2}" rx="${p.w/2}" ry="${p.h/2}" fill="${p.fill}"/>`;else if(p.geometry==='rightArrow')body=`<path d="M0 ${p.h*.3} H${Math.max(0,p.w-10)} V0 L${p.w} ${p.h/2} L${Math.max(0,p.w-10)} ${p.h} V${p.h*.7} H0Z" fill="${p.fill}"/>`;else body=`<rect width="${p.w}" height="${p.h}" rx="${p.geometry==='roundRect'?12:0}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="1.5"/>`;return `<g transform="translate(${p.x} ${p.y}) rotate(${p.rotation} ${p.w/2} ${p.h/2})">${body}</g>`;}).join('');}
function overflow(e,measure){return parts(e,measure).some(p=>p.text&&(p.lines.length*p.size*1.22>p.h+3||p.lines.some(l=>measure(l,p.size)>p.w-6)));}
const api={NAMES,PALETTES,create,parts,markup,overflow};if(typeof module!=='undefined')module.exports=api;else scope.Diagrams=api;
})(typeof window==='undefined'?globalThis:window);
