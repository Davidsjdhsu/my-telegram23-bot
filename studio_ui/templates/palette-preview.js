// The server uses production layouts; photos remain separate, untinted assets.
const cache=new Map();
let controller;
export function cancel(){controller?.abort();controller=null;}
export async function load(style,key,palette,page,background_mode='solid',diagram_layout='classic'){
 const query=new URLSearchParams({style,key,palette,page,background_mode,diagram_layout});const id=query.toString();
 cancel();
 if(cache.has(id)){const value=cache.get(id);cache.delete(id);cache.set(id,value);return value;}
 const active=new AbortController();controller=active;
 const timeout=setTimeout(()=>active.abort(new DOMException('Preview timed out','TimeoutError')),15000);
 try{
 const response=await fetch('palette-preview?'+query,{signal:active.signal});
 if(!response.ok)throw new Error('Preview unavailable');
 const {svg}=await response.json();
 const doc=new DOMParser().parseFromString(svg,'image/svg+xml');
 const element=doc.documentElement;
 if(element.localName!=='svg'||doc.querySelector('parsererror'))throw new Error('Invalid preview');
 // These scenes are generated from fixed examples, never from user HTML.
 const images=[...element.querySelectorAll('image')].map(im=>new Promise((resolve,reject)=>{
  const image=new Image();const done=fn=>value=>{active.signal.removeEventListener('abort',abort);image.onload=image.onerror=null;fn(value);};
  const abort=()=>{image.src='';done(reject)(active.signal.reason);};
  if(active.signal.aborted){reject(active.signal.reason);return;}
  active.signal.addEventListener('abort',abort,{once:true});
  image.onload=done(resolve);image.onerror=done(reject);image.src=im.getAttribute('href');
 }));
 await Promise.all(images);
 cache.set(id,element);while(cache.size>8)cache.delete(cache.keys().next().value);
 return element;
 }finally{clearTimeout(timeout);if(controller===active)controller=null;}
}
export function paint(element,hosts){
 hosts.forEach((host,index)=>{
  const svg=element.cloneNode(true);
  // Zoom and small preview have independent gradient/clip IDs.
  svg.querySelectorAll('[id]').forEach(n=>{const old=n.id;const next=old+'-view'+index;n.id=next;
   svg.querySelectorAll('*').forEach(child=>{for(const attr of [...child.attributes])if(attr.value.includes('url(#'+old+')'))child.setAttribute(attr.name,attr.value.replaceAll('url(#'+old+')','url(#'+next+')'));});
  });
  host.replaceChildren(svg);host.hidden=false;
 });
}
