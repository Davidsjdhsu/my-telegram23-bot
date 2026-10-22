(function(scope){
'use strict';
class PreviewLoader {
 constructor({fetchImage=(...args)=>fetch(...args),createURL=blob=>URL.createObjectURL(blob),revokeURL=url=>URL.revokeObjectURL(url),decode=async url=>{const im=new Image();im.src=url;await im.decode();return im;},limit=4,concurrency=2}={}){
  Object.assign(this,{fetchImage,createURL,revokeURL,decode,limit,concurrency});this.cache=new Map();this.tasks=new Map();this.active='';this.running=0;this.destroyed=false;
 }
 focus(url){
  this.active=url;
  for(const [key,t] of this.tasks)if(key!==url){t.controller.abort();if(!t.running){this.tasks.delete(key);t.reject(this.abortError());}}
  return this.load(url,0);
 }
 warm(urls){
  if(this.destroyed)return;
  for(const url of [...new Set(urls)].filter(u=>u!==this.active).slice(0,2))this.load(url,1).catch(()=>{});
 }
 abortError(){const e=new Error('Preview superseded');e.name='AbortError';return e;}
 load(url,priority){
  if(this.destroyed)return Promise.reject(this.abortError());
  if(this.cache.has(url)){const value=this.cache.get(url);this.cache.delete(url);this.cache.set(url,value);return Promise.resolve(value.url);}
  let t=this.tasks.get(url);
  if(t&&!t.controller.signal.aborted){t.priority=Math.min(t.priority,priority);return t.promise;}
  t={url,priority,controller:new AbortController(),running:false};
  t.promise=new Promise((resolve,reject)=>Object.assign(t,{resolve,reject}));
  this.tasks.set(url,t);this.pump();return t.promise;
 }
 pump(){
  if(this.destroyed)return;
  while(this.running<this.concurrency){
   const t=[...this.tasks.values()].filter(t=>!t.running&&!t.controller.signal.aborted).sort((a,b)=>a.priority-b.priority)[0];if(!t)return;
   t.running=true;this.running++;this.run(t);
  }
 }
 async run(t){
  let objectURL;
  try{
   const response=await this.fetchImage(t.url,{signal:t.controller.signal,cache:'force-cache',priority:t.priority===0?'high':'low'});
   if(!response.ok)throw new Error('Preview HTTP '+response.status);
   const blob=await response.blob();if(t.controller.signal.aborted)throw this.abortError();
   objectURL=this.createURL(blob);const decoded=await this.decode(objectURL);
   if(t.controller.signal.aborted||this.destroyed)throw this.abortError();
   this.cache.set(t.url,{url:objectURL,decoded});objectURL=null;
   while(this.cache.size>this.limit){const key=[...this.cache.keys()].find(k=>k!==this.active);const old=this.cache.get(key);this.revokeURL(old.url);this.cache.delete(key);}
   t.resolve(this.cache.get(t.url).url);
  }catch(e){if(objectURL)this.revokeURL(objectURL);t.reject(e);}
  finally{if(this.tasks.get(t.url)===t)this.tasks.delete(t.url);this.running--;this.pump();}
 }
 destroy(){
  this.destroyed=true;
  for(const t of this.tasks.values()){t.controller.abort();if(!t.running)t.reject(this.abortError());}
  this.tasks.clear();for(const v of this.cache.values())this.revokeURL(v.url);this.cache.clear();
 }
}
if(typeof module!=='undefined'&&module.exports)module.exports=PreviewLoader;
else scope.AlixPreviewLoader=PreviewLoader;
})(typeof window!=='undefined'?window:globalThis);
