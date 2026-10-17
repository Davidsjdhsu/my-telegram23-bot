const assert = require('node:assert/strict');
const create = require('./miniapp-session');
(async () => {
  const calls=[];
  let fail=false, release;
  let hold=false;
  const session=create({getTg:()=>({initData:'signed-fixture'}),apiBase:'https://test.invalid',fetchFn:async(url, options)=>{
    calls.push({url, body:JSON.parse(options.body)});
    if(hold) await new Promise(r=>release=r);
    return {ok:!fail,json:async()=>fail?{ok:false,error:'busy'}:{ok:true,profile:{name:'Тигран',credits:70}}};
  }});
  const p=await session.loadProfile();assert.equal(p.profile.name,'Тигран');
  assert.equal(calls[0].body.initData,'signed-fixture');
  fail=true;await assert.rejects(session.send('help'),/busy/);
  const retryId=calls.at(-1).body.requestId;
  fail=false;assert.equal(await session.send('help'),true);
  assert.equal(calls.at(-1).body.requestId,retryId);
  hold=true;const first=session.send('gen_presentation',{style:'designer_strict',slides:8});
  assert.equal(await session.send('gen_presentation',{}),false);
  release();assert.equal(await first,true);
  const blocked=create({getTg:()=>({}),apiBase:'',fetchFn:()=>{throw Error('must not fetch')}});
  await assert.rejects(blocked.send('help'),/telegram_auth_required/);
  console.log('PASS: signed profile, HTTP actions, failed request retry, duplicate-click guard, missing auth.');
})().catch(e=>{console.error(e);process.exitCode=1});
