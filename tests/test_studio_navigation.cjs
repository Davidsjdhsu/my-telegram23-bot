const assert=require('node:assert/strict');
const fs=require('node:fs');const vm=require('node:vm');
const source=fs.readFileSync('studio_ui/shared/mobile.js','utf8');
for(const platform of ['ios','android','tdesktop','unknown']){
 let closed=0;const target={href:''};
 const context={window:{Telegram:{WebApp:{platform,initData:'',close(){closed++},onEvent(){},ready(){},expand(){}}}},location:target,document:{documentElement:{style:{setProperty(){}}},querySelector(){return null}},matchMedia:()=>({matches:true}),addEventListener(){},cancelAnimationFrame(){},requestAnimationFrame(){return 1}};
 vm.runInNewContext(source,context);context.window.AlixMobile.openFiles();
 assert.equal(closed,platform==='unknown'?0:1);assert.equal(target.href,platform==='unknown'?'https://t.me/AlixDocBot':'');
}
console.log('PASS: iOS, Android, Desktop close Mini App; browser fallback opens bot chat.');
