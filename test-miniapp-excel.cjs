const fs=require('node:fs'), vm=require('node:vm'), assert=require('node:assert/strict');
const html=fs.readFileSync(__dirname+'/index.html','utf8');
const elements={};
const el=id=>elements[id] ||= {value:'',hidden:false,events:[],focus(){},addEventListener(_,f){this.events.push(f)},click(){this.events.forEach(f=>f())}};
el('excelKind').value='startup_model'; el('excelStyle').value='standard'; el('excelPalette').value='green';
const sent=[];
vm.runInNewContext(html.slice(html.indexOf('  // --- Отправка формы Excel'),html.indexOf('  // --- Отправка договора купли-продажи')),
 {document:{getElementById:el},activeChipValue:()=> 'user',sendAction:(action,payload)=>sent.push({action,payload})});
assert.equal(el('excelAppearanceBlock').hidden,false);
el('excelKind').value='expense_estimate';el('excelKind').click();
assert.equal(el('excelAppearanceBlock').hidden,false);assert.equal(el('excelPaletteBlock').hidden,true);
el('excelContent').value='Смета';el('excelSubmit').click();
assert.equal(sent.at(-1).payload.excel_style,'standard');assert.equal(sent.at(-1).payload.excel_palette,'green');
el('excelStyle').value='designer';el('excelStyle').click();el('excelPalette').value='copper';
assert.equal(el('excelPaletteBlock').hidden,false);el('excelSubmit').click();
assert.equal(sent.at(-1).payload.excel_palette,'copper');
el('excelStyle').value='standard';el('excelStyle').click();el('excelSubmit').click();
assert.equal(sent.at(-1).payload.excel_palette,'green');
for (const kind of ['family_budget','project_budget','price_list','calc_table']) {
 el('excelKind').value=kind;el('excelKind').click();
 assert.equal(el('excelAppearanceBlock').hidden,false);
 el('excelStyle').value='designer';el('excelStyle').click();el('excelSubmit').click();
 assert.equal(sent.at(-1).payload.excel_style,'designer');
 assert.equal(sent.at(-1).payload.excel_palette,'copper');
}
el('excelKind').value='startup_model';el('excelKind').click();el('excelSubmit').click();
assert.equal(el('excelAppearanceBlock').hidden,false);assert.equal(sent.at(-1).payload.excel_style,'designer');
assert.equal(sent.at(-1).payload.mode,'user');assert.equal(el('excelModeBlock').hidden,true);assert.equal(el('excelStartupHint').hidden,false);
el('excelKind').value='calc_table';el('excelKind').click();assert.equal(el('excelModeBlock').hidden,false);assert.equal(el('excelStartupHint').hidden,true);
assert.match(html,/data-cost="5"/);
console.log('PASS: estimate layout selection, palette visibility, standard fixed green, outgoing payload.');
