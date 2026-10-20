const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const html = fs.readFileSync(__dirname + '/index.html', 'utf8');
const ids = {};
const element = id => ids[id] ||= {value: '', textContent: '', style: {}, hidden: false, events: [],
  addEventListener(_, fn) { this.events.push(fn); }, click() { this.events.forEach(fn => fn()); }, focus() {}};
const sent = [];
element('wordKind').value = 'essay';
const ctx = vm.createContext({uiLang: 'ru', document: {getElementById: element},
  t: key => key, tpl: (_, vars) => String(vars.n), sendAction: (action, payload) => sent.push({action, payload})});
const stepper = html.slice(html.indexOf('  // --- Степпер числа страниц'), html.indexOf('  // --- Отправка формы презентации'));
const submit = html.slice(html.indexOf('  // --- Отправка формы Word'), html.indexOf('  // --- Отправка формы Excel'));
vm.runInContext(stepper + submit, ctx);
for (let i=0; i<4; i++) element('wordPagesPlus').click();
assert.equal(element('wordPagesValue').textContent, 6);
assert.equal(element('wordCoverPrefix').hidden, false);
assert.match(element('wordVolumeHint').textContent, /Всего: 7 стр/);
assert.equal(element('wordCostLabel').textContent, '30');
element('wordContent').value = 'Тема эссе';
element('wordSubmit').click();
assert.equal(sent.at(-1).payload.pages, 6);
assert.equal(sent.at(-1).payload.kind, 'essay');
for (const kind of ['referat','coursework','report','notes']) {
  element('wordKind').value = kind; element('wordKind').click();
  assert.equal(element('wordCoverPrefix').hidden, false);
  assert.match(element('wordVolumeHint').textContent, new RegExp('Всего: '+(['referat','coursework'].includes(kind)?8:7)+' стр'));
  assert.equal(element('wordCostLabel').textContent, '30');
}
element('wordKind').value = 'doc'; element('wordKind').click();
assert.equal(element('wordCoverPrefix').hidden, true);
assert.equal(element('wordVolumeHint').hidden, true);
element('wordKind').value = 'services'; element('wordKind').click();
assert.equal(element('wordPagesBlock').style.display, 'none');
element('wordSubmit').click();
assert.equal(sent.at(-1).payload.pages, 2);
assert.equal(element('wordCostLabel').textContent, '10');
console.log('PASS: academic cover/body labels, totals, unchanged credits and submitted body pages; other documents unaffected.');
