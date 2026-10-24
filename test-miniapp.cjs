// Run with: node test-miniapp.cjs. No browser, network or paid API calls.
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const html = fs.readFileSync(__dirname + '/index.html', 'utf8');
const chips = [...html.matchAll(/<(?:div|button)\b([^>]*data-group="[^"]+"[^>]*)>/g)].map(m => {
  const attr = name => m[1].match(new RegExp(name + '="([^"]*)"'))?.[1];
  const classes = new Set((attr('class') || '').split(' '));
  return {dataset: {group: attr('data-group'), value: attr('data-value')},
    hidden: /\bhidden\b/.test(m[1]), events: [],
    classList: {add: c => classes.add(c), remove: c => classes.delete(c),
      contains: c => classes.has(c), toggle: (c, v) => v ? classes.add(c) : classes.delete(c)},
    addEventListener(_, f) {this.events.push(f)}, click() {this.events.forEach(f => f())}};
});
const ids = {};
const element = id => ids[id] ||= {value: '', textContent: '', events: [],
  addEventListener(_, f) {this.events.push(f)}, click() {this.events.forEach(f => f())}, focus(){}};
function query(selector) {
  const group = selector.match(/data-group="([^"]+)"/)?.[1];
  return chips.filter(c => (!group || c.dataset.group === group) && (!selector.endsWith('.active') || c.classList.contains('active')));
}
const sent = [];
const ctx = vm.createContext({document: {querySelectorAll: query, querySelector: s => query(s)[0], getElementById: element},
  getTg: () => null, tpl: (_, v) => String(v.n), sendAction: (action, payload) => sent.push({action, payload})});
const generic = html.slice(html.indexOf("  document.querySelectorAll('.field-chip[data-group]')"), html.indexOf('  // --- Экран пополнения'));
const pricing = html.slice(html.indexOf('  // --- Степпер числа слайдов'), html.indexOf('  // --- Степпер числа страниц'));
const submit = html.slice(html.indexOf('  // --- Отправка формы презентации'), html.indexOf('  // --- Отправка формы Word'));
for (const script of html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)) new vm.Script(script[1]);
vm.runInContext(generic + pricing + submit, ctx);
const click = (group, value) => chips.find(c => c.dataset.group === group && c.dataset.value === value).click();
const price = n => assert.equal(element('presCostLabel').textContent, String(n));
price(128);
click('presKind', 'designer'); price(256);
assert.deepEqual(chips.filter(c => c.dataset.group === 'presStyle' && !c.hidden).map(c => c.dataset.value), ['designer_strict', 'designer_soft']);
element('presTopic').value = 'Проверка';
element('presSubmit').click();
assert.equal(sent.at(-1).payload.style, 'designer_strict');
click('presStyle', 'designer_soft'); price(256);
element('presSlidesPlus').click(); price(288);
click('presPhoto', 'own');
element('presSubmit').click();
assert.equal(sent.at(-1).action, 'gen_presentation');
assert.equal(sent.at(-1).payload.style, 'designer_soft');
assert.equal(sent.at(-1).payload.slides, 9);
assert.equal(sent.at(-1).payload.photo_mode, 'own');
click('presKind', 'standard'); price(144);
click('presStyle', 'business');
element('presSubmit').click(); assert.equal(sent.at(-1).payload.style, 'business');
click('presKind', 'designer'); price(288);
assert.equal(query('.field-chip[data-group="presStyle"].active')[0].dataset.value, 'designer_soft');
for (let i=0;i<30;i++) element('presSlidesMinus').click(); price(128);
for (let i=0;i<30;i++) element('presSlidesPlus').click(); price(640);
click('presKind', 'standard'); price(320);
assert.equal(query('.field-chip[data-group="presStyle"].active')[0].dataset.value, 'business');
console.log('PASS: JS syntax; styles, live prices, slide limits, selection memory and submitted payloads.');
