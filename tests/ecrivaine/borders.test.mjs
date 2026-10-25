// Paragraph borders join across neighbours that share them. A missing indent is zero, so the
// blank lines pressed in front of a bottom rule stay blank lines. node --test tests/ecrivaine/borders.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

function node() {
  const n = {
    style: {}, dataset: {}, children: [], className: '', lastChild: null,
    classList: { add() {}, remove() {}, contains: () => false },
    appendChild(c) { this.children.push(c); this.lastChild = c; return c; },
    querySelector: () => null,
    querySelectorAll: () => [],
  };
  return n;
}
const document = { createElement: () => node(), createElementNS: () => node(), createTextNode: (t) => ({ text: t }) };
const L = {
  fontKey: (s) => String(s || '').toLowerCase(),
  fontStack: (name) => name || 'Times New Roman',
  round: (n, d) => { const f = 10 ** (d || 0); return Math.round(n * f) / f; },
  clamp: (n, a, b) => Math.min(b, Math.max(a, n)),
  color: { hexToRgb: () => [0, 0, 0], mix: () => '#000000' },
  t: (s) => s,
};
const sandbox = { window: { L, document }, document, console, L };
sandbox.window.L = L;
vm.createContext(sandbox);
for (const file of ['dmodel.js', 'render.js']) {
  vm.runInContext(fs.readFileSync(new URL('../../public/ecrivaine/js/' + file, import.meta.url), 'utf8'), sandbox);
}
const { D, R } = sandbox.window.L;
R.lineBoxes = () => {};

const rule = { val: 'single', sz: 0.75, color: 'auto', space: 1 };
const doc = {
  styles: { Normal: { id: 'Normal', type: 'paragraph', basedOn: null, pPr: {}, rPr: {}, isDefault: true } },
  defaults: { pPr: {}, rPr: { font: 'Times New Roman', sz: 12 } },
  settings: {},
  numbering: { abs: {}, nums: {} },
};
const para = (borders, ind) => D.para([], Object.assign({ borders }, ind ? { ind } : {}));
const css = (p, prev, next) => R.para(p, 0, null, { doc, prev, next }).children[0].style.cssText || '';
const ruled = (s) => /border-bottom:(?!none)/.test(s);

test('blank lines in front of a bottom rule do not each draw the rule', () => {
  const a = para({ bottom: rule });
  const b = para({ bottom: rule });
  const c = para({ bottom: rule });
  const after = D.para();
  assert.equal(ruled(css(a, null, b)), false);
  assert.equal(ruled(css(b, a, c)), false);
  assert.equal(ruled(css(c, b, after)), true);
  assert.equal(ruled(css(after, c, null)), false);
});

test('a different left indent keeps its own bottom rule', () => {
  const a = para({ bottom: rule });
  const b = para({ bottom: rule }, { l: 36 });
  assert.equal(ruled(css(a, null, b)), true);
  assert.equal(ruled(css(b, a, null)), true);
});
