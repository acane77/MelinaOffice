const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
globalThis.document = { addEventListener() {} };
require('../../public/common/core.js');
require('../../public/quire/js/dmodel.js');
require('../../public/quire/js/ops.js');
L.layout = L.layout || {};
require('../../public/quire/js/editor.js');

const view = { top: 100, bottom: 1100, left: 0, right: 800 }; // 1000px tall, band line at 1000

test('a caret in the top nine tenths does not move the page', () => {
  assert.deepEqual(L.ed.caretScroll(view, { top: 400, bottom: 420, left: 40, right: 42 }).dy, 0);
  assert.deepEqual(L.ed.caretScroll(view, { top: 980, bottom: 1000, left: 40, right: 42 }).dy, 0);
});

test('crossing the bottom tenth parks the line on that line', () => {
  assert.equal(L.ed.caretScroll(view, { top: 990, bottom: 1010, left: 40, right: 42 }).dy, 10);
  assert.equal(L.ed.caretScroll(view, { top: 1400, bottom: 1420, left: 40, right: 42 }).dy, 420);
});

test('a line above the window comes back to the top edge only', () => {
  assert.equal(L.ed.caretScroll(view, { top: 40, bottom: 60, left: 40, right: 42 }).dy, -60);
  assert.equal(L.ed.caretScroll(view, { top: 90, bottom: 110, left: 40, right: 42 }).dy, -10);
});

test('horizontal follow is the overflow plus a small margin', () => {
  assert.equal(L.ed.caretScroll(view, { top: 400, bottom: 420, left: -5, right: -3 }).dx, -35);
  assert.equal(L.ed.caretScroll(view, { top: 400, bottom: 420, left: 810, right: 812 }).dx, 42);
});
