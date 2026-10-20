/* A click selects the row the grid painted under the pointer, including at the blank area past the last
   row that has any height. node --test tests/ledger/hit.test.js */
const { test } = require('node:test');
const assert = require('node:assert/strict');
require('../../public/common/core.js');
require('../../public/ledger/js/model.js');
require('../../public/ledger/js/layout.js');
require('../../public/ledger/js/render.js');
const L = globalThis.L;
const MAXR = L.model.MAXR, MAXC = L.model.MAXC;

function frame(rows, cols, z, pane) {
  const p = Object.assign({ r0: 0, c0: 0, rMax: MAXR, cMax: MAXC, x: 40, y: 22, w: 900, h: 700 }, pane);
  return { hdrW: 40, hdrH: 22, z, g: { rows, cols }, panes: [p] };
}
const cols = new L.layout.Axis(64, MAXC, new Map());
const hit = (fr, x, y) => L.render.hit(fr, x, y);

test('the centre of a painted cell is that cell, at every zoom', () => {
  const rows = new L.layout.Axis(20, MAXR, new Map());
  for (const z of [0.1, 0.25, 0.5, 1, 1.1, 1.5, 2]) {
    const fr = frame(rows, cols, z);
    const p = fr.panes[0];
    for (let r = 0; r < 40; r++) {
      const y1 = L.render.rowY(fr, p, r), y2 = L.render.rowY(fr, p, r + 1);
      if (y2 <= y1 || y1 < p.y || y2 > p.y + p.h) continue;
      const y = (y1 + y2 - 1) / 2;
      const x1 = L.render.colX(fr, p, 2), x2 = L.render.colX(fr, p, 3);
      const got = hit(fr, (x1 + x2 - 1) / 2, y);
      assert.equal(got.area, 'cell', 'z=' + z + ' r=' + r);
      assert.equal(got.r, r, 'z=' + z + ' y=' + y);
      assert.equal(got.c, 2, 'z=' + z);
    }
  }
});

test('a click on the rounded boundary stays on the cell that owns that pixel', () => {
  const over = new Map();
  for (let i = 0; i < 500; i += 7) over.set(i, i % 3 === 0 ? 0 : (5 + (i % 11) * 3));
  const rows = new L.layout.Axis(20, MAXR, over);
  const fr = frame(rows, cols, 0.1);
  const p = fr.panes[0];
  /* this point used to resolve two rows high, because hit testing undid the rounding the grid draws with */
  const y1 = L.render.rowY(fr, p, 43), y2 = L.render.rowY(fr, p, 44);
  assert.ok(y2 > y1);
  assert.equal(hit(fr, 80, y1 + (y2 - y1) * 0.1).r, 43);
  assert.equal(hit(fr, 80, y2 - 1).r, 43);
  assert.equal(hit(fr, 80, y2).r, 44);
});

test('the blank area under the last row that has a height is that row, not the end of the sheet', () => {
  const rows = new L.layout.Axis(0, MAXR, new Map([[0, 20], [4, 30]]));
  const fr = frame(rows, cols, 1);
  const p = fr.panes[0];
  const top = L.render.rowY(fr, p, 4);
  const bottom = L.render.rowY(fr, p, 5);
  assert.equal(hit(fr, 80, top + 10).r, 4);
  assert.equal(hit(fr, 80, bottom + 200).r, 4);
  assert.ok(hit(fr, 80, bottom + 200).r < 1000000);
  assert.equal(hit(fr, 20, bottom + 40).area, 'rowhdr');
  assert.equal(hit(fr, 20, bottom + 40).r, 4);
});

test('hidden rows do not steal a click on the next visible row', () => {
  const over = new Map();
  for (let r = 1; r <= 5; r++) over.set(r, 0);
  const rows = new L.layout.Axis(20, MAXR, over);
  const fr = frame(rows, cols, 1);
  const p = fr.panes[0];
  const y = L.render.rowY(fr, p, 6) + 8;
  assert.equal(hit(fr, 80, y).r, 6);
  assert.equal(hit(fr, 80, L.render.rowY(fr, p, 0) + 8).r, 0);
});

test('a scrolled pane and the column header use the same edges', () => {
  const rows = new L.layout.Axis(20, MAXR, new Map());
  const fr = frame(rows, cols, 1, { r0: 100, c0: 3 });
  const p = fr.panes[0];
  const y = L.render.rowY(fr, p, 103) + 5;
  const x = L.render.colX(fr, p, 5) + 5;
  assert.equal(hit(fr, x, y).r, 103);
  assert.equal(hit(fr, x, y).c, 5);
  assert.equal(hit(fr, x, 8).area, 'colhdr');
  assert.equal(hit(fr, x, 8).c, 5);
  assert.equal(hit(fr, 10, 8).area, 'corner');
});
