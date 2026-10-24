// Top-and-bottom wrap in a multi-column section clears the object's vertical span in every column it
// crosses (public/quire/js/layout.js, public/quire/js/drawing.js). A full-width box at the top of the page
// pushes both columns down; a box that stays inside one column leaves the other at the top. The IJCAI
// template (.test_docs/ijcai22.docx) is the same case: its title text box spans the page.
//   node --test tests/quire/columns-wrap.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const ijcai = new URL('../../.test_docs/ijcai22.docx', import.meta.url);

/** body-text rectangles in one column, and the floating shape that holds `mark` */
const MEASURE = `(() => {
  const pg = L.layout.pages[0];
  const boxEl = [...pg.el.querySelectorAll('.shp')].find((e) => e.textContent.includes(MARK));
  if (!boxEl) return { missing: true };
  const box = boxEl.getBoundingClientRect();
  const overlaps = (r) => r.width > 0 && r.height > 0 && r.left < box.right - 1 && r.right > box.left + 1 && r.top < box.bottom - 1 && r.bottom > box.top + 1;
  const cols = [...pg.body.querySelectorAll('.pg-col')].map((col) => {
    const texts = [];
    const w = document.createTreeWalker(col, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = w.nextNode())) {
      if (!n.data.trim() || n.parentElement.closest('.flt, .tbx, .shp, .anc')) continue;
      const r = n.parentElement.getBoundingClientRect();
      if (r.height > 0) texts.push({ top: r.top, bottom: r.bottom, left: r.left, right: r.right });
    }
    const colR = col.getBoundingClientRect();
    return {
      textTop: texts.length ? Math.min(...texts.map((t) => t.top)) : null,
      colTop: colR.top,
      n: texts.length,
      overlap: texts.filter(overlaps).length,
    };
  });
  return { box: { top: box.top, bottom: box.bottom }, cols };
})()`;

function built(wide) {
  return `(async () => {
    VO.opened = () => {}; L.layout.view = 'print'; L.layout.setZoom(1);
    const D = L.D, d = D.newDoc();
    d.sect.ml = 72; d.sect.mr = 72; d.sect.mt = 72; d.sect.mb = 72;
    d.sect.cols = { n: 2, space: 36, sep: false, eq: true, w: [] };
    const box = L.drawing.newShape('rect', ${wide ? 468 : 140}, 100, { textbox: true });
    box.tb.blocks = [D.para([D.text('Title box')])];
    box.line = { t: 'none' };
    L.drawing.setWrap(box, 'topBottom');
    box.float.posH = ${wide ? "{ rel: 'page', align: 'center' }" : "{ rel: 'column', off: 0 }"};
    box.float.posV = { rel: 'page', off: 72 };
    box.float.dist = { t: 0, b: 0, l: 0, r: 0 };
    const sentence = 'The quick brown fox jumps over the lazy dog beside the title. ';
    const blocks = [D.para([box, D.text('Left column starts here. ' + sentence)])];
    for (let i = 0; i < 14; i++) blocks.push(D.para([D.text('Body ' + i + ' ' + sentence.repeat(3))]));
    d.main.blocks = blocks;
    L.app.loadDoc(d, 'columns', { saved: true });
    await document.fonts.ready; L.layout.render(); await new Promise((r) => setTimeout(r, 200));
    const MARK = 'Title box';
    return ${MEASURE};
  })()`;
}

test('A full-width top-and-bottom object pushes every column it crosses below it', async () => {
  const stop = await ensureServer();
  let page;
  try {
    page = await openPage('quire', { persistence: false, automated: true });
    const got = await page.evaluate(built(true));
    assert.equal(got.missing, undefined, 'the text box is on the page');
    assert.equal(got.cols.length, 2);
    for (const col of got.cols) {
      assert.ok(col.n > 0, 'each column has body text');
      assert.ok(col.textTop >= got.box.bottom - 2, `column text at ${col.textTop} overlaps the box ending at ${got.box.bottom}`);
      assert.equal(col.overlap, 0);
    }
    assert.deepEqual(page.errors, []);
  } finally { if (page) await page.close(); stop(); }
});

test('A top-and-bottom object inside one column leaves the other column at the top', async () => {
  const stop = await ensureServer();
  let page;
  try {
    page = await openPage('quire', { persistence: false, automated: true });
    const got = await page.evaluate(built(false));
    assert.equal(got.missing, undefined, 'the text box is on the page');
    assert.equal(got.cols.length, 2);
    const [left, right] = got.cols;
    assert.ok(left.n > 0 && right.n > 0, 'each column has body text');
    assert.ok(left.textTop >= got.box.bottom - 2, `left column text at ${left.textTop} overlaps the box ending at ${got.box.bottom}`);
    assert.ok(right.textTop < got.box.bottom - 30, `right column was pushed to ${right.textTop}, box ends at ${got.box.bottom}`);
    assert.equal(right.overlap, 0);
    assert.deepEqual(page.errors, []);
  } finally { if (page) await page.close(); stop(); }
});

test('IJCAI title box: Abstract and Style and Format both start below it', { skip: !fs.existsSync(ijcai) && `no ${ijcai.pathname}` }, async () => {
  const stop = await ensureServer();
  let page;
  try {
    page = await openPage('quire', { persistence: false, automated: true });
    const b64 = fs.readFileSync(ijcai).toString('base64');
    const got = await page.evaluate(`(async () => {
      VO.opened = () => {}; L.layout.view = 'print'; L.layout.setZoom(1);
      const bytes = Uint8Array.from(atob(${JSON.stringify(b64)}), (c) => c.charCodeAt(0));
      const { doc } = await L.docx.read(bytes.buffer);
      L.app.loadDoc(doc, 'ijcai', { saved: true });
      await document.fonts.ready; L.layout.render(); await new Promise((r) => setTimeout(r, 300));
      const pg = L.layout.pages[0], body = pg.body;
      const boxEl = [...pg.el.querySelectorAll('.shp')].find((e) => e.textContent.includes('Formatting Instructions'));
      const find = (s) => {
        const w = document.createTreeWalker(body, NodeFilter.SHOW_TEXT);
        let n;
        while ((n = w.nextNode())) {
          if (n.data.includes(s) && !n.parentElement.closest('.flt, .tbx, .shp')) return n.parentElement.getBoundingClientRect().top;
        }
        return null;
      };
      const box = boxEl.getBoundingClientRect();
      return { boxBottom: box.bottom, abstractTop: find('Abstract'), styleTop: find('Style and Format'), errors: null };
    })()`);
    assert.ok(got.abstractTop != null && got.styleTop != null, 'both headings are on the first page');
    assert.ok(got.abstractTop >= got.boxBottom - 2, `Abstract at ${got.abstractTop}, title box ends at ${got.boxBottom}`);
    assert.ok(got.styleTop >= got.boxBottom - 2, `Style and Format at ${got.styleTop}, title box ends at ${got.boxBottom}`);
    assert.deepEqual(page.errors, []);
  } finally { if (page) await page.close(); stop(); }
});
