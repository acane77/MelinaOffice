// A document line grid stretches a line to whole pitches. Inside a table that happens only when
// the file sets w:adjustLineHeightInTable; otherwise the cell keeps the paragraph's own line
// spacing (LibreOffice SwTextFormatter::CalcRealHeight, tdf#88752 / tdf#167583). A font a little
// taller than one pitch used to jump to the next pitch, which is the empty band above and below
// the glyphs. node --test tests/ecrivaine/table-grid.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

function node() {
  const n = {
    nodeType: 1, nodeName: 'DIV', style: {}, dataset: {}, children: [], childNodes: [],
    className: '', lastChild: null, firstChild: null,
    classList: { add() {}, remove() {}, contains: () => false },
    appendChild(c) { this.childNodes.push(c); this.children.push(c); this.lastChild = c; if (!this.firstChild) this.firstChild = c; return c; },
    querySelector() { return null; },
    querySelectorAll: () => [],
  };
  return n;
}
const document = {
  createElement: (name) => { const n = node(); n.nodeName = String(name).toUpperCase(); return n; },
  createElementNS: (_ns, name) => { const n = node(); n.nodeName = String(name).toUpperCase(); return n; },
  createTextNode: (t) => ({ nodeType: 3, data: String(t), appendData(s) { this.data += s; } }),
};
const sandbox = { window: { document }, document, console };
sandbox.window.window = sandbox.window;
vm.createContext(sandbox);
for (const file of ['../../public/common/core.js', '../../public/ecrivaine/js/dmodel.js', '../../public/ecrivaine/js/render.js']) {
  vm.runInContext(fs.readFileSync(new URL(file, import.meta.url), 'utf8'), sandbox, { filename: file });
}
const { D, R } = sandbox.window.L;
R.lineBoxes = () => {};

const PITCH = 15.6;
const gridPx = R.px(PITCH);
const docOf = (compatXML) => ({
  styles: { Normal: { id: 'Normal', type: 'paragraph', basedOn: null, pPr: {}, rPr: {}, isDefault: true } },
  defaults: { pPr: {}, rPr: { font: 'Times New Roman', sz: 12 } },
  settings: compatXML == null ? {} : { compatXML },
  numbering: { abs: {}, nums: {} },
});
const drawn = (sz, font, ctx, pPr) => {
  const p = D.para([D.text('标题 Heading', { sz, font })], pPr || {});
  const outer = R.para(p, 0, null, Object.assign({ doc: ctx.doc, grid: gridPx }, ctx));
  const pc = outer.children[0];
  const span = pc.children.find((c) => c.dataset && c.dataset.o != null);
  const lh = (css) => (String(css || '').match(/line-height:([^;]+)/) || [])[1];
  return { para: lh(pc.style.cssText), run: span ? lh(span.style.cssText) : '' };
};
const pitches = (sz, font, mult) => R.gridLine(gridPx, mult, sz, font) / gridPx;

test('the flag is on only when adjustLineHeightInTable is present and not turned off', () => {
  assert.equal(R.adjustLineHeightInTable({ settings: {} }), false);
  assert.equal(R.adjustLineHeightInTable({ settings: { compatXML: '' } }), false);
  assert.equal(R.adjustLineHeightInTable({ settings: { compatXML: '<w:doNotUseHTMLParagraphAutoSpacing/>' } }), false);
  assert.equal(R.adjustLineHeightInTable({ settings: { compatXML: '<w:adjustLineHeightInTable/>' } }), true);
  assert.equal(R.adjustLineHeightInTable({ settings: { compatXML: '<w:adjustLineHeightInTable w:val="true"/>' } }), true);
  assert.equal(R.adjustLineHeightInTable({ settings: { compatXML: '<w:adjustLineHeightInTable w:val="1"/>' } }), true);
  assert.equal(R.adjustLineHeightInTable({ settings: { compatXML: '<w:adjustLineHeightInTable w:val="0"/>' } }), false);
  assert.equal(R.adjustLineHeightInTable({ settings: { compatXML: '<w:adjustLineHeightInTable w:val="false"/>' } }), false);
  assert.equal(R.adjustLineHeightInTable({ settings: { compatXML: '<w:adjustLineHeightInTable w:val="off"/>' } }), false);
});

test('a table cell uses the paragraph line spacing, at every size, unless the document asks for the grid', () => {
  const doc = docOf();
  const on = docOf('<w:compat><w:adjustLineHeightInTable/></w:compat>');
  for (const [sz, font] of [[10.5, 'SimSun'], [12, '宋体'], [13, 'Times New Roman'], [14, '黑体'], [16, 'SimHei'], [22, '黑体'], [28, 'Calibri'], [36, 'Times New Roman'], [48, 'Arial']]) {
    const cell = drawn(sz, font, { doc, inCell: true });
    const body = drawn(sz, font, { doc });
    const forced = drawn(sz, font, { doc: on, inCell: true });
    const factor = String(R.lineFactor(font));
    assert.equal(cell.para, factor, `cell ${sz}pt ${font} paragraph`);
    assert.equal(cell.run, factor, `cell ${sz}pt ${font} run`);
    assert.equal(body.para, gridPx + 'px', `body ${sz}pt ${font} still snaps`);
    assert.equal(forced.para, body.para, `flagged cell ${sz}pt ${font}`);
    assert.equal(forced.run, body.run, `flagged cell run ${sz}pt ${font}`);
    const need = pitches(sz, font, 1);
    if (need > 1) assert.equal(body.run, R.gridLine(gridPx, 1, sz, font) + 'px', `body ${sz}pt ${font} takes ${need} pitches`);
  }
});

test('multiple, exact and at-least spacing in a cell follow the paragraph, not the grid', () => {
  const doc = docOf();
  const multi = drawn(22, 'Times New Roman', { doc, inCell: true }, { sp: { line: 1.5, rule: 'auto' } });
  assert.equal(multi.para, String(Lround(1.149 * 1.5)));
  assert.equal(multi.run, multi.para);
  const body = drawn(22, 'Times New Roman', { doc }, { sp: { line: 1.5, rule: 'auto' } });
  assert.equal(body.para, R.gridLine(gridPx, 1.5, 0, 'Times New Roman') + 'px');
  const exact = drawn(22, '黑体', { doc, inCell: true }, { sp: { line: 20, rule: 'exact' } });
  assert.equal(exact.para, R.px(20) + 'px');
  const least = drawn(22, '黑体', { doc, inCell: true }, { sp: { line: 28, rule: 'atLeast' } });
  assert.equal(least.para, R.px(28) + 'px');
});

test('a paragraph that opts out of the grid stays off it in the body too', () => {
  const doc = docOf('<w:adjustLineHeightInTable/>');
  const out = drawn(22, 'Times New Roman', { doc }, { noSnap: true });
  assert.equal(out.para, '1.149');
  assert.equal(out.run, '1.149');
});

function Lround(v) { return String(Math.round(v * 10000) / 10000); }
