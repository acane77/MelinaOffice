const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/ecrivaine/js/dmodel.js');
const { D } = L;

function doc() {
  const d = D.newDoc();
  D.doc = d;
  return d;
}

test('a TOC entry keeps the Hyperlink style but draws automatic colour and no underline', () => {
  const d = doc();
  const p = D.para([D.text('Chapter', { style: 'Hyperlink', link: { anchor: '_Toc1' } })], { style: 'TOC1' });
  const r = D.rProps(d, p, p.runs[0]);
  assert.equal(r.style, 'Hyperlink');
  assert.equal(r.color, undefined);
  assert.equal(r.u, undefined);
  assert.equal(r.uColor, undefined);
});

test('the same Hyperlink style stays blue and underlined outside a contents paragraph', () => {
  const d = doc();
  const p = D.para([D.text('site', { style: 'Hyperlink', link: { url: 'https://example.test' } })]);
  const r = D.rProps(d, p, p.runs[0]);
  assert.equal(r.color, '0000FF');
  assert.equal(r.u, 'single');
});

test('a contents paragraph style colour shows through, and direct formatting still wins', () => {
  const d = doc();
  d.styles.TOC2.rPr = { color: 'AA0000', u: 'single' };
  D.stylesChanged();
  const plain = D.para([D.text('Section', { style: 'Hyperlink' })], { style: 'TOC2' });
  const over = D.para([D.text('Section', { style: 'Hyperlink', color: '00AA00', u: 'none' })], { style: 'TOC2' });
  const a = D.rProps(d, plain, plain.runs[0]);
  const b = D.rProps(d, over, over.runs[0]);
  assert.equal(a.color, 'AA0000');
  assert.equal(a.u, 'single');
  assert.equal(b.color, '00AA00');
  assert.equal(b.u, 'none');
});

test('FollowedHyperlink and a renamed style id follow the same contents rule', () => {
  const d = doc();
  d.styles.af = { id: 'af', name: 'Hyperlink', type: 'character', basedOn: null, pPr: {}, rPr: { color: '0563C1', u: 'single', uColor: '0563C1' } };
  d.styles.Toc3 = { id: 'Toc3', name: 'toc 3', type: 'paragraph', basedOn: 'Normal', pPr: {}, rPr: {} };
  D.stylesChanged();
  const followed = D.para([D.text('Visited', { style: 'FollowedHyperlink' })], { style: 'TableofFigures' });
  const renamed = D.para([D.text('Renamed', { style: 'af' })], { style: 'Toc3' });
  const f = D.rProps(d, followed, followed.runs[0]);
  const n = D.rProps(d, renamed, renamed.runs[0]);
  assert.equal(f.color, undefined);
  assert.equal(f.u, undefined);
  assert.equal(n.color, undefined);
  assert.equal(n.u, undefined);
  assert.equal(n.uColor, undefined);
});
