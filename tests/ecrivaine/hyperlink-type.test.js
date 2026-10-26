const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/ecrivaine/js/dmodel.js');
require('../../public/ecrivaine/js/ops.js');
const { D, O } = L;

function docWith(runs, mark) {
  const doc = D.newDoc();
  const p = D.para(runs, {}, mark || {});
  doc.main.blocks = [p];
  D.doc = doc;
  D.reindex(doc);
  return p;
}
const link = { url: 'https://example.com' };
const props = (r) => { const o = Object.assign({}, r); delete o.ins; delete o.del; return o; };

test('text typed after a hyperlink is not a hyperlink, and keeps other formatting', () => {
  const p = docWith([D.text('see ', {}), D.text('site', { style: 'Hyperlink', link, b: true, color: '0000FF', u: 'single' })]);
  const end = D.plen(p);
  const got = O.inheritRPr(D.pos(p, end));
  assert.equal(got.link, undefined);
  assert.equal(got.style, undefined);
  assert.equal(got.color, undefined);
  assert.equal(got.u, undefined);
  assert.equal(got.b, true);
  D.tx('Type', () => { O.insertText(D.pos(p, end), 'X'); });
  D.tx('Type', () => { O.insertText(D.pos(p, end + 1), 'Y'); });
  assert.equal(D.ptext(p), 'see siteXY');
  const typed = p.runs.filter((r) => r.text === 'XY' || r.text === 'X' || r.text === 'Y');
  assert.ok(typed.length);
  for (const r of typed) {
    assert.equal(r.rPr.link, undefined);
    assert.equal(r.rPr.style, undefined);
    assert.equal(r.rPr.b, true);
  }
});

test('typing between a link and the plain text after it does not extend the link', () => {
  const p = docWith([D.text('site', { style: 'Hyperlink', link }), D.text(' please', {})]);
  D.tx('Type', () => { O.insertText(D.pos(p, 4), 'Z'); });
  assert.equal(D.ptext(p), 'siteZ please');
  assert.equal(p.runs[0].text, 'site');
  assert.equal(p.runs[0].rPr.style, 'Hyperlink');
  const rest = p.runs.slice(1);
  assert.ok(rest.every((r) => !r.rPr.link && r.rPr.style !== 'Hyperlink'));
});

test('text typed inside a hyperlink stays in it', () => {
  const p = docWith([D.text('site', { style: 'Hyperlink', link })]);
  const got = O.inheritRPr(D.pos(p, 2));
  assert.equal(got.style, 'Hyperlink');
  assert.equal(got.link.url, link.url);
  D.tx('Type', () => { O.insertText(D.pos(p, 2), 'X'); });
  assert.equal(D.ptext(p), 'siXte');
  assert.ok(p.runs.every((r) => r.rPr && r.rPr.style === 'Hyperlink' && r.rPr.link && r.rPr.link.url === link.url));
});

test('text typed at the start of a leading hyperlink is not a hyperlink', () => {
  const p = docWith([D.text('site', { style: 'Hyperlink', link, i: true })]);
  const got = O.inheritRPr(D.pos(p, 0));
  assert.equal(got.link, undefined);
  assert.equal(got.style, undefined);
  assert.equal(got.i, true);
});

test('a renamed hyperlink style and FollowedHyperlink stop at the end of the link', () => {
  const doc = D.newDoc();
  doc.styles.af = { id: 'af', name: 'Hyperlink', type: 'character', basedOn: null, pPr: {}, rPr: { color: '0563C1', u: 'single' } };
  const p = D.para([
    D.text('a', { style: 'af', link, color: '0563C1', u: 'single' }),
    D.text('b', { style: 'FollowedHyperlink', link: { url: 'https://other.test' }, color: '800080', u: 'single' }),
  ]);
  doc.main.blocks = [p];
  D.doc = doc;
  D.reindex(doc);
  const afterA = props(O.inheritRPr(D.pos(p, 1)));
  assert.equal(afterA.style, undefined);
  assert.equal(afterA.link, undefined);
  assert.equal(afterA.color, undefined);
  assert.equal(afterA.u, undefined);
  const afterB = props(O.inheritRPr(D.pos(p, 2)));
  assert.equal(afterB.style, undefined);
  assert.equal(afterB.link, undefined);
  assert.equal(afterB.color, undefined);
});

test('a colour the user set on the link continues, the link does not', () => {
  const p = docWith([D.text('site', { style: 'Hyperlink', link, color: 'FF0000', u: 'single', sz: 18 })]);
  const got = props(O.inheritRPr(D.pos(p, 4)));
  assert.equal(got.link, undefined);
  assert.equal(got.style, undefined);
  assert.equal(got.u, undefined);
  assert.equal(got.color, 'FF0000');
  assert.equal(got.sz, 18);
});

test('Word and browser link colours do not continue after the link', () => {
  for (const color of ['0563C1', '0000EE', '467886', '954F72', '96607D']) {
    const p = docWith([D.text('site', { style: 'Hyperlink', link, color, u: 'single' })]);
    const got = props(O.inheritRPr(D.pos(p, 4)));
    assert.equal(got.color, undefined, color);
    assert.equal(got.style, undefined, color);
    assert.equal(got.link, undefined, color);
  }
});

test('making a whole paragraph a hyperlink does not put the link on the paragraph mark', () => {
  const doc = D.newDoc();
  const p = D.para([D.text('whole')]);
  doc.main.blocks = [p];
  D.doc = doc;
  D.reindex(doc);
  D.tx('Link', () => O.formatRange(D.pos(p, 0), D.pos(p, 5), (r) => Object.assign(r, { link, style: 'Hyperlink' })));
  assert.equal(p.runs[0].rPr.style, 'Hyperlink');
  assert.equal(p.rPr.link, undefined);
  assert.equal(p.rPr.style, undefined);
  const got = props(O.inheritRPr(D.pos(p, 5)));
  assert.equal(got.link, undefined);
  assert.equal(got.style, undefined);
});

test('an empty paragraph whose mark still carries a link does not type as a link', () => {
  const p = docWith([], { style: 'Hyperlink', link, b: true });
  const got = props(O.inheritRPr(D.pos(p, 0)));
  assert.equal(got.link, undefined);
  assert.equal(got.style, undefined);
  assert.equal(got.b, true);
});

test('the Hyperlink character style without a link still continues, like any other character style', () => {
  const p = docWith([D.text('note', { style: 'Hyperlink', i: true })]);
  const got = props(O.inheritRPr(D.pos(p, 4)));
  assert.equal(got.style, 'Hyperlink');
  assert.equal(got.i, true);
});
