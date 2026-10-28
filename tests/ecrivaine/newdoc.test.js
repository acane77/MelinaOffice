const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/common/locales.js');
require('../../public/ecrivaine/js/dmodel.js');

test('a new document is laid out as Microsoft Word 2013 and later', () => {
  assert.equal(L.D.newDoc().settings.compat, 15);
  assert.equal(L.D.newDoc({ modern: true }).settings.compat, 15);
});

test('a blank document names 等线 on every built-in style and on the theme', () => {
  const d = L.D.newDoc();
  assert.equal(d.defaults.rPr.font, '等线');
  assert.equal(d.defaults.rPr.fontEA, '等线');
  assert.deepEqual([d.theme.major, d.theme.minor, d.theme.majorEA, d.theme.minorEA], ['等线', '等线', '等线', '等线']);
  for (const s of Object.values(d.styles)) {
    if (s.id === 'DefaultParagraphFont' || s.type === 'numbering') continue;
    assert.equal(s.rPr.font, '等线', s.id);
    assert.equal(s.rPr.fontEA, '等线', s.id);
  }
  assert.equal(d.styles.DefaultParagraphFont.rPr.font, undefined);
  assert.equal(d.styles.Heading1.rPr.b, true);
  assert.equal(d.styles.Hyperlink.rPr.u, 'single');
});

test('a language with its own face uses it on those styles; a modern document keeps its old faces', () => {
  const ja = L.D.applyLocale(L.D.newDoc(), L.locale.get('ja'));
  assert.equal(ja.styles.Heading1.rPr.font, 'Century');
  assert.equal(ja.styles.Heading1.rPr.fontEA, 'ＭＳ 明朝');
  assert.equal(ja.styles.Hyperlink.rPr.font, 'Century');
  assert.equal(ja.styles.Normal.rPr.fontEA, 'ＭＳ 明朝');
  assert.equal(ja.theme.major, 'Century');
  assert.equal(ja.theme.majorEA, 'ＭＳ ゴシック');
  const zh = L.D.applyLocale(L.D.newDoc(), L.locale.get('zh-Hant'));
  assert.equal(zh.styles.Title.rPr.font, 'Times New Roman');
  assert.equal(zh.styles.Title.rPr.fontEA, '新細明體');
  const modern = L.D.newDoc({ modern: true });
  assert.equal(modern.defaults.rPr.font, 'Calibri');
  assert.equal(modern.styles.Normal.rPr.font, undefined);
  assert.equal(modern.styles.Heading1.rPr.font, 'Arial');
  assert.equal(modern.styles.PlainText.rPr.font, 'Courier New');
  assert.equal(modern.theme.minor, 'Calibri');
});
