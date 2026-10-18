// East Asian fonts (docs/i18n.md, Fonts): the names Japanese, Korean and Chinese Office write are the fonts every table
// knows by their English names, a run draws its East Asian characters from its East Asian font, and Ledger sizes rows
// for a Japanese or Korean Excel's default font:  node --test tests/fonts.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import fs from 'node:fs/promises';
import vm from 'node:vm';

const require = createRequire(import.meta.url);
globalThis.window = globalThis;
require('../public/common/core.js');
require('../public/ledger/js/model.js');
const L = globalThis.L;

test('localized names are the English fonts', () => {
  assert.equal(L.fontKey('ＭＳ 明朝'), 'ms mincho');
  assert.equal(L.fontKey('맑은 고딕'), 'malgun gothic');
  assert.equal(L.fontKey('宋体'), 'simsun');
  assert.equal(L.fontKey('Times New Roman'), 'times new roman');
  assert.ok(L.fontIsEA('游ゴシック') && L.fontIsEA('Batang') && !L.fontIsEA('Century'));
});

test('Simplified Chinese open-source font references retain their families and East Asian fallbacks', async () => {
  const source = process.env.BASE ? await (await fetch(process.env.BASE + '/common/core.js')).text()
    : await fs.readFile(new URL('../public/common/core.js', import.meta.url), 'utf8');
  const refs = [
    ['文泉驿微米黑', 'WenQuanYi Micro Hei', 'sans-serif'],
    ['思源黑体', 'Source Han Sans SC', 'sans-serif'],
    ['思源宋体', 'Source Han Serif SC', 'serif'],
    ['Noto Sans CJK SC', 'Noto Sans CJK SC', 'sans-serif'], ['Noto Sans SC', 'Noto Sans SC', 'sans-serif'],
    ['Noto Serif CJK SC', 'Noto Serif CJK SC', 'serif'], ['Noto Serif SC', 'Noto Serif SC', 'serif'],
  ];
  for (const lang of ['zh-CN', 'en', 'zh-Hant']) {
    const c = { document: { documentElement: { getAttribute: () => lang } } }; c.window = c;
    vm.runInNewContext(source, c);
    for (const [name, family, generic] of refs) {
      assert.ok(c.L.FONT_LIST.includes(name), name + ': ' + lang);
      assert.equal(c.L.fontKey(name), family.toLowerCase());
      assert.ok(c.L.fontIsEA(name) && c.L.fontIsEA(family), name);
      const stack = c.L.fontStack('Times New Roman', name);
      assert.ok(stack.includes('"' + name + '"') && stack.includes('"' + family + '"'), stack);
      assert.ok(stack.endsWith(',' + generic), stack);
      const serif = generic === 'serif', counterpart = serif ? '"Noto Serif CJK SC"' : '"Noto Sans CJK SC"';
      if (!family.startsWith('WenQuanYi')) assert.ok(stack.indexOf(counterpart) < stack.indexOf(serif ? '"SimSun"' : '"Microsoft YaHei"'), stack);      assert.ok(c.L.fontStack(name).startsWith('"' + name + '",'), 'the chosen family is tried first');
    }
    // Resolve an unavailable family with a controlled set of installed fonts, independently of this machine.
    c.L.fontInstalled = (f) => ['Noto Sans CJK SC', 'Noto Serif CJK SC', 'Microsoft YaHei', 'SimSun'].includes(f);
    assert.equal(c.L.fontUsed('思源黑体'), 'Noto Sans CJK SC');
    assert.equal(c.L.fontUsed('思源宋体'), 'Noto Serif CJK SC');
    assert.equal(c.L.fontUsed('文泉驿微米黑'), 'Microsoft YaHei');
    c.L.fontInstalled = (f) => ['WenQuanYi Micro Hei', 'Source Han Sans SC', 'Source Han Serif SC'].includes(f);
    assert.equal(c.L.fontUsed('文泉驿微米黑'), 'WenQuanYi Micro Hei');
    assert.equal(c.L.fontUsed('思源宋体'), 'Source Han Serif SC');
    assert.equal(c.L.fontUsed('宋体'), 'Source Han Serif SC');
    c.L.fontInstalled = () => false;
    assert.equal(c.L.fontUsed('思源黑体'), 'sans-serif');
    assert.equal(c.L.fontUsed('思源宋体'), 'serif');
  }
});

test('a run\'s letters come from its Latin font, its East Asian characters from its East Asian font', () => {
  const stack = L.fontStack('Century', 'ＭＳ 明朝');
  assert.ok(stack.startsWith('"Century",'), stack);
  assert.ok(stack.indexOf('"ＭＳ 明朝"') > stack.indexOf('"Century"'), 'the East Asian font after the Latin one');
  assert.ok(stack.includes('"MS Mincho"') && stack.endsWith(',serif'), stack);
  assert.equal(stack.match(/,serif/g).length, 1, 'one generic family, at the end');
  assert.equal(L.fontStack('Arial', 'Arial'), L.fontStack('Arial'), 'the same font once');
  assert.equal(L.fontStack('Arial', 'Wingdings'), L.fontStack('Arial'), 'not a font for East Asian text');
});

test('East Asian text is told from Latin text', () => {
  for (const s of ['日本語', 'ひらがな', 'カタカナ', '한국어', '中文', '（全角）']) assert.ok(L.hasEAText(s), s);
  for (const s of ['English', 'Français', 'Größe', '', null]) assert.ok(!L.hasEAText(s), String(s));
});

test('(Use Asian text font) is one East Asian family in both slots', () => {
  assert.equal(L.fontUsesEA('宋体', 'SimSun'), true);
  assert.equal(L.fontUsesEA('SimHei', '黑体'), true);
  assert.equal(L.fontUsesEA('Times New Roman', '宋体'), false);
  assert.equal(L.fontUsesEA('Times New Roman', 'Times New Roman'), false);
  assert.equal(L.fontUsesEA('宋体', ''), false);
  assert.equal(L.fontUsesEA('', '宋体'), false);
});

test('the toolbar Font box writes the slot the family belongs to', () => {
  assert.deepEqual(L.fontPick('Times New Roman', { font: '宋体' }), { font: 'Times New Roman', fontEA: '宋体' });
  assert.deepEqual(L.fontPick('Times New Roman', { font: '宋体', fontEA: '宋体' }), { font: 'Times New Roman', fontEA: '宋体' });
  assert.deepEqual(L.fontPick('Arial', { font: 'Times New Roman', fontEA: '宋体' }), { font: 'Arial' });
  assert.deepEqual(L.fontPick('黑体', { font: 'Times New Roman', fontEA: '宋体' }), { font: '黑体', fontEA: '黑体' });
  assert.deepEqual(L.fontPick('SimSun', { font: 'Arial', fontEA: '微软雅黑' }), { font: 'SimSun', fontEA: 'SimSun' });
  assert.deepEqual(L.fontPick('Times New Roman', { font: '黑体', fontEA: '黑体' }), { font: 'Times New Roman', fontEA: '黑体' });
  assert.equal(L.fontIsEA('华文中宋'), true);
  assert.equal(L.fontIsEA('Calibri'), false);
  assert.equal(L.fontIsEA('Century Gothic'), false);
});

test('a Font box completes to the closest family, recently used first', () => {
  const names = ['Arial', 'Tahoma', 'Times New Roman', 'Trebuchet MS', '宋体'];
  assert.equal(L.completeMatch(names, 'T'), 'Tahoma');
  assert.equal(L.completeMatch(names, 'ti'), 'Times New Roman');
  assert.equal(L.completeMatch(names, '宋'), '宋体');
  assert.equal(L.completeMatch(names, 'Z'), '');
  const orig = L.store.get;
  L.store.get = (k, d) => (k === 'recentFonts' ? ['Times New Roman', '宋体'] : orig.call(L.store, k, d));
  try {
    assert.equal(L.completeMatch(names, 'T', L.fontComplete), 'Times New Roman');
    assert.equal(L.completeMatch(names, 'A', L.fontComplete), 'Arial');
  } finally { L.store.get = orig; }
});

test('a cell keeps an East Asian face when a Latin font is chosen', () => {
  const song = { name: '宋体', sz: 11 };
  assert.deepEqual(L.scriptSlots('宋体', null), { font: '宋体', fontEA: '宋体' });
  assert.deepEqual(L.scriptSlots('Arial', null), { font: 'Arial', fontEA: '' });
  const split = L.scriptFontRuns('测试 Test', null, song, L.fontPick('Times New Roman', { font: '宋体', fontEA: '宋体' }));
  assert.equal(split.length, 2);
  assert.equal(split[0].t, '测试');
  assert.equal(split[0].font.name, '宋体');
  assert.equal(split[1].font.name, 'Times New Roman');
  assert.equal(L.scriptFontRuns('测试 Test', null, song, L.fontPick('黑体', { font: '宋体', fontEA: '宋体' })), null);
  const only = L.scriptFontRuns('中文', null, song, { font: 'Arial', fontEA: '宋体' });
  assert.equal(only.length, 1);
  assert.equal(only[0].font.name, '宋体');
  const again = L.scriptFontRuns('测试 Test', split, { name: 'Times New Roman', sz: 11 }, L.fontPick('Arial', L.scriptSlots('Times New Roman', split)));
  assert.equal(again[0].font.name, '宋体');
  assert.equal(again[1].font.name, 'Arial');
});

test('a run splits into East Asian characters and the rest', () => {
  assert.deepEqual(L.splitEA('测试 Test'), [{ ea: true, text: '测试' }, { ea: false, text: ' Test' }]);
  assert.deepEqual(L.splitEA('A中B'), [{ ea: false, text: 'A' }, { ea: true, text: '中' }, { ea: false, text: 'B' }]);
  assert.deepEqual(L.splitEA('Hello'), [{ ea: false, text: 'Hello' }]);
  assert.deepEqual(L.splitEA('中文'), [{ ea: true, text: '中文' }]);
  assert.deepEqual(L.splitEA(''), []);
});

test('a Chinese document draws the horizontal ellipsis as an East Asian character', () => {
  assert.equal(L.chineseLang('zh-CN'), true);
  assert.equal(L.chineseLang('zh-Hant'), true);
  assert.equal(L.chineseLang('zh-TW'), true);
  assert.equal(L.chineseLang('zh'), true);
  assert.equal(L.chineseLang('en-US'), false);
  assert.equal(L.chineseLang('ja-JP'), false);
  assert.equal(L.chineseLang(''), false);
  assert.equal(L.hasEAText('……'), false);
  assert.equal(L.hasEAText('……', true), true);
  assert.equal(L.hasEAText('...'), false);
  assert.equal(L.hasEAText('...', true), false);
  assert.deepEqual(L.splitEA('（a）……等'), [
    { ea: true, text: '（' }, { ea: false, text: 'a' }, { ea: true, text: '）' }, { ea: false, text: '……' }, { ea: true, text: '等' },
  ]);
  assert.deepEqual(L.splitEA('（a）……等', true), [
    { ea: true, text: '（' }, { ea: false, text: 'a' }, { ea: true, text: '）……等' },
  ]);
  assert.equal(L.chineseDoc({ settings: { themeLang: { eastAsia: 'zh-CN' } } }), true);
  assert.equal(L.chineseDoc({ defaults: { rPr: { langEA: 'zh-TW' } } }), true);
  assert.equal(L.chineseDoc({ lang: 'zh-CN' }), true);
  assert.equal(L.chineseDoc({ settings: { themeLang: { eastAsia: 'ja-JP' } }, lang: 'en-US' }), false);
  assert.equal(L.chineseDoc(null), false);
});

test('Ledger: the rows of a Japanese or Korean Excel\'s default font', () => {
  assert.equal(L.model.fontRowPt({ name: 'ＭＳ Ｐゴシック', sz: 11 }), 13.5);
  assert.equal(L.model.fontRowPt({ name: '游ゴシック', sz: 11 }), 18.75);
  assert.equal(L.model.fontRowPt({ name: '맑은 고딕', sz: 11 }), 16.5);
  assert.equal(L.model.fontRowPt({ name: 'Calibri', sz: 11 }), 15, 'Latin fonts as before');
  assert.equal(L.model.fontRowPt({ name: 'ＭＳ Ｐゴシック;MS PGothic', sz: 11 }), 13.5);
});

test('a font named as several families is drawn with the first one available', () => {
  assert.equal(L.fontFace('Times New Roman'), 'Times New Roman');
  assert.equal(L.fontFace('宋体;SimSun'), '宋体');
  assert.equal(L.fontFace('黑体;SimHei'), '黑体');
  assert.equal(L.fontFace('楷体_GB2312;楷体'), '楷体');
  assert.equal(L.fontFace('  宋体 ; SimSun  '), '宋体');
  assert.equal(L.fontFace('NotARealFont;Arial'), 'Arial');
  assert.equal(L.fontFace('NotARealFont;AlsoMissing'), 'NotARealFont');
  assert.equal(L.fontFace('Arial, Helvetica'), 'Arial');
  assert.equal(L.fontKey('宋体;SimSun'), 'simsun');
  assert.equal(L.fontKey('楷体_GB2312;楷体'), 'kaiti');
  assert.equal(L.fontUsesEA('宋体;SimSun', 'SimSun'), true);
  const song = L.fontStack('宋体;SimSun');
  assert.ok(song.startsWith('"宋体",'), song);
  assert.equal(song.includes(';'), false, song);
  assert.ok(song.includes('"SimSun"'), song);
  assert.ok(L.fontStack('楷体_GB2312;楷体').startsWith('"楷体",'));
  assert.ok(L.fontStack('Times New Roman', '黑体;SimHei').includes('"黑体"'));
  assert.equal(L.fontStack('Arial'), L.fontStack('Arial'), 'one name is unchanged');
  const orig = L.fontInstalled;
  L.fontInstalled = (f) => f === 'SimSun' || f === '楷体_GB2312';
  try {
    assert.equal(L.fontFace('Missing Face;SimSun'), 'SimSun');
    assert.equal(L.fontFace('楷体_GB2312;楷体'), '楷体_GB2312');
    assert.equal(L.fontFace('宋体;SimSun'), 'SimSun');
    assert.ok(L.fontStack('宋体;SimSun').startsWith('"SimSun",'));
  } finally {
    L.fontInstalled = orig;
  }
});
