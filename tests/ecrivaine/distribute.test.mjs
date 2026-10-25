// Distributed alignment: the gap between fullwidth characters, and a trailing punctuation mark
// that hangs instead of wrapping. node --test tests/ecrivaine/distribute.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const L = {
  xesc: (s) => String(s),
  pt2emu: (n) => n,
  fontKey: (s) => String(s || '').toLowerCase(),
  round: (n, d) => { const f = 10 ** (d || 0); return Math.round(n * f) / f; },
  preserve: { propertyXML: (_tag, _obj, inner) => inner || '', objectXML: () => null },
};
const sandbox = { window: { L }, console, L };
sandbox.window.L = L;
vm.createContext(sandbox);
for (const file of ['dmodel.js', 'render.js']) {
  vm.runInContext(fs.readFileSync(new URL('../../public/ecrivaine/js/' + file, import.meta.url), 'utf8'), sandbox);
}
const { R } = sandbox.window.L;
const em = 18.67; // 14pt

test('five fullwidth characters spread across the cell, seven narrow the colon', () => {
  const cell = 123.6; // 98.4pt column minus 2.85pt padding each side
  const five = R.distributeFit(Array(5).fill(em), cell, { chars: [...'作者姓名：'] });
  assert.ok(five.gap > 7 && five.gap < 8);
  const label = [...'申请学位级别：'];
  const seven = R.distributeFit(Array(7).fill(em), cell, { chars: label });
  assert.ok(seven.advances);
  assert.ok(Math.abs(seven.advances[6] - (em - (7 * em - cell))) < 0.05);
  for (let i = 0; i < 6; i++) assert.ok(Math.abs(seven.advances[i] - em) < 0.05);
  assert.equal(R.compressPunct('：'), 'middle');
  assert.equal(R.compressPunct('别'), '');
  assert.equal(R.distributeFit(Array(8).fill(em), cell, { chars: [...'答辩委员会主席：'] }), null);
});

test('justified text that already fits is not stretched, and compression can be turned off', () => {
  const cell = 123.6;
  assert.equal(R.distributeFit(Array(5).fill(em), cell, { chars: [...'作者姓名：'], mode: 'justify' }), null);
  const over = R.distributeFit(Array(7).fill(em), cell, { chars: [...'申请学位级别：'], mode: 'justify', compress: false });
  assert.equal(over, null);
});

test('a wider cell still distributes seven characters', () => {
  const wide = 146; // about 109.5pt
  const fit = R.distributeFit(Array(7).fill(em), wide);
  assert.ok(fit.gap > 2 && fit.gap < 3);
});

test('balanced spaces expand a Latin space to half an em and leave a CJK space alone', () => {
  assert.equal(R.balanceSBCS({ settings: { compatXML: '<w:balanceSingleByteDoubleByteWidth/>' } }), true);
  assert.equal(R.balanceSBCS({ settings: { compatXML: '' } }), false);
  assert.ok(0.5 - R.spaceEm('Times New Roman') > 0.2);
  assert.equal(R.spaceEm('SimSun'), 0.5);
  // 15pt Times: the space glyph is a quarter em, so letter-spacing supplies the other quarter (5px).
  assert.equal(R.halfEmSpace('Times New Roman', 15), 5);
  assert.equal(R.halfEmSpace('SimSun', 15), 0);
  assert.equal(R.halfEmSpace('DengXian', 15), 0);
});

test('heading numbers keep two ASCII spaces at half an em each', () => {
  const show = (text, cjk) => JSON.parse(JSON.stringify(R.balancedSpaceParts(text, cjk))).map((p) => p.half ? '[' + p.text + ']' : p.text);
  assert.deepEqual(show('1.3  '), ['1.3', '[  ]']);
  assert.deepEqual(show('.1  '), ['.1', '[  ]']);
  assert.deepEqual(show('1.3.1  '), ['1.3.1', '[  ]']);
  assert.deepEqual(show('1.1  '), ['1.1', '[  ]']);
  // a single space between Latin words is not widened; two spaces are
  assert.deepEqual(show('hello world'), ['hello world']);
  assert.deepEqual(show('hello  world'), ['hello', '[  ]', 'world']);
  // a space-only run, and a space at either end of the portion, still qualifies
  assert.deepEqual(show(' '), ['[ ]']);
  assert.deepEqual(show('   '), ['[   ]']);
  assert.deepEqual(show('a '), ['a', '[ ]']);
  // inside a CJK portion every ASCII space qualifies; an ideographic space does not
  assert.deepEqual(show('中 文', true), ['中', '[ ]', '文']);
  assert.deepEqual(show('中　文', false), ['中　文']);
});
