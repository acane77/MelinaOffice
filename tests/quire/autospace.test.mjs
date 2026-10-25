// Space between East Asian text and Latin letters or digits. Word leaves a quarter of the East
// Asian character's size; Quire used to draw them touching.
//   node --test tests/quire/autospace.test.mjs
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
  vm.runInContext(fs.readFileSync(new URL('../../public/quire/js/' + file, import.meta.url), 'utf8'), sandbox);
}
const { R } = sandbox.window.L;

const gaps = (text, sizes, flags) => {
  const chars = [...text];
  return R.autoSpaceGaps(chars, sizes || chars.map(() => 16), flags);
};

test('digits and Latin beside East Asian text take a quarter of the font size', () => {
  const q = 4;
  // 院|9 and 4|年 (digits), 年|A and A|字 (letters): the space follows the first of the two.
  assert.deepEqual(gaps('法院94年A字'), [0, q, 0, q, q, q, 0]);
  assert.deepEqual(gaps('法院94年A字', null, { de: true, dn: false }), [0, 0, 0, 0, q, q, 0]);
  assert.deepEqual(gaps('法院94年A字', null, { de: false, dn: true }), [0, q, 0, q, 0, 0, 0]);
  assert.deepEqual(gaps('法院94年A字', null, { de: false, dn: false }), [0, 0, 0, 0, 0, 0, 0]);
});

test('the space crosses sizes and takes the East Asian size', () => {
  assert.deepEqual(gaps('漢A漢', [20, 10, 20]), [5, 5, 0]);
  assert.equal(R.autoSpaceSide('院', '9', true, true), 'left');
  assert.equal(R.autoSpaceSide('A', 'カ', true, true), 'right');
  assert.equal(R.autoSpaceSide('院', '9', true, false), '');
  assert.equal(R.autoSpaceSide('A', '字', false, true), '');
});

test('spaces, punctuation, Hangul and Latin alone get nothing', () => {
  for (const text of ['例 a', '5 例', '5、', '、8', 'abc 123', '漢한', '한A', '漢%', '（A）', '漢１', '漢字']) {
    assert.ok(gaps(text).every((g) => g === 0), text);
  }
  assert.deepEqual(R.autoSpaceGaps(['漢', null, 'A'], [16, 0, 16]), [0, 0, 0]);
});

test('a character grid the paragraph snaps to turns the space off', () => {
  const flags = (pPr, ctx) => JSON.parse(JSON.stringify(R.autoSpaceFlags(pPr, ctx)));
  assert.deepEqual(flags({}, { charGrid: true }), { de: false, dn: false });
  assert.deepEqual(flags({ noSnap: true }, { charGrid: true }), { de: true, dn: true });
  assert.deepEqual(flags({ autoSpaceDE: false }, {}), { de: false, dn: true });
  assert.deepEqual(flags(null, null), { de: true, dn: true });
});
