// A paragraph whose only content is an inline equation is promoted to display by Word.
// Saving it with a preserved space afterwards keeps the equation inline. A second save must not add another space.
// node --test tests/ecrivaine/equation-inline-save.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const L = {
  xesc: (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])),
  pt2emu: (n) => n,
  fontKey: (s) => String(s || '').toLowerCase(),
  preserve: {
    propertyXML: (_tag, _obj, inner) => inner || '',
    objectXML: () => null,
  },
};
const sandbox = { window: { L }, console, L };
sandbox.window.L = L;
vm.createContext(sandbox);
for (const file of ['dmodel.js', 'docx-write.js', 'render.js']) {
  vm.runInContext(fs.readFileSync(new URL('../../public/ecrivaine/js/' + file, import.meta.url), 'utf8'), sandbox);
}
const { docx: W, R } = sandbox.window.L;
const MATH = '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>';
const DISP = '<m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>';
const SPACE = '<w:t xml:space="preserve"> </w:t>';
const para = (runs) => ({ t: 'p', pPr: {}, rPr: {}, runs });
const math = (xml, display) => ({ t: 'raw', math: true, para: !!display, xml });
const text = (s) => ({ t: 'text', text: s, rPr: {} });
const spaces = (xml) => xml.split(SPACE).length - 1;

test('a lone inline equation is saved with one preserved space, and a second save does not add another', () => {
  const once = W.oneParaXML(para([math(MATH)]), {});
  assert.equal(spaces(once), 1);
  assert.ok(once.indexOf(MATH) < once.indexOf(SPACE));
  const again = W.oneParaXML(para([math(MATH), text(' ')]), {});
  assert.equal(spaces(again), 1);
});

test('a display equation, an equation with text, and an equation number do not gain that space', () => {
  assert.equal(spaces(W.oneParaXML(para([math(DISP, true)]), {})), 0);
  assert.equal(spaces(W.oneParaXML(para([text('This is Inline Formula. '), math(MATH)]), {})), 0);
  assert.equal(spaces(W.oneParaXML(para([math(MATH), text('(1)')]), {})), 0);
});

test('a justified paragraph places a display equation by its m:jc', () => {
  assert.equal(R.isJustified('both'), true);
  assert.equal(R.isJustified('distribute'), true);
  assert.equal(R.isJustified('left'), false);
  assert.equal(R.isJustified('center'), false);
  assert.equal(R.mathAlign('left'), 'left');
  assert.equal(R.mathAlign('right'), 'right');
  assert.equal(R.mathAlign('center'), 'center');
  assert.equal(R.mathAlign('centerGroup'), 'center');
  assert.equal(R.mathAlign(undefined), 'center');
  const numbered = para([math(MATH, true), text('(1)')]);
  const prose = para([text('See '), math(MATH, true)]);
  assert.equal(R.mathNumbered(numbered), true);
  assert.equal(R.mathNumbered(para([math(MATH, true)])), false);
  assert.equal(R.mathNumbered(prose), false);
});

test('the space that keeps Word on the inline path is not drawn as a display equation', () => {
  assert.equal(R.mathSolo(para([math(MATH)])), true);
  assert.equal(R.mathSolo(para([math(MATH), text(' ')])), false);
  assert.equal(R.mathSolo(para([math(MATH), text('(1)')])), true);
});
