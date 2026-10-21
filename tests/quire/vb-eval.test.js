const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/quire/js/dmodel.js');
require('../../public/quire/js/ops.js');
require('../../public/quire/js/dom.js');
require('../../public/quire/js/dom-ir.js');
const { D } = L;

function fresh() {
  D.doc = D.newDoc();
  D.resetHistory();
  if (L.ed) L.ed.sel = null;
  ActiveDocument;
}

test('evalVB6 edits the document on one Immediate session', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  vm.runInThisContext(require('node:fs').readFileSync(buildEngine(), 'utf8'));
  require('../../public/quire/js/vb-host.js');
  fresh();
  try {
    assert.equal(evalVB6, L.evalVB6);
    assert.equal(await evalVB6('ActiveDocument.Paragraphs(1).Range.Text = "Hello"'), undefined);
    assert.equal(ActiveDocument.Paragraphs(1).Range.Text, 'Hello\r');
    assert.equal(await evalVB6('scratch = wdAlignParagraphCenter'), undefined);
    assert.equal(await evalVB6('?scratch'), 1);
    assert.equal(await evalVB6('?ActiveDocument.Paragraphs(1).Range.Text'), 'Hello\r');
    assert.equal(await evalVB6('?ActiveDocument.Name'), ActiveDocument.Name);
    assert.equal(await evalVB6('?wdAlignParagraphCenter'), 1);
    assert.equal(await evalVB6('Selection.TypeText("!")'), undefined);
    assert.ok(String(ActiveDocument.Content.Text).includes('!'));

    const host = await L.vb.session();
    assert.equal(await L.vb.session(), host);
    assert.equal(host.vm.immediateContext, true);
    assert.equal(host.vm.state, 'idle');
    assert.equal(host.vm.formInstances.size, 0);
    assert.equal(host.createForm, null);
    assert.equal(host.loadProject, null);
    assert.equal(await evalVB6('?scratch'), 1);

    await assert.rejects(() => evalVB6('?ActiveDocument.FullName'), (err) => {
      assert.equal(err && err.code, 'G-HOST');
      return true;
    });
    await assert.rejects(() => evalVB6('ActiveDocument.Paragraphs('));

    host.close();
    await evalVB6('n = 1');
    assert.notEqual(await L.vb.session(), host);
    assert.equal(await evalVB6('?scratch'), undefined);
    assert.equal(await evalVB6('?n'), 1);
  } finally {
    const host = await L.vb.session();
    host.close();
  }
});
