const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
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

test('an Immediate statement writes the document through the IR', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  const out = buildEngine();
  const code = fs.readFileSync(out, 'utf8');
  assert.equal(path.basename(path.dirname(out)), 'vb6');
  assert.match(fs.readFileSync(path.join(__dirname, '../../.gitignore'), 'utf8'), /\/public\/quire\/vendor\/vb6\//);
  assert.doesNotMatch(code, /\/\* host\.js \*\//);
  assert.doesNotMatch(code, /BrowserForm/);
  vm.runInThisContext(code, { filename: out });
  assert.equal(typeof VB6Engine.VirtualMachine, 'function');
  require('../../public/quire/js/vb-host.js');

  fresh();
  const host = await L.vb.open();
  try {
    assert.equal(host.vm.immediateContext, true);
    assert.equal(host.vm.state, 'idle');
    assert.equal(host.vm.formInstances.size, 0);
    for (const name of ['createForm', 'beginModal', 'msgBox', 'inputBox', 'openFile', 'saveFile', 'loadProject']) {
      assert.equal(host[name], null, name);
    }
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Hello"');
    assert.equal(ActiveDocument.Paragraphs(1).Range.Text, 'Hello\r');
    assert.equal(await host.eval('?ActiveDocument.Paragraphs(1).Range.Text'), 'Hello\r');
    await host.eval('scratch = wdAlignParagraphCenter');
    assert.equal(await host.eval('?scratch'), 1);
  } finally {
    host.close();
  }
  assert.equal(host.vm, null);
  await assert.rejects(() => host.eval('scratch'), /not running/);
});
