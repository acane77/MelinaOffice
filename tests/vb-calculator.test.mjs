/* The calculator sample on the form runtime. Node only: no browser.
   node --test tests/vb-calculator.test.mjs */
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const sample = path.join(root, 'examples/vb/calculator');

async function bootEcrivaine() {
  globalThis.window = globalThis;
  require('../public/common/host-audit.js');
  require('../public/common/core.js');
  require('../public/ecrivaine/js/dmodel.js');
  require('../public/ecrivaine/js/ops.js');
  require('../public/ecrivaine/js/dom.js');
  require('../public/ecrivaine/js/dom-ir.js');
  const { buildEngine } = await import('../tools/vb6/build-engine.mjs');
  const out = buildEngine();
  vm.runInThisContext(fs.readFileSync(out, 'utf8'), { filename: out });
  require('../public/common/vb-forms.js');
  require('../public/ecrivaine/js/vb-host.js');
  const { D } = globalThis.L;
  D.doc = D.newDoc();
  D.resetHistory();
}

async function press(form, name) {
  const control = form.formObject.controlMap.get(name);
  globalThis.VO.vbForms.element(control).click();
  await control._lastEvent;
}

function display(form) {
  return form.formObject.controlMap.get('display').Text;
}

test('the calculator sample adds, subtracts, multiplies and divides', async () => {
  await bootEcrivaine();
  const L = globalThis.L;
  L.ui = L.ui || {};
  const host = await L.vb.open();
  try {
    await host.loadProject({
      files: [
        { name: 'Calculator.frm', text: fs.readFileSync(path.join(sample, 'Calculator.frm'), 'utf8') },
        { name: 'Calculator.vbp', text: fs.readFileSync(path.join(sample, 'Calculator.vbp'), 'utf8') },
      ],
    });
    const form = host.formMachine.instances.get('calculator');
    assert.equal(display(form), '0');
    for (const name of ['cmd1', 'cmd2', 'cmdadd', 'cmd3', 'cmd4', 'cmdeq']) await press(form, name);
    assert.equal(display(form), '46');
    await press(form, 'cmdclear');
    for (const name of ['cmd8', 'cmdsub', 'cmd3', 'cmdeq']) await press(form, name);
    assert.equal(display(form), '5');
    await press(form, 'cmdclear');
    for (const name of ['cmd6', 'cmdmul', 'cmd7', 'cmdeq']) await press(form, name);
    assert.equal(display(form), '42');
    await press(form, 'cmdclear');
    for (const name of ['cmd1', 'cmddot', 'cmd5', 'cmdadd', 'cmd2', 'cmdeq']) await press(form, name);
    assert.equal(display(form), '3.5');
    await press(form, 'cmdclear');
    for (const name of ['cmd8', 'cmddiv', 'cmd2', 'cmdeq']) await press(form, name);
    assert.equal(display(form), '4');
    await press(form, 'cmdclear');
    for (const name of ['cmd1', 'cmddiv', 'cmd0', 'cmdeq']) await press(form, name);
    assert.equal(display(form), 'Error');
    await press(form, 'cmdinsert');
    await press(form, 'cmdclear');
    for (const name of ['cmd1', 'cmd2', 'cmdadd', 'cmd3', 'cmd4', 'cmdinsert']) await press(form, name);
    assert.equal(display(form), '46');
    const written = String(globalThis.ActiveDocument.Content.Text);
    assert.ok(written.includes('46'), written);
    assert.equal(written.includes('Error'), false);
  } finally {
    host.close();
  }
});
