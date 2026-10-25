const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/ecrivaine/js/dmodel.js');
require('../../public/ecrivaine/js/ops.js');
require('../../public/ecrivaine/js/dom.js');
require('../../public/ecrivaine/js/dom-ir.js');
if (!L.t) L.t = (s, ...args) => s.replace(/\{(\d+)\}/g, (m, i) => (args[i] != null ? String(args[i]) : m));
require('../../public/ecrivaine/js/macros.js');
const { D } = L;

function fresh() {
  D.doc = D.newDoc();
  D.resetHistory();
  if (L.ed) L.ed.sel = null;
  ActiveDocument;
}

const a = ['Public Sub A()', '  B', 'End Sub'].join('\n');
const b = ['Public Sub B()', '  ActiveDocument.Paragraphs(1).Range.Text = "called"', 'End Sub'].join('\n');

test('document macros call each other without replacing Immediate or the saved project', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  vmRun(buildEngine());
  require('../../public/ecrivaine/js/vb-host.js');
  fresh();
  assert.equal(await evalVB6('scratch = 7'), undefined);
  D.doc.macros = [
    { name: 'A', source: a, created: '2026-10-21', author: 'Miyuki' },
    { name: 'B', source: b, created: '2026-10-21', author: 'Miyuki' },
  ];
  await L.macros.sync();
  await L.macros.run('A');
  assert.equal(ActiveDocument.Paragraphs(1).Range.Text, 'called\r');
  assert.equal(await evalVB6('?scratch'), 7);

  const host = L.vb.current();
  const machine = host.macroMachine;
  await L.macros.runDraft('A', 'Public Sub A()\n  ActiveDocument.Paragraphs(1).Range.Text = "draft"\nEnd Sub\n');
  assert.equal(host.macroMachine, machine);
  assert.equal(ActiveDocument.Paragraphs(1).Range.Text, 'draft\r');
  await L.vb.runSaved('A');
  assert.equal(ActiveDocument.Paragraphs(1).Range.Text, 'called\r');
  assert.equal(await evalVB6('?scratch'), 7);
  await assert.rejects(() => L.vb.runSaved('Missing'), (err) => err && err.number === 35);

  D.doc.macros = [{ name: 'Count', source: ['Public n As Long', 'Public Sub Count()', '  n = n + 1', '  ActiveDocument.Paragraphs(1).Range.Text = CStr(n)', 'End Sub'].join('\n'), created: '', author: '' }];
  await L.macros.sync();
  await L.macros.run('Count');
  assert.equal(ActiveDocument.Paragraphs(1).Range.Text, '1\r');
  await L.macros.run('Count');
  assert.equal(ActiveDocument.Paragraphs(1).Range.Text, '2\r');
});

function vmRun(file) {
  require('node:vm').runInThisContext(require('node:fs').readFileSync(file, 'utf8'));
}
