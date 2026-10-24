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

const program = [
  'Function Twice(n As Long) As Long',
  '  Twice = n * 2',
  'End Function',
  'Sub Main()',
  '  Dim i As Long',
  '  Dim n As Long',
  '  n = 0',
  '  For i = 1 To 3',
  '    If i > 1 Then n = n + Twice(i)',
  '  Next',
  '  ActiveDocument.Paragraphs(1).Range.Text = CStr(n)',
  'End Sub',
].join('\n');

test('runVB6 compiles a module without disturbing Immediate locals', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  vm.runInThisContext(require('node:fs').readFileSync(buildEngine(), 'utf8'));
  require('../../public/quire/js/vb-host.js');
  fresh();
  try {
    assert.equal(runVB6, L.runVB6);
    assert.equal(await evalVB6('scratch = 7'), undefined);
    const before = ActiveDocument.Paragraphs(1).Range.Text;
    assert.equal(await runVB6(program, 'Twice', [21]), 42);
    assert.equal(ActiveDocument.Paragraphs(1).Range.Text, before);

    assert.equal(await runVB6(program), undefined);
    assert.equal(ActiveDocument.Paragraphs(1).Range.Text, '10\r');
    assert.equal(await evalVB6('?scratch'), 7);

    const host = await L.vb.session();
    assert.equal(host.vm.immediateContext, true);
    assert.equal(host.vm.state, 'idle');
    assert.equal(host.vm.formInstances.size, 0);
    assert.equal(host.script, null);
    assert.equal(host.createForm, null);
    assert.equal(typeof host.loadProject, 'function');
    assert.equal(await evalVB6('?scratch'), 7);

    await assert.rejects(() => runVB6('Sub Other()\nEnd Sub'), (err) => {
      assert.equal(err && err.number, 35);
      return true;
    });
    await assert.rejects(() => runVB6(program, 'Missing'), (err) => {
      assert.equal(err && err.number, 35);
      return true;
    });
    await assert.rejects(() => runVB6('Sub Main()\n  Dim x\n  x = ActiveDocument.FullName\nEnd Sub'), (err) => {
      assert.equal(err && err.code, 'G-HOST');
      return true;
    });
    await assert.rejects(() => runVB6('Sub Main(\nEnd Sub'));

    const ir = host.ir;
    const exec = ir.exec;
    let nestedEval;
    let nestedRun;
    ir.exec = function (req) {
      if (!nestedEval) {
        nestedEval = host.eval('x = 1');
        nestedRun = host.run('Sub Main()\nEnd Sub');
      }
      return exec.call(ir, req);
    };
    try {
      await runVB6('Sub Main()\n  ActiveDocument.Paragraphs(1).Range.Text = "ok"\nEnd Sub');
      await assert.rejects(() => nestedEval, /A VB script is running/);
      await assert.rejects(() => nestedRun, /A VB script is running/);
    } finally {
      ir.exec = exec;
    }
    assert.equal(ActiveDocument.Paragraphs(1).Range.Text, 'ok\r');
    assert.equal(await evalVB6('?scratch'), 7);

    const live = [];
    host.onPrint = () => { live.push(host.output[host.output.length - 1]); };
    let atBox = null;
    host.msgBox = () => {
      atBox = live.slice();
      return Promise.resolve(7);
    };
    try {
      await runVB6('Sub Main()\n  Debug.Print "one"\n  MsgBox "again?"\n  Debug.Print "two"\nEnd Sub');
      assert.deepEqual(atBox, ['one']);
      assert.deepEqual(live, ['one', 'two']);
    } finally {
      host.onPrint = null;
      host.msgBox = null;
    }
  } finally {
    const host = await L.vb.session();
    host.close();
  }
});
