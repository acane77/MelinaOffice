const test = require('node:test');
const assert = require('node:assert/strict');
const { boot, fresh } = require('./vb-boot.js');

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
  '  ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text = CStr(n)',
  'End Sub',
].join('\n');

test('runVB6 compiles a module without disturbing Immediate locals', { timeout: 120000 }, async () => {
  await boot();
  fresh();
  try {
    assert.equal(runVB6, L.runVB6);
    assert.equal(await evalVB6('scratch = 7'), undefined);
    const before = ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text;
    assert.equal(await runVB6(program, 'Twice', [21]), 42);
    assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, before);

    assert.equal(await runVB6(program), undefined);
    assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, '10');
    assert.equal(await evalVB6('?scratch'), 7);

    const host = await L.vb.session();
    assert.equal(host.vm.immediateContext, true);
    assert.equal(host.vm.state, 'idle');
    assert.equal(host.vm.formInstances.size, 0);
    assert.equal(host.script, null);
    assert.equal(host.createForm, null);
    assert.equal(host.loadProject, null);

    await assert.rejects(() => runVB6('Sub Other()\nEnd Sub'), (err) => {
      assert.equal(err && err.number, 35);
      return true;
    });
    await assert.rejects(() => runVB6(program, 'Missing'), (err) => {
      assert.equal(err && err.number, 35);
      return true;
    });
    await assert.rejects(() => runVB6('Sub Main()\n  Dim x\n  x = ActivePresentation.FullName\nEnd Sub'), (err) => {
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
      await runVB6('Sub Main()\n  ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "ok"\nEnd Sub');
      await assert.rejects(() => nestedEval, /A VB script is running/);
      await assert.rejects(() => nestedRun, /A VB script is running/);
    } finally {
      ir.exec = exec;
    }
    assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, 'ok');
    assert.equal(await evalVB6('?scratch'), 7);
  } finally {
    const host = await L.vb.session();
    host.close();
  }
});
