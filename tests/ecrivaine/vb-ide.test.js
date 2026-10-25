const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/ecrivaine/js/dmodel.js');
require('../../public/ecrivaine/js/ops.js');
require('../../public/ecrivaine/js/dom.js');
require('../../public/ecrivaine/js/dom-ir.js');
const { D } = L;

function fresh() {
  D.doc = D.newDoc();
  D.resetHistory();
  if (L.ed) L.ed.sel = null;
  ActiveDocument;
}

function pausesOf(machine) {
  const queue = [];
  let wake = null;
  machine.on('pause', (e) => {
    queue.push(e);
    if (wake) { const go = wake; wake = null; go(); }
  });
  return function next() {
    if (queue.length) return Promise.resolve(queue.shift());
    return new Promise((resolve) => { wake = () => resolve(queue.shift()); });
  };
}

function localValue(machine, name) {
  const row = machine.debugLocals({ includeFields: false }).find((entry) => String(entry.name).toLowerCase() === name);
  assert.ok(row, name);
  return row.value;
}

const source = [
  'Function Twice(n As Long) As Long',
  '  Twice = n * 2',
  'End Function',
  'Sub Main()',
  '  Dim x As Long',
  '  x = 1',
  '  Debug.Print "hi"',
  '  x = Twice(x)',
  '  ActiveDocument.Paragraphs(1).Range.Text = CStr(x)',
  'End Sub',
].join('\n');

test('the Visual Basic IDE VM pauses, steps, evaluates, and can run past breakpoints', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  vm.runInThisContext(require('node:fs').readFileSync(buildEngine(), 'utf8'));
  require('../../public/ecrivaine/js/vb-host.js');
  fresh();
  assert.equal(await evalVB6('scratch = 7'), undefined);
  const host = await L.vb.session();

  await assert.rejects(() => host.debugBegin('Sub Main(\nEnd Sub'), (err) => err && err.number === 1002);
  assert.equal(host.script || null, null);
  assert.equal(await evalVB6('?scratch'), 7);
  const printed = host.output.length;

  const debug = await host.debugBegin(source);
  assert.equal(debug.moduleName, 'Module1');
  const prints = [];
  debug.vm.on('output', (e) => prints.push(e.text));
  debug.vm.setBreakpoint('Module1', 8);
  const nextPause = pausesOf(debug.vm);
  const running = debug.vm.start();
  const paused = await nextPause();
  assert.equal(paused.instruction.line, 8);
  assert.equal(String(paused.instruction.source).toLowerCase(), 'module1');
  assert.equal(localValue(debug.vm, 'x'), '1');
  assert.deepEqual(prints, ['hi']);
  assert.equal(await debug.vm.immediate('?x'), 1);
  assert.deepEqual(prints, ['hi', '1']);
  assert.equal(host.output.length, printed);
  await assert.rejects(() => evalVB6('?scratch'), (err) => /A VB script is running/.test(err && err.message));

  debug.vm.resume('into');
  const inside = await nextPause();
  assert.equal(inside.instruction.line, 2);
  assert.equal(String(inside.instruction.procedure).toLowerCase(), 'twice');
  const stack = debug.vm.debugStack();
  assert.ok(stack.some((frame) => String(frame.module).toLowerCase() === 'module1' && frame.line === 8 && String(frame.procedure).toLowerCase() === 'main'));
  assert.ok(stack.some((frame) => frame.line === 2 && String(frame.procedure).toLowerCase() === 'twice'));

  debug.vm.resume('continue');
  await running;
  assert.equal(ActiveDocument.Paragraphs(1).Range.Text, '2\r');
  host.debugEnd();
  assert.equal(host.script, null);
  assert.equal(host.vm.immediateContext, true);
  assert.equal(await evalVB6('?scratch'), 7);

  ActiveDocument.Paragraphs(1).Range.Text = '0';
  const straight = await host.debugBegin(source);
  let sawPause = false;
  straight.vm.on('pause', () => { sawPause = true; straight.vm.resume('continue'); });
  await straight.vm.start();
  assert.equal(sawPause, false);
  assert.equal(ActiveDocument.Paragraphs(1).Range.Text, '2\r');
  host.debugEnd();
  assert.equal(await evalVB6('?scratch'), 7);

  ActiveDocument.Paragraphs(1).Range.Text = '0';
  const held = ActiveDocument.Paragraphs(1).Range.Text;
  const stopped = await host.debugBegin(source);
  stopped.vm.setBreakpoint('Module1', 8);
  const nextStop = pausesOf(stopped.vm);
  const pending = stopped.vm.start();
  await nextStop();
  stopped.vm.stop();
  await pending;
  host.debugEnd();
  assert.equal(ActiveDocument.Paragraphs(1).Range.Text, held);
  assert.equal(host.script, null);
  assert.equal(await evalVB6('?scratch'), 7);

  const watched = [
    'Sub Main()',
    '  Dim x As Long',
    '  x = 1',
    '  x = 2',
    'End Sub',
  ].join('\n');
  const skipped = await host.debugBegin(watched);
  let sawSkip = false;
  skipped.vm.on('pause', () => { sawSkip = true; skipped.vm.resume('continue'); });
  skipped.vm.setBreakpoint('Module1', 4, 'x = 0');
  await skipped.vm.start();
  assert.equal(sawSkip, false);
  host.debugEnd();

  const conditional = await host.debugBegin(watched);
  conditional.vm.setBreakpoint('Module1', 4, 'x = 1');
  const nextCondition = pausesOf(conditional.vm);
  const conditionRun = conditional.vm.start();
  const conditionPause = await nextCondition();
  assert.equal(conditionPause.reason, 'breakpoint');
  assert.equal(conditionPause.instruction.line, 4);
  assert.equal(localValue(conditional.vm, 'x'), '1');
  assert.equal(VB6Engine.describe(await conditional.vm.immediate('?x')), '1');
  conditional.vm.resume('continue');
  await conditionRun;
  host.debugEnd();

  const broken = await host.debugBegin(watched);
  broken.vm.setBreakpoint('Module1', 4, '1 / 0');
  const nextError = pausesOf(broken.vm);
  const errorRun = broken.vm.start();
  const failed = await nextError();
  assert.equal(failed.reason, 'breakpoint-condition');
  assert.match(String(failed.conditionError || ''), /./);
  broken.vm.resume('continue');
  await errorRun;
  host.debugEnd();
  assert.equal(await evalVB6('?scratch'), 7);

  const bad = 'Sub Main(\nEnd Sub';
  const good = 'Sub Main()\nEnd Sub\n';
  for (const compiled of [host.compileIde(bad), host.compileScript(bad), host.compileDraft([{ name: 'Hello', source: bad }])]) {
    assert.equal(compiled.valid, false);
    assert.ok(compiled.diagnostics.length);
    assert.ok(compiled.diagnostics[0].line >= 1);
    assert.ok(compiled.diagnostics[0].message);
    assert.equal(host.script || null, null);
  }
  assert.equal(host.compileIde(good).valid, true);
  assert.equal(host.compileScript(good).valid, true);
  assert.equal(host.script || null, null);
  assert.equal(await evalVB6('?scratch'), 7);
});
