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
  require('../../public/quire/js/vb-host.js');
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
});
