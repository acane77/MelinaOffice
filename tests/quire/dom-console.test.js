const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/quire/js/dmodel.js');
require('../../public/quire/js/ops.js');
require('../../public/quire/js/dom.js');
require('../../public/quire/js/dom-console.js');
const { D } = L;

function fresh() {
  D.doc = D.newDoc();
  D.resetHistory();
  if (L.ed) L.ed.sel = null;
  ActiveDocument;
}

test('an open bracket, string or comment waits, and a finished one does not', () => {
  assert.equal(L.domConsole.needsMore('function f() {'), true);
  assert.equal(L.domConsole.needsMore('function f() { return 1; }'), false);
  assert.equal(L.domConsole.needsMore('async function f() { return 1; }'), false);
  assert.equal(L.domConsole.needsMore("const s = '{"), true);
  assert.equal(L.domConsole.needsMore('const s = "{"'), false);
  assert.equal(L.domConsole.needsMore('const s = `a\nb`'), false);
  assert.equal(L.domConsole.needsMore('const s = `a ${'), true);
  assert.equal(L.domConsole.needsMore('/* {'), true);
  assert.equal(L.domConsole.needsMore('// {\n1'), false);
  assert.equal(L.domConsole.needsMore('1 + (2'), true);
  assert.equal(L.domConsole.needsMore('1 + (2)'), false);
});

test('a top-level function stays for the next line', async () => {
  const fn = await L.domConsole.evaluate('function greet(n) { return "hi " + n; }');
  assert.equal(typeof fn, 'function');
  assert.equal(await L.domConsole.evaluate('greet("Ada")'), 'hi Ada');
  delete globalThis.greet;
});

test('an async function can be awaited on the next line', async () => {
  await L.domConsole.evaluate('async function later() { return await Promise.resolve(4); }');
  assert.equal(await L.domConsole.evaluate('await later()'), 4);
  delete globalThis.later;
});

test('await of a value works', async () => {
  assert.equal(await L.domConsole.evaluate('await Promise.resolve(7)'), 7);
});

test('a function inside a string or another function is not published', async () => {
  await L.domConsole.evaluate('"function nope() { return 1 }"');
  assert.equal(await L.domConsole.evaluate('typeof nope'), 'undefined');
  await L.domConsole.evaluate('function outer(){ function inner(){ return 3 } return inner(); }');
  assert.equal(await L.domConsole.evaluate('outer()'), 3);
  assert.equal(await L.domConsole.evaluate('typeof inner'), 'undefined');
  assert.equal(await L.domConsole.evaluate('(function named(){ return 5 })()'), 5);
  assert.equal(await L.domConsole.evaluate('typeof named'), 'undefined');
  delete globalThis.outer;
  delete globalThis.nope;
  delete globalThis.inner;
  delete globalThis.named;
});

test('a VB line stays open for a continuation or an unfinished block', () => {
  assert.equal(L.domConsole.vbNeedsMore('x = 1 _'), true);
  assert.equal(L.domConsole.vbNeedsMore('x = 1 _\n + 2'), false);
  assert.equal(L.domConsole.joinVb('x = 1 _\n + 2'), 'x = 1 + 2');
  assert.equal(L.domConsole.vbNeedsMore('Sub Main()\n  x = 1'), true);
  assert.equal(L.domConsole.vbNeedsMore('Sub Main()\n  x = 1\nEnd Sub'), false);
  assert.equal(L.domConsole.vbNeedsMore('Function Twice(n As Long) As Long\n  Twice = n * 2'), true);
  assert.equal(L.domConsole.vbNeedsMore('Function Twice(n As Long) As Long\n  Twice = n * 2\nEnd Function'), false);
  assert.equal(L.domConsole.vbNeedsMore('If a Then'), true);
  assert.equal(L.domConsole.vbNeedsMore('If a Then b = 1'), false);
  assert.equal(L.domConsole.vbNeedsMore('If a Then\n  b = 1\nEnd If'), false);
  assert.equal(L.domConsole.vbNeedsMore('With ActiveDocument\nEnd With'), false);
  assert.equal(L.domConsole.vbNeedsMore('x = 1'), false);
});

test('an empty console script opens on the sample that types at the cursor', () => {
  const sample = L.domConsole.scriptSource('');
  assert.equal(sample, L.domConsole.scriptSource(null));
  assert.equal(sample, L.domConsole.scriptSource(' \n\t'));
  assert.match(sample, /^' Welcome to Visual Basic Playground\n/);
  assert.match(sample, /click Run to run and test it/);
  assert.match(sample, /https:\/\/learn\.microsoft\.com\/en-us\/office\/vba\/api\/overview\/Word\/object-model/);
  assert.match(sample, /' Your code should contain a Main procedure\nSub Main\(\)/);
  assert.match(sample, /Selection\.TypeText "Hello World from Visual Basic"/);
  assert.doesNotMatch(sample, /Documents\.Add/);
  const kept = 'Sub Main()\n  x = 1\nEnd Sub\n';
  assert.equal(L.domConsole.scriptSource(kept), kept);
});

test('a print is shown while the script is still waiting', async () => {
  const host = { output: ['earlier'] };
  const shown = [];
  let release;
  const paused = new Promise((resolve) => { release = resolve; });
  const stop = L.domConsole.followPrints(host, (line) => { shown.push(line); });
  const running = (async () => {
    host.output.push('0.3750879');
    host.onPrint();
    await paused;
    host.output.push('done');
    host.onPrint();
  })();
  assert.deepEqual(shown, ['0.3750879']);
  release();
  await running;
  stop();
  assert.deepEqual(shown, ['0.3750879', 'done']);
  host.output.push('after');
  assert.equal(host.onPrint, undefined);
  assert.deepEqual(shown, ['0.3750879', 'done']);
});

test('the page globals are visible', async () => {
  fresh();
  assert.equal(await L.domConsole.evaluate('ActiveDocument.Paragraphs.Count'), 1);
});
