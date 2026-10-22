const test = require('node:test');
const assert = require('node:assert/strict');
const { boot, fresh } = require('./vb-boot.js');

const title = 'ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text';
const a = ['Public Sub A()', '  B', 'End Sub'].join('\n');
const b = ['Public Sub B()', '  ' + title + ' = "called"', 'End Sub'].join('\n');

test('presentation macros call each other without replacing Immediate or the saved project', { timeout: 120000 }, async () => {
  await boot();
  if (!L.t) L.t = (s, ...args) => s.replace(/\{(\d+)\}/g, (m, i) => (args[i] != null ? String(args[i]) : m));
  require('../../public/lectern/js/macros.js');
  fresh();
  assert.equal(await evalVB6('scratch = 7'), undefined);
  L.pres.macros = [
    { name: 'A', source: a, created: '2026-10-21', author: 'Miyuki' },
    { name: 'B', source: b, created: '2026-10-21', author: 'Miyuki' },
  ];
  await L.macros.sync();
  await L.macros.run('A');
  assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, 'called');
  assert.equal(await evalVB6('?scratch'), 7);

  const host = L.vb.current();
  const machine = host.macroMachine;
  await L.macros.runDraft('A', 'Public Sub A()\n  ' + title + ' = "draft"\nEnd Sub\n');
  assert.equal(host.macroMachine, machine);
  assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, 'draft');
  await L.vb.runSaved('A');
  assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, 'called');
  assert.equal(await evalVB6('?scratch'), 7);
  await assert.rejects(() => L.vb.runSaved('Missing'), (err) => err && err.number === 35);

  L.pres.macros = [{ name: 'Count', source: ['Public n As Long', 'Public Sub Count()', '  n = n + 1', '  ' + title + ' = CStr(n)', 'End Sub'].join('\n'), created: '', author: '' }];
  await L.macros.sync();
  await L.macros.run('Count');
  assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, '1');
  await L.macros.run('Count');
  assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, '2');
});
