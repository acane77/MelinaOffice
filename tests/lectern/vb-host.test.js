const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { boot, fresh } = require('./vb-boot.js');

test('an Immediate statement writes the presentation through the IR', { timeout: 120000 }, async () => {
  const out = await boot();
  const code = fs.readFileSync(out, 'utf8');
  assert.equal(path.basename(path.dirname(out)), 'vb6');
  assert.equal(path.basename(path.dirname(path.dirname(path.dirname(out)))), 'quire');
  const lectern = path.join(__dirname, '../../public/lectern/vendor/vb6/vb6-engine.js');
  assert.equal(fs.readFileSync(lectern, 'utf8'), code);
  assert.match(fs.readFileSync(path.join(__dirname, '../../.gitignore'), 'utf8'), /\/public\/lectern\/vendor\/vb6\//);
  assert.doesNotMatch(code, /\/\* host\.js \*\//);
  assert.doesNotMatch(code, /BrowserForm/);
  assert.equal(typeof VB6Engine.VirtualMachine, 'function');

  fresh();
  const host = await L.vb.open();
  try {
    assert.equal(host.vm.immediateContext, true);
    assert.equal(host.vm.state, 'idle');
    assert.equal(host.vm.formInstances.size, 0);
    for (const name of ['createForm', 'beginModal', 'msgBox', 'inputBox', 'openFile', 'saveFile', 'loadProject']) {
      assert.equal(host[name], null, name);
    }
    await host.eval('ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "Hello"');
    assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, 'Hello');
    assert.equal(await host.eval('?ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text'), 'Hello');
    await host.eval('scratch = ppAlignCenter');
    assert.equal(await host.eval('?scratch'), 2);
  } finally {
    host.close();
  }
  assert.equal(host.vm, null);
  await assert.rejects(() => host.eval('scratch'), /not running/);
});
