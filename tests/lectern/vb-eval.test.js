const test = require('node:test');
const assert = require('node:assert/strict');
const { boot, fresh } = require('./vb-boot.js');

test('evalVB6 edits the presentation on one Immediate session', { timeout: 120000 }, async () => {
  await boot();
  fresh();
  try {
    assert.equal(evalVB6, L.evalVB6);
    assert.equal(await evalVB6('ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "Hello"'), undefined);
    assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, 'Hello');
    assert.equal(await evalVB6('scratch = ppAlignCenter'), undefined);
    assert.equal(await evalVB6('?scratch'), 2);
    assert.equal(await evalVB6('?ActivePresentation.Name'), ActivePresentation.Name);
    assert.equal(await evalVB6('?ppLayoutBlank'), 12);

    const host = await L.vb.session();
    assert.equal(await L.vb.session(), host);
    assert.equal(host.vm.immediateContext, true);
    assert.equal(host.vm.state, 'idle');
    assert.equal(host.vm.formInstances.size, 0);
    assert.equal(host.createForm, null);
    assert.equal(host.loadProject, null);
    assert.equal(await evalVB6('?scratch'), 2);

    await assert.rejects(() => evalVB6('?ActivePresentation.FullName'), (err) => {
      assert.equal(err && err.code, 'G-HOST');
      return true;
    });
    await assert.rejects(() => evalVB6('ActivePresentation.Slides('));

    host.close();
    await evalVB6('n = 1');
    assert.notEqual(await L.vb.session(), host);
    assert.equal(await evalVB6('?scratch'), undefined);
    assert.equal(await evalVB6('?n'), 1);
  } finally {
    const host = await L.vb.session();
    host.close();
  }
});
