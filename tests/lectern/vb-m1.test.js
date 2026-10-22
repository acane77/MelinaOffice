const test = require('node:test');
const assert = require('node:assert/strict');
const { boot, fresh, expectGap } = require('./vb-boot.js');

test('M1 Application, Presentations and Presentation are callable from VB', { timeout: 120000 }, async () => {
  await boot();
  fresh();
  const host = await L.vb.open();
  try {
    for (const key of Object.keys(L.dom)) {
      if (!/^(pp|mso)[A-Za-z][A-Za-z0-9]*$/.test(key) || !Number.isInteger(L.dom[key])) continue;
      assert.equal(await host.eval('?' + key), L.dom[key], key);
    }
    assert.equal(await host.eval('?ActivePresentation.Name'), ActivePresentation.Name);
    assert.equal(await host.eval('?Application.ActivePresentation.Name'), ActivePresentation.Name);
    assert.equal(await host.eval('?Presentations.Count'), 1);
    assert.equal(await host.eval('?Presentations(1).Name'), Presentations(1).Name);
    assert.equal(await host.eval('?Presentations.Item(1).Name'), Presentations(1).Name);
    assert.equal(await host.eval('?ActivePresentation.Parent Is Presentations'), -1);
    assert.equal(await host.eval('?ActivePresentation.Application Is Application'), -1);
    assert.equal(await host.eval('?Application.Parent Is Application'), -1);
    assert.equal(typeof await host.eval('?ActiveWindow.Selection'), 'object');
    assert.equal(await host.eval('?ActiveWindow.Presentation.Name'), ActivePresentation.Name);

    const name = ActivePresentation.Name;
    await assert.rejects(() => host.eval('ActivePresentation.Name = "Z"'));
    assert.equal(ActivePresentation.Name, name);

    await host.eval('ActivePresentation.Saved = False');
    assert.equal(ActivePresentation.Saved, false);
    await host.eval('ActivePresentation.Saved = True');
    assert.equal(ActivePresentation.Saved, true);
    await host.eval('ActivePresentation.BuiltInDocumentProperties("Title").Value = "M1"');
    assert.equal(ActivePresentation.BuiltInDocumentProperties('Title').Value, 'M1');
    assert.equal(await host.eval('?ActivePresentation.BuiltInDocumentProperties("Title").Name'), 'Title');
    await host.eval('ActivePresentation.CustomDocumentProperties("Note").Value = "x"');
    assert.equal(ActivePresentation.CustomDocumentProperties('Note').Value, 'x');

    await host.eval('ActivePresentation.PageSetup.SlideWidth = 720');
    assert.equal(ActivePresentation.PageSetup.SlideWidth, 720);
    assert.equal(await host.eval('?ActivePresentation.Undo'), -1);
    assert.equal(ActivePresentation.PageSetup.SlideWidth, 960);
    await host.eval('ActivePresentation.Redo');
    assert.equal(ActivePresentation.PageSetup.SlideWidth, 720);

    await expectGap(host, '?ActivePresentation.Creator', 'G-ENGINE');
    await expectGap(host, '?ActivePresentation.FullName', 'G-HOST');
    await expectGap(host, '?ActivePresentation.Path', 'G-HOST');
    await expectGap(host, 'Presentations.Open "x"', 'G-HOST');
    await expectGap(host, 'ActivePresentation.SaveAs "x"', 'G-HOST');
    await expectGap(host, 'ActivePresentation.PrintOut', 'G-HOST');
    await expectGap(host, '?ActiveWindow.View', 'G-HOST');
    await expectGap(host, 'Presentations.Add True', 'G-HOST');
  } finally {
    host.close();
  }
});
