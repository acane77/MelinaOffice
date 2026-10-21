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

async function expectGap(host, statement, code) {
  await assert.rejects(() => host.eval(statement), (err) => {
    assert.equal(err && err.code, code, statement + ': ' + (err && err.message));
    return true;
  });
}

test('M5 shapes, revisions and gap-only objects', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  vm.runInThisContext(require('node:fs').readFileSync(buildEngine(), 'utf8'));
  require('../../public/quire/js/vb-host.js');
  fresh();
  const host = await L.vb.open();
  try {
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Shape"');
    assert.equal(await host.eval('?ActiveDocument.InlineShapes.Count'), ActiveDocument.InlineShapes.Count);
    assert.equal(await host.eval('?ActiveDocument.Shapes.Count'), ActiveDocument.Shapes.Count);
    await expectGap(host, 'ActiveDocument.InlineShapes.AddPicture("picture.png")', 'G-HOST');
    await expectGap(host, 'ActiveDocument.Shapes.AddChart', 'G-ENGINE');
    let shaped = false;
    try {
      await host.eval('ActiveDocument.Shapes.AddShape(1, 10, 20, 72, 36)');
      shaped = true;
    } catch (err) {
      assert.equal(err && err.code, 'G-ENGINE');
    }
    if (shaped) {
      assert.ok(ActiveDocument.Shapes.Count >= 1);
      await host.eval('ActiveDocument.Shapes(1).Width = 90');
      assert.equal(ActiveDocument.Shapes(1).Width, 90);
      await host.eval('ActiveDocument.Shapes(1).Height = 40');
      await host.eval('ActiveDocument.Shapes(1).Fill.ForeColor.RGB = 255');
      assert.equal(ActiveDocument.Shapes(1).Fill.ForeColor.RGB, 255);
      await host.eval('ActiveDocument.Shapes(1).Line.Weight = 2');
      await expectGap(host, '?ActiveDocument.Shapes(1).Left', 'missing-live');
      await expectGap(host, '?ActiveDocument.Shapes(1).WrapFormat', 'missing-live');
      await expectGap(host, '?ActiveDocument.Shapes(1).OLEFormat', 'G-ENGINE');
      await host.eval('ActiveDocument.Shapes(1).Delete');
    }
    try {
      await host.eval('ActiveDocument.Shapes.AddTextbox(1, 0, 0, 72, 36)');
    } catch (err) {
      assert.equal(err && err.code, 'G-ENGINE');
    }

    await host.eval('ActiveDocument.TrackRevisions = True');
    assert.equal(await host.eval('?ActiveDocument.TrackRevisions'), -1);
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Tracked"');
    const revCount = ActiveDocument.Revisions.Count;
    assert.equal(await host.eval('?ActiveDocument.Revisions.Count'), revCount);
    if (revCount >= 1) {
      assert.equal(await host.eval('?ActiveDocument.Revisions(1).Type'), ActiveDocument.Revisions(1).Type);
      await expectGap(host, '?ActiveDocument.Revisions(1).Author', 'missing-live');
      await expectGap(host, '?ActiveDocument.Revisions(1).Date', 'missing-live');
      await expectGap(host, '?ActiveDocument.Revisions(1).Range', 'missing-live');
    }
    await host.eval('ActiveDocument.Revisions.AcceptAll');
    assert.equal(ActiveDocument.Revisions.Count, 0);
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Again"');
    if (ActiveDocument.Revisions.Count >= 1) await host.eval('ActiveDocument.Revisions(1).Reject');
    await host.eval('ActiveDocument.Revisions.RejectAll');
    await host.eval('ActiveDocument.TrackRevisions = False');

    const types = L.domIr.types;
    const gaps = Object.keys(types).filter((name) => types[name].gap).sort();
    assert.ok(gaps.length >= 10);
    for (const name of gaps) {
      assert.equal(await host.eval('?' + name + '.Application Is Application'), -1, name);
      assert.equal(await host.eval('?' + name + '.Parent Is Application'), -1, name);
      await expectGap(host, '?' + name + '.Creator', 'G-ENGINE');
      await expectGap(host, '?' + name + '.Count', types[name].gap);
    }
    await expectGap(host, 'Windows.Item(1)', 'G-HOST');
    await expectGap(host, 'Dialogs.Add', 'G-HOST');
    await expectGap(host, 'ContentControls.Add', 'G-ENGINE');
    await expectGap(host, '?MailMerge.DataSource', 'G-ENGINE');
    await expectGap(host, 'System.PrivateProfileString', 'G-HOST');
  } finally {
    host.close();
  }
});
