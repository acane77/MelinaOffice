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

test('M1 Application, Documents and Document are callable from VB', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  vm.runInThisContext(require('node:fs').readFileSync(buildEngine(), 'utf8'));
  require('../../public/quire/js/vb-host.js');
  fresh();
  const host = await L.vb.open();
  try {
    for (const key of Object.keys(L.dom)) {
      if (!/^wd[A-Za-z][A-Za-z0-9]*$/.test(key) || !Number.isInteger(L.dom[key])) continue;
      assert.equal(await host.eval('?' + key), L.dom[key], key);
    }

    assert.equal(await host.eval('?ActiveDocument.Name'), ActiveDocument.Name);
    assert.equal(await host.eval('?Application.ActiveDocument.Name'), ActiveDocument.Name);
    assert.equal(await host.eval('?Documents.Count'), Documents.Count);
    assert.equal(await host.eval('?Documents(1).Name'), Documents(1).Name);
    assert.equal(await host.eval('?Documents.Item(1).Name'), Documents.Item(1).Name);
    assert.equal(await host.eval('?Documents.First.Name'), Documents.First.Name);
    assert.equal(await host.eval('?Documents.Last.Name'), Documents.Last.Name);
    assert.equal(await host.eval('?ActiveDocument.Parent Is Documents'), -1);
    assert.equal(await host.eval('?ActiveDocument.Application Is Application'), -1);
    assert.equal(await host.eval('?Documents.Parent Is Application'), -1);
    assert.equal(await host.eval('?Application.Parent Is Application'), -1);
    assert.equal(typeof await host.eval('?Selection'), 'object');
    assert.equal(typeof await host.eval('?ActiveDocument.Content'), 'object');
    assert.equal(await host.eval('?ActiveDocument.Content.Parent.Name'), ActiveDocument.Name);
    for (const member of ['Paragraphs', 'Sentences', 'Words', 'Tables', 'Sections', 'Styles', 'Bookmarks', 'Fields', 'Hyperlinks', 'Comments', 'Footnotes', 'Endnotes', 'InlineShapes', 'Shapes', 'Lists', 'ListParagraphs', 'Revisions']) {
      assert.equal(typeof await host.eval('?ActiveDocument.' + member), 'object', member);
    }
    assert.equal(typeof await host.eval('?ActiveDocument.StoryRanges(wdMainTextStory)'), 'object');

    const name = ActiveDocument.Name;
    await assert.rejects(() => host.eval('ActiveDocument.Name = "Z"'));
    assert.equal(ActiveDocument.Name, name);

    await host.eval('ActiveDocument.Saved = False');
    assert.equal(ActiveDocument.Saved, false);
    await host.eval('ActiveDocument.Saved = True');
    assert.equal(ActiveDocument.Saved, true);
    await host.eval('ActiveDocument.TrackRevisions = True');
    assert.equal(ActiveDocument.TrackRevisions, -1);
    await host.eval('ActiveDocument.TrackRevisions = False');
    assert.equal(ActiveDocument.TrackRevisions, 0);

    await host.eval('ActiveDocument.BuiltInDocumentProperties("Title").Value = "M1"');
    assert.equal(ActiveDocument.BuiltInDocumentProperties('Title').Value, 'M1');
    assert.equal(await host.eval('?ActiveDocument.BuiltInDocumentProperties("Title").Name'), 'Title');
    await host.eval('ActiveDocument.CustomDocumentProperties("Note").Value = "x"');
    assert.equal(ActiveDocument.CustomDocumentProperties('Note').Value, 'x');

    assert.equal(await host.eval('?ActiveDocument.ComputeStatistics(wdStatisticParagraphs)'), ActiveDocument.ComputeStatistics(wdStatisticParagraphs));
    assert.equal(await host.eval('?ActiveDocument.ComputeStatistics(wdStatisticWords)'), ActiveDocument.ComputeStatistics(wdStatisticWords));
    const undone = await host.eval('?ActiveDocument.Undo');
    assert.equal(typeof undone === 'boolean' || typeof undone === 'number', true);
    await host.eval('ActiveDocument.Redo');
    await host.eval('ActiveDocument.Select');
    await host.eval('ActiveDocument.Activate');
    await host.eval('ActiveDocument.Save');

    await expectGap(host, '?ActiveDocument.Creator', 'G-ENGINE');
    await expectGap(host, '?Application.Creator', 'G-ENGINE');
    await expectGap(host, '?Documents.Creator', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.FullName', 'G-HOST');
    await expectGap(host, '?ActiveDocument.Path', 'G-HOST');
    assert.equal(typeof await host.eval('?ActiveDocument.Characters'), 'object');
    assert.equal(typeof await host.eval('?ActiveDocument.Range(0, 5)'), 'object');
    await expectGap(host, 'ActiveDocument.GoTo', 'G-STORY');
    await expectGap(host, 'ActiveDocument.UndoClear', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Protect', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Unprotect', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.ComputeStatistics(wdStatisticLines)', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.ComputeStatistics(wdStatisticPages)', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.BuiltInDocumentProperties("NoSuch").Value', 'G-ENGINE');
    await expectGap(host, 'Documents.Add("x")', 'G-ENGINE');
    await expectGap(host, 'Documents.Open("x")', 'G-HOST');
    await expectGap(host, 'Documents.Close', 'G-HOST');
    await expectGap(host, 'Documents.Save', 'G-HOST');
    await expectGap(host, 'ActiveDocument.SaveAs("x")', 'G-HOST');
    await expectGap(host, 'ActiveDocument.SaveAs2("x")', 'G-HOST');
    await expectGap(host, 'ActiveDocument.PrintOut', 'G-HOST');
    await expectGap(host, 'ActiveDocument.PrintPreview', 'G-HOST');
    await expectGap(host, 'ActiveDocument.Compare("x")', 'G-HOST');
    await expectGap(host, 'ActiveDocument.FollowHyperlink("x")', 'G-HOST');
    await expectGap(host, 'ActiveDocument.ExportAsFixedFormat', 'G-HOST');
    await expectGap(host, 'ActiveDocument.Close(wdPromptToSaveChanges)', 'G-HOST');
    await expectGap(host, 'ActiveDocument.Close(wdDoNotSaveChanges, 0)', 'G-HOST');
    assert.equal(await host.eval('?ActiveDocument.Name'), name);

    const before = Documents.Count;
    await host.eval('Documents.Add');
    assert.equal(Documents.Count, before + 1);
    await host.eval('ActiveDocument.Close(wdDoNotSaveChanges)');
    assert.equal(Documents.Count, before);
  } finally {
    host.close();
  }
});
