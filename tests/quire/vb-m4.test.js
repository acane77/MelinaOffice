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

test('M4 tables, find, bookmarks, fields, links and notes', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  vm.runInThisContext(require('node:fs').readFileSync(buildEngine(), 'utf8'));
  require('../../public/quire/js/vb-host.js');
  fresh();
  const host = await L.vb.open();
  try {
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Hello world."');
    const replaced = await host.eval('?Selection.Find.Execute("Hello", False, False, False, False, False, True, wdFindContinue, False, "Hi", wdReplaceOne)');
    assert.ok(replaced === true || replaced === -1, replaced);
    assert.ok(String(ActiveDocument.Paragraphs(1).Range.Text).indexOf('Hi') >= 0, ActiveDocument.Paragraphs(1).Range.Text);
    await host.eval('Selection.Find.Text = "Hi"');
    await host.eval('Selection.Find.MatchCase = False');
    await expectGap(host, 'Selection.Find.ClearFormatting', 'missing-live');
    assert.equal(await host.eval('?Selection.Find.Text'), 'Hi');
    await expectGap(host, 'Selection.Find.Execute("Hi", False, False, False, False, False, True, 2)', 'G-HOST');
    await expectGap(host, 'Selection.Find.Execute("Hi", False, False, False, False, False, True, wdFindContinue, False, "", wdReplaceNone, True)', 'G-ENGINE');
    await expectGap(host, 'Selection.Find.HitHighlight', 'G-ENGINE');
    await expectGap(host, '?Selection.Find.MatchPrefix', 'G-ENGINE');
    await expectGap(host, '?Selection.Find.Replacement.Font', 'G-ENGINE');

    await host.eval('ActiveDocument.Bookmarks.Add("Here", ActiveDocument.Paragraphs(1).Range)');
    assert.equal(await host.eval('?ActiveDocument.Bookmarks.Count'), ActiveDocument.Bookmarks.Count);
    assert.equal(await host.eval('?ActiveDocument.Bookmarks.Exists("Here")'), ActiveDocument.Bookmarks.Exists('Here'));
    assert.equal(await host.eval('?ActiveDocument.Bookmarks("Here").Name'), 'Here');
    assert.equal(await host.eval('?ActiveDocument.Bookmarks("Here").Empty'), ActiveDocument.Bookmarks('Here').Empty);
    assert.ok(String(await host.eval('?ActiveDocument.Bookmarks("Here").Range.Text')).indexOf('Hi') >= 0);
    await host.eval('ActiveDocument.Bookmarks("Here").Select');
    await expectGap(host, '?ActiveDocument.Bookmarks("Here").Start', 'G-STORY');
    await expectGap(host, '?ActiveDocument.Bookmarks("Here").End', 'G-STORY');
    await expectGap(host, 'ActiveDocument.Bookmarks("Here").Copy("There")', 'G-ENGINE');

    await host.eval('Selection.TypeParagraph');
    await host.eval('Selection.TypeText("link")');
    await host.eval('ActiveDocument.Hyperlinks.Add(Selection.Range, "https://vibeoffice.work", "", "", "Site")');
    assert.ok(ActiveDocument.Hyperlinks.Count >= 1);
    assert.equal(await host.eval('?ActiveDocument.Hyperlinks(1).Address'), ActiveDocument.Hyperlinks(1).Address);
    assert.equal(await host.eval('?ActiveDocument.Hyperlinks(1).TextToDisplay'), ActiveDocument.Hyperlinks(1).TextToDisplay);
    await host.eval('ActiveDocument.Hyperlinks(1).ScreenTip = "home"');
    assert.equal(ActiveDocument.Hyperlinks(1).ScreenTip, 'home');
    await expectGap(host, '?ActiveDocument.Hyperlinks(1).Range.Start', 'G-STORY');
    await expectGap(host, 'ActiveDocument.Hyperlinks(1).Follow', 'G-HOST');
    await expectGap(host, '?ActiveDocument.Hyperlinks(1).Type', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Hyperlinks.Add(Selection.Range, "https://vibeoffice.work", "", "", "x", "frame")', 'G-ENGINE');
    await host.eval('ActiveDocument.Hyperlinks(1).Delete');

    await host.eval('ActiveDocument.Comments.Add(ActiveDocument.Paragraphs(1).Range, "A note")');
    assert.equal(await host.eval('?ActiveDocument.Comments.Count'), ActiveDocument.Comments.Count);
    assert.equal(await host.eval('?ActiveDocument.Comments(1).Author'), ActiveDocument.Comments(1).Author);
    assert.equal(await host.eval('?ActiveDocument.Comments(1).Initial'), ActiveDocument.Comments(1).Initials);
    assert.ok(String(await host.eval('?ActiveDocument.Comments(1).Range.Text')).indexOf('A note') >= 0);
    await expectGap(host, '?ActiveDocument.Comments(1).Scope', 'missing-live');
    await expectGap(host, 'ActiveDocument.Comments(1).Edit', 'G-HOST');
    await expectGap(host, 'ActiveDocument.Comments.AddReply(ActiveDocument.Comments(1), "reply")', 'missing-live');
    await host.eval('ActiveDocument.Comments(1).Delete');
    assert.equal(ActiveDocument.Comments.Count, 0);

    await host.eval('ActiveDocument.Footnotes.Add(ActiveDocument.Paragraphs(1).Range, "Foot")');
    assert.equal(await host.eval('?ActiveDocument.Footnotes.Count'), ActiveDocument.Footnotes.Count);
    assert.equal(await host.eval('?ActiveDocument.Footnotes(1).Index'), ActiveDocument.Footnotes(1).Index);
    assert.ok(String(await host.eval('?ActiveDocument.Footnotes(1).Range.Text')).indexOf('Foot') >= 0);
    await expectGap(host, '?ActiveDocument.Footnotes(1).Reference', 'G-STORY');
    await expectGap(host, 'ActiveDocument.Footnotes.Convert', 'missing-live');
    await host.eval('ActiveDocument.Endnotes.Add(ActiveDocument.Paragraphs(1).Range, "End")');
    assert.equal(await host.eval('?ActiveDocument.Endnotes.Count'), 1);
    await expectGap(host, '?ActiveDocument.Endnotes(1).Reference', 'G-STORY');

    await host.eval('ActiveDocument.Tables.Add(ActiveDocument.Paragraphs(1).Range, 2, 2)');
    assert.equal(await host.eval('?ActiveDocument.Tables.Count'), 1);
    await host.eval('ActiveDocument.Tables(1).Cell(1, 1).Range.Text = "Name"');
    assert.ok(String(ActiveDocument.Tables(1).Cell(1, 1).Range.Text).indexOf('Name') >= 0);
    assert.equal(await host.eval('?ActiveDocument.Tables(1).Cell(1, 1).RowIndex'), 1);
    assert.equal(await host.eval('?ActiveDocument.Tables(1).Cell(1, 1).ColumnIndex'), 1);
    assert.equal(await host.eval('?ActiveDocument.Tables(1).Rows.Count'), 2);
    assert.equal(await host.eval('?ActiveDocument.Tables(1).Columns.Count'), 2);
    await host.eval('ActiveDocument.Tables(1).Cell(1, 1).Shading.BackgroundPatternColor = 65535');
    assert.equal(ActiveDocument.Tables(1).Cell(1, 1).Shading.BackgroundPatternColor, 65535);
    await host.eval('ActiveDocument.Tables(1).Cell(1, 1).Borders(-1).LineStyle = 1');
    assert.equal(ActiveDocument.Tables(1).Cell(1, 1).Borders(-1).LineStyle, 1);
    await host.eval('ActiveDocument.Tables(1).Rows(1).HeadingFormat = True');
    assert.equal(await host.eval('?ActiveDocument.Tables(1).Rows(1).HeadingFormat'), -1);
    await host.eval('ActiveDocument.Tables(1).Columns(1).Width = 80');
    assert.equal(ActiveDocument.Tables(1).Columns(1).Width, 80);
    await host.eval('ActiveDocument.Tables(1).ApplyStyleHeadingRows = True');
    assert.equal(await host.eval('?ActiveDocument.Tables(1).ApplyStyleHeadingRows'), ActiveDocument.Tables(1).ApplyStyleHeadingRows);
    await host.eval('ActiveDocument.Tables(1).AutoFitBehavior(1)');
    await host.eval('ActiveDocument.Tables(1).Rows(1).Cells.DistributeWidth');
    await expectGap(host, '?ActiveDocument.Tables(1).Uniform', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Tables(1).Spacing', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Tables.Add(ActiveDocument.Paragraphs(1).Range, 2, 2, 1)', 'G-ENGINE');
    await host.eval('ActiveDocument.Tables(1).Rows.Add');
    assert.equal(ActiveDocument.Tables(1).Rows.Count, 3);
    await host.eval('ActiveDocument.Tables(1).Columns.Add');
    assert.equal(ActiveDocument.Tables(1).Columns.Count, 3);
    await expectGap(host, 'ActiveDocument.Tables(1).Rows.Delete', 'missing-live');
    await host.eval('ActiveDocument.Tables(1).Rows(3).Delete');
    assert.equal(ActiveDocument.Tables(1).Rows.Count, 2);
    await host.eval('ActiveDocument.Tables(1).ConvertToText');
    assert.equal(ActiveDocument.Tables.Count, 0);
    await host.eval('ActiveDocument.Tables.Add(NumColumns:=2, NumRows:=2, Range:=ActiveDocument.Paragraphs(1).Range)');
    assert.equal(ActiveDocument.Tables.Count, 1);
    assert.equal(ActiveDocument.Tables(1).Columns.Count, 2);
    assert.equal(ActiveDocument.Tables(1).Rows.Count, 2);
    await host.eval('ActiveDocument.Tables(1).Delete');
    assert.equal(ActiveDocument.Tables.Count, 0);
    await host.eval('Selection.Tables.Add(NumColumns:=4, NumRows:=2, Range:=Selection.Range)');
    assert.equal(ActiveDocument.Tables.Count, 1);
    assert.equal(ActiveDocument.Tables(1).Columns.Count, 4);
    await host.eval('ActiveDocument.Tables(1).Delete');
    assert.equal(ActiveDocument.Tables.Count, 0);

    await host.eval('Selection.TypeParagraph');
    await host.eval('Selection.TypeText("field")');
    await host.eval('ActiveDocument.Fields.Add(Selection.Range, 1, "QUOTE hi")');
    assert.ok(ActiveDocument.Fields.Count >= 1);
    assert.equal(await host.eval('?ActiveDocument.Fields.Count'), ActiveDocument.Fields.Count);
    assert.equal(await host.eval('?ActiveDocument.Fields(1).Type'), ActiveDocument.Fields(1).Type);
    assert.ok(String(await host.eval('?ActiveDocument.Fields(1).Code.Text')).indexOf('QUOTE') >= 0);
    await expectGap(host, '?ActiveDocument.Fields(1).Code.Start', 'G-STORY');
    await host.eval('ActiveDocument.Fields.Update');
    await host.eval('ActiveDocument.Fields(1).Update');
    await expectGap(host, '?ActiveDocument.Fields(1).Locked', 'missing-live');
    await expectGap(host, 'ActiveDocument.Fields(1).Unlink', 'missing-live');
    await expectGap(host, 'ActiveDocument.Fields(1).Select', 'missing-live');
    await expectGap(host, 'ActiveDocument.Fields.ToggleShowCodes', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Fields(1).Copy', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Fields(1).DoClick', 'G-HOST');
    await host.eval('ActiveDocument.Fields(1).Delete');
    await host.eval('ActiveDocument.Bookmarks("Here").Delete');
    assert.equal(ActiveDocument.Bookmarks.Exists('Here'), false);
  } finally {
    host.close();
  }
});
