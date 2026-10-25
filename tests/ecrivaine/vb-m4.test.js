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

async function expectGap(host, statement, code) {
  await assert.rejects(() => host.eval(statement), (err) => {
    assert.equal(err && err.code, code, statement + ': ' + (err && err.message));
    return true;
  });
}

test('M4 tables, find, bookmarks, fields, links and notes', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  vm.runInThisContext(require('node:fs').readFileSync(buildEngine(), 'utf8'));
  require('../../public/ecrivaine/js/vb-host.js');
  fresh();
  const host = await L.vb.open();
  try {
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Hello world."');
    const replaced = await host.eval('?Selection.Find.Execute("Hello", False, False, False, False, False, True, wdFindContinue, False, "Hi", wdReplaceOne)');
    assert.ok(replaced === true || replaced === -1, replaced);
    assert.ok(String(ActiveDocument.Paragraphs(1).Range.Text).indexOf('Hi') >= 0, ActiveDocument.Paragraphs(1).Range.Text);
    await host.eval('Selection.Find.Text = "Hi"');
    await host.eval('Selection.Find.MatchCase = False');
    await host.eval('Selection.Find.ClearFormatting');
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
    assert.equal(typeof await host.eval('?ActiveDocument.Bookmarks("Here").Start'), 'number');
    assert.equal(typeof await host.eval('?ActiveDocument.Bookmarks("Here").End'), 'number');
    await expectGap(host, 'ActiveDocument.Bookmarks("Here").Copy("There")', 'G-ENGINE');

    await host.eval('Selection.TypeParagraph');
    await host.eval('Selection.TypeText("link")');
    await host.eval('ActiveDocument.Hyperlinks.Add(Selection.Range, "https://acane.net/melinaoffice", "", "", "Site")');
    assert.ok(ActiveDocument.Hyperlinks.Count >= 1);
    assert.equal(await host.eval('?ActiveDocument.Hyperlinks(1).Address'), ActiveDocument.Hyperlinks(1).Address);
    assert.equal(await host.eval('?ActiveDocument.Hyperlinks(1).TextToDisplay'), ActiveDocument.Hyperlinks(1).TextToDisplay);
    await host.eval('ActiveDocument.Hyperlinks(1).ScreenTip = "home"');
    assert.equal(ActiveDocument.Hyperlinks(1).ScreenTip, 'home');
    assert.equal(typeof await host.eval('?ActiveDocument.Hyperlinks(1).Range.Start'), 'number');
    await expectGap(host, 'ActiveDocument.Hyperlinks(1).Follow', 'G-HOST');
    await expectGap(host, '?ActiveDocument.Hyperlinks(1).Type', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Hyperlinks.Add(Selection.Range, "https://acane.net/melinaoffice", "", "", "x", "frame")', 'G-ENGINE');
    await host.eval('ActiveDocument.Hyperlinks(1).Delete');

    await host.eval('ActiveDocument.Comments.Add(ActiveDocument.Paragraphs(1).Range, "A note")');
    assert.equal(await host.eval('?ActiveDocument.Comments.Count'), ActiveDocument.Comments.Count);
    assert.equal(await host.eval('?ActiveDocument.Comments(1).Author'), ActiveDocument.Comments(1).Author);
    assert.equal(await host.eval('?ActiveDocument.Comments(1).Initial'), ActiveDocument.Comments(1).Initials);
    assert.ok(String(await host.eval('?ActiveDocument.Comments(1).Range.Text')).indexOf('A note') >= 0);
    assert.equal(typeof await host.eval('?ActiveDocument.Comments(1).Scope'), 'object');
    await expectGap(host, 'ActiveDocument.Comments(1).Edit', 'G-ENGINE');
    await host.eval('ActiveDocument.Comments.AddReply(ActiveDocument.Comments(1), "reply")');
    await host.eval('ActiveDocument.Comments(1).Delete');
    assert.equal(ActiveDocument.Comments.Count, 0);

    await host.eval('ActiveDocument.Footnotes.Add(ActiveDocument.Paragraphs(1).Range, "Foot")');
    assert.equal(await host.eval('?ActiveDocument.Footnotes.Count'), ActiveDocument.Footnotes.Count);
    assert.equal(await host.eval('?ActiveDocument.Footnotes(1).Index'), ActiveDocument.Footnotes(1).Index);
    assert.ok(String(await host.eval('?ActiveDocument.Footnotes(1).Range.Text')).indexOf('Foot') >= 0);
    assert.equal(typeof await host.eval('?ActiveDocument.Footnotes(1).Reference'), 'object');
    await host.eval('ActiveDocument.Footnotes.Convert');
    assert.equal(ActiveDocument.Footnotes.Count, 0);
    await host.eval('ActiveDocument.Endnotes.Add(ActiveDocument.Paragraphs(1).Range, "End")');
    assert.equal(await host.eval('?ActiveDocument.Endnotes.Count') >= 1, true);
    assert.equal(typeof await host.eval('?ActiveDocument.Endnotes(1).Reference'), 'object');

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
    await host.eval('ActiveDocument.Tables(1).Spacing = 6');
    assert.equal(await host.eval('?ActiveDocument.Tables(1).Spacing'), 6);
    assert.equal(typeof await host.eval('?ActiveDocument.Tables(1).Uniform'), 'number');
    await expectGap(host, 'ActiveDocument.Tables.Add(ActiveDocument.Paragraphs(1).Range, 2, 2, 9)', 'G-ENGINE');
    await host.eval('ActiveDocument.Tables(1).Rows.Add');
    assert.equal(ActiveDocument.Tables(1).Rows.Count, 3);
    await host.eval('ActiveDocument.Tables(1).Columns.Add');
    assert.equal(ActiveDocument.Tables(1).Columns.Count, 3);
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
    assert.equal(typeof await host.eval('?ActiveDocument.Fields(1).Code.Start'), 'number');
    await host.eval('ActiveDocument.Fields.Update');
    await host.eval('ActiveDocument.Fields(1).Update');
    await host.eval('ActiveDocument.Fields(1).Locked = True');
    assert.equal(await host.eval('?ActiveDocument.Fields(1).Locked'), -1);
    await host.eval('ActiveDocument.Fields(1).Select');
    await expectGap(host, 'ActiveDocument.Fields.ToggleShowCodes', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Fields(1).Copy', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Fields(1).DoClick', 'G-ENGINE');
    await host.eval('ActiveDocument.Fields(1).Unlink');
    await host.eval('ActiveDocument.Bookmarks("Here").Delete');
    assert.equal(ActiveDocument.Bookmarks.Exists('Here'), false);

    await host.eval('Selection.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:=4');
    const added = ActiveDocument.Tables(ActiveDocument.Tables.Count);
    await host.eval('ActiveDocument.Tables(ActiveDocument.Tables.Count).Borders.Enable = True');
    assert.equal(added.Borders.Enable, -1);
    for (const edge of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
      assert.equal(added._table.tblPr.borders[edge].val, 'single');
    }
    added.Borders(-1).LineWidth = 8;
    added.Borders(-1).Color = 255;
    assert.equal(added.Borders(-1).LineWidth, 8);
    assert.equal(added.Borders(-1).Color, 255);
    added.Cell(1, 1).Borders(-1).LineStyle = 1;
    assert.equal(added.Cell(1, 1).Borders(-1).LineStyle, 1);
    await host.eval('ActiveDocument.Tables(ActiveDocument.Tables.Count).Borders.Enable = False');
    assert.equal(added.Borders.Enable, 0);
    assert.equal(added._table.tblPr.borders.top.val, 'nil');
    const tableEnd = added.Range.End;
    Selection.SetRange(tableEnd + 1, tableEnd + 1);
    let seen = 0;
    let caretPara = null;
    D.walk(D.doc.main, (block) => {
      if (caretPara || block.t !== 'p') return;
      const mark = D.ptext(block).length + 1;
      if (Selection.Start < seen + mark) caretPara = block;
      seen += mark;
    });
    assert.ok(caretPara);
    assert.equal(D.cellOf(D.doc, caretPara), null);
    added.Rows.Delete();
    assert.equal(ActiveDocument.Tables.Count, 0);
    ActiveDocument.Sections(1).PageSetup.PaperSize = 7;
    assert.equal(ActiveDocument.Sections(1).PageSetup.PaperSize, 7);
  } finally {
    host.close();
  }
});
