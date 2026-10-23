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

test('M2 Selection, Range, paragraphs, sentences, words and characters', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  vm.runInThisContext(require('node:fs').readFileSync(buildEngine(), 'utf8'));
  require('../../public/quire/js/vb-host.js');
  fresh();
  const host = await L.vb.open();
  try {
    const nothing = await host.eval('?ActiveDocument.Paragraphs(1).Next');
    assert.equal(nothing && nothing.__nothing, true);
    assert.equal((await host.eval('?ActiveDocument.Paragraphs(1).Previous')).__nothing, true);

    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Hello world."');
    assert.equal(await host.eval('?ActiveDocument.Paragraphs(1).Range.Text'), ActiveDocument.Paragraphs(1).Range.Text);
    assert.equal(await host.eval('?ActiveDocument.Paragraphs.Count'), ActiveDocument.Paragraphs.Count);
    assert.equal(await host.eval('?ActiveDocument.Paragraphs(1).Range.Text'), ActiveDocument.Paragraphs.Item(1).Range.Text);
    assert.equal(await host.eval('?ActiveDocument.Paragraphs.First.Range.Text'), ActiveDocument.Paragraphs.First.Range.Text);
    assert.equal(await host.eval('?ActiveDocument.Paragraphs.Last.Range.Text'), ActiveDocument.Paragraphs.Last.Range.Text);
    assert.equal(await host.eval('?ActiveDocument.Words.Count'), ActiveDocument.Words.Count);
    assert.equal(await host.eval('?ActiveDocument.Sentences.Count'), ActiveDocument.Sentences.Count);
    assert.equal(await host.eval('?ActiveDocument.Paragraphs(1).Range.Words(1).Text'), ActiveDocument.Paragraphs(1).Range.Words(1).Text);
    assert.equal(await host.eval('?ActiveDocument.Paragraphs(1).Range.Characters.Count'), ActiveDocument.Paragraphs(1).Range.Characters.Count);
    assert.equal(typeof await host.eval('?ActiveDocument.Paragraphs(1).Range.Characters(1)'), 'object');
    assert.equal(typeof await host.eval('?ActiveDocument.Sentences.First'), 'object');
    assert.equal(typeof await host.eval('?ActiveDocument.Words.Last'), 'object');

    await host.eval('ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphCenter');
    assert.equal(ActiveDocument.Paragraphs(1).Alignment, wdAlignParagraphCenter);
    await host.eval('ActiveDocument.Paragraphs(1).LeftIndent = 36');
    assert.equal(ActiveDocument.Paragraphs(1).LeftIndent, 36);
    await host.eval('ActiveDocument.Paragraphs(1).Space1');
    assert.equal(ActiveDocument.Paragraphs(1).LineSpacingRule, 0);
    await host.eval('ActiveDocument.Paragraphs(1).Space15');
    assert.equal(ActiveDocument.Paragraphs(1).LineSpacingRule, 1);
    await host.eval('ActiveDocument.Paragraphs(1).Space2');
    assert.equal(ActiveDocument.Paragraphs(1).LineSpacingRule, 2);
    const indent = Number(ActiveDocument.Paragraphs(1).LeftIndent) || 0;
    await host.eval('ActiveDocument.Paragraphs(1).IndentCharWidth');
    assert.equal(ActiveDocument.Paragraphs(1).LeftIndent, indent + 36);
    assert.equal(await host.eval('?ActiveDocument.Paragraphs(1).Style'), ActiveDocument.Paragraphs(1).Style);
    assert.equal(typeof await host.eval('?ActiveDocument.Paragraphs(1).Format'), 'object');
    assert.equal(typeof await host.eval('?ActiveDocument.Paragraphs(1).TabStops'), 'object');
    assert.equal(await host.eval('?ActiveDocument.Paragraphs(1).Application Is Application'), -1);
    assert.equal(await host.eval('?ActiveDocument.Paragraphs(1).Parent.Name'), ActiveDocument.Name);
    await expectGap(host, '?ActiveDocument.Paragraphs(1).Creator', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Paragraphs(1).OutlineDemote', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Paragraphs(1).OutlinePromote', 'G-ENGINE');
    await expectGap(host, 'ActiveDocument.Paragraphs(1).OutlineDemoteToBody', 'G-ENGINE');
    assert.equal(typeof await host.eval('?ActiveDocument.Paragraphs(1).DropCap'), 'object');
    assert.equal((await host.eval('?ActiveDocument.Paragraphs(1).Next(2)')).__nothing, true);

    for (const member of ['Font', 'ParagraphFormat', 'Find', 'Characters', 'Paragraphs', 'Sentences', 'Words', 'Tables', 'ListFormat', 'Hyperlinks', 'Bookmarks', 'Fields', 'Comments', 'InlineShapes']) {
      assert.equal(typeof await host.eval('?ActiveDocument.Paragraphs(1).Range.' + member), 'object', member);
    }
    await host.eval('ActiveDocument.Paragraphs(1).Range.Bold = True');
    assert.equal(ActiveDocument.Paragraphs(1).Range.Bold, -1);
    await host.eval('ActiveDocument.Paragraphs(1).Range.Italic = False');
    assert.equal(ActiveDocument.Paragraphs(1).Range.Italic, 0);
    await host.eval('ActiveDocument.Paragraphs(1).Range.Underline = wdUnderlineSingle');
    assert.equal(ActiveDocument.Paragraphs(1).Range.Underline, wdUnderlineSingle);
    await host.eval('ActiveDocument.Paragraphs(1).Range.LanguageID = wdEnglishUS');
    assert.equal(ActiveDocument.Paragraphs(1).Range.LanguageID, wdEnglishUS);
    await host.eval('ActiveDocument.Paragraphs(1).Range.Case = wdUpperCase');
    assert.ok(ActiveDocument.Paragraphs(1).Range.Text.includes('HELLO'));
    assert.equal(await host.eval('?ActiveDocument.Paragraphs(1).Range.Application Is Application'), -1);
    await expectGap(host, '?ActiveDocument.Paragraphs(1).Range.Creator', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Paragraphs(1).Range.FormattedText', 'G-ENGINE');
    assert.equal(typeof await host.eval('?ActiveDocument.Paragraphs(1).Range.Duplicate'), 'object');
    await host.eval('ActiveDocument.Paragraphs(1).Range.InsertBefore("Hi ")');
    assert.ok(ActiveDocument.Paragraphs(1).Range.Text.startsWith('Hi '));
    await host.eval('ActiveDocument.Paragraphs(1).Range.Copy');

    await host.eval('Selection.TypeText("Z")');
    assert.equal(await host.eval('?Selection.Text'), Selection.Text);
    await host.eval('Selection.TypeParagraph');
    assert.ok(ActiveDocument.Paragraphs.Count >= 2);
    assert.equal(typeof await host.eval('?ActiveDocument.Paragraphs(1).Next'), 'object');
    await host.eval('Selection.TypeBackspace');
    await host.eval('Selection.Collapse(wdCollapseStart)');
    await host.eval('Selection.WholeStory');
    await host.eval('Selection.Select');
    assert.equal(await host.eval('?Selection.Type'), Selection.Type);
    assert.equal(typeof await host.eval('?Selection.Font'), 'object');
    assert.equal(typeof await host.eval('?Selection.ParagraphFormat'), 'object');
    assert.equal(typeof await host.eval('?Selection.Range'), 'object');
    assert.equal(typeof await host.eval('?Selection.Find'), 'object');
    assert.equal(await host.eval('?Selection.Style'), Selection.Style);
    await host.eval('Selection.InsertBefore("[")');
    await host.eval('Selection.InsertAfter("]")');
    await host.eval('Selection.InsertSymbol(65)');
    await host.eval('Selection.InsertBreak(wdPageBreak)');
    await host.eval('Selection.InsertDateTime("yyyy")');
    assert.ok(ActiveDocument.Content.Text.includes(String(new Date().getFullYear())));
    await host.eval('Selection.Copy');
    await host.eval('Selection.Paste');

    for (const statement of [
      'Selection.MoveUp', 'Selection.MoveDown',
      'Selection.HomeKey', 'Selection.EndKey', 'Selection.GoTo', 'Selection.GoToNext', 'Selection.GoToPrevious', 'Selection.Information(1)',
      'ActiveDocument.Paragraphs(1).Range.GoTo', 'ActiveDocument.Paragraphs(1).Range.Information(1)',
      'Selection.Delete(1, 1)',
    ]) await expectGap(host, statement, 'G-STORY');
    assert.equal(typeof await host.eval('?Selection.Start'), 'number');
    assert.equal(typeof await host.eval('?Selection.End'), 'number');
    await host.eval('Selection.MoveLeft');
    await host.eval('Selection.MoveRight');
    await host.eval('Selection.MoveLeft(Extend:=True, Count:=1, Unit:=1)');
    assert.equal(typeof await host.eval('?ActiveDocument.Paragraphs(1).Range.Start'), 'number');
    assert.equal(typeof await host.eval('?ActiveDocument.Paragraphs(1).Range.End'), 'number');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Move(1, 1)');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Expand(1)');
    await expectGap(host, 'Selection.InsertFile("x")', 'G-HOST');
    for (const statement of [
      'Selection.InsertCaption', 'Selection.PasteAndFormat', 'Selection.PasteSpecial', 'Selection.PasteExcelTable',
      'Selection.InsertRows', 'Selection.InsertColumns', 'Selection.InsertCells',
      'Selection.ConvertToTable(0, 1, 1, 0, 0, True)',
      'Selection.InsertDateTime("yyyy", False, True)',
    ]) await expectGap(host, statement, 'G-ENGINE');
  } finally {
    host.close();
  }
});
