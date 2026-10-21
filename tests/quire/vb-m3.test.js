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

test('M3 font, paragraph format, styles, lists and sections', { timeout: 120000 }, async () => {
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  vm.runInThisContext(require('node:fs').readFileSync(buildEngine(), 'utf8'));
  require('../../public/quire/js/vb-host.js');
  fresh();
  const host = await L.vb.open();
  try {
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Hello"');
    const font = 'ActiveDocument.Paragraphs(1).Range.Font';
    await host.eval(font + '.Name = "Arial"');
    await host.eval(font + '.NameFarEast = "SimSun"');
    await host.eval(font + '.Size = 14');
    await host.eval(font + '.Bold = True');
    await host.eval(font + '.Italic = True');
    await host.eval(font + '.Underline = wdUnderlineSingle');
    await host.eval(font + '.Color = wdColorRed');
    await host.eval(font + '.AllCaps = True');
    await host.eval(font + '.Superscript = True');
    assert.equal(ActiveDocument.Paragraphs(1).Range.Font.Name, 'Arial');
    assert.equal(ActiveDocument.Paragraphs(1).Range.Font.NameFarEast, 'SimSun');
    assert.equal(ActiveDocument.Paragraphs(1).Range.Font.Size, 14);
    assert.equal(ActiveDocument.Paragraphs(1).Range.Font.Bold, -1);
    assert.equal(ActiveDocument.Paragraphs(1).Range.Font.Italic, -1);
    assert.equal(ActiveDocument.Paragraphs(1).Range.Font.Underline, wdUnderlineSingle);
    assert.equal(await host.eval('?' + font + '.Color'), ActiveDocument.Paragraphs(1).Range.Font.Color);
    assert.equal(ActiveDocument.Paragraphs(1).Range.Font.AllCaps, -1);
    assert.equal(ActiveDocument.Paragraphs(1).Range.Font.Superscript, -1);
    assert.equal(await host.eval('?' + font + '.Application Is Application'), -1);
    await expectGap(host, '?' + font + '.Creator', 'G-ENGINE');
    for (const member of ['Animation', 'Borders', 'Shading', 'DiacriticColor', 'StylisticSet', 'NumberSpacing', 'NumberForm', 'Ligatures', 'ContextualAlternates']) {
      await expectGap(host, '?' + font + '.' + member, 'G-ENGINE');
    }

    const fmt = 'ActiveDocument.Paragraphs(1).Format';
    await host.eval(fmt + '.Alignment = wdAlignParagraphRight');
    await host.eval(fmt + '.RightIndent = 18');
    await host.eval(fmt + '.FirstLineIndent = 12');
    await host.eval(fmt + '.SpaceBefore = 6');
    await host.eval(fmt + '.SpaceAfter = 6');
    await host.eval(fmt + '.LineSpacingRule = wdLineSpaceDouble');
    await host.eval(fmt + '.KeepTogether = True');
    await host.eval(fmt + '.WidowControl = False');
    await host.eval(fmt + '.OutlineLevel = 1');
    await host.eval(fmt + '.ReadingOrder = wdReadingOrderRtl');
    await host.eval(fmt + '.DisableLineHeightGrid = True');
    assert.equal(ActiveDocument.Paragraphs(1).Format.Alignment, wdAlignParagraphRight);
    assert.equal(ActiveDocument.Paragraphs(1).Format.RightIndent, 18);
    assert.equal(ActiveDocument.Paragraphs(1).Format.FirstLineIndent, 12);
    assert.equal(ActiveDocument.Paragraphs(1).Format.SpaceBefore, 6);
    assert.equal(ActiveDocument.Paragraphs(1).Format.SpaceAfter, 6);
    assert.equal(ActiveDocument.Paragraphs(1).Format.LineSpacingRule, wdLineSpaceDouble);
    assert.equal(ActiveDocument.Paragraphs(1).Format.KeepTogether, -1);
    assert.equal(ActiveDocument.Paragraphs(1).Format.WidowControl, 0);
    assert.equal(ActiveDocument.Paragraphs(1).Format.OutlineLevel, 1);
    assert.equal(ActiveDocument.Paragraphs(1).Format.ReadingOrder, wdReadingOrderRtl);
    assert.equal(ActiveDocument.Paragraphs(1).Format.DisableLineHeightGrid, -1);
    for (const member of ['CharacterUnitLeftIndent', 'CharacterUnitRightIndent', 'CharacterUnitFirstLineIndent', 'MirrorIndents', 'AutoAdjustRightIndent', 'FarEastLineBreakControl', 'WordWrap', 'HangingPunctuation', 'HalfWidthPunctuationOnTopOfLine', 'AddSpaceBetweenFarEastAndAlpha', 'AddSpaceBetweenFarEastAndDigit']) {
      await expectGap(host, '?' + fmt + '.' + member, 'G-ENGINE');
    }
    await expectGap(host, '?' + fmt + '.Borders', 'missing-live');
    await expectGap(host, '?' + fmt + '.Shading', 'missing-live');

    await host.eval(fmt + '.TabStops.Add(72, 1, 1)');
    assert.equal(ActiveDocument.Paragraphs(1).Format.TabStops.Count, 1);
    assert.equal(await host.eval('?' + fmt + '.TabStops(1).Position'), 72);
    assert.equal(await host.eval('?' + fmt + '.TabStops.Item(1).Alignment'), ActiveDocument.Paragraphs(1).Format.TabStops.Item(1).Alignment);
    await host.eval(fmt + '.TabStops(1).Clear');
    assert.equal(ActiveDocument.Paragraphs(1).Format.TabStops.Count, 0);
    await host.eval(fmt + '.TabStops.Add(36, 0, 0)');
    await host.eval(fmt + '.TabStops.ClearAll');
    assert.equal(ActiveDocument.Paragraphs(1).Format.TabStops.Count, 0);

    assert.equal(await host.eval('?ActiveDocument.Styles.Count'), ActiveDocument.Styles.Count);
    assert.equal(await host.eval('?ActiveDocument.Styles(1).NameLocal'), ActiveDocument.Styles(1).NameLocal);
    assert.equal(await host.eval('?ActiveDocument.Styles(1).Type'), ActiveDocument.Styles(1).Type);
    assert.equal(await host.eval('?ActiveDocument.Styles(1).InUse'), ActiveDocument.Styles(1).InUse);
    assert.equal(typeof await host.eval('?ActiveDocument.Styles(1).Font'), 'object');
    assert.equal(typeof await host.eval('?ActiveDocument.Styles(1).ParagraphFormat'), 'object');
    await host.eval('ActiveDocument.Styles.Add("M3Style", wdStyleTypeParagraph)');
    assert.equal(ActiveDocument.Styles('M3Style').NameLocal, 'M3Style');
    await host.eval('ActiveDocument.Styles("M3Style").Font.Name = "Arial"');
    assert.equal(ActiveDocument.Styles('M3Style').Font.Name, 'Arial');
    await host.eval('ActiveDocument.Styles("M3Style").NameLocal = "M3Renamed"');
    assert.equal(ActiveDocument.Styles('M3Renamed').NameLocal, 'M3Renamed');
    await host.eval('ActiveDocument.Styles("M3Renamed").Delete');
    await expectGap(host, 'ActiveDocument.Styles(1).Delete', 'G-ENGINE');
    for (const member of ['LinkStyle', 'Visibility', 'UnhideWhenUsed', 'QuickStyle', 'Priority', 'Locked']) {
      await expectGap(host, '?ActiveDocument.Styles(1).' + member, 'G-ENGINE');
    }

    await host.eval('ActiveDocument.Paragraphs(1).Range.ListFormat.ApplyNumberDefault');
    assert.equal(await host.eval('?ActiveDocument.Paragraphs(1).Range.ListFormat.ListLevelNumber'), ActiveDocument.Paragraphs(1).Range.ListFormat.ListLevelNumber);
    await host.eval('ActiveDocument.Paragraphs(1).Range.ListFormat.ListLevelNumber = 2');
    assert.equal(ActiveDocument.Paragraphs(1).Range.ListFormat.ListLevelNumber, 2);
    await host.eval('ActiveDocument.Paragraphs(1).Range.ListFormat.ListIndent');
    await host.eval('ActiveDocument.Paragraphs(1).Range.ListFormat.ListOutdent');
    assert.ok(ActiveDocument.Lists.Count >= 1);
    assert.equal(await host.eval('?ActiveDocument.Lists.Count'), ActiveDocument.Lists.Count);
    assert.equal(typeof await host.eval('?ActiveDocument.Lists(1).ListParagraphs'), 'object');
    assert.ok(ActiveDocument.Lists(1).ListParagraphs.Count >= 1);
    await expectGap(host, '?ActiveDocument.Paragraphs(1).Range.ListFormat.ListType', 'missing-live');
    await expectGap(host, 'ActiveDocument.Paragraphs(1).Range.ListFormat.ApplyListTemplate', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Paragraphs(1).Range.ListFormat.CanContinuePreviousList', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Lists.ListTemplates', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Lists.ListTemplate', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Lists.ListLevels', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Lists.ListLevel', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Lists(1).Range', 'missing-live');
    await expectGap(host, 'ActiveDocument.Lists(1).CountNumberedItems', 'missing-live');
    await host.eval('ActiveDocument.Paragraphs(1).Range.ListFormat.RemoveNumbers');

    assert.equal(await host.eval('?ActiveDocument.Sections.Count'), ActiveDocument.Sections.Count);
    assert.equal(typeof await host.eval('?ActiveDocument.Sections(1).Range'), 'object');
    await host.eval('ActiveDocument.Sections(1).PageSetup.PageWidth = 612');
    await host.eval('ActiveDocument.Sections(1).PageSetup.PageHeight = 792');
    await host.eval('ActiveDocument.Sections(1).PageSetup.TopMargin = 72');
    await host.eval('ActiveDocument.Sections(1).PageSetup.LeftMargin = 72');
    await host.eval('ActiveDocument.Sections(1).PageSetup.Gutter = 0');
    await host.eval('ActiveDocument.Sections(1).PageSetup.HeaderDistance = 36');
    await host.eval('ActiveDocument.Sections(1).PageSetup.FooterDistance = 36');
    assert.equal(ActiveDocument.Sections(1).PageSetup.PageWidth, 612);
    assert.equal(ActiveDocument.Sections(1).PageSetup.PageHeight, 792);
    assert.equal(ActiveDocument.Sections(1).PageSetup.TopMargin, 72);
    assert.equal(await host.eval('?ActiveDocument.Sections(1).PageSetup.LeftMargin'), 72);
    await host.eval('ActiveDocument.Sections(1).PageSetup.Orientation = wdOrientLandscape');
    assert.equal(ActiveDocument.Sections(1).PageSetup.Orientation, wdOrientLandscape);
    await host.eval('ActiveDocument.Sections(1).PageSetup.TextColumns.Count = 2');
    await host.eval('ActiveDocument.Sections(1).PageSetup.TextColumns.Spacing = 36');
    await host.eval('ActiveDocument.Sections(1).PageSetup.TextColumns.EvenlySpaced = True');
    assert.equal(ActiveDocument.Sections(1).PageSetup.TextColumns.Count, 2);
    assert.equal(ActiveDocument.Sections(1).PageSetup.TextColumns.Spacing, 36);
    assert.equal(ActiveDocument.Sections(1).PageSetup.TextColumns.EvenlySpaced, -1);
    assert.equal(await host.eval('?ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Exists'), ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Exists);
    assert.equal(await host.eval('?ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text'), ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text);
    assert.equal(await host.eval('?ActiveDocument.Sections(1).Application Is Application'), -1);
    await expectGap(host, '?ActiveDocument.Sections(1).Creator', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Sections(1).PageSetup.PaperSize', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).LinkToPrevious', 'G-ENGINE');
    await expectGap(host, '?ActiveDocument.Sections(1).ProtectedForForms', 'G-ENGINE');
  } finally {
    host.close();
  }
});
