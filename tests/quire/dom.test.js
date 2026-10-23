const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/quire/js/dmodel.js');
require('../../public/quire/js/ops.js');
require('../../public/quire/js/dom.js');
const { D } = L;

function fresh() {
  D.doc = D.newDoc();
  D.resetHistory();
  if (L.ed) L.ed.sel = null;
  ActiveDocument;
}

test('Documents.Add starts with one paragraph', () => {
  fresh();
  Documents.Add();
  assert.equal(ActiveDocument.Paragraphs.Count, 1);
  assert.equal(ActiveDocument.Paragraphs(1)._block, ActiveDocument.Paragraphs.Item(1)._block);
});

test('Document.Range uses the story index', () => {
  fresh();
  Documents.Add();
  ActiveDocument.Paragraphs(1).Range.Text = 'Hello';
  const range = ActiveDocument.Range(0, 5);
  assert.equal(range.Text, 'Hello');
  assert.equal(range.Start, 0);
  assert.equal(range.End, 5);
});

test('Selection.Start is the story offset', () => {
  fresh();
  Documents.Add();
  assert.equal(typeof Selection.Start, 'number');
  assert.equal(typeof Selection.End, 'number');
  assert.equal(ActiveDocument.Paragraphs(1).Range.Start, 0);
  assert.equal(ActiveDocument.Paragraphs(1).Range.End, 0);
});

test('Creator throws G-ENGINE', () => {
  fresh();
  Documents.Add();
  assert.throws(() => ActiveDocument.Creator, (e) => e.code === 'G-ENGINE');
  assert.throws(() => Application.Creator, (e) => e.code === 'G-ENGINE');
  assert.equal(Application, L.dom.Application);
  assert.equal(window.Application, L.dom.Application);
});

test('Add arguments, Open and SaveAs throw the gap codes', () => {
  fresh();
  assert.throws(() => Documents.Add('template'), (e) => e.code === 'G-ENGINE');
  assert.throws(() => Documents.Open('a.docx'), (e) => e.code === 'G-HOST');
  const doc = Documents.Add();
  assert.throws(() => doc.SaveAs('a.docx'), (e) => e.code === 'G-HOST');
  assert.equal(wdAlignParagraphCenter, 1);
  assert.equal(L.dom.wdAlignParagraphCenter, 1);
  assert.equal(wdColorAutomatic, -16777216);
});

test('paragraph text stays one paragraph and undo restores it', () => {
  fresh();
  Documents.Add();
  const p = ActiveDocument.Paragraphs(1);
  p.Range.Text = 'Hello';
  assert.equal(p.Range.Text, 'Hello\r');
  assert.equal(ActiveDocument.model.main.blocks.length, 1);
  assert.equal(ActiveDocument.model.main.blocks[0], p._block);
  assert.equal(D.ptext(p._block), 'Hello');
  p.Range.Text = 'B';
  assert.equal(ActiveDocument.Undo(1), true);
  assert.equal(p.Range.Text, 'Hello\r');
  assert.equal(D.ptext(p._block), 'Hello');
});

test('alignment and style write the paragraph properties', () => {
  fresh();
  Documents.Add();
  const p = ActiveDocument.Paragraphs(1);
  p.Range.Style = 'Heading 1';
  assert.equal(p._block.pPr.style, 'Heading1');
  assert.equal(p.Style, 'Heading 1');
  p.Alignment = wdAlignParagraphCenter;
  assert.equal(p._block.pPr.jc, 'center');
  assert.equal(p.Format.Alignment, wdAlignParagraphCenter);
  assert.equal(p._block.pPr.style, 'Heading1');
});

test('font bold reads back as a stable truthy value', () => {
  fresh();
  Documents.Add();
  const p = ActiveDocument.Paragraphs(1);
  p.Range.Text = 'Hello';
  p.Range.Font.Bold = true;
  assert.ok(p.Range.Font.Bold);
  assert.equal(p.Range.Font.Bold, p.Range.Font.Bold);
  assert.equal(p._block.runs[0].rPr.b, true);
  p.Range.Font.Color = wdColorRed;
  assert.equal(p._block.runs[0].rPr.color, 'FF0000');
  assert.equal(p.Range.Font.Color, wdColorRed);
});

test('TypeText appends and TypeParagraph adds a paragraph', () => {
  fresh();
  Documents.Add();
  Selection.TypeText('A');
  Selection.TypeText('B');
  assert.equal(D.ptext(ActiveDocument.Paragraphs(1)._block), 'AB');
  Selection.TypeParagraph();
  assert.equal(ActiveDocument.Paragraphs.Count, 2);
});

test('collapsed font is used by the next TypeText', () => {
  fresh();
  Documents.Add();
  Selection.Collapse(wdCollapseStart);
  Selection.Font.Bold = true;
  Selection.TypeText('A');
  assert.equal(ActiveDocument.Paragraphs(1)._block.runs[0].rPr.b, true);
  assert.ok(Selection.Font.Bold);
});

test('typing stays in the heading when the range edit did not move the selection', () => {
  fresh();
  Documents.Add();
  const p = ActiveDocument.Paragraphs(1);
  p.Style = 'Heading 1';
  p.Alignment = wdAlignParagraphCenter;
  p.Range.Text = 'DOM Test';
  p.Range.Font.Color = wdColorDarkBlue;
  p.Range.Collapse(wdCollapseEnd);
  Selection.TypeParagraph();
  Selection.TypeText('Quire exposes a Word object model. ');
  Selection.Font.Bold = true;
  Selection.TypeText('This sentence is bold.');
  assert.equal(ActiveDocument.Paragraphs.Count, 2);
  assert.equal(D.ptext(ActiveDocument.Paragraphs(1)._block), '');
  assert.equal(ActiveDocument.Paragraphs(1)._block.pPr.style, 'Heading1');
  const merged = ActiveDocument.Paragraphs(2)._block;
  assert.equal(D.ptext(merged), 'Quire exposes a Word object model. This sentence is bold.DOM Test');
  assert.equal(merged.pPr.style, 'Heading1');
  assert.equal(merged.pPr.jc, 'center');
  const runs = merged.runs.filter((r) => r.t === 'text');
  assert.ok(runs.length && runs.every((r) => r.rPr && r.rPr.color === '000080'));
  assert.equal(runs.filter((r) => r.rPr && r.rPr.b).map((r) => r.text).join(''), 'This sentence is bold.');
});

test('a heading stays when the selection is moved to its end first', () => {
  fresh();
  Documents.Add();
  const p = ActiveDocument.Paragraphs(1);
  p.Style = 'Heading 1';
  p.Alignment = wdAlignParagraphCenter;
  p.Range.Text = 'DOM Test';
  p.Range.Font.Color = wdColorDarkBlue;
  p.Range.Select();
  Selection.Collapse(wdCollapseEnd);
  Selection.TypeParagraph();
  Selection.TypeText('Quire exposes a Word object model. ');
  Selection.Font.Bold = true;
  Selection.TypeText('This sentence is bold.');
  assert.equal(D.ptext(ActiveDocument.Paragraphs(1)._block), 'DOM Test');
  assert.equal(ActiveDocument.Paragraphs(1)._block.pPr.style, 'Heading1');
  assert.equal(ActiveDocument.Paragraphs(1)._block.pPr.jc, 'center');
  assert.equal(ActiveDocument.Paragraphs(1)._block.runs.find((r) => r.t === 'text').rPr.color, '000080');
  const body = ActiveDocument.Paragraphs(2)._block;
  assert.equal(D.ptext(body), 'Quire exposes a Word object model. This sentence is bold.');
  assert.notEqual(body.pPr.style, 'Heading1');
  assert.equal(body.runs.filter((r) => r.t === 'text' && r.rPr && r.rPr.b).map((r) => r.text).join(''), 'This sentence is bold.');
});

test('Latin and East Asian faces are stored on the same run', () => {
  fresh();
  Documents.Add();
  const p = ActiveDocument.Paragraphs(1);
  p.Range.Text = 'DOM 标题';
  p.Range.Font.Name = 'Times New Roman';
  p.Range.Font.NameFarEast = '宋体';
  const run = p._block.runs.find((r) => r.t === 'text');
  assert.equal(run.rPr.font, 'Times New Roman');
  assert.equal(run.rPr.fontEA, '宋体');
  assert.equal(p.Range.Font.Name, 'Times New Roman');
  assert.equal(p.Range.Font.NameFarEast, '宋体');
  p.Range.Select();
  Selection.Collapse(wdCollapseEnd);
  Selection.TypeParagraph();
  Selection.Font.Name = 'Times New Roman';
  Selection.Font.NameFarEast = '宋体';
  Selection.TypeText('Hello 宋体');
  const typed = ActiveDocument.Paragraphs(2)._block.runs.find((r) => r.t === 'text');
  assert.equal(typed.rPr.font, 'Times New Roman');
  assert.equal(typed.rPr.fontEA, '宋体');
  assert.equal(D.ptext(p._block), 'DOM 标题');
});

test('a table inserted in a numbered paragraph continues the list through every cell', () => {
  fresh();
  Documents.Add();
  Selection.TypeText('Item');
  Selection.Range.ListFormat.ApplyNumberDefault();
  Selection.TypeParagraph();
  const table = ActiveDocument.Tables.Add(Selection.Range, 2, 2);
  const labels = D.computeLists(D.doc);
  const item = ActiveDocument.Paragraphs(1)._block;
  assert.equal(labels.get(item.id).text, '1.');
  const nums = [];
  for (const row of table._table.rows) for (const cell of row.cells) {
    const p = cell.blocks[0];
    assert.equal(p.pPr.num.id, item.pPr.num.id);
    nums.push(labels.get(p.id).text);
  }
  assert.deepEqual(nums, ['2.', '3.', '4.', '5.']);
  const after = D.doc.main.blocks[D.doc.main.blocks.length - 1];
  assert.equal(labels.get(after.id).text, '6.');
});

test('clearing the list before the table keeps the cells plain and the following paragraph numbered', () => {
  fresh();
  Documents.Add();
  Selection.TypeText('Item');
  Selection.Range.ListFormat.ApplyNumberDefault();
  Selection.TypeParagraph();
  Selection.Range.ListFormat.RemoveNumbers();
  const table = ActiveDocument.Tables.Add(Selection.Range, 2, 2);
  ActiveDocument.Paragraphs.Last.Range.ListFormat.ApplyNumberDefault();
  for (const row of table._table.rows) for (const cell of row.cells) {
    const num = cell.blocks[0].pPr.num;
    assert.ok(!num || num.id === '0');
  }
  const labels = D.computeLists(D.doc);
  assert.equal(labels.get(ActiveDocument.Paragraphs(1)._block.id).text, '1.');
  assert.equal(labels.get(ActiveDocument.Paragraphs.Last._block.id).text, '2.');
});

test('Tables.Add inserts a grid and cell text keeps the paragraph', () => {
  fresh();
  Documents.Add();
  const t = ActiveDocument.Tables.Add(Selection.Range, 2, 2);
  assert.equal(t.Rows.Count, 2);
  assert.equal(t.Columns.Count, 2);
  assert.equal(ActiveDocument.Tables.Count, 1);
  assert.equal(ActiveDocument.Tables(1)._table, ActiveDocument.Tables.Item(1)._table);
  t.Cell(1, 1).Range.Text = 'Name';
  t.Cell(1, 1).Range.Font.Name = 'Times New Roman';
  t.Cell(1, 1).Range.Font.NameFarEast = '宋体';
  assert.equal(t.Cell(1, 1).Range.Text.startsWith('Name'), true);
  assert.equal(t.Cell(1, 1)._cell.blocks[0].runs[0].rPr.font, 'Times New Roman');
  assert.equal(t.Cell(1, 1)._cell.blocks[0].runs[0].rPr.fontEA, '宋体');
  assert.equal(t.Cell(1, 1).Range.Text, 'Name\r');
  const paras = t.Cell(1, 1)._cell.blocks.filter((b) => b.t === 'p');
  assert.equal(paras.length, 1);
  assert.equal(D.ptext(paras[0]), 'Name');
  assert.throws(() => ActiveDocument.Tables.Add(Selection.Range, 2, 2, 9), (e) => e.code === 'G-ENGINE');
  const fromSelection = Selection.Tables.Add(undefined, 2, 4);
  assert.equal(fromSelection.Rows.Count, 2);
  assert.equal(fromSelection.Columns.Count, 4);
  assert.equal(ActiveDocument.Tables.Count, 2);
  assert.equal(Selection.Document, ActiveDocument);
  Selection.Bold = -1;
  assert.equal(Selection.Bold, -1);
  assert.throws(() => Selection.Range.Tables.Add(Selection.Range, 1, 1, 9), (e) => e.code === 'G-ENGINE');
});

test('a bookmark keeps the paragraph text', () => {
  fresh();
  Documents.Add();
  ActiveDocument.Paragraphs(1).Range.Text = 'Hello';
  ActiveDocument.Bookmarks.Add('Here', ActiveDocument.Paragraphs(1).Range);
  assert.equal(ActiveDocument.Bookmarks.Exists('Here'), true);
  assert.equal(ActiveDocument.Bookmarks('Here').Range.Text.includes('Hello'), true);
  assert.equal(ActiveDocument.Bookmarks.Item('Here').Name, 'Here');
  assert.equal(ActiveDocument.Paragraphs(1).Range.Bookmarks.Count, 1);
  assert.equal(typeof ActiveDocument.Bookmarks('Here').Start, 'number');
});

test('Find.Execute replaces one literal match', () => {
  fresh();
  Documents.Add();
  ActiveDocument.Paragraphs(1).Range.Text = 'Hello';
  assert.equal(Selection.Find.Execute('Hello', true, false, false, false, false, true, wdFindContinue, false, 'Hi', wdReplaceOne), true);
  assert.equal(D.ptext(ActiveDocument.Paragraphs(1)._block), 'Hi');
  assert.throws(() => Selection.Find.Execute('Hi', true, false, false, false, false, true, wdFindContinue, false, '', wdReplaceNone, true), (e) => e.code === 'G-ENGINE');
  assert.throws(() => Selection.Find.Execute('Hi', true, false, false, false, false, true, wdFindAsk), (e) => e.code === 'G-HOST');
});

test('Delete with a unit throws G-STORY', () => {
  fresh();
  Documents.Add();
  assert.throws(() => Selection.Delete(1, 1), (e) => e.code === 'G-STORY');
});

test('paragraph format, tabs and list level write the paragraph', () => {
  fresh();
  Documents.Add();
  const p = ActiveDocument.Paragraphs(1);
  p.Range.Text = 'Item';
  p.LeftIndent = 36;
  p.SpaceAfter = 8;
  p.Format.LineSpacingRule = wdLineSpaceDouble;
  p.Hyphenation = false;
  p.DisableLineHeightGrid = true;
  p.ReadingOrder = wdReadingOrderRtl;
  p.OutlineLevel = 1;
  p.KeepWithNext = true;
  p.TabStops.Add(72, 0, 1);
  assert.equal(p._block.pPr.ind.l, 36);
  assert.equal(p._block.pPr.sp.a, 8);
  assert.equal(p._block.pPr.sp.line, 2);
  assert.equal(p.Format.LineSpacingRule, wdLineSpaceDouble);
  assert.equal(p._block.pPr.noHyphen, true);
  assert.equal(p._block.pPr.noSnap, true);
  assert.equal(p._block.pPr.bidi, true);
  assert.equal(p._block.pPr.outline, 0);
  assert.equal(p._block.pPr.keepNext, true);
  assert.equal(p._block.pPr.tabs[0].pos, 72);
  assert.equal(p._block.pPr.tabs[0].leader, 'dot');
  assert.throws(() => { p.Format.CharacterUnitLeftIndent = 1; }, (e) => e.code === 'G-ENGINE');
  p.Range.ListFormat.ApplyNumberDefault();
  assert.equal(p._block.pPr.num.id, '2');
  p.Range.ListFormat.ListLevelNumber = 2;
  assert.equal(p._block.pPr.num.lvl, 1);
  assert.equal(p.Range.Characters.Count, 4);
  assert.equal(p.Range.Paragraphs.Count, 1);
  assert.equal(p.Range.Paragraphs(1)._block, p._block);
  assert.equal(ActiveDocument.Words.Count, 1);
});

test('document session, styles, page setup and statistics', () => {
  fresh();
  Documents.Add();
  ActiveDocument.Paragraphs(1).Range.Text = 'Hello';
  assert.equal(ActiveDocument.Saved, false);
  ActiveDocument.Saved = true;
  assert.equal(D.doc.dirty, false);
  assert.equal(D.ptext(ActiveDocument.Paragraphs(1)._block), 'Hello');
  assert.equal(ActiveDocument.Content.Text, 'Hello\r');
  assert.equal(ActiveDocument.ComputeStatistics(wdStatisticWords), 1);
  assert.equal(ActiveDocument.ComputeStatistics(wdStatisticParagraphs), 1);
  assert.equal(ActiveDocument.ComputeStatistics(wdStatisticCharacters), 5);
  assert.throws(() => ActiveDocument.ComputeStatistics(wdStatisticLines), (e) => e.code === 'G-ENGINE');
  assert.throws(() => ActiveDocument.ComputeStatistics(wdStatisticPages), (e) => e.code === 'G-ENGINE');
  ActiveDocument.BuiltInDocumentProperties('Title').Value = 'DOM Test';
  assert.equal(D.doc.props.title, 'DOM Test');
  assert.equal(ActiveDocument.BuiltInDocumentProperties('Title').Value, 'DOM Test');
  ActiveDocument.CustomDocumentProperties.Add('Key', false, wdPropertyTypeString, 'Value');
  assert.equal(D.doc.custom.Key, 'Value');
  assert.throws(() => ActiveDocument.CustomDocumentProperties.Add('N', false, 1, 5), (e) => e.code === 'G-ENGINE');
  const st = ActiveDocument.Styles.Add('SampleStyle', wdStyleTypeParagraph);
  st.Font.Bold = true;
  assert.equal(D.doc.styles[st._id].rPr.b, true);
  ActiveDocument.Paragraphs(1).Style = 'SampleStyle';
  assert.equal(ActiveDocument.Paragraphs(1)._block.pPr.style, st._id);
  assert.equal(ActiveDocument.Styles('Heading 1').NameLocal, 'Heading 1');
  st.Delete();
  assert.equal(D.doc.styles[st._id], undefined);
  assert.equal(ActiveDocument.Sections.Count, 1);
  const ps = ActiveDocument.Sections(1).PageSetup;
  ps.Orientation = wdOrientLandscape;
  assert.equal(ps.Orientation, wdOrientLandscape);
  assert.equal(D.doc.sect.orient, 'landscape');
  assert.equal(D.doc.sect.pgW, 792);
  assert.throws(() => { ps.PaperSize = 1; }, (e) => e.code === 'G-ENGINE');
  assert.equal(ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Exists, 0);
  assert.throws(() => { ActiveDocument.Sections(1).Headers(1).LinkToPrevious = true; }, (e) => e.code === 'G-ENGINE');
  ActiveDocument.TrackRevisions = true;
  assert.equal(D.doc.settings.track, true);
  assert.throws(() => ActiveDocument.FullName, (e) => e.code === 'G-HOST');
  assert.throws(() => ActiveDocument.Paragraphs(1).Range.Font.Animation, (e) => e.code === 'G-ENGINE');
  assert.throws(() => ActiveDocument.Shapes.AddShape(1, 0, 0, 72, 72), (e) => e.code === 'G-ENGINE');
  assert.throws(() => ActiveDocument.InlineShapes.AddPicture('x'), (e) => e.code === 'G-HOST');
});

test('case, language, break, copy and fields', () => {
  fresh();
  Documents.Add();
  ActiveDocument.Paragraphs(1).Range.Text = 'Hello';
  ActiveDocument.Paragraphs(1).Range.Case = wdLowerCase;
  assert.equal(D.ptext(ActiveDocument.Paragraphs(1)._block), 'hello');
  ActiveDocument.Paragraphs(1).Range.LanguageID = wdEnglishUS;
  const run = ActiveDocument.Paragraphs(1)._block.runs.find((r) => r.t === 'text');
  assert.equal(run.rPr.lang, 'en-US');
  ActiveDocument.Paragraphs(1).Range.Select();
  Selection.Collapse(wdCollapseEnd);
  Selection.InsertBreak(wdLineBreak);
  Selection.TypeText('B');
  assert.equal(ActiveDocument.Paragraphs(1)._block.runs.some((r) => r.t === 'br'), true);
  ActiveDocument.Paragraphs(1).Range.Copy();
  Selection.Paste();
  assert.equal(D.ptext(ActiveDocument.Paragraphs(1)._block).includes('hello'), true);
  const field = ActiveDocument.Fields.Add(Selection.Range, -1, 'DATE');
  assert.equal(ActiveDocument.Fields.Count, 1);
  assert.equal(field.Code.Text.includes('DATE'), true);
  assert.throws(() => field.DoClick(), (e) => e.code === 'G-ENGINE');
});

test('hyperlinks, comments, footnotes and tables', () => {
  fresh();
  Documents.Add();
  ActiveDocument.Paragraphs(1).Range.Text = 'Hello';
  const link = ActiveDocument.Hyperlinks.Add(ActiveDocument.Paragraphs(1).Range, 'https://vibeoffice.work');
  assert.equal(link.Address, 'https://vibeoffice.work');
  assert.equal(ActiveDocument.Paragraphs(1)._block.runs.some((r) => r.rPr && r.rPr.link && r.rPr.link.url === 'https://vibeoffice.work'), true);
  assert.throws(() => link.Follow(), (e) => e.code === 'G-HOST');
  const note = ActiveDocument.Comments.Add(ActiveDocument.Paragraphs(1).Range, 'Note');
  assert.equal(ActiveDocument.Comments.Count, 1);
  assert.equal(note.Range.Text.includes('Note'), true);
  assert.throws(() => note.Edit(), (e) => e.code === 'G-ENGINE');
  const fn = ActiveDocument.Footnotes.Add(ActiveDocument.Paragraphs(1).Range, '', 'A note');
  assert.equal(ActiveDocument.Footnotes.Count, 1);
  assert.equal(fn.Range.Text.includes('A note'), true);
  assert.equal(typeof fn.Reference.Start, 'number');
  const t = ActiveDocument.Tables.Add(Selection.Range, 2, 2);
  t.Cell(1, 1).Range.Text = 'B';
  t.Cell(2, 1).Range.Text = 'A';
  t.Rows.Add();
  assert.equal(t.Rows.Count, 3);
  t.Columns.Add();
  assert.equal(t.Columns.Count, 3);
  t.Cell(1, 1).Shading.BackgroundPatternColor = wdColorYellow;
  assert.equal(t.Cell(1, 1)._cell.tcPr.shd.fill, 'FFFF00');
  t.Rows(1).HeadingFormat = true;
  assert.equal(t._table.rows[0].trPr.header, true);
  t.Sort();
  assert.equal(t.Cell(1, 1).Range.Text.startsWith('A') || t.Cell(2, 1).Range.Text.startsWith('A'), true);
  t.AutoFitBehavior(wdAutoFitWindow);
  assert.equal(t._table.tblPr.layout, 'autofit');
  assert.equal(typeof t.Uniform, 'number');
});

test('a document that is not active allows Name and Activate only', () => {
  fresh();
  const first = Documents.Add();
  const name = first.Name;
  Documents.Add();
  assert.equal(first.Name, name);
  assert.throws(() => first.Paragraphs.Count, (e) => e.code === 'inactive');
  first.Activate();
  assert.equal(ActiveDocument, first);
  assert.equal(first.Paragraphs.Count, 1);
});

test('AddPicture embeds at a range and floats from Shapes', async () => {
  fresh();
  Documents.Add();
  ActiveDocument.Paragraphs(1).Range.Text = 'AAA';
  ActiveDocument.Paragraphs(1).Range.InsertParagraphAfter();
  ActiveDocument.Paragraphs(2).Range.Text = 'BBB';
  const prevAudit = L.hostAudit;
  const prevApp = L.app;
  L.hostAudit = {
    has(name) { return name === 'Picture'; },
    readBytes() { return Promise.resolve(new Uint8Array([1, 2, 3])); },
  };
  L.app = {
    insertPictureFiles(files, opts) {
      const it = D.item('img', { media: 'm', w: 10, h: 8, name: files[0].name, alt: '' });
      if (opts && opts.prepare) opts.prepare(it);
      D.tx('Insert Picture', () => L.O.insertItem(opts.pos, it));
      return [it];
    },
  };
  try {
    assert.throws(() => ActiveDocument.InlineShapes.AddPicture('a.png', true), (e) => /LinkToFile/.test(e.message));
    assert.throws(() => ActiveDocument.InlineShapes.AddPicture('a.png', false, 0), (e) => /SaveWithDocument/.test(e.message));
    const inline = await ActiveDocument.Paragraphs(2).Range.InlineShapes.AddPicture('pic.png');
    assert.equal(inline.Type, 3);
    assert.equal(inline.WrapFormat.Type, 7);
    assert.equal(ActiveDocument.InlineShapes.Count, 1);
    assert.equal(ActiveDocument.Shapes.Count, 0);
    assert.equal(ActiveDocument.Paragraphs(2)._block.runs.some((it) => it.t === 'img'), true);
    assert.equal(ActiveDocument.Paragraphs(1)._block.runs.some((it) => it.t === 'img'), false);
    inline.ConvertToShape();
    assert.equal(inline.WrapFormat.Type, 0);
    assert.equal(ActiveDocument.Shapes.Count, 1);
    assert.equal(ActiveDocument.InlineShapes.Count, 0);
    inline.ConvertToInlineShape();
    assert.equal(ActiveDocument.InlineShapes.Count, 1);
    const floated = await ActiveDocument.Shapes.AddPicture('float.png', false, true, 72, 36, 40, 20);
    assert.equal(floated.Left, 72);
    assert.equal(floated.Top, 36);
    assert.equal(floated.Width, 40);
    assert.equal(floated.Height, 20);
    assert.equal(floated.WrapFormat.Type, 0);
    const again = await ActiveDocument.Paragraphs(1).Range.Shapes.AddPicture2('c.png', false, true, 12, 4, 30, 15);
    assert.equal(again.Left, 12);
    assert.equal(again.Width, 30);
    assert.equal(ActiveDocument.Paragraphs(1)._block.runs.some((it) => it.t === 'img' && it.float), true);
    Selection.SetRange(ActiveDocument.Paragraphs(2).Range.Start, ActiveDocument.Paragraphs(2).Range.Start);
    await Selection.InlineShapes.AddPicture('sel.png');
    assert.equal(ActiveDocument.Paragraphs(2)._block.runs.filter((it) => it.t === 'img' && !it.float).length, 2);
  } finally {
    L.hostAudit = prevAudit;
    L.app = prevApp;
  }
});
