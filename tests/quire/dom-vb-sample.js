/*
Run this in Quire. It builds the same document as tests/quire/dom-sample.js.
Each document call is a VB statement on the Immediate host (L.vb).

  1. Open Quire. The engine must already be built (make).
  2. Either press F12, or open Help → Quire DOM Console.
  3. Paste this file.
  4. Run: await buildSampleVb()

The console still evaluates JavaScript. This function sends the VB.
 */
async function buildSampleVb() {
  if (buildSampleVb.host) {
    try { buildSampleVb.host.close(); } catch (e) { /* the previous run already stopped */ }
  }
  const host = await L.vb.open();
  buildSampleVb.host = host;
  const vb = (text) => host.eval(text);
  const read = (text) => host.eval('?' + text);

  await vb('Documents.Add');
  await vb('ActiveDocument.BuiltInDocumentProperties("Title").Value = "DOM Test"');

  await vb('ActiveDocument.Paragraphs(1).Style = "Heading 1"');
  await vb('ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphCenter');
  await vb('ActiveDocument.Paragraphs(1).Range.Text = "DOM Test 标题"');
  await vb('ActiveDocument.Paragraphs(1).Range.Font.Color = wdColorDarkBlue');
  await vb('ActiveDocument.Paragraphs(1).Range.Font.Name = "Times New Roman"');
  await vb('ActiveDocument.Paragraphs(1).Range.Font.NameFarEast = "宋体"');

  await vb('ActiveDocument.Paragraphs(1).Range.Select');
  await vb('Selection.Collapse(wdCollapseEnd)');
  await vb('Selection.TypeParagraph');
  await vb('Selection.Font.Name = "Times New Roman"');
  await vb('Selection.Font.NameFarEast = "宋体"');
  await vb('Selection.TypeText("Quire exposes a Word object model. 中文用宋体。")');
  await vb('Selection.Font.Bold = True');
  await vb('Selection.TypeText("This sentence is bold.")');
  await vb('Selection.Font.Bold = False');
  await vb('Selection.TypeParagraph');

  for (let i = 1; i <= 10; i++) {
    await vb('Selection.TypeText("Hello world ' + i + '")');
    await vb('Selection.TypeParagraph');
  }

  await vb('ActiveDocument.Paragraphs(2).LeftIndent = 18');
  await vb('ActiveDocument.Paragraphs(2).SpaceAfter = 8');
  await vb('ActiveDocument.Paragraphs(2).Range.Hyperlinks.Add(ActiveDocument.Paragraphs(2).Range, "https://vibeoffice.work", "", "VibeOffice")');

  await vb('Selection.TypeText("A numbered item from ListFormat")');
  await vb('Selection.Range.ListFormat.ApplyNumberDefault');
  await vb('Selection.TypeParagraph');
  await vb('Selection.Range.ListFormat.RemoveNumbers');

  await vb('ActiveDocument.Tables.Add(Selection.Range, 2, 2)');
  await vb('ActiveDocument.Paragraphs.Last.Range.ListFormat.ApplyNumberDefault');
  const table = 'ActiveDocument.Tables.Last';
  await vb(table + '.Cell(1, 1).Range.Text = "Name"');
  await vb(table + '.Cell(1, 2).Range.Text = "Value"');
  await vb(table + '.Cell(2, 1).Range.Text = "DOM"');
  await vb(table + '.Cell(2, 2).Range.Text = "Word object model 对象模型"');
  await vb(table + '.Rows(1).HeadingFormat = True');
  await vb(table + '.Cell(1, 1).Shading.BackgroundPatternColor = wdColorYellow');
  await vb(table + '.Cell(1, 2).Shading.BackgroundPatternColor = wdColorYellow');
  for (let r = 1; r <= 2; r++) {
    for (let c = 1; c <= 2; c++) {
      await vb(table + '.Cell(' + r + ', ' + c + ').Range.Font.Name = "Times New Roman"');
      await vb(table + '.Cell(' + r + ', ' + c + ').Range.Font.NameFarEast = "宋体"');
    }
  }

  await vb('ActiveDocument.Bookmarks.Add("Sample", ActiveDocument.Paragraphs(1).Range)');
  const found = await read('Selection.Find.Execute("Hello world 1", True, False, False, False, False, True, wdFindContinue, False, "Hi world 1", wdReplaceOne)');
  await vb('ActiveDocument.Comments.Add(ActiveDocument.Paragraphs(1).Range, "Sample note")');

  const report = {
    paragraphs: await read('ActiveDocument.Paragraphs.Count'),
    title: await read('ActiveDocument.Paragraphs(1).Range.Text'),
    documentTitle: await read('ActiveDocument.BuiltInDocumentProperties("Title").Value'),
    words: await read('ActiveDocument.ComputeStatistics(wdStatisticWords)'),
    indent: await read('ActiveDocument.Paragraphs(2).LeftIndent'),
    rows: await read(table + '.Rows.Count'),
    columns: await read(table + '.Columns.Count'),
    name: await read(table + '.Cell(1, 1).Range.Text'),
    value: await read(table + '.Cell(2, 2).Range.Text'),
    bookmark: await read('ActiveDocument.Bookmarks("Sample").Name'),
    bookmarkText: await read('ActiveDocument.Bookmarks("Sample").Range.Text'),
    comments: await read('ActiveDocument.Comments.Count'),
    found: found,
  };
  console.log(report);
  return report;
}
