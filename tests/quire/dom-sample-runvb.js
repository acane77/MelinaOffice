/*
Run this in Quire. It builds the same document as tests/quire/dom-sample.js.
The steps run as one compiled standard module (runVB6), including the loops.

  1. Open Quire. The engine must already be built (make).
  2. Press F12, or open Help → Quire DOM Console.
  3. Paste this file.
  4. Run: await buildSampleRunVb()

The console still evaluates JavaScript. This function sends the VB.
 */
async function buildSampleRunVb() {
  const found = await runVB6([
    'Function Build()',
    '  Dim title As Object',
    '  Dim body As Object',
    '  Dim table As Object',
    '  Dim sample As Object',
    '  Dim i As Long',
    '  Dim r As Long',
    '  Dim c As Long',
    '  Documents.Add',
    '  ActiveDocument.BuiltInDocumentProperties("Title").Value = "DOM Test"',
    '  Set title = ActiveDocument.Paragraphs(1)',
    '  title.Style = "Heading 1"',
    '  title.Alignment = wdAlignParagraphCenter',
    '  title.Range.Text = "DOM Test 标题"',
    '  title.Range.Font.Color = wdColorDarkBlue',
    '  title.Range.Font.Name = "Times New Roman"',
    '  title.Range.Font.NameFarEast = "宋体"',
    '  title.Range.Select',
    '  Selection.Collapse wdCollapseEnd',
    '  Selection.TypeParagraph',
    '  Selection.Font.Name = "Times New Roman"',
    '  Selection.Font.NameFarEast = "宋体"',
    '  Selection.TypeText "Quire exposes a Word object model. 中文用宋体。"',
    '  Selection.Font.Bold = True',
    '  Selection.TypeText "This sentence is bold."',
    '  Selection.Font.Bold = False',
    '  Selection.TypeParagraph',
    '  For i = 1 To 10',
    '    Selection.TypeText "Hello world " & i',
    '    Selection.TypeParagraph',
    '  Next',
    '  Set body = ActiveDocument.Paragraphs(2)',
    '  body.LeftIndent = 18',
    '  body.SpaceAfter = 8',
    '  body.Range.Hyperlinks.Add body.Range, "https://vibeoffice.work", "", "VibeOffice"',
    '  Selection.TypeText "A numbered item from ListFormat"',
    '  Selection.Range.ListFormat.ApplyNumberDefault',
    '  Selection.TypeParagraph',
    '  Selection.Range.ListFormat.RemoveNumbers',
    '  Set table = ActiveDocument.Tables.Add(Selection.Range, 2, 2)',
    '  ActiveDocument.Paragraphs.Last.Range.ListFormat.ApplyNumberDefault',
    '  table.Cell(1, 1).Range.Text = "Name"',
    '  table.Cell(1, 2).Range.Text = "Value"',
    '  table.Cell(2, 1).Range.Text = "DOM"',
    '  table.Cell(2, 2).Range.Text = "Word object model 对象模型"',
    '  table.Rows(1).HeadingFormat = True',
    '  table.Cell(1, 1).Shading.BackgroundPatternColor = wdColorYellow',
    '  table.Cell(1, 2).Shading.BackgroundPatternColor = wdColorYellow',
    '  For r = 1 To 2',
    '    For c = 1 To 2',
    '      table.Cell(r, c).Range.Font.Name = "Times New Roman"',
    '      table.Cell(r, c).Range.Font.NameFarEast = "宋体"',
    '    Next',
    '  Next',
    '  Set sample = ActiveDocument.Bookmarks.Add("Sample", ActiveDocument.Paragraphs(1).Range)',
    '  Build = Selection.Find.Execute("Hello world 1", True, False, False, False, False, True, wdFindContinue, False, "Hi world 1", wdReplaceOne)',
    '  ActiveDocument.Comments.Add ActiveDocument.Paragraphs(1).Range, "Sample note"',
    'End Function',
  ].join('\n'), 'Build');

  const title = ActiveDocument.Paragraphs(1);
  const body = ActiveDocument.Paragraphs(2);
  const table = ActiveDocument.Tables(1);
  const sample = ActiveDocument.Bookmarks('Sample');
  const report = {
    paragraphs: ActiveDocument.Paragraphs.Count,
    title: title.Range.Text,
    documentTitle: ActiveDocument.BuiltInDocumentProperties('Title').Value,
    words: ActiveDocument.ComputeStatistics(wdStatisticWords),
    indent: body.LeftIndent,
    rows: table.Rows.Count,
    columns: table.Columns.Count,
    name: table.Cell(1, 1).Range.Text,
    value: table.Cell(2, 2).Range.Text,
    bookmark: sample.Name,
    bookmarkText: sample.Range.Text,
    comments: ActiveDocument.Comments.Count,
    found: !!found,
  };
  console.log(report);
  return report;
}
