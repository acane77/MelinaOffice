/* 
Run this code in Quire:
   1. Open Quire App
   2. Open: Help -> Quire DOM Console
   3. Paste the following code into the console
   4. Run "await buildSample()"
See how document changes!
 */
async function buildSample() {
  await Documents.Add();

  ActiveDocument.BuiltInDocumentProperties('Title').Value = 'DOM Test';

  const title = ActiveDocument.Paragraphs(1);
  title.Style = 'Heading 1';
  title.Alignment = wdAlignParagraphCenter;
  title.Range.Text = 'DOM Test 标题';
  title.Range.Font.Color = wdColorDarkBlue;
  title.Range.Font.Name = 'Times New Roman';
  title.Range.Font.NameFarEast = '宋体';

  title.Range.Select();
  Selection.Collapse(wdCollapseEnd);
  Selection.TypeParagraph();
  Selection.Font.Name = 'Times New Roman';
  Selection.Font.NameFarEast = '宋体';
  Selection.TypeText('Quire exposes a Word object model. 中文用宋体。');
  Selection.Font.Bold = true;
  Selection.TypeText('This sentence is bold.');
  Selection.Font.Bold = false;
  Selection.TypeParagraph();

  for (let i = 1; i <= 10; i++) {
    Selection.TypeText('Hello world ' + i);
    Selection.TypeParagraph();
  }

  const body = ActiveDocument.Paragraphs(2);
  body.LeftIndent = 18;
  body.SpaceAfter = 8;
  body.Range.Hyperlinks.Add(body.Range, 'https://vibeoffice.work', '', 'VibeOffice');

  Selection.TypeText('A numbered item from ListFormat');
  Selection.Range.ListFormat.ApplyNumberDefault();
  Selection.TypeParagraph();
  // A numbered anchor would number every cell. Clear it for the insert, then number the paragraph that follows the table.
  Selection.Range.ListFormat.RemoveNumbers();
  const table = ActiveDocument.Tables.Add(Selection.Range, 2, 2);
  ActiveDocument.Paragraphs.Last.Range.ListFormat.ApplyNumberDefault();
  table.Cell(1, 1).Range.Text = 'Name';
  table.Cell(1, 2).Range.Text = 'Value';
  table.Cell(2, 1).Range.Text = 'DOM';
  table.Cell(2, 2).Range.Text = 'Word object model 对象模型';
  table.Rows(1).HeadingFormat = true;
  table.Cell(1, 1).Shading.BackgroundPatternColor = wdColorYellow;
  table.Cell(1, 2).Shading.BackgroundPatternColor = wdColorYellow;
  for (let r = 1; r <= 2; r++) {
    for (let c = 1; c <= 2; c++) {
      const cell = table.Cell(r, c).Range;
      cell.Font.Name = 'Times New Roman';
      cell.Font.NameFarEast = '宋体';
    }
  }

  const sample = ActiveDocument.Bookmarks.Add('Sample', ActiveDocument.Paragraphs(1).Range);
  const found = Selection.Find.Execute('Hello world 1', true, false, false, false, false, true, wdFindContinue, false, 'Hi world 1', wdReplaceOne);
  ActiveDocument.Comments.Add(ActiveDocument.Paragraphs(1).Range, 'Sample note');

  console.log({
    paragraphs: ActiveDocument.Paragraphs.Count,
    title: ActiveDocument.Paragraphs(1).Range.Text,
    documentTitle: ActiveDocument.BuiltInDocumentProperties('Title').Value,
    words: ActiveDocument.ComputeStatistics(wdStatisticWords),
    indent: body.LeftIndent,
    rows: table.Rows.Count,
    columns: table.Columns.Count,
    name: table.Cell(1, 1).Range.Text,
    value: table.Cell(2, 2).Range.Text,
    bookmark: sample.Name,
    bookmarkText: ActiveDocument.Bookmarks('Sample').Range.Text,
    comments: ActiveDocument.Comments.Count,
    found,
  });
}
