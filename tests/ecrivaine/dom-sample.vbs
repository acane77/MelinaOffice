' Same document as tests/ecrivaine/dom-sample.js and tests/ecrivaine/dom-sample-runvb.js.
' Paste this into Help → Ecrivaine DOM Console → Run Script, then click Run.
' The engine must already be built (make). Run starts Sub Main.
' Documents.Add opens a new document.

Sub Main()
  Dim title As Object
  Dim body As Object
  Dim table As Object
  Dim sample As Object
  Dim found
  Dim i As Long
  Dim r As Long
  Dim c As Long
  Documents.Add
  ActiveDocument.BuiltInDocumentProperties("Title").Value = "DOM Test"
  Set title = ActiveDocument.Paragraphs(1)
  title.Style = "Heading 1"
  title.Alignment = wdAlignParagraphCenter
  title.Range.Text = "DOM Test 标题"
  title.Range.Font.Color = wdColorDarkBlue
  title.Range.Font.Name = "Times New Roman"
  title.Range.Font.NameFarEast = "宋体"
  title.Range.Select
  Selection.Collapse wdCollapseEnd
  Selection.TypeParagraph
  Selection.Font.Name = "Times New Roman"
  Selection.Font.NameFarEast = "宋体"
  Selection.TypeText "Ecrivaine exposes a Word object model. 中文用宋体。"
  Selection.Font.Bold = True
  Selection.TypeText "This sentence is bold."
  Selection.Font.Bold = False
  Selection.TypeParagraph
  For i = 1 To 10
    Selection.TypeText "Hello world " & i
    Selection.TypeParagraph
  Next
  Set body = ActiveDocument.Paragraphs(2)
  body.LeftIndent = 18
  body.SpaceAfter = 8
  body.Range.Hyperlinks.Add body.Range, "https://acane.net/melinaoffice", "", "Melina Office"
  Selection.TypeText "A numbered item from ListFormat"
  Selection.Range.ListFormat.ApplyNumberDefault
  Selection.TypeParagraph
  Selection.Range.ListFormat.RemoveNumbers
  Set table = ActiveDocument.Tables.Add(Selection.Range, 2, 2)
  ActiveDocument.Paragraphs.Last.Range.ListFormat.ApplyNumberDefault
  table.Cell(1, 1).Range.Text = "Name"
  table.Cell(1, 2).Range.Text = "Value"
  table.Cell(2, 1).Range.Text = "DOM"
  table.Cell(2, 2).Range.Text = "Word object model 对象模型"
  table.Rows(1).HeadingFormat = True
  table.Cell(1, 1).Shading.BackgroundPatternColor = wdColorYellow
  table.Cell(1, 2).Shading.BackgroundPatternColor = wdColorYellow
  For r = 1 To 2
    For c = 1 To 2
      table.Cell(r, c).Range.Font.Name = "Times New Roman"
      table.Cell(r, c).Range.Font.NameFarEast = "宋体"
    Next
  Next
  Set sample = ActiveDocument.Bookmarks.Add("Sample", ActiveDocument.Paragraphs(1).Range)
  found = Selection.Find.Execute("Hello world 1", True, False, False, False, False, True, wdFindContinue, False, "Hi world 1", wdReplaceOne)
  ActiveDocument.Comments.Add ActiveDocument.Paragraphs(1).Range, "Sample note"
  Debug.Print "paragraphs: " & ActiveDocument.Paragraphs.Count
  Debug.Print "title: " & ActiveDocument.Paragraphs(1).Range.Text
  Debug.Print "documentTitle: " & ActiveDocument.BuiltInDocumentProperties("Title").Value
  Debug.Print "words: " & ActiveDocument.ComputeStatistics(wdStatisticWords)
  Debug.Print "indent: " & body.LeftIndent
  Debug.Print "rows: " & table.Rows.Count
  Debug.Print "columns: " & table.Columns.Count
  Debug.Print "name: " & table.Cell(1, 1).Range.Text
  Debug.Print "value: " & table.Cell(2, 2).Range.Text
  Debug.Print "bookmark: " & sample.Name
  Debug.Print "bookmarkText: " & sample.Range.Text
  Debug.Print "comments: " & ActiveDocument.Comments.Count
  Debug.Print "found: " & found
End Sub
