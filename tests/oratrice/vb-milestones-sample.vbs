' Run Script equivalent of tests/oratrice/vb-milestones-sample.js.
' Paste this file into Oratrice DOM Console → Run Script, then Run.
' Each check opens a new presentation. The last one stays open and shows the
' title, the text box, the table and the link. Presentations that this script
' opened for the checks are closed. The presentation that was already open stays.
' Save and Presentations.Open stay in the Node tests: in the browser they open a dialog.
' The Node test checks every pp* and mso* constant. This script checks the ones it uses.

Dim failures As String
Dim started As Long

Sub Fail(message As String)
  If failures <> "" Then failures = failures & vbCrLf
  failures = failures & message
End Sub

Sub Same(label As String, actual, expected)
  If actual <> expected Then Fail label & " = " & CStr(actual) & ", expected " & CStr(expected)
End Sub

Sub ExpectCode(label As String, code As String)
  If Err.Number = 0 Then
    Fail label & " did not fail"
  ElseIf InStr(Err.Description, "(" & code & ")") = 0 Then
    Fail label & " -> " & Err.Description
  End If
  Err.Clear
End Sub

Sub Blank()
  Presentations.Add
End Sub

Sub CloseAdded(keepFinal As String)
  Dim i As Long
  Dim guard As Long
  guard = 0
  i = Presentations.Count
  Do While i > started And guard < 30
    guard = guard + 1
    If Presentations(i).Name = keepFinal Then
      i = i - 1
    Else
      Presentations(i).Close
      i = Presentations.Count
    End If
  Loop
End Sub

Sub GapObject(obj As Object, label As String, code As String)
  If Not obj.Parent Is Application Then Fail label & ".Parent"
  On Error Resume Next
  Dim x
  Err.Clear
  x = obj.Creator
  ExpectCode label & ".Creator", "G-ENGINE"
  Err.Clear
  x = obj.Count
  ExpectCode label & ".Count", code
  On Error GoTo 0
End Sub

Sub CheckApplication()
  Blank
  Same "ppLayoutBlank", ppLayoutBlank, 12
  Same "ppSelectionShapes", ppSelectionShapes, 2
  Same "msoTrue", msoTrue, -1
  Same "msoFalse", msoFalse, 0
  Same "ppAlignLeft", ppAlignLeft, 1
  Same "msoShapeRectangle", msoShapeRectangle, 1
  Same "msoShapeOval", msoShapeOval, 9
  Same "ppEffectFadeSmoothly", ppEffectFadeSmoothly, 3849
  If Not ActivePresentation.Parent Is Presentations Then Fail "Parent"
  If Not ActivePresentation.Application Is Application Then Fail "Application"
  Same "Presentations.Count", Presentations.Count, started + 1
  Dim name As String
  name = ActivePresentation.Name
  On Error Resume Next
  ActivePresentation.Name = "Z"
  If Err.Number = 0 Then Fail "Name assignment did not fail"
  Err.Clear
  On Error GoTo 0
  If ActivePresentation.Name <> name Then Fail "Name changed"
  ActivePresentation.Saved = False
  If ActivePresentation.Saved Then Fail "Saved stayed True"
  ActivePresentation.Saved = True
  If Not ActivePresentation.Saved Then Fail "Saved stayed False"
  ActivePresentation.BuiltInDocumentProperties("Title").Value = "M1"
  Same "Title", ActivePresentation.BuiltInDocumentProperties("Title").Value, "M1"
  ActivePresentation.PageSetup.SlideWidth = 720
  Same "SlideWidth", ActivePresentation.PageSetup.SlideWidth, 720
  On Error Resume Next
  Dim x
  Err.Clear
  x = ActivePresentation.Creator
  ExpectCode "Creator", "G-ENGINE"
  Err.Clear
  x = ActivePresentation.FullName
  ExpectCode "FullName", "G-HOST"
  Err.Clear
  Presentations.Open "x"
  ExpectCode "Open", "G-HOST"
  Err.Clear
  ActivePresentation.SaveAs "x"
  ExpectCode "SaveAs", "G-HOST"
  Err.Clear
  ActivePresentation.PrintOut
  ExpectCode "PrintOut", "G-HOST"
  Err.Clear
  x = ActiveWindow.View
  ExpectCode "View", "G-HOST"
  On Error GoTo 0
End Sub

Sub CheckSlides()
  Blank
  ActivePresentation.Slides.Add 2, ppLayoutBlank
  Same "slides", ActivePresentation.Slides.Count, 2
  ActivePresentation.Slides(2).Shapes.AddShape msoShapeRectangle, 10, 20, 40, 30
  ActivePresentation.Slides(2).Shapes(1).Left = 15
  Same "Left", ActivePresentation.Slides(2).Shapes(1).Left, 15
  ActivePresentation.Slides(2).Shapes(1).Select
  Same "selection", ActiveWindow.Selection.Type, ppSelectionShapes
  ActiveWindow.Selection.ShapeRange.Left = 18
  Same "ShapeRange.Left", ActivePresentation.Slides(2).Shapes(1).Left, 18
End Sub

Sub CheckText()
  Blank
  Dim r As Object
  Set r = ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange
  r.Text = "Hello"
  r.Font.Bold = msoTrue
  r.ParagraphFormat.Alignment = ppAlignLeft
  Same "text", r.Text, "Hello"
  Same "bold", r.Font.Bold, -1
  Same "count", r.Characters.Count, 5
  On Error Resume Next
  Dim x
  Err.Clear
  x = r.Characters(1)
  ExpectCode "Characters", "G-TEXT"
  Err.Clear
  x = r.Start
  ExpectCode "Start", "G-TEXT"
  Err.Clear
  r.Find "H"
  ExpectCode "Find", "G-TEXT"
  On Error GoTo 0
End Sub

Sub CheckTable()
  Blank
  ActivePresentation.Slides.Add 2, ppLayoutBlank
  ActivePresentation.Slides(2).Shapes.AddTable 2, 2, 20, 80, 200
  ActivePresentation.Slides(2).Shapes(1).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "A1"
  Same "cell", ActivePresentation.Slides(2).Shapes(1).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text, "A1"
  ActivePresentation.Slides(2).Shapes.AddShape msoShapeOval, 10, 10, 40, 40
  ActivePresentation.Slides(2).Shapes(2).Fill.ForeColor.RGB = 255
  Same "fill", ActivePresentation.Slides(2).Shapes(2).Fill.ForeColor.RGB, 255
  Dim r As Object
  Set r = ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange
  r.Text = "Site"
  r.Hyperlinks.Add Nothing, "https://acane.net/melinaoffice"
  Same "link", r.Hyperlinks(1).Address, "https://acane.net/melinaoffice"
  ActivePresentation.Slides(2).SlideShowTransition.EntryEffect = ppEffectFadeSmoothly
  Same "effect", ActivePresentation.Slides(2).SlideShowTransition.EntryEffect, ppEffectFadeSmoothly
  ActivePresentation.SlideShowSettings.Run
  ActivePresentation.SlideShowWindows(1).View.Next
  Same "show", ActivePresentation.SlideShowWindows(1).View.CurrentShowPosition, 2
  ActivePresentation.SlideShowWindows(1).View.Exit
  On Error Resume Next
  Err.Clear
  ActivePresentation.Slides(2).Shapes.AddPicture "x", msoFalse, msoFalse, 0, 0
  ExpectCode "AddPicture", "G-HOST"
  Err.Clear
  Dim x
  x = ActivePresentation.Slides(1).TimeLine
  ExpectCode "TimeLine", "G-ENGINE"
  On Error GoTo 0
End Sub

Sub CheckGaps()
  GapObject TimeLine, "TimeLine", "G-ENGINE"
  GapObject Sequence, "Sequence", "G-ENGINE"
  GapObject Effect, "Effect", "G-ENGINE"
  GapObject Chart, "Chart", "G-ENGINE"
  GapObject ChartData, "ChartData", "G-ENGINE"
  GapObject Diagram, "Diagram", "G-ENGINE"
  GapObject DiagramNode, "DiagramNode", "G-ENGINE"
  GapObject SmartArt, "SmartArt", "G-ENGINE"
  GapObject OLEFormat, "OLEFormat", "G-ENGINE"
  GapObject MediaFormat, "MediaFormat", "G-ENGINE"
  GapObject CustomerData, "CustomerData", "G-ENGINE"
  GapObject NotesPage, "NotesPage", "G-ENGINE"
  GapObject Comments, "Comments", "G-ENGINE"
  GapObject Comment, "Comment", "G-ENGINE"
  GapObject Tags, "Tags", "G-ENGINE"
  GapObject PrintOptions, "PrintOptions", "G-ENGINE"
  GapObject Panes, "Panes", "G-HOST"
  GapObject Pane, "Pane", "G-HOST"
  GapObject View, "View", "G-HOST"
  GapObject CommandBars, "CommandBars", "G-HOST"
  GapObject FileDialog, "FileDialog", "G-HOST"
  GapObject Dialogs, "Dialogs", "G-HOST"
  GapObject ActionSettings, "ActionSettings", "G-ENGINE"
  GapObject NamedSlideShows, "NamedSlideShows", "G-ENGINE"
  GapObject SectionProperties, "SectionProperties", "G-ENGINE"
  GapObject ColorSchemes, "ColorSchemes", "G-ENGINE"
  GapObject AnimationSettings, "AnimationSettings", "G-ENGINE"
End Sub

Sub PresentDeck()
  Blank
  Dim r As Object
  Set r = ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange
  r.Text = "Hello from Visual Basic"
  r.Font.Name = "Arial"
  r.Font.Size = 32
  r.Font.Bold = msoTrue
  r.Hyperlinks.Add Nothing, "https://acane.net/melinaoffice"
  ActivePresentation.Slides.Add 2, ppLayoutBlank
  ActivePresentation.Slides(2).Shapes.AddTextbox ppTextOrientationHorizontal, 72, 36, 240, 40
  ActivePresentation.Slides(2).Shapes(1).TextFrame.TextRange.Text = "A text box"
  ActivePresentation.Slides(2).Shapes.AddTable 2, 2, 72, 100, 240
  ActivePresentation.Slides(2).Shapes(2).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Name"
  ActivePresentation.Slides(2).Shapes(2).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Value"
  ActivePresentation.Slides(2).SlideShowTransition.EntryEffect = ppEffectFadeSmoothly
End Sub

Sub Main()
  failures = ""
  started = Presentations.Count
  CheckApplication
  CheckSlides
  CheckText
  CheckTable
  CheckGaps
  PresentDeck
  CloseAdded ActivePresentation.Name
  If failures = "" Then
    Debug.Print "ok"
  Else
    Debug.Print failures
  End If
End Sub
