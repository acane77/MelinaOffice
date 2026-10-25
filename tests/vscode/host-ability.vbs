' Host ability check for the VS Code editor. One file for Ecrivaine, Matheuse, and Oratrice.
' Open a saved .docx, .xlsx, or .pptx (a blank tab has no Path, Print, or FixedFormat).
' Then either:
'   Help, DOM Console, Run Script, paste this file, Run
' or pass this file to melinaoffice_run_vbs (mode module, entry Main).
' Reload the tab after public/ changes.
'
' A user run (DOM console) does not ask Allow / Deny. An agent run does, once per ability.
' Shell is refused (E-DENIED). XPS fails before any write. Both are expected.
' Every local file step opens a dialog first. Cancel skips that step.
' PrintOut writes a PDF beside the open file and does not ask for a path.
' In-app windows (MsgBox, InputBox, an Office dialog) do not pause the agent's one-minute timer.

Dim doc As Object
Dim appName As String
Dim fullName As String
Dim folder As String
Dim sep As String

Public Sub Main()
  Debug.Print "=== host ability ==="
  ShowAbilities
  Bind
  QuietChecks
  Dim answer
  answer = MsgBox("Quiet checks are done. Continue with dialogs and file calls? Cancel skips a later step.", vbYesNo, "Host ability")
  Debug.Print "MsgBox = " & CStr(answer)
  If answer = vbYes Then Interactive
  Debug.Print "=== done ==="
End Sub

Sub ShowAbilities()
  Ability "Path"
  Ability "Save"
  Ability "SaveAs"
  Ability "Open"
  Ability "Print"
  Ability "FixedFormat"
  Ability "Hyperlink"
  Ability "Picture"
  Ability "FileDialog"
  Ability "Dir"
  Ability "Shell"
  Ability "InsertFile"
  Ability "Export"
  Ability "SheetSaveAs"
  Ability "Declare"
  Ability "NoSuch"
End Sub

Sub Ability(name As String)
  Dim n
  n = HasHostAbility(name)
  If n = 0 Then
    Debug.Print "NO   " & name
  Else
    Debug.Print "YES  " & name
  End If
End Sub

Sub Bind()
  appName = "other"
  If HasHostAbility("InsertFile") <> 0 Then appName = "Ecrivaine"
  If HasHostAbility("SheetSaveAs") <> 0 Then appName = "Matheuse"
  If HasHostAbility("Export") <> 0 Then appName = "Oratrice"
  Debug.Print "app  " & appName
  On Error Resume Next
  Err.Clear
  Set doc = Nothing
  If appName = "Ecrivaine" Then Set doc = Application.ActiveDocument
  If appName = "Matheuse" Then Set doc = Application.ActiveWorkbook
  If appName = "Oratrice" Then Set doc = Application.ActivePresentation
  If Err.Number <> 0 Then
    Debug.Print "FAIL active file -> " & Err.Description
    Err.Clear
    Set doc = Nothing
  End If
  fullName = ""
  folder = ""
  sep = "/"
  If Not doc Is Nothing Then
    Err.Clear
    fullName = doc.FullName & ""
    If Err.Number <> 0 Then
      Debug.Print "FAIL FullName -> " & Err.Description
      Err.Clear
      fullName = ""
    Else
      Debug.Print "OK   FullName = " & fullName
    End If
    Err.Clear
    folder = doc.Path & ""
    If Err.Number <> 0 Then
      Debug.Print "FAIL Path -> " & Err.Description
      Err.Clear
      folder = ""
    Else
      Debug.Print "OK   Path = " & folder
    End If
    If InStr(fullName, "\") > 0 And InStr(fullName, "/") = 0 Then sep = "\"
  End If
  On Error GoTo 0
  If folder = "" Then Debug.Print "Open a saved file before Print, Save, or ExportAsFixedFormat."
End Sub

Sub QuietChecks()
  On Error Resume Next
  Err.Clear
  HostShell "cmd"
  ExpectFail "HostShell", "refused"
  Err.Clear
  Shell "cmd"
  ExpectFail "Shell", "refused"
  If doc Is Nothing Then
    Debug.Print "SKIP document checks (no active file)"
  Else
    Err.Clear
    doc.FollowHyperlink "ftp://example.com"
    ExpectFail "FollowHyperlink scheme", "address"
    Err.Clear
    doc.FollowHyperlink "https://acane.net/melinaoffice", "x"
    ExpectFail "FollowHyperlink SubAddress", "SubAddress"
    Err.Clear
    LinkPicture "unused.png", True
    ExpectFail "AddPicture LinkToFile", "LinkToFile"
    Err.Clear
    ExportXps
    If HasHostAbility("FixedFormat") = 0 Then
      ExpectFail "ExportAsFixedFormat XPS", "G-HOST"
    Else
      ExpectFail "ExportAsFixedFormat XPS", "XPS"
    End If
    Err.Clear
    doc.SaveAs "unused.docx", 0, "x"
    ExpectFail "SaveAs extra argument", "is not available"
  End If
  If appName = "Ecrivaine" Then
    Err.Clear
    Application.Selection.InsertFile "unused.txt", "x"
    ExpectFail "InsertFile extra argument", "is not available"
  End If
  If appName = "Oratrice" Then
    Err.Clear
    Application.ActivePresentation.Export "unused", "gif"
    ExpectFail "Export gif", "FilterName"
  End If
  On Error GoTo 0
End Sub

Sub Interactive()
  StepInput
  StepHyperlink
  StepOfficeDialog
  StepFileDialogs
  StepDir
  StepPicture
  StepInsertFile
  StepOpen
  StepSaveCopy
  StepSheetSaveAs
  StepPrint
  StepFixedFormat
  StepSlideExport
  StepSave
End Sub

Sub StepInput()
  Dim typed
  typed = InputBox("Type anything. Cancel returns an empty string.", "Host ability", "sample")
  Debug.Print "OK   InputBox = [" & typed & "]"
End Sub

Sub StepHyperlink()
  If doc Is Nothing Then
    Debug.Print "SKIP FollowHyperlink"
  Else
    Dim address
    address = InputBox("Address to open. Cancel skips. VS Code may ask before an external link.", "Host ability", "https://acane.net/melinaoffice")
    If address = "" Then
      Debug.Print "SKIP FollowHyperlink"
    Else
      On Error Resume Next
      Err.Clear
      doc.FollowHyperlink address
      Report "FollowHyperlink " & address
      On Error GoTo 0
    End If
  End If
End Sub

Sub StepOfficeDialog()
  On Error Resume Next
  Dim n
  If appName = "Ecrivaine" Then
    Debug.Print "--- Dialogs file open (cancel is 0)"
    Err.Clear
    n = Application.Dialogs.Item(80).Show
    ReportValue "Dialogs 80 Show", n
    Debug.Print "--- Dialogs font window. Close it."
    Err.Clear
    n = Application.Dialogs.Item(174).Show
    ReportValue "Dialogs 174 Show", n
  ElseIf appName = "Matheuse" Then
    Debug.Print "--- Dialogs file open (cancel is 0)"
    Err.Clear
    n = Application.Dialogs.Item(1).Show
    ReportValue "Dialogs 1 Show", n
    Debug.Print "--- Dialogs zoom window. Close it."
    Err.Clear
    n = Application.Dialogs.Item(256).Show
    ReportValue "Dialogs 256 Show", n
  Else
    Debug.Print "SKIP Dialogs (Oratrice has no Dialogs index)"
  End If
  On Error GoTo 0
End Sub

Sub StepFileDialogs()
  ShowFileDialog msoFileDialogOpen, "Open"
  ShowFileDialog msoFileDialogSaveAs, "SaveAs"
  ShowFileDialog msoFileDialogFilePicker, "FilePicker"
  ShowFileDialog msoFileDialogFolderPicker, "FolderPicker"
  Debug.Print "--- GetOpenFilename MultiSelect. Cancel is an empty string."
  On Error Resume Next
  Err.Clear
  Dim picked
  picked = GetOpenFilename("", 0, "", "", True)
  If Err.Number <> 0 Then
    Report "GetOpenFilename MultiSelect"
  Else
    Debug.Print "OK   GetOpenFilename MultiSelect"
  End If
  On Error GoTo 0
End Sub

Sub ShowFileDialog(kind As Long, label As String)
  Debug.Print "--- FileDialog " & label & ". Cancel returns 0."
  On Error Resume Next
  Err.Clear
  Dim fd As Object
  Dim n
  Set fd = Application.FileDialog(kind)
  If Err.Number <> 0 Then
    Report "FileDialog " & label
  Else
    n = fd.Show
    ReportValue "FileDialog " & label, n
  End If
  On Error GoTo 0
End Sub

Sub StepDir()
  If folder = "" Then
    Debug.Print "SKIP Dir (the open file has no folder)"
  Else
    Debug.Print "--- Dir lists " & folder
    On Error Resume Next
    Err.Clear
    Dim name1
    Dim name2
    name1 = HostDir(folder & sep & "*")
    If Err.Number <> 0 Then
      Report "HostDir"
    Else
      Debug.Print "OK   HostDir = [" & name1 & "]"
      Err.Clear
      name2 = Dir()
      If Err.Number <> 0 Then
        Report "Dir"
      Else
        Debug.Print "OK   Dir continue = [" & name2 & "]"
      End If
    End If
    On Error GoTo 0
  End If
End Sub

Sub StepPicture()
  Dim path
  path = PickOpen("Choose a picture (PNG, JPEG, EMF, or WMF; Matheuse also converts SVG). Cancel skips.")
  If path <> "" Then
    On Error Resume Next
    Err.Clear
    LinkPicture path, False
    Report "AddPicture " & path
    On Error GoTo 0
  End If
End Sub

Sub LinkPicture(path, linkToFile)
  If appName = "Ecrivaine" Then
    Application.ActiveDocument.InlineShapes.AddPicture path, linkToFile
  ElseIf appName = "Matheuse" Then
    Application.ActiveSheet.Shapes.AddPicture path, linkToFile
  ElseIf appName = "Oratrice" Then
    Application.ActivePresentation.Slides(1).Shapes.AddPicture path, linkToFile
  End If
End Sub

Sub StepInsertFile()
  If appName <> "Ecrivaine" Then
    Debug.Print "SKIP InsertFile"
  Else
    Dim path
    path = PickOpen("Choose a .txt, .htm, .html, .docx, .docm, or .dotx to insert. Cancel skips.")
    If path <> "" Then
      On Error Resume Next
      Err.Clear
      Application.Selection.InsertFile path
      Report "InsertFile " & path
      On Error GoTo 0
    End If
  End If
End Sub

Sub StepOpen()
  Dim path
  path = PickOpen("Choose a file to open in a new tab. Cancel skips.")
  If path <> "" Then
    On Error Resume Next
    Err.Clear
    If appName = "Ecrivaine" Then Application.Documents.Open path
    If appName = "Matheuse" Then Application.Workbooks.Open path
    If appName = "Oratrice" Then Application.Presentations.Open path
    Report "Open " & path
    On Error GoTo 0
  End If
End Sub

Sub StepSaveCopy()
  Dim path
  If doc Is Nothing Then
    Debug.Print "SKIP SaveAs"
  Else
    path = PickSave("Save As: choose a new name. Cancel skips. The original tab keeps its unsaved edits.")
    If path <> "" Then
      On Error Resume Next
      Err.Clear
      doc.SaveAs path
      Report "SaveAs " & path
      On Error GoTo 0
    End If
    path = PickSave("SaveAs2 is the same method. Choose another name, or cancel to skip.")
    If path <> "" Then
      On Error Resume Next
      Err.Clear
      doc.SaveAs2 path
      Report "SaveAs2 " & path
      On Error GoTo 0
    End If
  End If
End Sub

Sub StepSheetSaveAs()
  If appName <> "Matheuse" Then
    Debug.Print "SKIP Worksheet.SaveAs"
  Else
    Dim path
    path = PickSave("Worksheet SaveAs as CSV. Choose the file name. Cancel skips.")
    If path <> "" Then
      On Error Resume Next
      Err.Clear
      Application.ActiveSheet.SaveAs path, 6
      Report "Worksheet.SaveAs csv " & path
      On Error GoTo 0
    End If
    path = PickSave("Worksheet SaveAs as tab-separated text. Choose the file name. Cancel skips.")
    If path <> "" Then
      On Error Resume Next
      Err.Clear
      Application.ActiveSheet.SaveAs path, -4158
      Report "Worksheet.SaveAs txt " & path
      On Error GoTo 0
    End If
  End If
End Sub

Sub StepPrint()
  If doc Is Nothing Then
    Debug.Print "SKIP PrintOut"
  Else
    Dim answer
    answer = MsgBox("PrintOut writes a PDF beside the open file (same name, .pdf). Continue?", vbYesNo, "Host ability")
    If answer <> vbYes Then
      Debug.Print "SKIP PrintOut"
    Else
      On Error Resume Next
      Err.Clear
      doc.PrintOut
      Report "PrintOut"
      If appName = "Matheuse" Then
        Err.Clear
        Application.ActiveSheet.PrintOut
        Report "Worksheet.PrintOut"
        Err.Clear
        Application.ActiveSheet.PageSetup.PrintOut
        Report "PageSetup.PrintOut"
      End If
      On Error GoTo 0
    End If
  End If
End Sub

Sub StepFixedFormat()
  If doc Is Nothing Then
    Debug.Print "SKIP ExportAsFixedFormat"
  Else
    Dim path
    path = PickSave("ExportAsFixedFormat: choose where to write the PDF. Cancel skips.")
    If path <> "" Then
      On Error Resume Next
      Err.Clear
      ExportPdf path
      Report "ExportAsFixedFormat " & path
      On Error GoTo 0
    End If
    If appName = "Matheuse" Then
      path = PickSave("Worksheet ExportAsFixedFormat: choose a PDF for this sheet. Cancel skips.")
      If path <> "" Then
        On Error Resume Next
        Err.Clear
        Application.ActiveSheet.ExportAsFixedFormat 0, path
        Report "Worksheet.ExportAsFixedFormat " & path
        On Error GoTo 0
      End If
    End If
  End If
End Sub

Sub ExportPdf(path)
  If appName = "Oratrice" Then
    doc.ExportAsFixedFormat path, 2
  ElseIf appName = "Matheuse" Then
    doc.ExportAsFixedFormat 0, path
  Else
    doc.ExportAsFixedFormat 17, path
  End If
End Sub

Sub ExportXps()
  If appName = "Oratrice" Then
    doc.ExportAsFixedFormat "unused.xps", 1
  ElseIf appName = "Matheuse" Then
    doc.ExportAsFixedFormat 1, "unused.xps"
  Else
    doc.ExportAsFixedFormat 18, "unused.xps"
  End If
End Sub

Sub StepSlideExport()
  If appName <> "Oratrice" Then
    Debug.Print "SKIP Export"
  Else
    Dim path
    Dim dest
    path = PickSave("Export slides: choose any file name in the destination folder. The script uses the folder and writes Slide1.png. Cancel skips.")
    If path <> "" Then
      dest = FolderOf(path)
      On Error Resume Next
      Err.Clear
      Application.ActivePresentation.Export dest, "png"
      Report "Export " & dest
      On Error GoTo 0
    End If
  End If
End Sub

Sub StepSave()
  If doc Is Nothing Then
    Debug.Print "SKIP Save"
  Else
    Dim answer
    answer = MsgBox("Save writes the open file itself. Continue?", vbYesNo, "Host ability")
    If answer <> vbYes Then
      Debug.Print "SKIP Save"
    Else
      On Error Resume Next
      Err.Clear
      doc.Save
      Report "Save"
      On Error GoTo 0
    End If
  End If
End Sub

Function PickOpen(hint As String) As String
  Debug.Print "--- " & hint
  On Error Resume Next
  Err.Clear
  Dim path
  path = ""
  path = GetOpenFilename()
  If Err.Number <> 0 Then
    Debug.Print "FAIL GetOpenFilename -> " & Err.Description
    Err.Clear
    PickOpen = ""
  Else
    PickOpen = path & ""
    If PickOpen = "" Then
      Debug.Print "SKIP cancelled"
    Else
      Debug.Print "picked " & PickOpen
    End If
  End If
  On Error GoTo 0
End Function

Function PickSave(hint As String) As String
  Debug.Print "--- " & hint
  On Error Resume Next
  Err.Clear
  Dim path
  path = ""
  path = GetSaveAsFilename()
  If Err.Number <> 0 Then
    Debug.Print "FAIL GetSaveAsFilename -> " & Err.Description
    Err.Clear
    PickSave = ""
  Else
    PickSave = path & ""
    If PickSave = "" Then
      Debug.Print "SKIP cancelled"
    Else
      Debug.Print "picked " & PickSave
    End If
  End If
  On Error GoTo 0
End Function

Function FolderOf(full As String) As String
  Dim i As Long
  Dim slash As Long
  slash = 0
  For i = 1 To Len(full)
    If Mid(full, i, 1) = "/" Or Mid(full, i, 1) = "\" Then slash = i
  Next
  If slash > 1 Then
    FolderOf = Left(full, slash - 1)
  Else
    FolderOf = full
  End If
End Function

Sub ExpectFail(label As String, piece As String)
  If Err.Number = 0 Then
    Debug.Print "FAIL " & label & " (expected an error)"
  ElseIf piece <> "" And InStr(Err.Description, piece) = 0 Then
    Debug.Print "FAIL " & label & " -> " & Err.Description
  Else
    Debug.Print "OK   " & label & " -> " & Err.Description
  End If
  Err.Clear
End Sub

Sub Report(label As String)
  If Err.Number = 0 Then
    Debug.Print "OK   " & label
  Else
    Debug.Print "FAIL " & label & " -> " & Err.Description
    Err.Clear
  End If
End Sub

Sub ReportValue(label As String, value)
  If Err.Number = 0 Then
    Debug.Print "OK   " & label & " = " & CStr(value)
  Else
    Debug.Print "FAIL " & label & " -> " & Err.Description
    Err.Clear
  End If
End Sub
