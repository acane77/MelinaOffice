Attribute VB_Name = "ModWordInsert"
' Insert table at the selection using Word document object model (in-process)

Public Sub InsertLeverageTableToWord(ByVal mag As Long, ByVal decPct As Double, ByVal lostPct As Double)
    Dim doc As Document
    Dim tbl As Table
    Dim rng As Range

    On Error GoTo Fail

    Set doc = ActiveDocument
    Set rng = Selection.Range
    Set tbl = doc.Tables.Add(Range:=rng, NumRows:=2, NumColumns:=3)

    With tbl
        .Borders.Enable = True
        .Cell(1, 1).Range.Text = "Leverage Mag"
        .Cell(1, 2).Range.Text = "Decline %"
        .Cell(1, 3).Range.Text = "Lost %"
        .Cell(2, 1).Range.Text = CStr(mag)
        .Cell(2, 2).Range.Text = Format$(decPct, "0.##") & "%"
        .Cell(2, 3).Range.Text = Format$(lostPct, "0.0") & "%"
        .Rows(1).Range.Font.Bold = True
        .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Rows(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With

    rng.SetRange tbl.Range.End, tbl.Range.End
    rng.InsertParagraphAfter
    Exit Sub

Fail:
    MsgBox "Could not insert table." & vbCrLf & Err.Description, vbCritical, "Insert table"
End Sub
