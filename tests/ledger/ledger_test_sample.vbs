' Ledger test sample.
' Run it from Help or File, Ledger DOM Console, Run Script.
' Paste this file into the editor and click Run.
' Sub Main writes a short table on the active sheet.
'
' After it runs, Sheet1 should show:
'   A1  Ledger sample
'   A2:D2  Item, Qty, Price, Line   (bold)
'   Pens   4   12.50   50.00
'   Paper  2    7.50   15.00
'   Total              65.00
'   A7  65     (WorksheetFunction.Sum of the line amounts)
' The defined name LineTotal refers to =Sheet1!$D$5.
'
' Excel VBA reference:
'   https://learn.microsoft.com/en-us/office/vba/api/overview/excel

Sub Main()
  Dim sh As Object
  Set sh = ActiveSheet

  sh.Range("A1").Value = "Ledger sample"
  sh.Range("A1").Font.Size = 14
  sh.Range("A1").Font.Bold = True

  sh.Range("A2").Value = "Item"
  sh.Range("B2").Value = "Qty"
  sh.Range("C2").Value = "Price"
  sh.Range("D2").Value = "Line"
  sh.Range("A2:D2").Font.Bold = True

  sh.Range("A3").Value = "Pens"
  sh.Range("B3").Value = 4
  sh.Range("C3").Value = 12.5
  sh.Range("D3").Formula = "=B3*C3"

  sh.Range("A4").Value = "Paper"
  sh.Range("B4").Value = 2
  sh.Range("C4").Value = 7.5
  sh.Range("D4").Formula = "=B4*C4"

  sh.Range("A5").Value = "Total"
  sh.Range("D5").Formula = "=SUM(D3:D4)"
  sh.Range("C3:D5").NumberFormat = "0.00"
  sh.Range("A:A").ColumnWidth = 16

  ActiveWorkbook.Names.Add "LineTotal", "=Sheet1!$D$5"
  sh.Range("A7").Value = WorksheetFunction.Sum(sh.Range("D3:D4"))
End Sub
