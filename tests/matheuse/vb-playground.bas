' Welcome to Visual Basic Playground
' Write your own Visual Basic code in this editor,
' then click Run to run and test it.
'
' See also: Office VBA Documentation:
'    https://learn.microsoft.com/en-us/office/vba/api/overview/excel

' Your code should contain a Main procedure
Sub Main()
  ' Write a short list and a total on the active sheet.
  ActiveSheet.Range("A1").Value = "Item"
  ActiveSheet.Range("B1").Value = "Amount"
  ActiveSheet.Range("A2").Value = "Pens"
  ActiveSheet.Range("B2").Value = 12.5
  ActiveSheet.Range("A3").Value = "Paper"
  ActiveSheet.Range("B3").Value = 7.5
  ActiveSheet.Range("B4").Formula = "=SUM(B2:B3)"
  ActiveSheet.Range("A1:B1").Font.Bold = True
  ActiveSheet.Range("B2:B4").NumberFormat = "0.00"
  ActiveWorkbook.Names.Add "Total", "=Sheet1!$B$4"
  ActiveSheet.Range("A6").Value = WorksheetFunction.Sum(ActiveSheet.Range("B2:B3"))
End Sub
