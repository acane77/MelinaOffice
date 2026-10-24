Attribute VB_Name = "Main"
Public Sub Main()
    Dim f As Calculator
    Set f = New Calculator
    f.Caption = "Second calculator"
    f.Display.Text = "Ready"
    f.cmdInsert.Caption = "Write"
    f.Show
End Sub