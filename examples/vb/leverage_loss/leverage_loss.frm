VERSION 5.00
Begin VB.Form leverage_loss 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Volatility Decay on Leverages"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInsertWord 
      Caption         =   "Calculate and Insert Table to Word"
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtDeclinePct 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "10"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtLeverageMag 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Text            =   "2"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblLostValue 
      Alignment       =   1  'Right Justify
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Lost %:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2820
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Decline %:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2340
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Leverage Mag:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1740
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Calculate how much you lost from the oscillation of leverage."
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "VOLATILITY DECAY ON LEVERAGES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   4935
   End
End
Attribute VB_Name = "leverage_loss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Incompatible = False
Attribute VB_ShowModal = False
Option Explicit

Private Sub cmdCalculate_Click()
    Dim mag As Long
    Dim dec As Double
    Dim lost As Double
    If Not RunCalculation(mag, dec, lost) Then Exit Sub
    lblLostValue.Caption = Format$(lost, "0.0")
End Sub

Private Sub cmdInsertWord_Click()
    Dim mag As Long
    Dim dec As Double
    Dim lost As Double
    If Not RunCalculation(mag, dec, lost) Then Exit Sub
    lblLostValue.Caption = Format$(lost, "0.0")
    InsertLeverageTableToWord mag, dec, lost
End Sub

Private Function RunCalculation(ByRef mag As Long, ByRef dec As Double, ByRef lost As Double) As Boolean
    Dim errMsg As String

    If Not ValidateLeverageInputs(txtLeverageMag.Text, txtDeclinePct.Text, mag, dec, errMsg) Then
        MsgBox errMsg, vbExclamation, "Invalid input"
        If InStr(1, errMsg, "Leverage", vbTextCompare) > 0 Or InStr(1, errMsg, "Magnification", vbTextCompare) > 0 Then
            txtLeverageMag.SetFocus
        Else
            txtDeclinePct.SetFocus
        End If
        Exit Function
    End If

    On Error GoTo CalcErr
    lost = Volidec(mag, dec)
    RunCalculation = True
    Exit Function

CalcErr:
    MsgBox "Calculation error: " & Err.Description, vbExclamation, "Invalid input"
End Function

Private Sub Form_Load()
    lblLostValue.Caption = "--"
End Sub

Private Sub txtDeclinePct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCalculate_Click
End Sub

Private Sub txtLeverageMag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCalculate_Click
End Sub
