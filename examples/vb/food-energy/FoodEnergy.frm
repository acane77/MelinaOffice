VERSION 5.00
Begin VB.Form FoodEnergy
   Caption         =   "Food Energy Calculator"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4680
   Begin VB.TextBox txtPer100
      Height          =   345
      Left            =   2400
      TabIndex        =   0
      Top             =   150
      Width           =   2100
   End
   Begin VB.TextBox txtWeight
      Height          =   345
      Left            =   2400
      TabIndex        =   1
      Top             =   630
      Width           =   2100
   End
   Begin VB.TextBox txtKj
      Alignment       =   1
      Height          =   345
      Left            =   2400
      Locked          =   -1
      TabIndex        =   4
      TabStop         =   0
      Top             =   1170
      Width           =   2100
   End
   Begin VB.TextBox txtKcal
      Alignment       =   1
      Height          =   345
      Left            =   2400
      Locked          =   -1
      TabIndex        =   5
      TabStop         =   0
      Top             =   1650
      Width           =   2100
   End
   Begin VB.CommandButton cmdInsertKj
      Caption         =   "Insert kJ to Document"
      Height          =   420
      Left            =   180
      TabIndex        =   2
      Top             =   2160
      Width           =   4320
   End
   Begin VB.CommandButton cmdInsertKcal
      Caption         =   "Insert kCal to Document"
      Height          =   420
      Left            =   180
      TabIndex        =   3
      Top             =   2640
      Width           =   4320
   End
   Begin VB.Label lblPer100
      Caption         =   "Energy per 100 g (kJ)"
      Height          =   300
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   2160
   End
   Begin VB.Label lblWeight
      Caption         =   "Weight (g)"
      Height          =   300
      Left            =   180
      TabIndex        =   7
      Top             =   660
      Width           =   2160
   End
   Begin VB.Label lblKj
      Caption         =   "Energy (kJ)"
      Height          =   300
      Left            =   180
      TabIndex        =   8
      Top             =   1200
      Width           =   2160
   End
   Begin VB.Label lblKcal
      Caption         =   "Energy (kCal)"
      Height          =   300
      Left            =   180
      TabIndex        =   9
      Top             =   1680
      Width           =   2160
   End
End
Attribute VB_Name = "FoodEnergy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 1 kCal = 4.184 kJ. The portion is the per-100 g figure times the weight.
Private Const KjPerKcal As Double = 4.184

Private Sub Form_Load()
    ShowResult
End Sub

Private Sub txtPer100_Change()
    ShowResult
End Sub

Private Sub txtWeight_Change()
    ShowResult
End Sub

Private Function PortionKj() As Double
    PortionKj = -1
    If Not IsNumeric(txtPer100.Text) Or Not IsNumeric(txtWeight.Text) Then Exit Function
    If Val(txtPer100.Text) < 0 Or Val(txtWeight.Text) < 0 Then Exit Function
    PortionKj = Val(txtPer100.Text) * Val(txtWeight.Text) / 100
End Function

Private Function OnePlace(ByVal N As Double) As String
    OnePlace = Trim$(Str$(Int(N * 10 + 0.5) / 10))
End Function

Private Sub ShowResult()
    Dim Kj As Double
    Kj = PortionKj()
    If Kj < 0 Then
        txtKj.Text = ""
        txtKcal.Text = ""
    Else
        txtKj.Text = OnePlace(Kj)
        txtKcal.Text = OnePlace(Kj / KjPerKcal)
    End If
End Sub

Private Sub cmdInsertKj_Click()
    ShowResult
    If txtKj.Text = "" Then Exit Sub
    InsertText txtKj.Text & " kJ"
End Sub

Private Sub cmdInsertKcal_Click()
    ShowResult
    If txtKcal.Text = "" Then Exit Sub
    InsertText txtKcal.Text & " kCal"
End Sub
