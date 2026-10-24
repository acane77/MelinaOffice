VERSION 5.00
Begin VB.Form Calculator
   Caption         =   "Calculator"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3480
   Begin VB.TextBox Display
      Alignment       =   1
      Height          =   560
      Left            =   120
      Locked          =   -1
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   3240
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   16
      EndProperty
   End
   Begin VB.CommandButton cmdClear
      Caption         =   "C"
      Height          =   640
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   720
   End
   Begin VB.CommandButton cmdDiv
      Caption         =   "/"
      Height          =   640
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   720
   End
   Begin VB.CommandButton cmdMul
      Caption         =   "*"
      Height          =   640
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   720
   End
   Begin VB.CommandButton cmdSub
      Caption         =   "-"
      Height          =   640
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   720
   End
   Begin VB.CommandButton cmd7
      Caption         =   "7"
      Height          =   640
      Left            =   120
      TabIndex        =   5
      Top             =   1600
      Width           =   720
   End
   Begin VB.CommandButton cmd8
      Caption         =   "8"
      Height          =   640
      Left            =   960
      TabIndex        =   6
      Top             =   1600
      Width           =   720
   End
   Begin VB.CommandButton cmd9
      Caption         =   "9"
      Height          =   640
      Left            =   1800
      TabIndex        =   7
      Top             =   1600
      Width           =   720
   End
   Begin VB.CommandButton cmdAdd
      Caption         =   "+"
      Height          =   640
      Left            =   2640
      TabIndex        =   8
      Top             =   1600
      Width           =   720
   End
   Begin VB.CommandButton cmd4
      Caption         =   "4"
      Height          =   640
      Left            =   120
      TabIndex        =   9
      Top             =   2360
      Width           =   720
   End
   Begin VB.CommandButton cmd5
      Caption         =   "5"
      Height          =   640
      Left            =   960
      TabIndex        =   10
      Top             =   2360
      Width           =   720
   End
   Begin VB.CommandButton cmd6
      Caption         =   "6"
      Height          =   640
      Left            =   1800
      TabIndex        =   11
      Top             =   2360
      Width           =   720
   End
   Begin VB.CommandButton cmdEq
      Caption         =   "="
      Height          =   640
      Left            =   2640
      TabIndex        =   12
      Top             =   2360
      Width           =   720
   End
   Begin VB.CommandButton cmd1
      Caption         =   "1"
      Height          =   640
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   720
   End
   Begin VB.CommandButton cmd2
      Caption         =   "2"
      Height          =   640
      Left            =   960
      TabIndex        =   14
      Top             =   3120
      Width           =   720
   End
   Begin VB.CommandButton cmd3
      Caption         =   "3"
      Height          =   640
      Left            =   1800
      TabIndex        =   15
      Top             =   3120
      Width           =   720
   End
   Begin VB.CommandButton cmdDot
      Caption         =   "."
      Height          =   640
      Left            =   2640
      TabIndex        =   16
      Top             =   3120
      Width           =   720
   End
   Begin VB.CommandButton cmd0
      Caption         =   "0"
      Height          =   640
      Left            =   120
      TabIndex        =   17
      Top             =   3880
      Width           =   1560
   End
   Begin VB.CommandButton cmdInsert
      Caption         =   "Insert"
      Height          =   640
      Left            =   1800
      TabIndex        =   18
      Top             =   3880
      Width           =   1560
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Stored As Double
Private Pending As String
Private Fresh As Boolean

Private Sub Form_Load()
    Display.Text = "0"
    Stored = 0
    Pending = ""
    Fresh = True
End Sub

Private Sub TypeDigit(ByVal Digit As String)
    If Display.Text = "Error" Then Exit Sub
    If Fresh Or Display.Text = "0" Then
        If Digit = "." Then
            Display.Text = "0."
        Else
            Display.Text = Digit
        End If
    ElseIf Digit = "." Then
        If InStr(Display.Text, ".") = 0 Then Display.Text = Display.Text & "."
    Else
        Display.Text = Display.Text & Digit
    End If
    Fresh = False
End Sub

Private Sub Commit(ByVal NextOp As String)
    Dim Current As Double
    If Display.Text = "Error" Then
        Pending = ""
        Fresh = True
        Exit Sub
    End If
    If Fresh Then
        If NextOp <> "" Then Pending = NextOp
        Exit Sub
    End If
    Current = Val(Display.Text)
    If Pending = "+" Then
        Current = Stored + Current
    ElseIf Pending = "-" Then
        Current = Stored - Current
    ElseIf Pending = "*" Then
        Current = Stored * Current
    ElseIf Pending = "/" Then
        If Current = 0 Then
            Display.Text = "Error"
            Stored = 0
            Pending = ""
            Fresh = True
            Exit Sub
        End If
        Current = Stored / Current
    End If
    Stored = Current
    Display.Text = Trim$(Str$(Stored))
    Pending = NextOp
    Fresh = True
End Sub

Private Sub cmdClear_Click()
    Display.Text = "0"
    Stored = 0
    Pending = ""
    Fresh = True
End Sub

Private Sub cmd0_Click(): TypeDigit "0": End Sub
Private Sub cmd1_Click(): TypeDigit "1": End Sub
Private Sub cmd2_Click(): TypeDigit "2": End Sub
Private Sub cmd3_Click(): TypeDigit "3": End Sub
Private Sub cmd4_Click(): TypeDigit "4": End Sub
Private Sub cmd5_Click(): TypeDigit "5": End Sub
Private Sub cmd6_Click(): TypeDigit "6": End Sub
Private Sub cmd7_Click(): TypeDigit "7": End Sub
Private Sub cmd8_Click(): TypeDigit "8": End Sub
Private Sub cmd9_Click(): TypeDigit "9": End Sub
Private Sub cmdDot_Click(): TypeDigit ".": End Sub
Private Sub cmdAdd_Click(): Commit "+": End Sub
Private Sub cmdSub_Click(): Commit "-": End Sub
Private Sub cmdMul_Click(): Commit "*": End Sub
Private Sub cmdDiv_Click(): Commit "/": End Sub
Private Sub cmdEq_Click(): Commit "": End Sub

Private Sub cmdInsert_Click()
    If Pending <> "" And Not Fresh Then Commit ""
    If Display.Text = "Error" Then Exit Sub
    InsertText Display.Text
End Sub
