Attribute VB_Name = "ModVolidec"
' Ecrivaine / Word 内嵌: 使用标准 VB.Form 的 leverage_loss.frm（非 Word UserForm GUID 格式）
' 入口: ShowLeverageLoss
' Volatility decay on leveraged products (same formula as leverage_loss.cpp)

Public Sub ShowLeverageLoss()
    leverage_loss.Show vbModeless
End Sub

Public Function Volidec(ByVal mag As Long, ByVal decperc As Double) As Double
    If decperc >= 100 Then
        Volidec = 100
        Exit Function
    End If
    If decperc < 0 Then decperc = 0
    If mag < 0 Then mag = 0
    Dim osci As Double
    osci = 100 - (1 + (100 / (100 - decperc) - 1) * mag) * (100 - decperc * mag)
    If osci > 100 Then osci = 100
    Volidec = osci
End Function

Public Function ValidateLeverageInputs( _
        ByVal leverageText As String, _
        ByVal declineText As String, _
        ByRef mag As Long, _
        ByRef dec As Double, _
        ByRef errMsg As String) As Boolean

    ValidateLeverageInputs = False
    errMsg = vbNullString

    If Not ParseIntegerMag(leverageText, mag, errMsg) Then Exit Function

    If Not ParseDeclinePct(declineText, dec, errMsg) Then
        If Len(errMsg) = 0 Then errMsg = "Decline % requires a number >= 0 and <= 100."
        Exit Function
    End If
    If dec < 0 Or dec > 100 Then
        errMsg = "Decline percentage could not be smaller than 0 or greater than 100."
        Exit Function
    End If

    ValidateLeverageInputs = True
End Function

Private Function ParseIntegerMag(ByVal s As String, ByRef mag As Long, ByRef errMsg As String) As Boolean
    Dim t As String
    Dim d As Double

    ParseIntegerMag = False
    t = Trim$(s)
    If Len(t) = 0 Then
        errMsg = "Leverage Mag requires an integer >= 1."
        Exit Function
    End If
    If Not IsNumeric(t) Then
        errMsg = "Leverage Mag requires an integer >= 1."
        Exit Function
    End If

    d = Val(t)
    If d < 1 Or d > 2147483647# Then
        errMsg = "Magnification could not be smaller than 1."
        Exit Function
    End If
    If d <> Fix(d) Then
        errMsg = "Leverage Mag must be a whole number (no decimals)."
        Exit Function
    End If

    mag = CLng(d)
    ParseIntegerMag = True
End Function

Public Function ParseDeclinePct(ByVal s As String, ByRef outVal As Double, Optional ByRef errMsg As String) As Boolean
    Dim t As String
    Dim d As Double

    ParseDeclinePct = False
    errMsg = vbNullString
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    If Right$(t, 1) = "%" Then t = Trim$(Left$(t, Len(t) - 1))
    If Len(t) = 0 Then Exit Function
    If Not IsNumeric(t) Then Exit Function

    d = Val(t)
    If Abs(d) > 1E+308 Then
        errMsg = "Decline % is out of range."
        Exit Function
    End If
    outVal = d
    ParseDeclinePct = True
End Function
