Attribute VB_Name = "modMain"
'modMain - modMain.bas
'   Main module for the WizEdit Application...
'   Copyright © 2000, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   08/26/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const adoNullError = &H80040E21
Public Function ChkINumber(cnum As Integer, _
                Optional NoNegative As Boolean = False) As Integer
    If NoNegative And cnum = 45 Or _
        (cnum <> 8 And cnum <> 32 And cnum <> 45 And (cnum < 48 Or cnum > 57)) Then
        ChkINumber = 7
    Else
        ChkINumber = cnum
    End If
End Function
Public Sub DtoI6test(ByVal x As Double)
    Dim Data(1 To 3) As Integer
    Dim xBytes(1 To 6) As Byte
    
    Data(3) = CInt(x / 100000000#)
    x = x - (Data(3) * 100000000#)
    Data(2) = CInt(x / 10000#)
    x = x - (Data(2) * 10000#)
    Data(1) = x

    Debug.Print "Wizardry Says: " & (Data(3) * 100000000#) + (Data(2) * 10000#) + Data(1)
End Sub
Public Sub DtoI6test2()
    Dim Data(1 To 3) As Integer
    Dim x As Double
    Data(3) = 0
    Data(2) = 3
    Data(1) = 2898

    x = I6toD(Data)
    
    Data(3) = 257
    Data(2) = 257
    Data(1) = 257

    x = I6toD(Data)
    
    Data(3) = 0
    Data(2) = 214
    Data(1) = -4000

    x = I6toD(Data)
    
    x = 32898
    Call DtoI6(x, Data)
    Debug.Print Data(3) & Data(2) & Data(1)
End Sub
Public Sub DtoI6(ByVal x As Double, Data() As Integer)
    Data(3) = CInt(x / 100000000#)
    x = x - (Data(3) * 100000000#)
    Data(2) = CInt(x / 10000#)
    x = x - (Data(2) * 10000#)
    Data(1) = x
End Sub
Public Function I6toD(Data() As Integer) As Double
    Dim r1 As Double
    Dim r2 As Double
    Dim r3 As Double
    
    r1 = Data(1)
    r2 = Data(2) * 10000#
    r3 = Data(3) * 100000000#
    
    I6toD = r1 + r2 + r3
End Function
Public Sub initCommonDialog()
    frmMain.cdgMain.CancelError = False
    frmMain.cdgMain.FileName = vbNullString
    frmMain.cdgMain.Filter = vbNullString
    frmMain.cdgMain.FilterIndex = 0
    frmMain.cdgMain.flags = 0
    'Any additional fields used by any Form in FiRRe must be added to this list...
End Sub
Public Function strHex(ByRef xBytes() As Byte, nBytes As Integer) As String
    Dim i As Integer

    strHex = ""
    For i = 1 To nBytes
        strHex = strHex & Format(Hex(xBytes(i)), "00") ' & " "
        If i Mod 4 = 0 Then strHex = strHex & " "
        If i Mod 32 = 0 Then strHex = strHex & vbCrLf
    Next i
End Function
Public Function ValidateByte(Optional x As Control = Nothing) As Boolean
    Const iLimit As Byte = 99
    Dim fCancel As Boolean
    Dim ctl As Control
    
    On Error Resume Next
    fCancel = False
    If x Is Nothing Then Set ctl = Screen.ActiveControl Else Set ctl = x
    With ctl
        If .Text = vbNullString Then .Text = "0"
        If Val(.Text) < 0 Or Val(.Text) > iLimit Then
            fCancel = True
            Call Beep
            .Text = vbNullString
        Else
            '.Text = Format(.Text, "00")
        End If
    End With
    Set ctl = Nothing
    ValidateByte = fCancel
End Function
Public Function ValidateI2(Optional x As Control = Nothing) As Boolean
    Dim iLimit As Long
    Dim fCancel As Boolean
    Dim ctl As Control
    
    On Error Resume Next
    fCancel = False
    If x Is Nothing Then Set ctl = Screen.ActiveControl Else Set ctl = x
    With ctl
        If .Text = vbNullString Then .Text = "0"
        iLimit = (2 ^ 16) - 1
        If Val(.Text) < 0 Or CLng(Val(.Text)) > iLimit Then
            fCancel = True
            Call Beep
            .Text = vbNullString
        Else
            .Text = Format(.Text, "#,##0")
        End If
    End With
    Set ctl = Nothing
    ValidateI2 = fCancel
End Function
Public Function ValidateI4(Optional x As Control = Nothing) As Boolean
    Dim iLimit As Long
    Dim fCancel As Boolean
    Dim ctl As Control
    
    On Error Resume Next
    fCancel = False
    If x Is Nothing Then Set ctl = Screen.ActiveControl Else Set ctl = x
    With ctl
        If .Text = vbNullString Then .Text = "0"
        iLimit = (2 ^ 31) - 1
        If Val(.Text) < 0 Or CLng(Val(.Text)) > iLimit Then
            fCancel = True
            Call Beep
            .Text = vbNullString
        Else
            .Text = Format(.Text, "#,##0")
        End If
    End With
    Set ctl = Nothing
    ValidateI4 = fCancel
End Function
Public Function UpCase(uKey As Integer) As Integer
    If uKey > 96 And uKey < 123 Then
        UpCase = uKey - 32
    Else
        UpCase = uKey
    End If
End Function

