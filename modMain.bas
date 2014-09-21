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
Public CommandLineArgs() As Variant
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
    Dim r1 As Double
    Dim r2 As Double
    Dim r3 As Double
    
    r3 = x \ 100000000#
    Data(3) = CInt(r3)
    x = x - (r3 * 100000000#)
    r2 = x \ 10000#
    Data(2) = r2
    x = x - (r2 * 10000#)
    r1 = x
    Data(1) = r1
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
    frmMain.cdgMain.Flags = 0
    'Any additional fields used by any Form in FiRRe must be added to this list...
End Sub
Public Function IsPrintable(ByVal xByte As Byte) As Boolean
    If xByte < 32 Then
        IsPrintable = False
    Else
        Select Case xByte
            Case 127, 129, 141, 143, 144, 157
                IsPrintable = False
            Case Else
                IsPrintable = True
        End Select
    End If
End Function
Public Sub HexDump(xInput As String, xOutput As String)
    Dim i As Long
    Dim iChar As Integer
    Dim Offset As Long
    Dim iUnit As Integer
    Dim oUnit As Integer
    Dim Data(0 To 15) As Byte
    Dim BytesReadSoFar As Long
    Dim BytesLeft As Long
    Dim MaxBytes As Long
    Dim errorCode As Long
    Dim aOutput As String
    Dim hOutput As String
    Dim oOutput As String
    
    'On Error GoTo ErrorHandler
    iUnit = FreeFile
    MaxBytes = FileLen(xInput)
    Open xInput For Binary Access Read Lock Read As #iUnit
    oUnit = FreeFile
    Open xOutput For Output Access Write Lock Read Write As #oUnit
    Print #oUnit, String(100, "=")
    Print #oUnit, "HexDump of " & xInput & " Generated on " & Format(Date, "Long Date")
    Print #oUnit, String(100, "=")
    
    Offset = 0
    Do While Not EOF(iUnit) And BytesReadSoFar < MaxBytes
        Get #iUnit, , Data
        BytesReadSoFar = Seek(iUnit) - 1
        BytesLeft = BytesReadSoFar Mod 16
    
        aOutput = vbNullString
        hOutput = vbNullString
        For i = 0 To 15 - BytesLeft
            oOutput = Hex(Data(i))
            oOutput = String(2 - Len(oOutput), "0") & oOutput
            hOutput = hOutput & oOutput & " "
            If IsPrintable(Data(i)) Then aOutput = aOutput & Chr(CLng(Data(i))) Else aOutput = aOutput & "."
            Data(i) = 0
        Next i
        If BytesLeft > 0 Then
            For i = 1 To BytesLeft
                hOutput = hOutput & "   "
                aOutput = aOutput & " "
            Next i
        End If
        oOutput = Hex(Offset)
        If Len(oOutput) < 8 Then oOutput = String(8 - Len(oOutput), "0") & oOutput
        
        Print #oUnit, oOutput & vbTab & hOutput & vbTab & aOutput
        Offset = Offset + 16
    Loop
    
    Print #oUnit, String(100, "=")
    Print #oUnit, "Total Bytes Dumped: " & BytesReadSoFar
    'Debug.Print "Total Bytes Dumped: " & BytesReadSoFar
    Print #oUnit, String(100, "=")

ExitSub:
    Close #iUnit
    Close #oUnit
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Test"
    Exit Sub
    Resume Next
    
'    strHex = ""
'    For i = 1 To nBytes
'        strHex = strHex & Format(Hex(xBytes(i)), "00") ' & " "
'        If i Mod 4 = 0 Then strHex = strHex & " "
'        If i Mod 32 = 0 Then strHex = strHex & vbCrLf
'    Next i
End Sub
Public Sub Main()
    frmMain.Show vbModal
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

