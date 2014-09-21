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

Global Const gstrRegPath As String = "Software\Sage Software\WizEdit"
Public Function ChkINumber(cnum As Integer, _
                Optional NoNegative As Boolean = False) As Integer
    If NoNegative And cnum = 45 Or _
        (cnum <> 8 And cnum <> 32 And cnum <> 45 And (cnum < 48 Or cnum > 57)) Then
        ChkINumber = 7
    Else
        ChkINumber = cnum
    End If
End Function
Public Function GetWizEditSetting(Key As String, Value As String, vDefault As Variant) As Variant
    GetWizEditSetting = VbRegQueryValue(HKEY_CURRENT_USER, gstrRegPath & "\" & Key, Value)
    If GetWizEditSetting = vbNullString Then GetWizEditSetting = vDefault
End Function
Public Sub initCommonDialog()
    frmMain.cdgMain.CancelError = False
    frmMain.cdgMain.FileName = vbNullString
    frmMain.cdgMain.Filter = vbNullString
    frmMain.cdgMain.FilterIndex = 0
    frmMain.cdgMain.flags = 0
    'Any additional fields used by any Form in FiRRe must be added to this list...
End Sub
Public Sub SaveWizEditSetting(Key As String, Value As String, Data As Variant)
    Dim CurrentValue As Variant
    Dim Token As String
    Dim KeyPath As String
    Dim i As Integer
    
    CurrentValue = GetWizEditSetting(Key, Value, vbNullString)
    KeyPath = gstrRegPath & "\" & Key
    If CurrentValue = vbNullString Then
        Call VbRegCreateKey(HKEY_CURRENT_USER, KeyPath)
    End If
    
    If CurrentValue <> Data Then
        'Although this routine accepts a Variant as the data, it really stores the
        'data in the registry as a string (REG_SZ)...
        Call VbRegSetValue(HKEY_CURRENT_USER, KeyPath, Value, REG_SZ, Data)
    End If
End Sub
Public Function ValidateByte(Optional x As Control = Nothing) As Boolean
    Const iLimit As Byte = 99
    Dim fCancel As Boolean
    Dim ctl As Control
    fCancel = False
    If x Is Nothing Then Set ctl = Screen.ActiveControl Else Set ctl = x
    With ctl
        If Val(.Text) < 1 Or Val(.Text) > iLimit Then
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
    fCancel = False
    If x Is Nothing Then Set ctl = Screen.ActiveControl Else Set ctl = x
    With ctl
        iLimit = (2 ^ 16) - 1
        If Val(.Text) < 1 Or CLng(Val(.Text)) > iLimit Then
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
    fCancel = False
    If x Is Nothing Then Set ctl = Screen.ActiveControl Else Set ctl = x
    With ctl
        iLimit = (2 ^ 31) - 1
        If Val(.Text) < 1 Or CLng(Val(.Text)) > iLimit Then
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

