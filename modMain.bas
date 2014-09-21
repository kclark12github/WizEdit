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

