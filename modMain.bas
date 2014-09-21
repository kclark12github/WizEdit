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
Global gstrUWAPath1 As String
Global gstrUWAPath2 As String
Global gstrUWAPath3 As String
Global gstrUWAPath4 As String
Global gstrUWAPath5 As String
Global gstrUWAPath6 As String
Global gstrUWAPath7 As String
Global gstrUWAPath7g As String
Public Function GetWizEditSetting(Key As String, Value As String, vDefault As Variant) As Variant
    GetWizEditSetting = VbRegQueryValue(HKEY_CURRENT_USER, gstrRegPath & "\" & Key, Value)
    If GetWizEditSetting = vbNullString Then GetWizEditSetting = vDefault
End Function
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

