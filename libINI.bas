Attribute VB_Name = "libINI"
'libINI - INI.bas
'   Library INI Module...
'Public domain, taken from "The Waite Group's Visual Basic Source Library"/SAMS Publishing...
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   02/18/99    None        Ken Clark       Incorporated into FiRRe;
'=================================================================================================================================
Option Explicit
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Sub SaveINIKey(ByVal File As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
    Dim fRet As Long
    
    fRet = WritePrivateProfileString(Section, Key, Value, File)
End Sub
Function GetINIKey(ByVal szFile As String, ByVal szSection As String, ByVal szKey As String, ByVal szDefault) As String
    Dim szValue As String
    Dim nLen As Integer
    
    '---prepare string buffers
    szValue = Space$(250)
    'szDefault = ""
    
    '---call WINAPI
    nLen = GetPrivateProfileString(szSection, ByVal szKey, szDefault, szValue, Len(szValue), szFile)

    '---trim null char
    If nLen = 0 Then
        GetINIKey = ""
    Else
        GetINIKey = Trim$(Left$(szValue, nLen))
    End If
End Function


