Attribute VB_Name = "libRegistry"
'libRegistry - VbRegMod.bas
'   Library Registry Manipulation Module...
'Public domain, taken from "The Waite Group's Visual Basic Source Library"/SAMS Publishing...
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   03/04/99    None        Ken Clark       Elaborated error messages to include better detail;
'   03/02/99    None        Ken Clark       Incorporated into FiRRe;
'=================================================================================================================================
Option Explicit

'---------------------------------------------------
'-- VbRegMod.Bas
'-- A Visual Basic 32-Bit Module For Accessing
'-- The Windows Registry.
'--
'-- Date: Sunday, May 17, 1998
'-- By Custom Software Designers.
'-- Programmer Raymond L. King
'---------------------------------------------------

'-- Windows Registry Root Key Constants.
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

'-- Windows Registry Key Type Constants.
Public Const REG_OPTION_NON_VOLATILE = 0        ' Key is preserved when system is rebooted
Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)

'-- Function Error Constants.
Public Const ERROR_SUCCESS = 0
Public Const ERROR_REG = 1
Public Const ERROR_BADKEY = 2

'-- Registry Access Rights.
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

'-- Windows Registry API Declarations.
'-- Registry API To Open A Key.
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long

'-- Registry API To Create A New Key.
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
  ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
  ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

'-- Registry API To Query A String Value.
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
  lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
  ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'-- Registry API To Query A Long (DWORD) Value.
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
  lpType As Long, lpData As Long, lpcbData As Long) As Long

'-- Registry API To Query A NULL Value.
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
  lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

'-- Registry API To Set A String Value.
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
  ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
  ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'-- Registry API To Set A Long (DWORD) Value.
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
  ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

'-- Registry API To Delete A Key.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long

'-- Registry API To Delete A Key Value.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
  (ByVal hKey As Long, ByVal lpValueName As String) As Long

'-- Registry API To Close A Key.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'-- Constants For Error Messages.
Public Const OpenErr = "Error: Opening Registry Key!"
Public Const DeleteErr = "Error: Deleteing Key!"
Public Const CreateErr = "Error: Creating Key!"
Public Const QueryErr = "Error: Querying Value!"
Private Function cvtKey(RootKey As Long) As String
    Select Case RootKey
        Case HKEY_CLASSES_ROOT
            cvtKey = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_CONFIG
            cvtKey = "HKEY_CURRENT_CONFIG"
        Case HKEY_CURRENT_USER
            cvtKey = "HKEY_CURRENT_USER"
        Case HKEY_DYN_DATA
            cvtKey = "HKEY_DYN_DATA"
        Case HKEY_LOCAL_MACHINE
            cvtKey = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            cvtKey = "HKEY_USERS"
    End Select
End Function
Private Function cvtKeyType(KeyType As Integer) As String
    Select Case KeyType
        Case REG_OPTION_NON_VOLATILE
            cvtKeyType = "REG_OPTION_NON_VOLATILE"
        Case REG_EXPAND_SZ
            cvtKeyType = "REG_EXPAND_SZ"
        Case REG_DWORD
            cvtKeyType = "REG_DWORD"
        Case REG_SZ
            cvtKeyType = "REG_SZ"
        Case REG_BINARY
            cvtKeyType = "REG_BINARY"
        Case REG_DWORD_BIG_ENDIAN
            cvtKeyType = "REG_DWORD_BIG_ENDIAN"
        Case REG_DWORD_LITTLE_ENDIAN
            cvtKeyType = "REG_DWORD_LITTLE_ENDIAN"
    End Select
End Function
'-------------------------------------------------------------
'-- Procedure   : Public Method VbRegDeleteKey
'-- Programmer  : Raymond L. King
'-- Created On  : Sunday, May 17, 1998 11:03:04 AM
'-- Module      : VbRegMod
'-- Module File : VbRegMod.bas
'-- Project     : Project1
'-- Project File: Project1.vbp
'-- Parameters  :
'-- RootKey     : The Root Key To Open, EG: HKEY_CURRENT_USER
'-- KeyName     : The Key Name To Open
'--             : Example: MySettings\Settings
'-- SubKey      : The Sub Key Under KeyName To Delete
'-------------------------------------------------------------
Public Sub VbRegDeleteKey(RootKey As Long, KeyName As String, SubKey As String)
    Dim lRtn    As Long      '-- API Return Value
    Dim hKey    As Long      '-- Handle Of Key
  
    '-- Open The Specified Registry Key.
    lRtn = RegOpenKeyEx(RootKey, KeyName, 0&, KEY_ALL_ACCESS, hKey)
    If lRtn <> ERROR_SUCCESS Then
        MsgBox GetErrorText(CInt(lRtn)) & vbCr & _
            GetErrorTitle(CInt(lRtn)) & " (" & lRtn & ") returned from RegOpenKeyEx accessing" & vbCr & _
            cvtKey(RootKey) & "\" & KeyName & "\" & SubKey
        RegCloseKey (hKey)
        Exit Sub
    End If
  
    '-- Delete The Registry SubKey.
    lRtn = RegDeleteKey(hKey, SubKey)
    If lRtn <> ERROR_SUCCESS Then
        MsgBox GetErrorText(CInt(lRtn)) & vbCr & _
            GetErrorTitle(CInt(lRtn)) & " (" & lRtn & ") returned from RegDeleteKey accessing" & vbCr & _
            cvtKey(RootKey) & "\" & KeyName & "\" & SubKey
    End If
    RegCloseKey (hKey)
End Sub
'-------------------------------------------------------------
'-- Procedure   : Public Method VbRegCreateKey
'-- Programmer  : Raymond L. King
'-- Created On  : Sunday, May 17, 1998 11:03:18 AM
'-- Module      : VbRegMod
'-- Module File : VbRegMod.bas
'-- Project     : Project1
'-- Project File: Project1.vbp
'-- Parameters  :
'-- RootKey     : The Root Key To Open, EG: HKEY_CURRENT_USER
'-- KeyName     : The New Key Name To Create
'--             : Example: MySettings\Settings
'-------------------------------------------------------------
'
Public Sub VbRegCreateKey(RootKey As Long, KeyName As String)
    Dim lRtn    As Long     '-- Registry API Return Value
    Dim hKey    As Long     '-- Handle Of Open Key
    
    '-- Create The New Registry Key.
    lRtn = RegCreateKeyEx(RootKey, KeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_WRITE, 0&, hKey, lRtn)
    If lRtn <> ERROR_SUCCESS Then
        MsgBox GetErrorText(CInt(lRtn)) & vbCr & _
            GetErrorTitle(CInt(lRtn)) & " (" & lRtn & ") returned from RegCreateKeyEx accessing" & vbCr & _
            cvtKey(RootKey) & "\" & KeyName
    End If
    RegCloseKey (hKey)
End Sub
'-------------------------------------------------------------
'-- Procedure   : Public Method VbRegQueryValue
'-- Programmer  : Raymond L. King
'-- Created On  : Sunday, May 17, 1998 11:03:29 AM
'-- Module      : VbRegMod
'-- Module File : VbRegMod.bas
'-- Project     : Project1
'-- Project File: Project1.vbp
'-- Parameters  :
'-- RootKey     : The Root Key To Open, EG: HKEY_CURRENT_USER
'-- KeyName     : The Key Name To Open
'--             : Example: MySettings\Settings
'-- ValueName   : The Value Name To Query
'-------------------------------------------------------------
'
Public Function VbRegQueryValue(RootKey As Long, KeyName As String, ValueName As String) As Variant
    Dim lRtn    As Long     '-- API Return Code
    Dim hKey    As Long     '-- Handle Of Open Key
    Dim lCdata  As Long     '-- The Data
    Dim lValue  As Long     '-- Long (DWORD) Value
    Dim sValue  As String   '-- String Value
    Dim lRtype  As Long     '-- Type Returned String Or DWORD
  
    '-- Open The Registry Key.
    lRtn = RegOpenKeyEx(RootKey, KeyName, 0&, KEY_QUERY_VALUE, hKey)
    If lRtn <> ERROR_SUCCESS Then
        If lRtn <> ERROR_BADKEY Then
            MsgBox GetErrorText(CInt(lRtn)) & vbCr & _
                GetErrorTitle(CInt(lRtn)) & " (" & lRtn & ") returned from RegOpenKeyEx accessing" & vbCr & _
                cvtKey(RootKey) & "\" & KeyName & "\" & ValueName
        End If
        RegCloseKey (hKey)
        Exit Function
    End If
  
    '-- Query Registry Key For Value Type.
    lRtn = RegQueryValueExNULL(hKey, ValueName, 0&, lRtype, 0&, lCdata)
    If lRtn <> ERROR_SUCCESS Then
        If lRtn <> ERROR_BADKEY Then
            MsgBox GetErrorText(CInt(lRtn)) & vbCr & _
                GetErrorTitle(CInt(lRtn)) & " (" & lRtn & ") returned from RegQueryValueExNULL accessing" & vbCr & _
                cvtKey(RootKey) & "\" & KeyName & "\" & ValueName
        End If
        RegCloseKey (hKey)
        Exit Function
    End If
  
    '-- Get The Key Value By Type.
    Select Case lRtype
        Case 1    '-- REG_SZ (String)
            sValue = String(lCdata, 0)
            '-- Get Registry String Value.
            lRtn = RegQueryValueExString(hKey, ValueName, 0&, lRtype, sValue, lCdata)
            If lRtn = ERROR_SUCCESS Then
                VbRegQueryValue = Replace(sValue, vbNullChar, vbNullString)
            Else
                VbRegQueryValue = Empty
            End If
        Case 4    '-- REG_DWORD
            '-- Get Registry Long (DWORD) Value.
            lRtn = RegQueryValueExLong(hKey, ValueName, 0&, lRtype, lValue, lCdata)
            If lRtn = ERROR_SUCCESS Then
                VbRegQueryValue = lValue
            Else
                VbRegQueryValue = Empty
            End If
    End Select
    RegCloseKey (hKey)
End Function
'-------------------------------------------------------------
'-- Procedure   : Public Method VbRegSetValue
'-- Programmer  : Raymond L. King
'-- Created On  : Sunday, May 17, 1998 11:03:42 AM
'-- Module      : VbRegMod
'-- Module File : VbRegMod.bas
'-- Project     : Project1
'-- Project File: Project1.vbp
'-- Parameters  :
'-- RootKey     : The Root Key To Open, EG: HKEY_CURRENT_USER
'-- KeyName     : The Key Name To Open
'--             : Example: MySettings\Settings
'-- ValueName   : The Value Name To Open
'-- KeyType     : The Key Type, EG: REG_SZ Or REG_DWORD
'-- KeyValue    : The Value To Set Under ValueName
'-------------------------------------------------------------
'
Public Sub VbRegSetValue(RootKey As Long, KeyName As String, ValueName As String, KeyType As Integer, KeyValue As Variant)
    Dim lRtn    As Long     '-- Returned Value From API Registry Call
    Dim hKey    As Long     '-- Handle Of The Opened Key
    Dim lValue  As Long     '-- Setting A Long Data Value
    Dim sValue  As String   '-- Setting A String Data Value
    Dim lSize   As Long     '-- Size Of String Data To Set
    
    '-- Open The Registry Key.
    lRtn = RegOpenKeyEx(RootKey, KeyName, 0, KEY_SET_VALUE, hKey)
    If lRtn <> ERROR_SUCCESS Then
        MsgBox GetErrorText(CInt(lRtn)) & vbCr & _
            GetErrorTitle(CInt(lRtn)) & " (" & lRtn & ") returned from RegOpenKeyEx accessing" & vbCr & _
            cvtKey(RootKey) & "\" & KeyName
        RegCloseKey (hKey)
        Exit Sub
    End If
  
    '-- Select The Key Type.
    Select Case KeyType
        Case 1    '-- REG_SZ (String)
            sValue = KeyValue        '-- Assign Key Value
            lSize = Len(sValue)      '-- Get Size Of String
            '-- Set String Value.
            lRtn = RegSetValueExString(hKey, ValueName, 0&, REG_SZ, sValue, lSize)
            If lRtn <> ERROR_SUCCESS Then
                MsgBox GetErrorText(CInt(lRtn)) & vbCr & _
                    GetErrorTitle(CInt(lRtn)) & " (" & lRtn & ") returned from RegSetValueExString accessing" & vbCr & _
                    cvtKey(RootKey) & "\" & KeyName & "\" & ValueName & " = " & KeyValue & "(" & cvtKeyType(KeyType) & ")"
                RegCloseKey (hKey)
                Exit Sub
            End If
        Case 4    '-- REG_DWORD
            lValue = KeyValue    '-- Assign The Long Value.
            '-- Set The Long Value (DWORD).
            lRtn = RegSetValueExLong(hKey, ValueName, 0&, REG_DWORD, lValue, 4)
            If lRtn <> ERROR_SUCCESS Then
                MsgBox GetErrorText(CInt(lRtn)) & vbCr & _
                    GetErrorTitle(CInt(lRtn)) & " (" & lRtn & ") returned from RegSetValueExLong accessing" & vbCr & _
                    cvtKey(RootKey) & "\" & KeyName & "\" & ValueName & " = " & KeyValue & "(" & cvtKeyType(KeyType) & ")"
            End If
    End Select
    RegCloseKey (hKey)
End Sub
'-------------------------------------------------------------
'-- Procedure   : Public Method VbRegDeleteValue
'-- Programmer  : Raymond L. King
'-- Created On  : Sunday, May 17, 1998 11:03:49 AM
'-- Module      : VbRegMod
'-- Module File : VbRegMod.bas
'-- Project     : Project1
'-- Project File: Project1.vbp
'-- Parameters  :
'-- RootKey     : The Root Key To Open, EG: HKEY_CURRENT_USER
'-- KeyName     : The Key Name To Open
'--             : Example: MySettings\Settings
'-- ValueName   : The Value Name To Delete
'-------------------------------------------------------------
'
Public Sub VbRegDeleteValue(RootKey As Long, KeyName As String, ValueName As String)
    Dim lRtn    As Long        '-- API Call Returned Value
    Dim hKey    As Long        '-- Handle Of Opened Key
  
    '-- Open Registry Key...
    lRtn = RegOpenKeyEx(RootKey, KeyName, 0&, KEY_ALL_ACCESS, hKey)
    If lRtn <> ERROR_SUCCESS Then
        MsgBox GetErrorText(CInt(lRtn)) & vbCr & _
            GetErrorTitle(CInt(lRtn)) & " (" & lRtn & ") returned from RegOpenKeyEx accessing" & vbCr & _
            cvtKey(RootKey) & "\" & KeyName
        RegCloseKey (hKey)
        Exit Sub
    End If
  
    '-- Delete Opened Key Value Name...
    lRtn = RegDeleteValue(hKey, ValueName)
    If lRtn <> ERROR_SUCCESS Then
        MsgBox GetErrorText(CInt(lRtn)) & vbCr & _
            GetErrorTitle(CInt(lRtn)) & " (" & lRtn & ") returned from RegDeleteValue accessing" & vbCr & _
        cvtKey(RootKey) & "\" & KeyName & "\" & ValueName
    End If
    RegCloseKey (hKey)
End Sub

