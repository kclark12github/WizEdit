Attribute VB_Name = "libShell"
'libShell - libShell.bas
'   Shell Interface...
'   Copyright © 2000, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   05/04/00    None        Ken Clark       Incorporated into FiRRe;
'=================================================================================================================================
Option Explicit
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF       '  Infinite timeout
Private Const DEBUG_PROCESS = &H1
Private Const DEBUG_ONLY_THIS_PROCESS = &H2

Private Const CREATE_SUSPENDED = &H4

Private Const DETACHED_PROCESS = &H8

Private Const CREATE_NEW_CONSOLE = &H10

Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const REALTIME_PRIORITY_CLASS = &H100

Private Const CREATE_NEW_PROCESS_GROUP = &H200

Private Const CREATE_NO_WINDOW = &H8000000

Private Const WAIT_FAILED = -1&
Private Const WAIT_OBJECT_0 = 0
Private Const WAIT_ABANDONED = &H80&
Private Const WAIT_ABANDONED_0 = &H80&

Private Const WAIT_TIMEOUT = &H102&

Private Const SW_SHOW = 5

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function AllocConsole Lib "kernel32" () As Long
Declare Function FreeConsole Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SleepEx& Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long)

Public Const STD_OUTPUT_HANDLE = -11&
Dim hConsole As Long
Sub WaitForProcess(pid&)
    Dim phnd&
    Dim lMilliseconds As Long
    Dim lAlertable As Long
    
    lAlertable = False
    lMilliseconds = 10000
    While True
        phnd = OpenProcess(SYNCHRONIZE, 0, pid)
        If phnd <> 0 Then
            'Call WaitForSingleObject(phnd, INFINITE)
            Call WaitForSingleObject(phnd, lMilliseconds)
            DoEvents
            Call CloseHandle(phnd)
        Else
            Exit Sub
        End If
    Wend
End Sub
