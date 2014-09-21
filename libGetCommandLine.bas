Attribute VB_Name = "libGetCommandLine"
'libGetCommandLine - libGetCommandLine.bas
'   Library Get Command Line Module...
'   Copyright © 2000, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   05/04/00    None        Ken Clark       Incorporated into FiRRe;
'=================================================================================================================================
Option Explicit
Function GetCommandLine(Optional MaxArgs)
   'Declare variables.
   Dim C, CmdLine, CmdLnLen, InArg, I, NumArgs
   
   'See if MaxArgs was provided.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   'Make array of the correct size.
   ReDim Args(MaxArgs)
   
   NumArgs = 0
   InArg = False
   'Get command line arguments.
   
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   'Go thru command line one character
   'at a time.
   For I = 1 To CmdLnLen
      C = Mid(CmdLine, I, 1)
      'Test for space or tab.
      If (C <> " " And C <> vbTab) Then
         'Neither space nor tab.
         'Test if already in argument.
         If Not InArg Then
         'New argument begins.
         'Test for too many arguments.
            If NumArgs = MaxArgs Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
         End If
         'Concatenate character to current argument.
         Args(NumArgs) = Args(NumArgs) & C
      Else
         'Found a space or tab.
         'Set InArg flag to False.
         InArg = False
      End If
   Next I
   'Resize array just enough to hold arguments.
   ReDim Preserve Args(NumArgs)
   'Return Array in Function name.
   GetCommandLine = Args()
End Function


