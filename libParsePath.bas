Attribute VB_Name = "libParsePath"
'libParsePath - ParsePath.bas
'   Library ParsePath Module...
'Public domain, taken from "The Waite Group's Visual Basic Source Library"/SAMS Publishing...
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   02/18/99    None        Ken Clark       Incorporated into FiRRe;
'=================================================================================================================================
Option Explicit

Public Enum ParseParts
  DrvOnly = 1
  DirOnly = 2
  DirOnlyNoSlash = -2
  DrvDir = 3
  DrvDirNoSlash = -3
  FileNameBase = 4
  FileNameExt = 5
  FileNameExtNoDot = -5
  FileNameBaseExt = 6
  DrvDirFileNameBase = 7
End Enum

Public Function ParsePath(ByVal strPath As String, intPart As ParseParts) As String

  '*******************************************************************************
  '
  ' DESCRIPTION
  '     Extracts the specified portion of a pathname.
  '
  ' ARGUMENTS
  '     strPath = Path to parse.
  '     intPart = DrvOnly            - Drive only (with colon).
  '             = DirOnly            - Directory only.
  '             = DirOnlyNoSlash     - Same as above but not terminated with backslash.
  '             = DrvDir             - Drive and directory.
  '             = DrvDirNoSlash      - Same as above but not terminated with backslash.
  '             = FileNameBase       - Filename base.
  '             = FileNameExt        - Filename extension.
  '             = FileNameExtNoDot   - Same as above but without the dot.
  '             = FileNameBaseExt    - Filename base and extension.
  '             = DrvDirFileNameBase - Drive, directory, and filename base.
  '
  ' RETURNS
  '     Specifed path portion. NULL string if intPart is invalid.
  '
  ' DEPENDENCIES
  '     None
  '
  ' REMARKS
  '     When passing just a DIRECTORY, make sure it is terminated with "\".
  '
  '*******************************************************************************
  
  On Local Error Resume Next
  
  Dim intCPos As Integer
  Dim intLPos As Integer
  Dim intTemp As Integer
  Dim intPathStart As Integer
  Dim intPathLen As Integer
  Dim strPart1 As String
  Dim strPart2 As String
  Dim strPart4 As String
  Dim strPart5 As String
  
  intPathLen = Len(strPath)
  
  '-------------------
  ' Get drive portion.
  '-------------------
  intCPos = InStr(strPath, ":")
   
  If intCPos Then
    strPart1 = Left(strPath, intCPos)
  End If
 
  '------------------
  ' Get path portion.
  '------------------
  intLPos = InStr(1, strPath, "\")
  If Right(strPath, 1) = "\" Then
          
    '----------------------------
    ' strPath contains no filename.
    '----------------------------
    If intPathLen > intLPos Then
      If intPart < 0 Then
        strPart2 = Mid(strPath, intLPos, intPathLen - intLPos)
      Else
        strPart2 = Mid(strPath, intLPos)
      End If
    Else
      strPart2 = "\"
    End If
   
  Else
       
    If intLPos Then
    
      '------------------------------
      'strPath must contain a filename.
      '------------------------------
      intPathStart = intLPos
      intLPos = intLPos + 1
            
      Do
        intCPos = InStr(intLPos, strPath, "\")
        If intCPos Then
          intLPos = intCPos + 1
        End If
      Loop While intCPos
          
      If intPart < 0 Then
        strPart2 = Mid(strPath, intPathStart, intLPos - intPathStart - 1)
      Else
        strPart2 = Mid(strPath, intPathStart, intLPos - intPathStart)
      End If
      
    Else
      '-------------------
      ' No path was found.
      '-------------------
      
      If Len(strPart1) Then
        
        '------------------------------------------------------------------
        ' If drive spec, start at position 3 when getting filename portion.
        '------------------------------------------------------------------
        intLPos = 3
        
      Else
        intLPos = 1
        
      End If
      
    End If
       
    strPart4 = Mid(strPath, intLPos)
      
    '--------------------------------------
    ' Check if filename base has extension.
    '--------------------------------------
    intCPos = 1
    Do
      intTemp = InStr(intCPos + 1, strPart4, ".")
      If intTemp Then intCPos = intTemp
    Loop While intTemp
    
    If intCPos > 1 Then
               
      '--------------------------------
      ' Get filename extension portion.
      '--------------------------------
      If InStr(CStr(intPart), "-") Then ' Check if it's "negative".
        strPart5 = Mid(strPart4, intCPos + 1)
      Else
        strPart5 = Mid(strPart4, intCPos)
      End If
      
      '---------------------------
      ' Get filename base portion.
      '---------------------------
      strPart4 = Left(strPart4, intCPos - 1)
        
    End If
      
  End If
   
  Select Case intPart
    
    '----------------------
    ' Return drive portion.
    '----------------------
    Case DrvOnly
      ParsePath = strPart1
   
    '---------------------
    ' Return path portion.
    '---------------------
    Case DirOnly, DirOnlyNoSlash
      ParsePath = strPart2
      
    '-------------------------------
    ' Return drive and path portion.
    '-------------------------------
    Case DrvDir, DrvDirNoSlash
      ParsePath = strPart1 & strPart2
            
    '------------------------------
    ' Return filename base portion.
    '------------------------------
    Case FileNameBase
      ParsePath = strPart4
    
    '-----------------------------------
    ' Return filename extension portion.
    '-----------------------------------
    Case FileNameExt, FileNameExtNoDot
      ParsePath = strPart5
    
    '--------------------------------------------
    ' Return filename base and extension portion.
    '--------------------------------------------
    Case FileNameBaseExt
      ParsePath = strPart4 & strPart5
      
    '--------------------------------------------
    ' Return filename base and extension portion.
    '--------------------------------------------
    Case DrvDirFileNameBase
      ParsePath = strPart1 & strPart2 & strPart4
    
  End Select
  
End Function




