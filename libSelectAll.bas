Attribute VB_Name = "libSelectAll"
'SelectAll - SelectAll.bas
'   Generic TextBox Control Routines...
'   Copyright � 2000, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   05/04/00    None        Ken Clark       Incorporated into FiRRe;
'=================================================================================================================================
Option Explicit
Public Sub TextSelected()
    Dim i As Integer
    Dim ctl As Control
   
    On Error Resume Next
    Set ctl = Screen.ActiveControl
    Select Case TypeName(ctl)
        Case "TextBox", "ComboBox", "DataCombo"
            i = Len(ctl.Text)
            ctl.SelStart = 0
            ctl.SelLength = i
    End Select
    'Debug.Print ctl.Name & " Selected..."
    Set ctl = Nothing
End Sub
Public Sub KeyPressUcase(KeyAscii As Integer)
    Dim Char As String
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub
Public Sub KeyPressInteger(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 45 And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 7
    End If
End Sub
Public Sub KeyPressReal(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 45 And (KeyAscii < 46 Or KeyAscii > 57) Then
        KeyAscii = 7
    End If
End Sub
