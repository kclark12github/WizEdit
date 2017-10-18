'WizEditBase.cls
'   Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/11/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Imports WizEdit.Character04
Public Class frmWizardry04
    Public Sub New(ByVal base As WizEditBase, ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image)
        MyBase.New(base, Caption, Icon, BoxArt)
    End Sub
End Class