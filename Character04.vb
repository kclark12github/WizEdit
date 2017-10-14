'Character04.cls
'   Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/14/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class Character04
    Inherits CharacterBase
    Public Sub New(ByVal Base As Wizardry04)
        MyBase.New(Base)
    End Sub
End Class