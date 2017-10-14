'Character05.cls
'   Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/14/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class Character05
    Inherits CharacterBase
    Public Sub New(ByVal Base As Wizardry05)
        MyBase.New(Base)
    End Sub
End Class