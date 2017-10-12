Public Class frmWizardry15Base
    Public Sub New(ByVal base As WizEditBase, ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mBase = base
        Me.Text = Caption
        Me.Icon = Icon
        Me.pbBoxArt.Image = BoxArt
    End Sub
#Region "Properties"
#Region "Declarations"
    Private mBase As WizEditBase
#End Region
#End Region
#Region "Methods"
#End Region
#Region "Event Handlers"
#End Region
End Class