'frmWizardry05.vb
'   Scenario-Specific Form for Wizardry05...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/21/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class frmWizardry05
    Public Sub New(ByVal base As WizEditBase, ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image)
        MyBase.New(base, Caption, Icon, BoxArt)
        InitializeComponent()
        'These properties are set in the base class, but our InitializeComponent will have most likely overwritten them...
        Me.Text = Caption
        Me.Icon = Icon
        Me.pbBoxArt.Image = BoxArt
    End Sub
    Protected Overrides Sub cbCharacter_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            Me.epScenario.SetError(sender, "")
            MyBase.cbCharacter_SelectedIndexChanged(sender, e)
            'Populate our controls with Character data...
            With CType(mCharacter, Character05)
                'TODO: Scenario-Specific screen population...
                Me.nudMarks.Value = .Marks
                Me.nudRIP.Value = .RIP
            End With
        Catch ex As Exception : Debug.WriteLine(ex.ToString)
            Me.epScenario.SetError(sender, ex.ToString)
        End Try
    End Sub
    Protected Overrides Sub cmdSave_Click(sender As Object, e As EventArgs)
        Try
            Me.epScenario.SetError(sender, "")
            MyBase.cmdSave_Click(sender, e)
            With CType(mCharacter, Character05)
                'TODO: Scenario-Specific character property assignment...
                .Marks = Me.nudMarks.Value
                .RIP = Me.nudRIP.Value
            End With
            MyBase.cmdSave_Click(sender, e)
        Catch ex As Exception : Debug.WriteLine(ex.ToString)
            Me.epScenario.SetError(sender, ex.ToString)
        End Try
    End Sub
End Class