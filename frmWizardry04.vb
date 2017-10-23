'frmWizardry04.vb
'   Scenario-Specific Form for Wizardry04...
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
        InitializeComponent()
        'These properties are set in the base class, but our InitializeComponent will have most likely overwritten them...
        Me.Text = Caption
        Me.Icon = Icon
        Me.pbBoxArt.Image = BoxArt
    End Sub
    Protected Sub ClearGroups()
        For i As Short = 1 To 3
            InitGroup(i, 0)
        Next i
    End Sub
    Protected Function FindGroup(ByVal cbGroup As ComboBox, ByVal GroupCode As UShort) As Integer
        FindGroup = -1
        For i As Integer = 0 To cbGroup.Items.Count - 1
            If CType(cbGroup.Items(i), MonsterGroupData).Code = GroupCode Then Return i
        Next i
        Return -1
    End Function
    Protected Sub InitGroup(ByVal iGroup As Short, ByVal GroupCode As UShort)
        Dim cbGroup As ComboBox = Nothing
        Select Case iGroup
            Case 1 : cbGroup = Me.cbGroup1
            Case 2 : cbGroup = Me.cbGroup2
            Case 3 : cbGroup = Me.cbGroup3
        End Select
        cbGroup.SelectedIndex = FindGroup(cbGroup, GroupCode)
    End Sub
    Protected Overrides Sub RefreshData()
        MyBase.RefreshData()
        For iChar As Short = 0 To mBase.Characters.Length - 1
            cbCharacter.Items(iChar) = String.Format("Save Game {0}", iChar + 1)
        Next iChar
        Me.cbGroup1.Items.Clear() : Me.cbGroup1.Items.AddRange(CType(mBase, Wizardry04).MasterMonsterGroupList)
        Me.cbGroup2.Items.Clear() : Me.cbGroup2.Items.AddRange(CType(mBase, Wizardry04).MasterMonsterGroupList)
        Me.cbGroup3.Items.Clear() : Me.cbGroup3.Items.AddRange(CType(mBase, Wizardry04).MasterMonsterGroupList)
    End Sub
    Protected Friend Overrides Sub ToggleEditMode(ByVal EditMode As Boolean)
        MyBase.ToggleEditMode(EditMode)
        'Under no circumstances should we ever allow change of txtName or txtPassword...
        MyBase.EnableControl(MyBase.txtName, False) : MyBase.EnableControl(MyBase.txtPassword, False)
    End Sub
    Protected Overrides Sub cbCharacter_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            Me.epScenario.SetError(sender, "")
            MyBase.cbCharacter_SelectedIndexChanged(sender, e)
            'Populate our controls with Character data...
            With CType(mCharacter, Character04)
                'Groups...
                Me.ClearGroups()
                Me.InitGroup(1, .SummonedMonsterCode(0))
                Me.InitGroup(2, .SummonedMonsterCode(1))
                Me.InitGroup(3, .SummonedMonsterCode(2))
            End With
        Catch ex As Exception : Debug.WriteLine(ex.ToString)
            Me.epScenario.SetError(sender, ex.ToString)
        End Try
    End Sub
    Protected Overrides Sub cmdSave_Click(sender As Object, e As EventArgs)
        Try
            Me.epScenario.SetError(sender, "")
            MyBase.cmdSave_Click(sender, e)
            With CType(mCharacter, Character04)
                'Groups...
                .SummonedMonsterCount(0) = Me.nudGroup1.Value
                .SummonedMonsterCount(1) = Me.nudGroup2.Value
                .SummonedMonsterCount(2) = Me.nudGroup3.Value
                .SummonedMonsterCode(0) = CType(Me.cbGroup1.SelectedItem, MonsterGroupData).Code
                .SummonedMonsterCode(1) = CType(Me.cbGroup2.SelectedItem, MonsterGroupData).Code
                .SummonedMonsterCode(2) = CType(Me.cbGroup3.SelectedItem, MonsterGroupData).Code
                .SummonedMonsterName(0) = CType(Me.cbGroup1.SelectedItem, MonsterGroupData).Text
                .SummonedMonsterName(1) = CType(Me.cbGroup2.SelectedItem, MonsterGroupData).Text
                .SummonedMonsterName(2) = CType(Me.cbGroup3.SelectedItem, MonsterGroupData).Text
            End With
            MyBase.cmdSave_Click(sender, e)
        Catch ex As Exception : Debug.WriteLine(ex.ToString)
            Me.epScenario.SetError(sender, ex.ToString)
        End Try
    End Sub
End Class