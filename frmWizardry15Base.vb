'frmWizardry15Base.vb
'   Scenario-Specific Base Form for Wizardry 1-5...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/11/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class frmWizardry15Base
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByVal base As WizEditBase, ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mBase = base
        Me.Text = Caption
        Me.Icon = Icon
        Me.pbBoxArt.Image = BoxArt

        Me.cbAlignment.SelectedValue = -1
        Me.cbProfession.SelectedValue = -1
        Me.cbRace.SelectedValue = -1
        Me.cbStatus.SelectedValue = -1
        Me.nudAC.Value = 0
        Me.nudAgeInWeeks.Value = 0
        Me.nudAgeInYears.Value = 0
        Me.nudAgility.Value = 0
        Me.nudHitPoints.Value = 0
        Me.nudIntelligence.Value = 0
        Me.nudLevel.Value = 0
        Me.nudLuck.Value = 0
        Me.nudPiety.Value = 0
        Me.nudStrength.Value = 0
        Me.nudVitality.Value = 0
        Me.nudLocationEast.Value = 0
        Me.nudLocationNorth.Value = 0
        Me.nudLocationDown.Value = 0
        Me.nudEXP.Value = 0
        Me.nudGold.Value = 0
        Me.txtName.Text = ""
        Me.txtPassword.Text = ""
    End Sub

#Region "Properties"
#Region "Declarations"
    Protected mBase As WizEditBase
    Protected mChanged As Boolean = False
    Protected mCharacter As CharacterBase = Nothing
    Protected mEditMode As Boolean = False
#End Region
#End Region
#Region "Methods"
    Protected Sub ClearItems()
        For i As Short = 1 To 8
            InitItem(i, False, False, False, 0)
        Next i
    End Sub
    Protected Friend Sub EnableControl(ByVal ctl As Control, ByVal Enable As Boolean)
        EnableControl(ctl, Enable, False)
    End Sub
    Protected Friend Sub EnableControl(ByVal ctl As Control, ByVal Enable As Boolean, ByVal Clear As Boolean)
        Dim strTag As String = IIf(ctl.Tag Is Nothing, "", CType(ctl.Tag, String))
        Dim tagParams() As String = strTag.Split(",".ToCharArray)
        For i As Integer = 0 To tagParams.Length - 1
            Select Case tagParams(i).ToUpper
                Case "IGNORE" : Exit Sub
            End Select
        Next i

        Dim ForeColor As Color = IIf(Enable, System.Drawing.Color.Black, System.Drawing.Color.Yellow)
        Dim BackColor As Color = IIf(Enable, System.Drawing.Color.Gray, System.Drawing.Color.DarkGray)
        Select Case TypeName(ctl)
            Case "Button" : CType(ctl, Button).Enabled = Enable
            Case "CheckBox"
                With CType(ctl, CheckBox)
                    .Enabled = Enable
                    If Clear Then .CheckState = IIf(.ThreeState, CheckState.Indeterminate, CheckState.Unchecked)
                End With
            Case "CheckedListBox"
                With CType(ctl, CheckedListBox)
                    .Enabled = Enable
                    If Clear Then .ClearSelected()
                    .ForeColor = ForeColor : .BackColor = BackColor
                End With
            Case "ComboBox"
                With CType(ctl, Windows.Forms.ComboBox)
                    .Enabled = Enable
                    If Not .Focused Then .SelectionLength = 0
                    If Clear Then .SelectedIndex = -1
                    .ForeColor = ForeColor : .BackColor = BackColor
                End With
            Case "DateTimePicker"
                With CType(ctl, DateTimePicker)
                    .Enabled = Enable
                    .CalendarMonthBackground = BackColor
                    .CalendarTitleBackColor = BackColor
                End With
            Case "Form" : CType(ctl, Form).Enabled = Enable
            Case "GroupBox" : EnableControls(CType(ctl, GroupBox).Controls, Enable, Clear)
            Case "HScrollBar"
            Case "Label"
            Case "NumericUpDown"
                With CType(ctl, Windows.Forms.NumericUpDown)
                    .Enabled = Enable
                    '.ReadOnly = Not Enable
                    If Clear Then .Value = .Minimum
                    .ForeColor = ForeColor : .BackColor = BackColor
                End With
            Case "PictureBox" : CType(ctl, PictureBox).Enabled = Enable
            Case "RichTextBox"
                With CType(ctl, Windows.Forms.RichTextBox)
                    '.Enabled = Enable  'Allow scrolling, etc.
                    '.ReadOnly = Not Enable
                    If Clear Then .Text = ""
                    .ForeColor = ForeColor : .BackColor = BackColor
                End With
            Case "StatusBar", "StatusStrip", "MenuStrip"
            Case "TabControl" : EnableControls(CType(ctl, TabControl).Controls, Enable, Clear)
            Case "TabPage" : EnableControls(CType(ctl, TabPage).Controls, Enable, Clear)
            Case "TextBox"
                With CType(ctl, Windows.Forms.TextBox)
                    .Enabled = Enable
                    '.ReadOnly = Not Enable
                    If Clear Then .Text = ""
                    .ForeColor = ForeColor : .BackColor = BackColor
                End With
            Case "TreeView"
                With CType(ctl, TreeView)
                    .Enabled = Enable
                    .ForeColor = ForeColor : .BackColor = BackColor
                End With
            Case "ToolBar"
            Case "VScrollBar"
            Case Else : Throw New Exception(String.Format("Unexpected control type ({0}) encountered in {1}(). Control: {2}", TypeName(ctl), "EnableControl", ctl.Name))
        End Select
    End Sub
    Protected Friend Sub EnableControls(ByVal pControls As Control.ControlCollection, ByVal Enable As Boolean, ByVal Clear As Boolean)
        For Each ctl As Control In pControls
            EnableControl(ctl, Enable, Clear)
        Next
    End Sub
    Protected Function FindItem(ByVal cbItem As ComboBox, ByVal ItemCode As Short) As Integer
        FindItem = -1
        For i As Integer = 0 To cbItem.Items.Count - 1
            If CType(cbItem.Items(i), ItemData).ItemCode = ItemCode Then Return i
        Next i
    End Function
    Protected Sub InitItem(ByVal iItem As Short, ByVal Equipped As Boolean, ByVal Cursed As Boolean, ByVal IDed As Boolean, ByVal ItemCode As Short)
        Dim chkEquipped As CheckBox = Nothing
        Dim chkCursed As CheckBox = Nothing
        Dim chkID As CheckBox = Nothing
        Dim cbItem As ComboBox = Nothing
        Select Case iItem
            Case 1 : chkEquipped = Me.chkEquipped1 : chkCursed = Me.chkCursed1 : chkID = Me.chkID1 : cbItem = Me.cbItem1
            Case 2 : chkEquipped = Me.chkEquipped2 : chkCursed = Me.chkCursed2 : chkID = Me.chkID2 : cbItem = Me.cbItem2
            Case 3 : chkEquipped = Me.chkEquipped3 : chkCursed = Me.chkCursed3 : chkID = Me.chkID3 : cbItem = Me.cbItem3
            Case 4 : chkEquipped = Me.chkEquipped4 : chkCursed = Me.chkCursed4 : chkID = Me.chkID4 : cbItem = Me.cbItem4
            Case 5 : chkEquipped = Me.chkEquipped5 : chkCursed = Me.chkCursed5 : chkID = Me.chkID5 : cbItem = Me.cbItem5
            Case 6 : chkEquipped = Me.chkEquipped6 : chkCursed = Me.chkCursed6 : chkID = Me.chkID6 : cbItem = Me.cbItem6
            Case 7 : chkEquipped = Me.chkEquipped7 : chkCursed = Me.chkCursed7 : chkID = Me.chkID7 : cbItem = Me.cbItem7
            Case 8 : chkEquipped = Me.chkEquipped8 : chkCursed = Me.chkCursed8 : chkID = Me.chkID8 : cbItem = Me.cbItem8
        End Select
        chkEquipped.Checked = Equipped : chkCursed.Checked = Cursed : chkID.Checked = IDed : cbItem.SelectedIndex = FindItem(cbItem, ItemCode)
    End Sub
    Protected Overridable Sub ProtectItems(ByVal Count As Short)
        Dim chkEquipped As CheckBox = Nothing
        Dim chkCursed As CheckBox = Nothing
        Dim chkID As CheckBox = Nothing
        Dim cbItem As ComboBox = Nothing
        For i As Short = 1 To 8
            Select Case i
                Case 1 : chkEquipped = Me.chkEquipped1 : chkCursed = Me.chkCursed1 : chkID = Me.chkID1 : cbItem = Me.cbItem1
                Case 2 : chkEquipped = Me.chkEquipped2 : chkCursed = Me.chkCursed2 : chkID = Me.chkID2 : cbItem = Me.cbItem2
                Case 3 : chkEquipped = Me.chkEquipped3 : chkCursed = Me.chkCursed3 : chkID = Me.chkID3 : cbItem = Me.cbItem3
                Case 4 : chkEquipped = Me.chkEquipped4 : chkCursed = Me.chkCursed4 : chkID = Me.chkID4 : cbItem = Me.cbItem4
                Case 5 : chkEquipped = Me.chkEquipped5 : chkCursed = Me.chkCursed5 : chkID = Me.chkID5 : cbItem = Me.cbItem5
                Case 6 : chkEquipped = Me.chkEquipped6 : chkCursed = Me.chkCursed6 : chkID = Me.chkID6 : cbItem = Me.cbItem6
                Case 7 : chkEquipped = Me.chkEquipped7 : chkCursed = Me.chkCursed7 : chkID = Me.chkID7 : cbItem = Me.cbItem7
                Case 8 : chkEquipped = Me.chkEquipped8 : chkCursed = Me.chkCursed8 : chkID = Me.chkID8 : cbItem = Me.cbItem8
            End Select
            EnableControl(chkEquipped, False) : EnableControl(chkCursed, False) : EnableControl(chkID, False) : EnableControl(cbItem, False)
            If i <= Count Then
                EnableControl(chkEquipped, True) : EnableControl(chkCursed, True) : EnableControl(chkID, True) : EnableControl(cbItem, True)
            End If
        Next i
    End Sub
    Protected Friend Overridable Sub ToggleEditMode(ByVal EditMode As Boolean)
        mEditMode = EditMode
        Me.EnableControls(Me.Controls, EditMode, False)
        Me.EnableControl(Me.cbCharacter, Not EditMode)
        Me.EnableControl(Me.cmdEdit, Not EditMode) : Me.cmdEdit.Visible = Not EditMode
        Me.EnableControl(Me.cmdExit, Not EditMode) : Me.cmdExit.Visible = Not EditMode
        Me.EnableControl(Me.cmdCancel, EditMode) : Me.cmdCancel.Visible = EditMode
        Me.EnableControl(Me.cmdSave, EditMode) : Me.cmdSave.Visible = EditMode
        If EditMode Then ProtectItems(Me.nudItemCount.Value)
        tsslStatus.Visible = EditMode
    End Sub
#End Region
#Region "Event Handlers"
    Protected Overridable Sub cbCharacter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbCharacter.SelectedIndexChanged
        Try
            Me.epMain.SetError(sender, "")
            mCharacter = mBase.GetCharacter(cbCharacter.Items(cbCharacter.SelectedIndex))
            'Populate our controls with Character data...
            With mCharacter
                'Statistics...
                Me.txtName.Text = .Name
                Me.txtPassword.Text = .Password

                Me.cbAlignment.SelectedIndex = .Alignment
                Me.cbProfession.SelectedIndex = .Profession
                Me.cbRace.SelectedIndex = .Race '- 1 'Starts with 1 rather than the normal zero
                Me.cbStatus.SelectedIndex = .StatusCode

                Me.nudStrength.Value = .Strength
                Me.nudIntelligence.Value = .Intelligence
                Me.nudPiety.Value = .Piety
                Me.nudVitality.Value = .Vitality
                Me.nudAgility.Value = .Agility
                Me.nudLuck.Value = .Luck

                Me.nudAC.Value = .ArmorClass : Me.ttMain.SetToolTip(Me.lblAC, "Armor Class") : Me.ttMain.SetToolTip(Me.nudAC, "Armor Class")
                Me.nudAgeInWeeks.Value = .AgeInWeeks : Me.nudAgeInYears.Value = .Age
                Me.nudHitPoints.Value = .HP.Current : Me.nudHitPointsMax.Value = .HP.Maximum
                Me.nudLevel.Value = .LVL.Current : Me.nudLevelMax.Value = .LVL.Maximum
                Me.ttMain.SetToolTip(Me.lblLocation, .Location)
                Me.nudLocationEast.Value = .LocationEast : Me.ttMain.SetToolTip(Me.lblLocation, .Location)
                Me.nudLocationNorth.Value = .LocationNorth : Me.ttMain.SetToolTip(Me.lblLocation, .Location)
                Me.nudLocationDown.Value = .LocationDown : Me.ttMain.SetToolTip(Me.lblLocation, .Location)
                Me.nudEXP.Value = .Experience
                Me.nudGold.Value = .Gold

                'Items...
                Me.ClearItems()
                Me.nudItemCount.Value = .ItemCount
                Me.InitItem(1, .Items(0).Equipped, .Items(0).Cursed, .Items(0).Identified, .Items(0).ItemCode)
                Me.InitItem(2, .Items(1).Equipped, .Items(1).Cursed, .Items(1).Identified, .Items(1).ItemCode)
                Me.InitItem(3, .Items(2).Equipped, .Items(2).Cursed, .Items(2).Identified, .Items(2).ItemCode)
                Me.InitItem(4, .Items(3).Equipped, .Items(3).Cursed, .Items(3).Identified, .Items(3).ItemCode)
                Me.InitItem(5, .Items(4).Equipped, .Items(4).Cursed, .Items(4).Identified, .Items(4).ItemCode)
                Me.InitItem(6, .Items(5).Equipped, .Items(5).Cursed, .Items(5).Identified, .Items(5).ItemCode)
                Me.InitItem(7, .Items(6).Equipped, .Items(6).Cursed, .Items(6).Identified, .Items(6).ItemCode)
                Me.InitItem(8, .Items(7).Equipped, .Items(7).Cursed, .Items(7).Identified, .Items(7).ItemCode)

                'SpellBooks...
                For iSpell As Short = 1 To mBase.MageSpellBook.GetUpperBound(0)
                    If .MageSpellBook(iSpell) Then Me.clbMageSpells.SetItemChecked(iSpell - 1, True)
                Next
                For iSpell As Short = 1 To mBase.PriestSpellBook.GetUpperBound(0) + 1
                    If .PriestSpellBook(iSpell) Then Me.clbPriestSpells.SetItemChecked(iSpell - 1, True)
                Next
                Me.nudMageSP1.Value = .MageSpellPoints(0) : Me.nudMageSP2.Value = .MageSpellPoints(1) : Me.nudMageSP3.Value = .MageSpellPoints(2) : Me.nudMageSP4.Value = .MageSpellPoints(3) : Me.nudMageSP5.Value = .MageSpellPoints(4) : Me.nudMageSP6.Value = .MageSpellPoints(5) : Me.nudMageSP7.Value = .MageSpellPoints(6)
                Me.nudPriestSP1.Value = .PriestSpellPoints(0) : Me.nudPriestSP2.Value = .PriestSpellPoints(1) : Me.nudPriestSP3.Value = .PriestSpellPoints(2) : Me.nudPriestSP4.Value = .PriestSpellPoints(3) : Me.nudPriestSP5.Value = .PriestSpellPoints(4) : Me.nudPriestSP6.Value = .PriestSpellPoints(5) : Me.nudPriestSP7.Value = .PriestSpellPoints(6)

                ToggleEditMode(False)
            End With
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Me.epMain.SetError(sender, ex.ToString)
        End Try
    End Sub
    Private Sub cbItem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbItem1.SelectedIndexChanged, cbItem2.SelectedIndexChanged, cbItem3.SelectedIndexChanged, cbItem4.SelectedIndexChanged, cbItem5.SelectedIndexChanged, cbItem6.SelectedIndexChanged, cbItem7.SelectedIndexChanged, cbItem8.SelectedIndexChanged
        Dim cbItem As ComboBox = CType(sender, ComboBox)
        Select Case cbItem.Name
            Case "cbItem1" : Me.ttMain.SetToolTip(Me.lblItem1, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem2" : Me.ttMain.SetToolTip(Me.lblItem2, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem3" : Me.ttMain.SetToolTip(Me.lblItem3, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem4" : Me.ttMain.SetToolTip(Me.lblItem4, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem5" : Me.ttMain.SetToolTip(Me.lblItem5, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem6" : Me.ttMain.SetToolTip(Me.lblItem6, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem7" : Me.ttMain.SetToolTip(Me.lblItem7, CType(cbItem.SelectedItem, ItemData).Tag)
        End Select
    End Sub
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        ToggleEditMode(False)
        'Use our event handler to re-read the character...
        cbCharacter_SelectedIndexChanged(cbCharacter, e)
    End Sub
    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        ToggleEditMode(True)
    End Sub
    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub
    Protected Overridable Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        Try
            Me.epMain.SetError(sender, "")
            mCharacter = mBase.GetCharacter(cbCharacter.Items(cbCharacter.SelectedIndex))
            'Populate our controls with Character data...
            With mCharacter
                'Statistics...
                .Name = Me.txtName.Text.Trim
                .Password = Me.txtPassword.Text.Trim

                .Alignment = Me.cbAlignment.SelectedIndex
                .Profession = Me.cbProfession.SelectedIndex
                .Race = Me.cbRace.SelectedIndex '+ 1 'Starts with 1 rather than the normal zero
                .StatusCode = Me.cbStatus.SelectedIndex

                .Strength = Me.nudStrength.Value
                .Intelligence = Me.nudIntelligence.Value
                .Piety = Me.nudPiety.Value
                .Vitality = Me.nudVitality.Value
                .Agility = Me.nudAgility.Value
                .Luck = Me.nudLuck.Value

                .ArmorClass = Me.nudAC.Value
                .HP.Current = Me.nudHitPoints.Value : .HP.Maximum = Me.nudHitPointsMax.Value
                .LVL.Current = Me.nudLevel.Value : .LVL.Maximum = Me.nudLevelMax.Value
                .LocationEast = Me.nudLocationEast.Value
                .LocationNorth = Me.nudLocationNorth.Value
                .LocationDown = Me.nudLocationDown.Value

                .AgeInWeeks = Me.nudAgeInWeeks.Value

                .Experience = Me.nudEXP.Value
                .Gold = Me.nudGold.Value

                'Items...
                .ItemCount = Me.nudItemCount.Value
                If .ItemCount >= 1 Then .Items(0).Equipped = Me.chkEquipped1.Checked : .Items(0).Cursed = Me.chkCursed1.Checked : .Items(0).Identified = Me.chkID1.Checked : .Items(0).ItemCode = CType(Me.cbItem1.SelectedItem, ItemData).ItemCode
                If .ItemCount >= 2 Then .Items(1).Equipped = Me.chkEquipped2.Checked : .Items(1).Cursed = Me.chkCursed2.Checked : .Items(1).Identified = Me.chkID2.Checked : .Items(1).ItemCode = CType(Me.cbItem2.SelectedItem, ItemData).ItemCode
                If .ItemCount >= 3 Then .Items(2).Equipped = Me.chkEquipped3.Checked : .Items(2).Cursed = Me.chkCursed3.Checked : .Items(2).Identified = Me.chkID3.Checked : .Items(2).ItemCode = CType(Me.cbItem3.SelectedItem, ItemData).ItemCode
                If .ItemCount >= 4 Then .Items(3).Equipped = Me.chkEquipped4.Checked : .Items(3).Cursed = Me.chkCursed4.Checked : .Items(3).Identified = Me.chkID4.Checked : .Items(3).ItemCode = CType(Me.cbItem4.SelectedItem, ItemData).ItemCode
                If .ItemCount >= 5 Then .Items(4).Equipped = Me.chkEquipped5.Checked : .Items(4).Cursed = Me.chkCursed5.Checked : .Items(4).Identified = Me.chkID5.Checked : .Items(4).ItemCode = CType(Me.cbItem5.SelectedItem, ItemData).ItemCode
                If .ItemCount >= 6 Then .Items(5).Equipped = Me.chkEquipped6.Checked : .Items(5).Cursed = Me.chkCursed6.Checked : .Items(5).Identified = Me.chkID6.Checked : .Items(5).ItemCode = CType(Me.cbItem6.SelectedItem, ItemData).ItemCode
                If .ItemCount >= 7 Then .Items(6).Equipped = Me.chkEquipped7.Checked : .Items(6).Cursed = Me.chkCursed7.Checked : .Items(6).Identified = Me.chkID7.Checked : .Items(6).ItemCode = CType(Me.cbItem7.SelectedItem, ItemData).ItemCode
                If .ItemCount >= 8 Then .Items(7).Equipped = Me.chkEquipped8.Checked : .Items(7).Cursed = Me.chkCursed8.Checked : .Items(7).Identified = Me.chkID8.Checked : .Items(7).ItemCode = CType(Me.cbItem8.SelectedItem, ItemData).ItemCode

                'SpellBooks...
                For iSpell As Short = 1 To mBase.MageSpellBook.GetUpperBound(0)
                    .MageSpellBook(iSpell) = Me.clbMageSpells.GetItemChecked(iSpell - 1)
                Next
                For iSpell As Short = 1 To mBase.PriestSpellBook.GetUpperBound(0) + 1
                    .PriestSpellBook(iSpell) = Me.clbPriestSpells.GetItemChecked(iSpell - 1)
                Next
                .MageSpellPoints(0) = Me.nudMageSP1.Value : .MageSpellPoints(1) = Me.nudMageSP2.Value : .MageSpellPoints(2) = Me.nudMageSP3.Value : .MageSpellPoints(3) = Me.nudMageSP4.Value : .MageSpellPoints(4) = Me.nudMageSP5.Value : .MageSpellPoints(5) = Me.nudMageSP6.Value : .MageSpellPoints(6) = Me.nudMageSP7.Value
                .PriestSpellPoints(0) = Me.nudPriestSP1.Value : .PriestSpellPoints(1) = Me.nudPriestSP2.Value : .PriestSpellPoints(2) = Me.nudPriestSP3.Value : .PriestSpellPoints(3) = Me.nudPriestSP4.Value : .PriestSpellPoints(4) = Me.nudPriestSP5.Value : .PriestSpellPoints(5) = Me.nudPriestSP6.Value : .PriestSpellPoints(6) = Me.nudPriestSP7.Value

                mChanged = True
                ToggleEditMode(False)
            End With
        Catch ex As Exception : Debug.WriteLine(ex.ToString)
            Me.epMain.SetError(sender, ex.ToString)
        End Try
    End Sub
    Protected Overridable Sub Form_Closing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If DesignMode Then Exit Sub   'Form Designer
        If mEditMode Then Beep() : e.Cancel = True : Exit Sub
        If mChanged Then
            Select Case MessageBox.Show(Me, "Update Save Game file?", "Save?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                Case DialogResult.Yes : mBase.Save()
                Case DialogResult.No
                Case DialogResult.Cancel : e.Cancel = True
            End Select
        End If
    End Sub
    Protected Overridable Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        If DesignMode Then Exit Sub   'Form Designer
        Try
            'Populate our List controls here - not specific to any given Character...
            For iChar As Short = 0 To mBase.Characters.Length - 1
                If mBase.Characters(iChar).Name <> "" Then cbCharacter.Items.Add(mBase.Characters(iChar).Tag)
            Next iChar
            Me.cbAlignment.Items.AddRange(mBase.AlignmentList)
            Me.cbProfession.Items.AddRange(mBase.ProfessionList)
            Me.cbRace.Items.AddRange(mBase.RaceList)
            Me.cbStatus.Items.AddRange(mBase.StatusList)
            Me.clbHonors.Items.AddRange(mBase.HonorsList)

            Me.cbItem1.Items.AddRange(mBase.MasterItemList)
            Me.cbItem2.Items.AddRange(mBase.MasterItemList)
            Me.cbItem3.Items.AddRange(mBase.MasterItemList)
            Me.cbItem4.Items.AddRange(mBase.MasterItemList)
            Me.cbItem5.Items.AddRange(mBase.MasterItemList)
            Me.cbItem6.Items.AddRange(mBase.MasterItemList)
            Me.cbItem7.Items.AddRange(mBase.MasterItemList)
            Me.cbItem8.Items.AddRange(mBase.MasterItemList)

            Me.clbMageSpells.Items.AddRange(mBase.MageSpellList)
            Me.clbPriestSpells.Items.AddRange(mBase.PriestSpellList)

            ToggleEditMode(False)
            tsslMessage.Text = mBase.ScenarioDataPath
            timMain_Tick(Nothing, Nothing)
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Throw
        End Try
    End Sub
    Private Sub nudAge_ValueChanged(sender As Object, e As EventArgs) Handles nudAgeInWeeks.ValueChanged, nudAgeInYears.ValueChanged
        Dim nud As NumericUpDown = sender
        If nud Is Me.nudAgeInWeeks AndAlso nud.Focused Then Me.nudAgeInYears.Value = CInt(nud.Value * 52)
        If nud Is Me.nudAgeInYears AndAlso nud.Focused Then Me.nudAgeInWeeks.Value = CInt(nud.Value \ 52)
    End Sub
    Private Sub nudItemCount_ValueChanged(sender As Object, e As EventArgs) Handles nudItemCount.ValueChanged
        If mEditMode Then ProtectItems(Me.nudItemCount.Value)
    End Sub
    Private Sub timMain_Tick(sender As Object, e As EventArgs) Handles timMain.Tick
        tsslTime.Text = String.Format("{0:hh:mm tt}", Now)
    End Sub
#End Region
End Class