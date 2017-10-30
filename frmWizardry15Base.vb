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
    Protected Sub ClearSpellBooks(ByVal Mage As Boolean, ByVal Priest As Boolean)
        If Mage Then
            For iSpell As Short = 1 To mBase.MageSpellBook.GetUpperBound(0)
                Me.clbMageSpells.SetItemChecked(iSpell - 1, False)
            Next
            Me.nudMageSP1.Value = 0 : Me.nudMageSP2.Value = 0 : Me.nudMageSP3.Value = 0 : Me.nudMageSP4.Value = 0 : Me.nudMageSP5.Value = 0 : Me.nudMageSP6.Value = 0 : Me.nudMageSP7.Value = 0
        End If
        If Priest Then
            For iSpell As Short = 1 To mBase.PriestSpellBook.GetUpperBound(0) + 1
                Me.clbPriestSpells.SetItemChecked(iSpell - 1, False)
            Next
            Me.nudPriestSP1.Value = 0 : Me.nudPriestSP2.Value = 0 : Me.nudPriestSP3.Value = 0 : Me.nudPriestSP4.Value = 0 : Me.nudPriestSP5.Value = 0 : Me.nudPriestSP6.Value = 0 : Me.nudPriestSP7.Value = 0
        End If
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
        Dim chkEquipped As CheckBox = Nothing, chkCursed As CheckBox = Nothing, chkID As CheckBox = Nothing, cbItem As ComboBox = Nothing
        For iItem As Short = 1 To 8
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
            Me.EnableControl(chkEquipped, False) : Me.EnableControl(chkCursed, False) : Me.EnableControl(chkID, False) : Me.EnableControl(cbItem, False)
            If iItem <= Count Then
                Me.EnableControl(chkEquipped, True) : Me.EnableControl(chkCursed, True) : Me.EnableControl(chkID, True) : Me.EnableControl(cbItem, True)
            End If
        Next iItem
    End Sub
    Protected Overridable Sub ProtectSpellBooks(ByVal Profession As CharacterBase.enumProfession)
        Dim gotMage As Boolean = False
        Dim gotPriest As Boolean = False
        Select Case Profession
            Case CharacterBase.enumProfession.Fighter, CharacterBase.enumProfession.Thief, CharacterBase.enumProfession.Ninja
                gotMage = False : gotPriest = False
            Case CharacterBase.enumProfession.Mage, CharacterBase.enumProfession.Samurai
                gotMage = True : gotPriest = False
            Case CharacterBase.enumProfession.Priest
                gotMage = False : gotPriest = True
            Case CharacterBase.enumProfession.Bishop
                gotMage = True : gotPriest = True
            Case CharacterBase.enumProfession.Lord
                gotMage = False : gotPriest = True
                'Reset Max on Priest Spell Points...
                Me.cmdAllSP_Click(Me.cmdAllPSP, New EventArgs())
        End Select
        Me.ClearSpellBooks(Not gotMage, Not gotPriest)
        Me.EnableControl(Me.clbMageSpells, gotMage)
        Me.EnableControl(Me.cmdAllMS, gotMage) : Me.EnableControl(Me.cmdNoneMS, gotMage)
        Me.EnableControl(Me.nudMageSP1, gotMage) : Me.EnableControl(Me.nudMageSP2, gotMage) : Me.EnableControl(Me.nudMageSP3, gotMage) : Me.EnableControl(Me.nudMageSP4, gotMage) : Me.EnableControl(Me.nudMageSP5, gotMage) : Me.EnableControl(Me.nudMageSP6, gotMage) : Me.EnableControl(Me.nudMageSP7, gotMage)
        Me.EnableControl(Me.clbPriestSpells, gotPriest)
        Me.EnableControl(Me.cmdAllPS, gotPriest) : Me.EnableControl(Me.cmdNonePS, gotPriest)
        Me.EnableControl(Me.nudPriestSP1, gotPriest) : Me.EnableControl(Me.nudPriestSP2, gotPriest) : Me.EnableControl(Me.nudPriestSP3, gotPriest) : Me.EnableControl(Me.nudPriestSP4, gotPriest) : Me.EnableControl(Me.nudPriestSP5, gotPriest) : Me.EnableControl(Me.nudPriestSP6, gotPriest) : Me.EnableControl(Me.nudPriestSP7, gotPriest)
    End Sub
    Protected Overridable Sub RefreshData()
        'Populate our List controls here - not specific to any given Character...
        mBase.Read()
        Me.cbCharacter.Items.Clear()
        For iChar As Short = 0 To mBase.Characters.Length - 1
            If mBase.Characters(iChar).Name <> "" Then Me.cbCharacter.Items.Add(mBase.Characters(iChar).Tag)
        Next iChar
        Me.cbAlignment.Items.Clear() : Me.cbAlignment.Items.AddRange(mBase.AlignmentList)
        Me.cbProfession.Items.Clear() : Me.cbProfession.Items.AddRange(mBase.ProfessionList)
        Me.cbRace.Items.Clear() : Me.cbRace.Items.AddRange(mBase.RaceList)
        Me.cbStatus.Items.Clear() : Me.cbStatus.Items.AddRange(mBase.StatusList)
        Me.clbHonors.Items.Clear() : Me.clbHonors.Items.AddRange(mBase.HonorsList)
        Me.cbItem1.Items.Clear() : Me.cbItem1.Items.AddRange(mBase.MasterItemList)
        Me.cbItem2.Items.Clear() : Me.cbItem2.Items.AddRange(mBase.MasterItemList)
        Me.cbItem3.Items.Clear() : Me.cbItem3.Items.AddRange(mBase.MasterItemList)
        Me.cbItem4.Items.Clear() : Me.cbItem4.Items.AddRange(mBase.MasterItemList)
        Me.cbItem5.Items.Clear() : Me.cbItem5.Items.AddRange(mBase.MasterItemList)
        Me.cbItem6.Items.Clear() : Me.cbItem6.Items.AddRange(mBase.MasterItemList)
        Me.cbItem7.Items.Clear() : Me.cbItem7.Items.AddRange(mBase.MasterItemList)
        Me.cbItem8.Items.Clear() : Me.cbItem8.Items.AddRange(mBase.MasterItemList)
        Me.clbMageSpells.Items.Clear() : Me.clbMageSpells.Items.AddRange(mBase.MageSpellList)
        Me.clbPriestSpells.Items.Clear() : Me.clbPriestSpells.Items.AddRange(mBase.PriestSpellList)
    End Sub
    Protected Friend Overridable Sub ToggleEditMode(ByVal EditMode As Boolean)
        mEditMode = EditMode
        Me.EnableControls(Me.Controls, EditMode, False)
        Me.EnableControl(Me.cbCharacter, Not EditMode)
        Me.EnableControl(Me.cmdEdit, Not EditMode) : Me.cmdEdit.Visible = Not EditMode
        Me.EnableControl(Me.cmdExit, Not EditMode) : Me.cmdExit.Visible = Not EditMode
        Me.EnableControl(Me.cmdCancel, EditMode) : Me.cmdCancel.Visible = EditMode
        Me.EnableControl(Me.cmdSave, EditMode) : Me.cmdSave.Visible = EditMode
        If EditMode Then Me.ProtectItems(Me.nudItemCount.Value) : Me.ProtectSpellBooks(Me.cbProfession.SelectedIndex)
        Me.tsslStatus.Visible = EditMode
    End Sub
    Protected Function UpdateChangedData() As Boolean
        If mChanged Then
            Select Case MessageBox.Show(Me, "Update Save Game file?", "Save?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                Case DialogResult.Yes : mBase.Save()
                Case DialogResult.No
                Case DialogResult.Cancel : Return False
            End Select
        End If
        Return True
    End Function
#End Region
#Region "Event Handlers"
    Protected Overridable Sub cbCharacter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbCharacter.SelectedIndexChanged
        Try
            Me.epScenario.SetError(sender, "")
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

                Me.nudAC.Value = .ArmorClass : Me.ttScenario.SetToolTip(Me.lblAC, "Armor Class") : Me.ttScenario.SetToolTip(Me.nudAC, "Armor Class")
                Me.nudAgeInWeeks.Value = .AgeInWeeks : Me.nudAgeInYears.Value = .Age
                Me.nudHitPoints.Value = .HP.Current : Me.nudHitPointsMax.Value = .HP.Maximum
                Me.nudLevel.Value = .LVL.Current : Me.nudLevelMax.Value = .LVL.Maximum
                Me.ttScenario.SetToolTip(Me.lblLocation, .Location)
                Me.nudLocationEast.Value = .LocationEast : Me.ttScenario.SetToolTip(Me.lblLocation, .Location)
                Me.nudLocationNorth.Value = .LocationNorth : Me.ttScenario.SetToolTip(Me.lblLocation, .Location)
                Me.nudLocationDown.Value = .LocationDown : Me.ttScenario.SetToolTip(Me.lblLocation, .Location)
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
                    Me.clbMageSpells.SetItemChecked(iSpell - 1, .MageSpellBook(iSpell))
                Next
                For iSpell As Short = 1 To mBase.PriestSpellBook.GetUpperBound(0) + 1
                    Me.clbPriestSpells.SetItemChecked(iSpell - 1, .PriestSpellBook(iSpell))
                Next
                Me.nudMageSP1.Value = .MageSpellPoints(0) : Me.nudMageSP2.Value = .MageSpellPoints(1) : Me.nudMageSP3.Value = .MageSpellPoints(2) : Me.nudMageSP4.Value = .MageSpellPoints(3) : Me.nudMageSP5.Value = .MageSpellPoints(4) : Me.nudMageSP6.Value = .MageSpellPoints(5) : Me.nudMageSP7.Value = .MageSpellPoints(6)
                Me.nudPriestSP1.Value = .PriestSpellPoints(0) : Me.nudPriestSP2.Value = .PriestSpellPoints(1) : Me.nudPriestSP3.Value = .PriestSpellPoints(2) : Me.nudPriestSP4.Value = .PriestSpellPoints(3) : Me.nudPriestSP5.Value = .PriestSpellPoints(4) : Me.nudPriestSP6.Value = .PriestSpellPoints(5) : Me.nudPriestSP7.Value = .PriestSpellPoints(6)

                ToggleEditMode(False)
            End With
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Me.epScenario.SetError(sender, ex.ToString)
        End Try
    End Sub
    Private Sub cbItem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbItem1.SelectedIndexChanged, cbItem2.SelectedIndexChanged, cbItem3.SelectedIndexChanged, cbItem4.SelectedIndexChanged, cbItem5.SelectedIndexChanged, cbItem6.SelectedIndexChanged, cbItem7.SelectedIndexChanged, cbItem8.SelectedIndexChanged
        Dim cbItem As ComboBox = CType(sender, ComboBox)
        If cbItem.SelectedIndex = -1 Then Exit Sub
        Select Case cbItem.Name
            Case "cbItem1" : Me.ttScenario.SetToolTip(Me.lblItem1, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem2" : Me.ttScenario.SetToolTip(Me.lblItem2, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem3" : Me.ttScenario.SetToolTip(Me.lblItem3, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem4" : Me.ttScenario.SetToolTip(Me.lblItem4, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem5" : Me.ttScenario.SetToolTip(Me.lblItem5, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem6" : Me.ttScenario.SetToolTip(Me.lblItem6, CType(cbItem.SelectedItem, ItemData).Tag)
            Case "cbItem7" : Me.ttScenario.SetToolTip(Me.lblItem7, CType(cbItem.SelectedItem, ItemData).Tag)
        End Select
    End Sub
    Private Sub cbProfession_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbProfession.SelectedIndexChanged
        Const defaultMax As Short = 9
        Me.ProtectSpellBooks(Me.cbProfession.SelectedIndex)
        Select Case CType(Me.cbProfession.SelectedIndex, CharacterBase.enumProfession)
            Case CharacterBase.enumProfession.Lord
                Me.nudPriestSP1.Maximum = defaultMax
                Me.nudPriestSP2.Maximum = 7
                Me.nudPriestSP3.Maximum = 4
                Me.nudPriestSP4.Maximum = 4
                Me.nudPriestSP5.Maximum = 6
                Me.nudPriestSP6.Maximum = 4
                Me.nudPriestSP7.Maximum = 2
            Case Else
                Me.nudPriestSP1.Maximum = defaultMax
                Me.nudPriestSP2.Maximum = defaultMax
                Me.nudPriestSP3.Maximum = defaultMax
                Me.nudPriestSP4.Maximum = defaultMax
                Me.nudPriestSP5.Maximum = defaultMax
                Me.nudPriestSP6.Maximum = defaultMax
                Me.nudPriestSP7.Maximum = defaultMax
        End Select
    End Sub
    Private Sub cmdAllSpells_Click(sender As Object, e As EventArgs) Handles cmdAllMS.Click, cmdAllPS.Click
        If sender Is cmdAllMS Then
            For iSpell As Short = 1 To mBase.MageSpellBook.GetUpperBound(0)
                Me.clbMageSpells.SetItemChecked(iSpell - 1, True)
            Next
        Else
            For iSpell As Short = 1 To mBase.PriestSpellBook.GetUpperBound(0) + 1
                Me.clbPriestSpells.SetItemChecked(iSpell - 1, True)
            Next
        End If
    End Sub
    Private Sub cmdAllSP_Click(sender As Object, e As EventArgs) Handles cmdAllMSP.Click, cmdAllPSP.Click
        If sender Is cmdAllMSP Then
            Me.nudMageSP1.Value = Me.nudMageSP1.Maximum : Me.nudMageSP2.Value = Me.nudMageSP2.Maximum : Me.nudMageSP3.Value = Me.nudMageSP3.Maximum : Me.nudMageSP4.Value = Me.nudMageSP4.Maximum : Me.nudMageSP5.Value = Me.nudMageSP5.Maximum : Me.nudMageSP6.Value = Me.nudMageSP6.Maximum : Me.nudMageSP7.Value = Me.nudMageSP7.Maximum
        Else
            Me.nudPriestSP1.Value = Me.nudPriestSP1.Maximum : Me.nudPriestSP2.Value = Me.nudPriestSP2.Maximum : Me.nudPriestSP3.Value = Me.nudPriestSP3.Maximum : Me.nudPriestSP4.Value = Me.nudPriestSP4.Maximum : Me.nudPriestSP5.Value = Me.nudPriestSP5.Maximum : Me.nudPriestSP6.Value = Me.nudPriestSP6.Maximum : Me.nudPriestSP7.Value = Me.nudPriestSP7.Maximum
        End If
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
    Private Sub cmdHPReset_Click(sender As Object, e As EventArgs) Handles cmdHPReset.Click
        Me.nudHitPoints.Value = Me.nudHitPointsMax.Value
    End Sub
    Private Sub cmdLVLReset_Click(sender As Object, e As EventArgs) Handles cmdLVLReset.Click
        Me.nudLevel.Value = Me.nudLevelMax.Value
    End Sub
    Private Sub cmdNoneMS_Click(sender As Object, e As EventArgs) Handles cmdNoneMS.Click, cmdNonePS.Click
        If sender Is cmdAllMS Then
            Me.ClearSpellBooks(True, False)
        Else
            Me.ClearSpellBooks(False, True)
        End If
    End Sub
    Protected Overridable Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        Try
            Me.epScenario.SetError(sender, "")
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
                .MageSpellPoints(0) = CShort(Me.nudMageSP1.Value) : .MageSpellPoints(1) = CShort(Me.nudMageSP2.Value) : .MageSpellPoints(2) = CShort(Me.nudMageSP3.Value) : .MageSpellPoints(3) = CShort(Me.nudMageSP4.Value) : .MageSpellPoints(4) = CShort(Me.nudMageSP5.Value) : .MageSpellPoints(5) = CShort(Me.nudMageSP6.Value) : .MageSpellPoints(6) = CShort(Me.nudMageSP7.Value)
                .PriestSpellPoints(0) = CShort(Me.nudPriestSP1.Value) : .PriestSpellPoints(1) = CShort(Me.nudPriestSP2.Value) : .PriestSpellPoints(2) = CShort(Me.nudPriestSP3.Value) : .PriestSpellPoints(3) = CShort(Me.nudPriestSP4.Value) : .PriestSpellPoints(4) = CShort(Me.nudPriestSP5.Value) : .PriestSpellPoints(5) = CShort(Me.nudPriestSP6.Value) : .PriestSpellPoints(6) = CShort(Me.nudPriestSP7.Value)

                mChanged = True
                ToggleEditMode(False)
                'Since we cannot reset just our character's Tag, we have to reload the entire list based on our updated data
                Dim saveIndex As Integer = Me.cbCharacter.SelectedIndex
                Me.cbCharacter.Items.Clear()
                For iChar As Short = 0 To mBase.Characters.Length - 1
                    If mBase.Characters(iChar).Name <> "" Then Me.cbCharacter.Items.Add(mBase.Characters(iChar).Tag)
                Next iChar
                Me.cbCharacter.SelectedIndex = saveIndex
            End With
        Catch ex As Exception : Debug.WriteLine(ex.ToString)
            Me.epScenario.SetError(sender, ex.ToString)
        End Try
    End Sub
    Protected Overridable Sub Form_Closing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If DesignMode Then Exit Sub   'Form Designer
        If mEditMode Then Beep() : e.Cancel = True : Exit Sub
        If Not UpdateChangedData() Then e.Cancel = True : Exit Sub
    End Sub
    Protected Overridable Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        If DesignMode Then Exit Sub   'Form Designer
        Try
            RefreshData()
            ToggleEditMode(False)
            tsslMessage.Text = mBase.ScenarioDataPath
            timScenario_Tick(Nothing, Nothing)
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Throw
        End Try
    End Sub
    Private Sub nudAge_ValueChanged(sender As Object, e As EventArgs) Handles nudAgeInWeeks.ValueChanged, nudAgeInYears.ValueChanged
        Dim nud As NumericUpDown = sender
        If nud Is Me.nudAgeInWeeks AndAlso nud.Focused Then Me.nudAgeInYears.Value = CInt(nud.Value \ 52)
        If nud Is Me.nudAgeInYears AndAlso nud.Focused Then Me.nudAgeInWeeks.Value = CInt(nud.Value * 52)
    End Sub
    Private Sub nudLevel_ValueChanged(sender As Object, e As EventArgs) Handles nudLevel.ValueChanged ', nudEXP.ValueChanged
        Dim nud As NumericUpDown = sender
        If nud.Value > Me.nudLevelMax.Value Then Me.nudLevelMax.Value = nud.Value

        'Allow change of Level to affect Experience Points, but not vice-versa as there are still unknown fields within the
        'Character structure that may be updated when leveling through the Adventurer's Inn...
        Dim EXPNeeded4LVL As Decimal = mBase.EXPRequiredForNextLevel(CInt(nud.Value) - 2, Me.cbProfession.SelectedIndex)
        If Me.nudEXP.Value < EXPNeeded4LVL Then Me.nudEXP.Value = EXPNeeded4LVL
    End Sub
    Private Sub nudHitPoints_ValueChanged(sender As Object, e As EventArgs) Handles nudHitPoints.ValueChanged
        Dim nud As NumericUpDown = sender
        If nud.Value > Me.nudHitPointsMax.Value Then Me.nudHitPointsMax.Value = nud.Value
    End Sub
    Private Sub nudItemCount_ValueChanged(sender As Object, e As EventArgs) Handles nudItemCount.ValueChanged
        If mEditMode Then ProtectItems(Me.nudItemCount.Value)
    End Sub
    Private Sub timScenario_Tick(sender As Object, e As EventArgs) Handles timScenario.Tick
        tsslTime.Text = String.Format("{0:hh:mm tt}", Now)
    End Sub
    Protected Sub tsmiOptionsOpen_Click(sender As Object, e As EventArgs) Handles tsmiOptionsOpen.Click
        Dim Path As String = ""
        Dim fi As FileInfo = Nothing
        Try
            Try : Path = mBase.ScenarioDataPath : Catch ex As FileNotFoundException : End Try
            With ofdScenario
                If Path <> "" Then fi = New FileInfo(Path) : .InitialDirectory = fi.DirectoryName
                .FileName = ""
                Dim Scenario As String = Me.Text.Substring(9, 3).Trim.ToUpper
                Select Case Scenario
                    Case "01", "02", "03", "04", "05"
                        .Filter = String.Format("{0} Saved Games (SAVE{1}.DSK)|SAVE{1}.DSK", Me.Text, CShort(Scenario))
                    Case "06", "07"
                        .Filter = String.Format("{0} Saved Games (SCENARIO.DBS)|SCENARIO.DBS", Me.Text)
                    Case "07G"
                        .Filter = String.Format("{0} Saved Games (SCENARIO.GLD)|SCENARIO.GLD", Me.Text)
                End Select
                .FilterIndex = 0
                .ShowReadOnly = False
                .Multiselect = False
                .CheckPathExists = True
                If .ShowDialog(Me) = DialogResult.Cancel Then
                    MessageBox.Show(Me, "Operation canceled at User's request.", "Restore", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Try
                End If
                mBase.ScenarioDataPath = .FileName
                UpdateChangedData()
                RefreshData()
                tsslMessage.Text = mBase.ScenarioDataPath
            End With
        Catch ex As Exception
            MessageBox.Show(Me, String.Format("{0}{1}{2}", ex.Message, vbCrLf, ex.StackTrace), ex.GetType.Name, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Protected Sub tsmiOptionsRestore_Click(sender As Object, e As EventArgs) Handles tsmiOptionsRestore.Click
        Try
            Dim fi As FileInfo = New FileInfo(mBase.ScenarioDataPath)
            Dim fs() As String = fi.Name.Split({"."c})
            With ofdScenario
                .InitialDirectory = fi.DirectoryName
                .FileName = ""
                Dim Scenario As String = Me.Text.Substring(9, 3).Trim.ToUpper
                Select Case Scenario
                    Case "01", "02", "03", "04", "05"
                        .Filter = String.Format("{0} Saved Games (SAVE{1}.*.DSK)|SAVE{1}.*.DSK|All Files (*.*)|*.*", Me.Text, CShort(Scenario))
                    Case "06", "07"
                        .Filter = String.Format("{0} Saved Games (SCENARIO.*.DBS)|SCENARIO.*.DBS|All Files (*.*)|*.*", Me.Text)
                    Case "07G"
                        .Filter = String.Format("{0} Saved Games (SCENARIO.*.GLD)|SCENARIO.*.GLD|All Files (*.*)|*.*", Me.Text)
                End Select
                .FilterIndex = 0
                .ShowReadOnly = False
                .Multiselect = False
                .CheckPathExists = True
                If .ShowDialog(Me) = DialogResult.Cancel Then
                    MessageBox.Show(Me, "Operation canceled at User's request.", "Restore", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Try
                End If
                MessageBox.Show(Me, String.Format("Are you sure you want to replace {0} with {1}?", fi.Name, .FileName), "Restore", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'Do not use mBase.ScenarioDataPath as it checks for the existence of the file and that would screw us up attempting to replace it...
                File.Delete(fi.FullName) : File.Copy(.FileName, fi.FullName)
                RefreshData()
            End With
        Catch ex As Exception
            MessageBox.Show(Me, String.Format("{0}{1}{2}", ex.Message, vbCrLf, ex.StackTrace), ex.GetType.Name, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Protected Sub tsmiOptionsExport_Click(sender As Object, e As EventArgs) Handles tsmiOptionsExport.Click
        Dim textWriter As StreamWriter = Nothing
        Try
            Dim fi As FileInfo = New FileInfo(mBase.ScenarioDataPath)
            Dim fs() As String = fi.Name.Split({"."c})
            With sfdScenario
                .InitialDirectory = fi.DirectoryName
                Dim Scenario As String = Me.Text.Substring(9, 3).Trim.ToUpper
                .Filter = String.Format("Text (Export{0}.txt)|Export{0}.txt|All Files (*.*)|*.*", CShort(Scenario))
                .FilterIndex = 0
                .FileName = String.Format("Export{0}.txt", CShort(Scenario))
                .CheckPathExists = False
                If .ShowDialog(Me) = DialogResult.Cancel Then
                    MessageBox.Show(Me, "Operation canceled at User's request.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Try
                End If
                Dim sep1 As String = New String("="c, 132)
                Dim sep2 As String = New String("-"c, 132)
                textWriter = New StreamWriter(File.Open(.FileName, FileMode.Create))
                textWriter.WriteLine(sep1)
                For iChar As Short = 0 To cbCharacter.Items.Count - 1
                    If iChar > 0 Then textWriter.WriteLine(sep2)
                    If CShort(Scenario) = 4 Then textWriter.WriteLine(String.Format("Save Game #{0}", iChar + 1))
                    textWriter.WriteLine(mBase.GetCharacter(cbCharacter.Items(iChar)).ToString)
                Next iChar
                textWriter.WriteLine(sep1)
            End With
        Catch ex As Exception
            MessageBox.Show(Me, String.Format("{0}{1}{2}", ex.Message, vbCrLf, ex.StackTrace), ex.GetType.Name, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            If textWriter IsNot Nothing Then textWriter.Close() : textWriter = Nothing
        End Try
    End Sub
#End Region
End Class