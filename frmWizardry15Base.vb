Public Class frmWizardry15Base
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
        Me.txtAgeInWeeks.Text = 0
        Me.txtAgeInYears.Text = 0
        Me.txtExperience.Text = 0
        Me.txtGold.Text = 0
        Me.txtName.Text = ""
        Me.txtPassword.Text = ""
    End Sub

#Region "Properties"
#Region "Declarations"
    Private mBase As WizEditBase
    Private mCharacter As WizEditBase.Character = Nothing
#End Region
#End Region
#Region "Methods"
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
    Protected Friend Sub ToggleEditMode(ByVal EditMode As Boolean)
        Me.EnableControls(Me.Controls, EditMode, False)
        Me.EnableControl(Me.cbCharacter, Not EditMode)
        Me.EnableControl(Me.cmdEdit, Not EditMode) : Me.cmdEdit.Visible = Not EditMode
        Me.EnableControl(Me.cmdExit, Not EditMode) : Me.cmdExit.Visible = Not EditMode
        Me.EnableControl(Me.cmdCancel, EditMode) : Me.cmdCancel.Visible = EditMode
        Me.EnableControl(Me.cmdSave, EditMode) : Me.cmdSave.Visible = EditMode
    End Sub
#End Region
#Region "Event Handlers"
    Private Sub cbCharacter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbCharacter.SelectedIndexChanged
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
                Me.cbRace.SelectedIndex = .Race - 1 'Starts with 1 rather than the normal zero
                Me.cbStatus.SelectedIndex = .StatusCode

                Me.nudStrength.Value = .Strength
                Me.nudIntelligence.Value = .Intelligence
                Me.nudPiety.Value = .Piety
                Me.nudVitality.Value = .Vitality
                Me.nudAgility.Value = .Agility
                Me.nudLuck.Value = .Luck

                Me.nudAC.Value = .ArmorClass : Me.ttMain.SetToolTip(Me.lblAC, "Armor Class") : Me.ttMain.SetToolTip(Me.nudAC, "Armor Class")
                Me.nudHitPoints.Value = .HP.Current : Me.nudHitPointsMax.Value = .HP.Maximum
                Me.nudLevel.Value = .LVL.Current : Me.nudLevelMax.Value = .LVL.Maximum
                Me.ttMain.SetToolTip(Me.lblLocation, .Location)
                Me.nudLocationEast.Value = .LocationEast : Me.ttMain.SetToolTip(Me.lblLocation, .Location)
                Me.nudLocationNorth.Value = .LocationNorth : Me.ttMain.SetToolTip(Me.lblLocation, .Location)
                Me.nudLocationDown.Value = .LocationDown : Me.ttMain.SetToolTip(Me.lblLocation, .Location)

                Me.txtAgeInWeeks.Text = .AgeInWeeks
                Me.txtAgeInYears.Text = .Age

                Me.txtExperience.Text = .Experience
                Me.txtGold.Text = .Gold

                'Items...
                Me.chkEquipped1.Checked = .Items(0).Equipped : Me.chkCursed1.Checked = .Items(0).Cursed : Me.chkID1.Checked = .Items(0).Identified : Me.cbItem1.SelectedIndex = .Items(0).ItemCode
                Me.chkEquipped2.Checked = .Items(1).Equipped : Me.chkCursed2.Checked = .Items(1).Cursed : Me.chkID2.Checked = .Items(1).Identified : Me.cbItem2.SelectedIndex = .Items(1).ItemCode
                Me.chkEquipped3.Checked = .Items(2).Equipped : Me.chkCursed3.Checked = .Items(2).Cursed : Me.chkID3.Checked = .Items(2).Identified : Me.cbItem3.SelectedIndex = .Items(2).ItemCode
                Me.chkEquipped4.Checked = .Items(3).Equipped : Me.chkCursed4.Checked = .Items(3).Cursed : Me.chkID4.Checked = .Items(3).Identified : Me.cbItem4.SelectedIndex = .Items(3).ItemCode
                Me.chkEquipped5.Checked = .Items(4).Equipped : Me.chkCursed5.Checked = .Items(4).Cursed : Me.chkID5.Checked = .Items(4).Identified : Me.cbItem5.SelectedIndex = .Items(4).ItemCode
                Me.chkEquipped6.Checked = .Items(5).Equipped : Me.chkCursed6.Checked = .Items(5).Cursed : Me.chkID6.Checked = .Items(5).Identified : Me.cbItem6.SelectedIndex = .Items(5).ItemCode
                Me.chkEquipped7.Checked = .Items(6).Equipped : Me.chkCursed7.Checked = .Items(6).Cursed : Me.chkID7.Checked = .Items(6).Identified : Me.cbItem7.SelectedIndex = .Items(6).ItemCode
                Me.chkEquipped8.Checked = .Items(7).Equipped : Me.chkCursed8.Checked = .Items(7).Cursed : Me.chkID8.Checked = .Items(7).Identified : Me.cbItem8.SelectedIndex = .Items(7).ItemCode

                'SpellBooks...
                For i As Integer = 0 To Me.clbMageSpells.Items.Count - 1
                    If .MageSpellBook(i).Known Then Me.clbMageSpells.SetItemChecked(i, True)
                Next i
                For i As Integer = 0 To Me.clbPriestSpells.Items.Count - 1
                    If .PriestSpellBook(i).Known Then Me.clbPriestSpells.SetItemChecked(i, True)
                Next i
                Me.nudMageSP1.Value = .MageSpellPoints(0) : Me.nudMageSP2.Value = .MageSpellPoints(1) : Me.nudMageSP3.Value = .MageSpellPoints(2) : Me.nudMageSP4.Value = .MageSpellPoints(3) : Me.nudMageSP5.Value = .MageSpellPoints(4) : Me.nudMageSP6.Value = .MageSpellPoints(5) : Me.nudMageSP7.Value = .MageSpellPoints(6)
                Me.nudPriestSP1.Value = .PriestSpellPoints(0) : Me.nudPriestSP2.Value = .PriestSpellPoints(1) : Me.nudPriestSP3.Value = .PriestSpellPoints(2) : Me.nudPriestSP4.Value = .PriestSpellPoints(3) : Me.nudPriestSP5.Value = .PriestSpellPoints(4) : Me.nudPriestSP6.Value = .PriestSpellPoints(5) : Me.nudPriestSP7.Value = .PriestSpellPoints(6)

                ToggleEditMode(False)
            End With
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Me.epMain.SetError(sender, ex.ToString)
        End Try
    End Sub
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        ToggleEditMode(False)
        cbCharacter_SelectedIndexChanged(cbCharacter, e)
    End Sub
    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        ToggleEditMode(True)
    End Sub
    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
    End Sub
    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
    End Sub
    Private Sub frmWizardry15Base_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'Populate our List controls here - not specific to any given Character...
            For iChar As Short = 0 To mBase.Characters.Length - 1
                If mBase.Characters(iChar).Name <> "" Then cbCharacter.Items.Add(mBase.Characters(iChar).Name)
            Next iChar
            Me.cbAlignment.Items.AddRange(mBase.AlignmentList)
            Me.cbProfession.Items.AddRange(mBase.ProfessionList)
            Me.cbRace.Items.AddRange(mBase.RaceList)
            Me.cbStatus.Items.AddRange(mBase.StatusList)
            Me.clbHonors.Items.AddRange(mBase.HonorsList)

            Me.cbItem1.Items.AddRange(mBase.ItemList)
            Me.cbItem2.Items.AddRange(mBase.ItemList)
            Me.cbItem3.Items.AddRange(mBase.ItemList)
            Me.cbItem4.Items.AddRange(mBase.ItemList)
            Me.cbItem5.Items.AddRange(mBase.ItemList)
            Me.cbItem6.Items.AddRange(mBase.ItemList)
            Me.cbItem7.Items.AddRange(mBase.ItemList)
            Me.cbItem8.Items.AddRange(mBase.ItemList)

            Me.clbMageSpells.Items.AddRange(mBase.MageSpellList)
            Me.clbPriestSpells.Items.AddRange(mBase.PriestSpellList)

            ToggleEditMode(False)
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Throw
        End Try
    End Sub

#End Region
End Class