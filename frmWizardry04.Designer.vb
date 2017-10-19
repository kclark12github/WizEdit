<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWizardry04
    Inherits WizEdit.frmWizardry15Base

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tpGroups = New System.Windows.Forms.TabPage()
        Me.lblGroups = New System.Windows.Forms.Label()
        Me.lblGroupCount = New System.Windows.Forms.Label()
        Me.lblGroup3 = New System.Windows.Forms.Label()
        Me.lblGroup2 = New System.Windows.Forms.Label()
        Me.lblGroup1 = New System.Windows.Forms.Label()
        Me.nudGroup3 = New System.Windows.Forms.NumericUpDown()
        Me.nudGroup2 = New System.Windows.Forms.NumericUpDown()
        Me.nudGroup1 = New System.Windows.Forms.NumericUpDown()
        Me.cbGroup3 = New System.Windows.Forms.ComboBox()
        Me.cbGroup2 = New System.Windows.Forms.ComboBox()
        Me.cbGroup1 = New System.Windows.Forms.ComboBox()
        CType(Me.pbBoxArt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbWizardry, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpStats.SuspendLayout()
        CType(Me.nudLevel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudHitPoints, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudLocationDown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudLocationNorth, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudLocationEast, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudIntelligence, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPiety, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudVitality, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudAgility, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudLuck, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudStrength, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudAC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpItems.SuspendLayout()
        Me.tpSpellBooks.SuspendLayout()
        CType(Me.nudLevelMax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudHitPointsMax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.epMain, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudMageSP2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPriestSP1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPriestSP2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPriestSP3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPriestSP4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudMageSP4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPriestSP5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudMageSP5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPriestSP6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudMageSP6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudMageSP7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPriestSP7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudMageSP1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudMageSP3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudItemCount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tcWiz.SuspendLayout()
        Me.tpGroups.SuspendLayout()
        CType(Me.nudGroup3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudGroup2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pbBoxArt
        '
        Me.pbBoxArt.Image = Global.WizEdit.My.Resources.Resources.Wiz04Box
        Me.pbBoxArt.Location = New System.Drawing.Point(515, 33)
        Me.pbBoxArt.Size = New System.Drawing.Size(113, 168)
        '
        'lblExperience
        '
        Me.lblExperience.Location = New System.Drawing.Point(245, 292)
        Me.lblExperience.Size = New System.Drawing.Size(38, 16)
        Me.lblExperience.Text = "Keys"
        '
        'clbHonors
        '
        Me.clbHonors.Visible = False
        '
        'lblHonors
        '
        Me.lblHonors.Visible = False
        '
        'tcWiz
        '
        Me.tcWiz.Controls.Add(Me.tpGroups)
        Me.tcWiz.Controls.SetChildIndex(Me.tpSpellBooks, 0)
        Me.tcWiz.Controls.SetChildIndex(Me.tpItems, 0)
        Me.tcWiz.Controls.SetChildIndex(Me.tpStats, 0)
        Me.tcWiz.Controls.SetChildIndex(Me.tpGroups, 0)
        '
        'tpGroups
        '
        Me.tpGroups.BackgroundImage = Global.WizEdit.My.Resources.Resources.backgnd3b
        Me.tpGroups.Controls.Add(Me.lblGroups)
        Me.tpGroups.Controls.Add(Me.lblGroupCount)
        Me.tpGroups.Controls.Add(Me.lblGroup3)
        Me.tpGroups.Controls.Add(Me.lblGroup2)
        Me.tpGroups.Controls.Add(Me.lblGroup1)
        Me.tpGroups.Controls.Add(Me.nudGroup3)
        Me.tpGroups.Controls.Add(Me.nudGroup2)
        Me.tpGroups.Controls.Add(Me.nudGroup1)
        Me.tpGroups.Controls.Add(Me.cbGroup3)
        Me.tpGroups.Controls.Add(Me.cbGroup2)
        Me.tpGroups.Controls.Add(Me.cbGroup1)
        Me.tpGroups.Location = New System.Drawing.Point(4, 25)
        Me.tpGroups.Margin = New System.Windows.Forms.Padding(4)
        Me.tpGroups.Name = "tpGroups"
        Me.tpGroups.Padding = New System.Windows.Forms.Padding(4)
        Me.tpGroups.Size = New System.Drawing.Size(436, 467)
        Me.tpGroups.TabIndex = 3
        Me.tpGroups.Text = "Groups"
        Me.tpGroups.UseVisualStyleBackColor = True
        '
        'lblGroups
        '
        Me.lblGroups.AutoSize = True
        Me.lblGroups.Location = New System.Drawing.Point(126, 85)
        Me.lblGroups.Name = "lblGroups"
        Me.lblGroups.Size = New System.Drawing.Size(52, 16)
        Me.lblGroups.TabIndex = 10
        Me.lblGroups.Text = "Groups"
        '
        'lblGroupCount
        '
        Me.lblGroupCount.AutoSize = True
        Me.lblGroupCount.Location = New System.Drawing.Point(72, 85)
        Me.lblGroupCount.Name = "lblGroupCount"
        Me.lblGroupCount.Size = New System.Drawing.Size(42, 16)
        Me.lblGroupCount.TabIndex = 9
        Me.lblGroupCount.Text = "Count"
        '
        'lblGroup3
        '
        Me.lblGroup3.AutoSize = True
        Me.lblGroup3.Location = New System.Drawing.Point(13, 171)
        Me.lblGroup3.Name = "lblGroup3"
        Me.lblGroup3.Size = New System.Drawing.Size(55, 16)
        Me.lblGroup3.TabIndex = 8
        Me.lblGroup3.Text = "Group 3"
        '
        'lblGroup2
        '
        Me.lblGroup2.AutoSize = True
        Me.lblGroup2.Location = New System.Drawing.Point(14, 139)
        Me.lblGroup2.Name = "lblGroup2"
        Me.lblGroup2.Size = New System.Drawing.Size(55, 16)
        Me.lblGroup2.TabIndex = 7
        Me.lblGroup2.Text = "Group 2"
        '
        'lblGroup1
        '
        Me.lblGroup1.AutoSize = True
        Me.lblGroup1.Location = New System.Drawing.Point(14, 106)
        Me.lblGroup1.Name = "lblGroup1"
        Me.lblGroup1.Size = New System.Drawing.Size(55, 16)
        Me.lblGroup1.TabIndex = 6
        Me.lblGroup1.Text = "Group 1"
        '
        'nudGroup3
        '
        Me.nudGroup3.BackColor = System.Drawing.Color.LightGray
        Me.nudGroup3.Location = New System.Drawing.Point(75, 169)
        Me.nudGroup3.Maximum = New Decimal(New Integer() {9999, 0, 0, 0})
        Me.nudGroup3.Name = "nudGroup3"
        Me.nudGroup3.Size = New System.Drawing.Size(47, 22)
        Me.nudGroup3.TabIndex = 5
        Me.nudGroup3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'nudGroup2
        '
        Me.nudGroup2.BackColor = System.Drawing.Color.LightGray
        Me.nudGroup2.Location = New System.Drawing.Point(75, 137)
        Me.nudGroup2.Maximum = New Decimal(New Integer() {9999, 0, 0, 0})
        Me.nudGroup2.Name = "nudGroup2"
        Me.nudGroup2.Size = New System.Drawing.Size(47, 22)
        Me.nudGroup2.TabIndex = 4
        Me.nudGroup2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'nudGroup1
        '
        Me.nudGroup1.BackColor = System.Drawing.Color.LightGray
        Me.nudGroup1.Location = New System.Drawing.Point(75, 104)
        Me.nudGroup1.Maximum = New Decimal(New Integer() {9999, 0, 0, 0})
        Me.nudGroup1.Name = "nudGroup1"
        Me.nudGroup1.Size = New System.Drawing.Size(47, 22)
        Me.nudGroup1.TabIndex = 3
        Me.nudGroup1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'cbGroup3
        '
        Me.cbGroup3.BackColor = System.Drawing.Color.Gray
        Me.cbGroup3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbGroup3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbGroup3.FormattingEnabled = True
        Me.cbGroup3.Location = New System.Drawing.Point(129, 167)
        Me.cbGroup3.Margin = New System.Windows.Forms.Padding(4)
        Me.cbGroup3.Name = "cbGroup3"
        Me.cbGroup3.Size = New System.Drawing.Size(288, 24)
        Me.cbGroup3.TabIndex = 2
        '
        'cbGroup2
        '
        Me.cbGroup2.BackColor = System.Drawing.Color.Gray
        Me.cbGroup2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbGroup2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbGroup2.FormattingEnabled = True
        Me.cbGroup2.Location = New System.Drawing.Point(129, 135)
        Me.cbGroup2.Margin = New System.Windows.Forms.Padding(4)
        Me.cbGroup2.Name = "cbGroup2"
        Me.cbGroup2.Size = New System.Drawing.Size(288, 24)
        Me.cbGroup2.TabIndex = 1
        '
        'cbGroup1
        '
        Me.cbGroup1.BackColor = System.Drawing.Color.Gray
        Me.cbGroup1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbGroup1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbGroup1.FormattingEnabled = True
        Me.cbGroup1.Location = New System.Drawing.Point(129, 103)
        Me.cbGroup1.Margin = New System.Windows.Forms.Padding(4)
        Me.cbGroup1.Name = "cbGroup1"
        Me.cbGroup1.Size = New System.Drawing.Size(288, 24)
        Me.cbGroup1.TabIndex = 0
        '
        'frmWizardry04
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.ClientSize = New System.Drawing.Size(644, 627)
        Me.Name = "frmWizardry04"
        Me.Text = "frmWizardry04"
        CType(Me.pbBoxArt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbWizardry, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpStats.ResumeLayout(False)
        Me.tpStats.PerformLayout()
        CType(Me.nudLevel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudHitPoints, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudLocationDown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudLocationNorth, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudLocationEast, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudIntelligence, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPiety, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudVitality, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudAgility, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudLuck, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudStrength, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudAC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpItems.ResumeLayout(False)
        Me.tpItems.PerformLayout()
        Me.tpSpellBooks.ResumeLayout(False)
        Me.tpSpellBooks.PerformLayout()
        CType(Me.nudLevelMax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudHitPointsMax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.epMain, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudMageSP2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPriestSP1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPriestSP2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPriestSP3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPriestSP4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudMageSP4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPriestSP5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudMageSP5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPriestSP6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudMageSP6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudMageSP7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPriestSP7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudMageSP1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudMageSP3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudItemCount, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tcWiz.ResumeLayout(False)
        Me.tpGroups.ResumeLayout(False)
        Me.tpGroups.PerformLayout()
        CType(Me.nudGroup3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudGroup2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudGroup1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tpGroups As TabPage
    Friend WithEvents cbGroup1 As ComboBox
    Friend WithEvents nudGroup1 As NumericUpDown
    Friend WithEvents cbGroup3 As ComboBox
    Friend WithEvents cbGroup2 As ComboBox
    Friend WithEvents lblGroup1 As Label
    Friend WithEvents nudGroup3 As NumericUpDown
    Friend WithEvents nudGroup2 As NumericUpDown
    Friend WithEvents lblGroups As Label
    Friend WithEvents lblGroupCount As Label
    Friend WithEvents lblGroup3 As Label
    Friend WithEvents lblGroup2 As Label
End Class
