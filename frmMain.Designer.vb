<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.lblSelectedScenario = New System.Windows.Forms.Label()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.ssStatus = New System.Windows.Forms.StatusStrip()
        Me.tsslMessage = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslTime = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pbUWA = New System.Windows.Forms.PictureBox()
        Me.pbSirTech = New System.Windows.Forms.PictureBox()
        Me.imgIcons32 = New System.Windows.Forms.ImageList(Me.components)
        Me.pbWiz01 = New System.Windows.Forms.PictureBox()
        Me.ttMain = New System.Windows.Forms.ToolTip(Me.components)
        Me.pbWiz02 = New System.Windows.Forms.PictureBox()
        Me.pbWiz03 = New System.Windows.Forms.PictureBox()
        Me.pbWiz04 = New System.Windows.Forms.PictureBox()
        Me.pbWiz05 = New System.Windows.Forms.PictureBox()
        Me.pbWiz06 = New System.Windows.Forms.PictureBox()
        Me.pbWiz07 = New System.Windows.Forms.PictureBox()
        Me.pbWiz07g = New System.Windows.Forms.PictureBox()
        Me.ofdBrowse = New System.Windows.Forms.OpenFileDialog()
        Me.sfdSave = New System.Windows.Forms.SaveFileDialog()
        Me.cmdBrowse = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.txtFile = New System.Windows.Forms.TextBox()
        Me.epMain = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.timMain = New System.Windows.Forms.Timer(Me.components)
        Me.ssStatus.SuspendLayout()
        CType(Me.pbUWA, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbSirTech, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbWiz01, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbWiz02, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbWiz03, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbWiz04, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbWiz05, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbWiz06, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbWiz07, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbWiz07g, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.epMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblSelectedScenario
        '
        Me.lblSelectedScenario.AutoSize = True
        Me.lblSelectedScenario.BackColor = System.Drawing.Color.Transparent
        Me.lblSelectedScenario.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectedScenario.ForeColor = System.Drawing.Color.Yellow
        Me.lblSelectedScenario.Location = New System.Drawing.Point(103, 316)
        Me.lblSelectedScenario.Name = "lblSelectedScenario"
        Me.lblSelectedScenario.Size = New System.Drawing.Size(151, 20)
        Me.lblSelectedScenario.TabIndex = 0
        Me.lblSelectedScenario.Text = "Selected Scenario..."
        '
        'lblFile
        '
        Me.lblFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblFile.AutoSize = True
        Me.lblFile.BackColor = System.Drawing.Color.Transparent
        Me.lblFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFile.ForeColor = System.Drawing.Color.Yellow
        Me.lblFile.Location = New System.Drawing.Point(182, 339)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(183, 20)
        Me.lblFile.TabIndex = 2
        Me.lblFile.Text = "Select Save Game File..."
        '
        'ssStatus
        '
        Me.ssStatus.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsslMessage, Me.tsslTime})
        Me.ssStatus.Location = New System.Drawing.Point(0, 471)
        Me.ssStatus.Name = "ssStatus"
        Me.ssStatus.Size = New System.Drawing.Size(1132, 22)
        Me.ssStatus.TabIndex = 4
        Me.ssStatus.Text = "StatusStrip1"
        '
        'tsslMessage
        '
        Me.tsslMessage.Name = "tsslMessage"
        Me.tsslMessage.Size = New System.Drawing.Size(1083, 17)
        Me.tsslMessage.Spring = True
        '
        'tsslTime
        '
        Me.tsslTime.Name = "tsslTime"
        Me.tsslTime.Size = New System.Drawing.Size(34, 17)
        Me.tsslTime.Text = "Time"
        '
        'pbUWA
        '
        Me.pbUWA.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.pbUWA.Image = Global.WizEdit.My.Resources.Resources.UWA50
        Me.pbUWA.Location = New System.Drawing.Point(416, 12)
        Me.pbUWA.Name = "pbUWA"
        Me.pbUWA.Size = New System.Drawing.Size(300, 99)
        Me.pbUWA.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbUWA.TabIndex = 5
        Me.pbUWA.TabStop = False
        '
        'pbSirTech
        '
        Me.pbSirTech.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbSirTech.Image = Global.WizEdit.My.Resources.Resources.SirTech50
        Me.pbSirTech.Location = New System.Drawing.Point(1055, 12)
        Me.pbSirTech.Name = "pbSirTech"
        Me.pbSirTech.Size = New System.Drawing.Size(65, 35)
        Me.pbSirTech.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbSirTech.TabIndex = 6
        Me.pbSirTech.TabStop = False
        '
        'imgIcons32
        '
        Me.imgIcons32.ImageStream = CType(resources.GetObject("imgIcons32.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgIcons32.TransparentColor = System.Drawing.Color.Transparent
        Me.imgIcons32.Images.SetKeyName(0, "Wiz1.ico")
        Me.imgIcons32.Images.SetKeyName(1, "Wiz2.ico")
        Me.imgIcons32.Images.SetKeyName(2, "Wiz3.ico")
        Me.imgIcons32.Images.SetKeyName(3, "Wiz4.ico")
        Me.imgIcons32.Images.SetKeyName(4, "Wiz5.ico")
        Me.imgIcons32.Images.SetKeyName(5, "Wiz6.ico")
        Me.imgIcons32.Images.SetKeyName(6, "Wiz7.ico")
        Me.imgIcons32.Images.SetKeyName(7, "Wiz7g.ico")
        '
        'pbWiz01
        '
        Me.pbWiz01.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.pbWiz01.Image = Global.WizEdit.My.Resources.Resources.Wiz01Box
        Me.pbWiz01.Location = New System.Drawing.Point(106, 136)
        Me.pbWiz01.Name = "pbWiz01"
        Me.pbWiz01.Size = New System.Drawing.Size(111, 167)
        Me.pbWiz01.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbWiz01.TabIndex = 7
        Me.pbWiz01.TabStop = False
        Me.ttMain.SetToolTip(Me.pbWiz01, "Wizardry 01 - Proving Grounds of the Mad Overlord")
        '
        'pbWiz02
        '
        Me.pbWiz02.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.pbWiz02.Image = Global.WizEdit.My.Resources.Resources.Wiz02Box
        Me.pbWiz02.Location = New System.Drawing.Point(217, 135)
        Me.pbWiz02.Name = "pbWiz02"
        Me.pbWiz02.Size = New System.Drawing.Size(113, 168)
        Me.pbWiz02.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbWiz02.TabIndex = 8
        Me.pbWiz02.TabStop = False
        Me.ttMain.SetToolTip(Me.pbWiz02, "Wizardry 02 - Knight of Diamonds")
        '
        'pbWiz03
        '
        Me.pbWiz03.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.pbWiz03.Image = Global.WizEdit.My.Resources.Resources.Wiz03Box
        Me.pbWiz03.Location = New System.Drawing.Point(330, 136)
        Me.pbWiz03.Name = "pbWiz03"
        Me.pbWiz03.Size = New System.Drawing.Size(112, 168)
        Me.pbWiz03.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbWiz03.TabIndex = 9
        Me.pbWiz03.TabStop = False
        Me.ttMain.SetToolTip(Me.pbWiz03, "Wizardry 03 - Legacy of Llylgamyn")
        '
        'pbWiz04
        '
        Me.pbWiz04.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.pbWiz04.Image = Global.WizEdit.My.Resources.Resources.Wiz04Box
        Me.pbWiz04.Location = New System.Drawing.Point(442, 136)
        Me.pbWiz04.Name = "pbWiz04"
        Me.pbWiz04.Size = New System.Drawing.Size(113, 168)
        Me.pbWiz04.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbWiz04.TabIndex = 10
        Me.pbWiz04.TabStop = False
        Me.ttMain.SetToolTip(Me.pbWiz04, "Wizardry 04 - Return of Werdna")
        '
        'pbWiz05
        '
        Me.pbWiz05.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.pbWiz05.Image = Global.WizEdit.My.Resources.Resources.Wiz05Box
        Me.pbWiz05.Location = New System.Drawing.Point(555, 136)
        Me.pbWiz05.Name = "pbWiz05"
        Me.pbWiz05.Size = New System.Drawing.Size(109, 169)
        Me.pbWiz05.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbWiz05.TabIndex = 11
        Me.pbWiz05.TabStop = False
        Me.ttMain.SetToolTip(Me.pbWiz05, "Wizardry 05 - Heart of the Maelstrom")
        '
        'pbWiz06
        '
        Me.pbWiz06.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.pbWiz06.Image = Global.WizEdit.My.Resources.Resources.Wiz06Box
        Me.pbWiz06.Location = New System.Drawing.Point(664, 137)
        Me.pbWiz06.Name = "pbWiz06"
        Me.pbWiz06.Size = New System.Drawing.Size(112, 168)
        Me.pbWiz06.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbWiz06.TabIndex = 12
        Me.pbWiz06.TabStop = False
        Me.ttMain.SetToolTip(Me.pbWiz06, "Wizardry 06 - Bane of the Cosmic Forge")
        '
        'pbWiz07
        '
        Me.pbWiz07.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.pbWiz07.Image = Global.WizEdit.My.Resources.Resources.Wiz07Box
        Me.pbWiz07.Location = New System.Drawing.Point(776, 136)
        Me.pbWiz07.Name = "pbWiz07"
        Me.pbWiz07.Size = New System.Drawing.Size(112, 168)
        Me.pbWiz07.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbWiz07.TabIndex = 13
        Me.pbWiz07.TabStop = False
        Me.ttMain.SetToolTip(Me.pbWiz07, "Wizardry 07 - Crusaders of the Dark Savant")
        '
        'pbWiz07g
        '
        Me.pbWiz07g.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.pbWiz07g.Image = Global.WizEdit.My.Resources.Resources.Wiz07GBox
        Me.pbWiz07g.Location = New System.Drawing.Point(888, 135)
        Me.pbWiz07g.Name = "pbWiz07g"
        Me.pbWiz07g.Size = New System.Drawing.Size(139, 169)
        Me.pbWiz07g.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbWiz07g.TabIndex = 14
        Me.pbWiz07g.TabStop = False
        Me.ttMain.SetToolTip(Me.pbWiz07g, "Wizardry 07G - Wizardry Gold")
        '
        'ofdBrowse
        '
        Me.ofdBrowse.FileName = "OpenFileDialog1"
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowse.Image = Global.WizEdit.My.Resources.Resources.RockBrowse
        Me.cmdBrowse.Location = New System.Drawing.Point(910, 362)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(102, 33)
        Me.cmdBrowse.TabIndex = 15
        Me.cmdBrowse.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = Global.WizEdit.My.Resources.Resources.RockCancel
        Me.cmdCancel.Location = New System.Drawing.Point(1018, 410)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(102, 33)
        Me.cmdCancel.TabIndex = 16
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Image = Global.WizEdit.My.Resources.Resources.RockExit
        Me.cmdExit.Location = New System.Drawing.Point(1018, 410)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(102, 33)
        Me.cmdExit.TabIndex = 17
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'txtFile
        '
        Me.txtFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFile.BackColor = System.Drawing.Color.DimGray
        Me.txtFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFile.ForeColor = System.Drawing.Color.Gold
        Me.txtFile.Location = New System.Drawing.Point(185, 362)
        Me.txtFile.Name = "txtFile"
        Me.txtFile.Size = New System.Drawing.Size(703, 26)
        Me.txtFile.TabIndex = 18
        '
        'epMain
        '
        Me.epMain.ContainerControl = Me
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOK.Image = Global.WizEdit.My.Resources.Resources.RockOK
        Me.cmdOK.Location = New System.Drawing.Point(910, 410)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(102, 33)
        Me.cmdOK.TabIndex = 19
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'timMain
        '
        Me.timMain.Enabled = True
        Me.timMain.Interval = 1000
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.WizEdit.My.Resources.Resources.backgnd3b
        Me.ClientSize = New System.Drawing.Size(1132, 493)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.txtFile)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.pbWiz07g)
        Me.Controls.Add(Me.pbWiz07)
        Me.Controls.Add(Me.pbWiz06)
        Me.Controls.Add(Me.pbWiz05)
        Me.Controls.Add(Me.pbWiz04)
        Me.Controls.Add(Me.pbWiz03)
        Me.Controls.Add(Me.pbWiz02)
        Me.Controls.Add(Me.pbWiz01)
        Me.Controls.Add(Me.pbSirTech)
        Me.Controls.Add(Me.pbUWA)
        Me.Controls.Add(Me.ssStatus)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.lblSelectedScenario)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WizEdit 2017© for The Ultimate Wizardry Archives®"
        Me.ssStatus.ResumeLayout(False)
        Me.ssStatus.PerformLayout()
        CType(Me.pbUWA, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbSirTech, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbWiz01, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbWiz02, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbWiz03, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbWiz04, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbWiz05, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbWiz06, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbWiz07, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbWiz07g, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.epMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblSelectedScenario As Label
    Friend WithEvents lblFile As Label
    Friend WithEvents ssStatus As StatusStrip
    Friend WithEvents pbUWA As PictureBox
    Friend WithEvents pbSirTech As PictureBox
    Friend WithEvents imgIcons32 As ImageList
    Friend WithEvents pbWiz01 As PictureBox
    Friend WithEvents ttMain As ToolTip
    Friend WithEvents pbWiz02 As PictureBox
    Friend WithEvents pbWiz03 As PictureBox
    Friend WithEvents pbWiz04 As PictureBox
    Friend WithEvents pbWiz05 As PictureBox
    Friend WithEvents pbWiz06 As PictureBox
    Friend WithEvents pbWiz07 As PictureBox
    Friend WithEvents pbWiz07g As PictureBox
    Friend WithEvents ofdBrowse As OpenFileDialog
    Friend WithEvents sfdSave As SaveFileDialog
    Friend WithEvents cmdBrowse As Button
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdExit As Button
    Friend WithEvents txtFile As TextBox
    Friend WithEvents tsslMessage As ToolStripStatusLabel
    Friend WithEvents epMain As ErrorProvider
    Friend WithEvents cmdOK As Button
    Friend WithEvents tsslTime As ToolStripStatusLabel
    Friend WithEvents timMain As Timer
End Class
