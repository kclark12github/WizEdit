'frmMain.frm
'   Main form for the WizEdit Application...
'   Copyright © 2000-2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/09/17    Ken Clark       Upgraded to VS2017;
'   08/26/00    Ken Clark       Created;
'=================================================================================================================================
Public Class frmMain
    Private SelectedScenario As WizEditBase = Nothing
    Private Scenario As String = ""
    Dim dFilter As String = ""
    Private Sub cmdBrowse_Click(sender As Object, e As EventArgs) Handles cmdBrowse.Click
        Try
            epMain.SetError(sender, "")
            With ofdBrowse
                '.InitialDirectory = dPath
                .FileName = txtFile.Text
                .Filter = dFilter
                .FilterIndex = 0
                .ShowReadOnly = False
                .Multiselect = False
                .CheckPathExists = True
                If .ShowDialog(Me) = DialogResult.Cancel Then
                    MessageBox.Show(Me, "Operation canceled at User's request.", "Browse", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    cmdCancel_Click(sender, e)
                    Exit Try
                End If
                txtFile.Text = .FileName
            End With
            'Parse the directory from the rest of the path and save it in the scenario...
            If txtFile.Text <> "" Then SelectedScenario.ScenarioDataPath = txtFile.Text : cmdOK.Focus()
        Catch ex As Exception : epMain.SetError(sender, ex.Message)
        End Try
    End Sub
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Try
            epMain.SetError(sender, "")
            cmdExit.Visible = True

            cmdOK.Visible = False : cmdCancel.Visible = False
            cmdExit.Enabled = True

            lblSelectedScenario.Text = "Please select a scenario by clicking on the appropriate game box above..."

            lblFile.Visible = False : txtFile.Visible = False : cmdBrowse.Visible = False

            SelectedScenario = Nothing
            tsslMessage.Text = ""
        Catch ex As Exception : epMain.SetError(sender, ex.Message)
        End Try
    End Sub
    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Try
            Me.Close()
        Catch ex As Exception : epMain.SetError(sender, ex.Message)
        End Try
    End Sub
    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Dim bHidden As Boolean = False
        Try
            epMain.SetError(sender, "")

            tsslMessage.Text = SelectedScenario.ScenarioDataPath    'Will trigger NotFoundException if appropriate...
            Me.Hide() : bHidden = True

            SelectedScenario.Read()
            SelectedScenario.Show()

            tsslMessage.Text = ""
        Catch ex As FileNotFoundException : tsslMessage.Text = ex.Message : epMain.SetError(sender, ex.Message)
            cmdBrowse_Click(sender, e)
        Catch ex As System.IO.IOException : tsslMessage.Text = ex.Message : epMain.SetError(sender, ex.Message)
        Catch ex As Exception : tsslMessage.Text = ex.Message : epMain.SetError(sender, ex.Message)
            Debug.WriteLine(ex.ToString)
        Finally
            If bHidden Then Me.Show() : bHidden = False
            'Use our cmdCancel to reset the screen...
            cmdCancel_Click(sender, e)
        End Try
    End Sub
    Private Function ToBit(ByVal iByte As Byte) As String
        ToBit = ""
        For iBit As Short = 0 To 7
            ToBit &= IIf(CBool((iByte And 2 ^ iBit) = 2 ^ iBit), "1", "0")
        Next iBit
    End Function

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'Handle Command Line Arguments (if any)...
            'CommandLineArgs = GetCommandLine()
            cmdOK.Visible = False
            cmdCancel.Visible = False
            cmdExit.Enabled = True

            lblSelectedScenario.Text = "Please select a scenario by clicking on the appropriate game box above..."

            lblFile.Visible = False
            txtFile.Visible = False
            cmdBrowse.Visible = False

            With pbWiz01 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz02 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz03 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz04 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz05 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz06 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz07 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz07g : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With

            tsslMessage.Text = ""
            timMain_Tick(Nothing, Nothing)
        Catch ex As Exception : MessageBox.Show(Me, ex.Message, ex.GetType.Name)
        End Try
    End Sub
    Private Sub pb_Click(sender As Object, e As EventArgs)
        Dim Caption As String = ttMain.GetToolTip(CType(sender, PictureBox))
        Dim dFileName As String = ""
        Try
            If SelectedScenario IsNot Nothing Then Beep() : Exit Try
            Scenario = Caption.Substring(9, 3).Trim.ToUpper
            If Scenario = "07G" Then    'Drop off the "G" for display purposes...
                lblSelectedScenario.Text = Caption.Replace("07G", "07")
            Else
                lblSelectedScenario.Text = Caption
            End If
            lblSelectedScenario.Visible = True

            cmdOK.Visible = True
            cmdExit.Visible = False
            cmdCancel.Visible = True
            lblFile.Visible = True
            txtFile.Visible = True
            cmdBrowse.Visible = True

            'dFileName = GetRegSetting("Environment", "Wiz" & Scenario & "DataFile", "")
            If dFileName = "" Then
                Select Case Scenario
                    Case "01", "02", "03", "04", "05" : dFileName = String.Format("SAVE{0}.DSK", Scenario.Substring(1, 1))
                    Case "06", "07" : dFileName = "SCENARIO.DBS"
                    Case "07G" : dFileName = "SCENARIO.GLD"
                End Select
            End If
            txtFile.Text = dFileName

            'See if we know where UWA is installed... If not, use defaults from UWA...
            Select Case Scenario
                Case "01", "02", "03", "04", "05"
                    Select Case Scenario
                        Case "01" : SelectedScenario = New Wizardry01(ttMain.GetToolTip(pbWiz01), Global.WizEdit.My.Resources.Resources.Wiz1, Global.WizEdit.My.Resources.Resources.Wiz01Box, Me)
                        Case "02" : SelectedScenario = New Wizardry02(ttMain.GetToolTip(pbWiz02), Global.WizEdit.My.Resources.Resources.Wiz2, Global.WizEdit.My.Resources.Resources.Wiz02Box, Me)
                        Case "03" : SelectedScenario = New Wizardry03(ttMain.GetToolTip(pbWiz03), Global.WizEdit.My.Resources.Resources.Wiz3, Global.WizEdit.My.Resources.Resources.Wiz03Box, Me)
                        Case "04" : SelectedScenario = New Wizardry04(ttMain.GetToolTip(pbWiz04), Global.WizEdit.My.Resources.Resources.Wiz4, Global.WizEdit.My.Resources.Resources.Wiz04Box, Me)
                        Case "05" : SelectedScenario = New Wizardry05(ttMain.GetToolTip(pbWiz05), Global.WizEdit.My.Resources.Resources.Wiz5, Global.WizEdit.My.Resources.Resources.Wiz05Box, Me)
                    End Select
                    dFilter = String.Format("{0} Saved Games (SAVE{1}.DSK)|SAVE{1}.DSK|All Files (*.*)|*.*", SelectedScenario.ScenarioName, CInt(Scenario))
                Case "06"
                    'SelectedScenario = New Wizardry06(ttMain.GetToolTip(pbWiz06), Global.WizEdit.My.Resources.Resources.Wiz6, Global.WizEdit.My.Resources.Resources.Wiz06Box, Me)
                    'With SelectedScenario : dPath = .GetRegSetting("Environment", .RegDataPath, "C:\BANE") : End With
                    'dFilter = "Saved Games (*.DBS)|*.DBS|All Files (*.*)|*.*"
                Case "07"
                    'SelectedScenario = New Wizardry07(ttMain.GetToolTip(pbWiz07), Global.WizEdit.My.Resources.Resources.Wiz7, Global.WizEdit.My.Resources.Resources.Wiz07Box, Me)
                    'With SelectedScenario : dPath = .GetRegSetting("Environment", .RegDataPath, "C:\DSAVANT") : End With
                    'dFilter = "Saved Games (*.DBS)|*.DBS|All Files (*.*)|*.*"
                Case "07G"
                    'SelectedScenario = New Wizardry07G(ttMain.GetToolTip(pbWiz07G), Global.WizEdit.My.Resources.Resources.Wiz7G, Global.WizEdit.My.Resources.Resources.Wiz07GBox, Me)
                    'With SelectedScenario : dPath = .GetRegSetting("Environment", .RegDataPath, "C:\Sirtech\WizGold") : End With
                    'dFilter = "Saved Games (*.GLD)|*.GLD|All Files (*.*)|*.*"
            End Select
            If SelectedScenario IsNot Nothing Then Me.Icon = SelectedScenario.Icon
        Catch ex As Exception : MessageBox.Show(Me, ex.Message, ex.GetType.Name)
        End Try
    End Sub
    Private Sub pb_GotFocus(sender As Object, e As EventArgs)
        tsslMessage.Text = ttMain.GetToolTip(CType(sender, PictureBox))
    End Sub
    Private Sub pb_LostFocus(sender As Object, e As EventArgs)
        tsslMessage.Text = ""
    End Sub
    Private Sub timMain_Tick(sender As Object, e As EventArgs) Handles timMain.Tick
        tsslTime.Text = String.Format("{0:hh:mm tt}", Now)
    End Sub
End Class