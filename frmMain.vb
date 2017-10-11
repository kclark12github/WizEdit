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
    Private fScenarioSelected As Boolean = False
    Private Scenario As String = ""
    Dim dFilter As String = ""
    Dim dPath As String = ""
    Private Sub EnableFields(ByVal Caption As String)
        Dim dFileName As String = ""
        Try
            If fScenarioSelected Then Beep() : Exit Try
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

            fScenarioSelected = True

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
                    dPath = GetRegSetting("Environment", String.Format("UWAPath{0}", Scenario), "C:\WIZARD15")
                    dPath &= String.Format("\{0}", GetRegSetting("Environment", String.Format("Wiz{0}DataFile", Scenario), dFileName))
                    dFilter = "Saved Games (SAVE?.DSK)|SAVE?.DSK|All Files (*.*)|*.*"
                Case "06"
                    dPath = GetRegSetting("Environment", String.Format("UWAPath{0}", Scenario), "C:\BANE")
                    dPath &= String.Format("\{0}", GetRegSetting("Environment", String.Format("Wiz{0}DataFile", Scenario), dFileName))
                    dFilter = "Saved Games (*.DBS)|*.DBS|All Files (*.*)|*.*"
                Case "07"
                    dPath = GetRegSetting("Environment", String.Format("UWAPath{0}", Scenario), "C:\DSAVANT")
                    dPath &= String.Format("\{0}", GetRegSetting("Environment", String.Format("Wiz{0}DataFile", Scenario), dFileName))
                    dFilter = "Saved Games (*.DBS)|*.DBS|All Files (*.*)|*.*"
                Case "07G"
                    dPath = GetRegSetting("Environment", String.Format("UWAPath{0}", Scenario), "C:\Sirtech\WizGold")
                    dPath &= String.Format("\{0}", GetRegSetting("Environment", String.Format("Wiz{0}DataFile", Scenario), dFileName))
                    dFilter = "Saved Games (*.GLD)|*.GLD|All Files (*.*)|*.*"
            End Select
        Catch ex As Exception : MessageBox.Show(Me, ex.Message, ex.GetType.Name)
        End Try
    End Sub
    Private Function GetRegSetting(ByVal KeyName As String, ByVal ValueName As String, ByVal [Default] As Object) As Object
        Dim Reg As RegistryKey = Nothing
        GetRegSetting = Nothing
        Try
            Reg = Registry.CurrentUser.OpenSubKey(String.Format("Software\KClark Software\WizEdit\{0}", KeyName)) : If IsNothing(Reg) Then Exit Try
            GetRegSetting = Reg.GetValue(ValueName, [Default])
        Catch ex As System.Exception
        Finally : If Not IsNothing(Reg) Then Reg.Close()
        End Try
    End Function
    Private Sub SaveRegSetting(ByVal KeyName As String, ByVal ValueName As String, ByVal Value As Object)
        Dim Reg As RegistryKey = Nothing
        Dim CurrentValue As Object = Nothing
        Try
            If KeyName Is Nothing Then Throw New ArgumentException("KeyName must be provided.")
            If ValueName Is Nothing Then Throw New ArgumentException("ValueName must be provided.")
            If Value Is Nothing Then Throw New ArgumentException("Value must be provided.")

            KeyName = String.Format("Software\KClark Software\WizEdit\{0}", KeyName)
            Reg = Registry.CurrentUser.OpenSubKey(KeyName, True)
            If Reg Is Nothing Then
                'Iterate through the KeyName making sure each sub-key exists (create as necessary)...
                Dim SubKeys() As String = KeyName.Split("\")
                Dim Key As String = SubKeys(0)
                For i As Short = 1 To SubKeys.Length - 1
                    Dim SubKey As String = String.Format("{0}\{1}", Key, SubKeys(i))
                    Reg = Registry.CurrentUser.OpenSubKey(SubKey)
                    If Reg Is Nothing Then
                        Reg = Registry.CurrentUser.OpenSubKey(Key, True)
                        Reg.CreateSubKey(SubKeys(i))
                    End If
                    Reg.Close() : Reg = Nothing
                    Key = SubKey
                Next i
                Reg = Registry.CurrentUser.OpenSubKey(KeyName, True)
            End If
            CurrentValue = Reg.GetValue(ValueName)
            If CurrentValue Is Nothing OrElse CurrentValue.ToString <> Value.ToString Then Reg.SetValue(ValueName, Value)
        Catch ex As System.Exception
        Finally : If Reg IsNot Nothing Then Reg.Close()
        End Try
    End Sub
    Private Sub cmdBrowse_Click(sender As Object, e As EventArgs) Handles cmdBrowse.Click
        Try
            epMain.SetError(sender, "")
            If Directory.Exists(dPath) Then ChDir(dPath)
            With ofdBrowse
                '.InitialDirectory = dPath
                .FileName = txtFile.Text
                .Filter = dFilter
                .FilterIndex = 1
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
            'Parse the directory from the rest of the path and save it in the registry...
            dPath = txtFile.Text
            If dPath <> "" Then
                Dim fi As FileInfo = New FileInfo(dPath)
                dPath = fi.DirectoryName
                txtFile.Text = fi.Name
                dPath = String.Format("{0}\{1}", dPath, txtFile.Text)
                cmdOK.Focus()
            End If
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

            fScenarioSelected = False
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
        Dim fi As FileInfo = Nothing
        Try
            epMain.SetError(sender, "")
            If dPath = "" Then Throw New FileNotFoundException("Save Game file not found!")
            If Not File.Exists(dPath) Then
                'Before we bail, try the  current directory...
                fi = New FileInfo(dPath)
                dPath = String.Format("{0}\{1}", CurDir(), fi.Name)
                If Not File.Exists(dPath) Then Throw New FileNotFoundException("Save Game file not found!")
            End If

            fi = New FileInfo(dPath)
            Select Case Scenario
                Case "01"
                    Dim wiz01 As clsWiz01 = New clsWiz01(dPath)
                    SaveRegSetting("Environment", "UWAPath01", fi.DirectoryName)
                    SaveRegSetting("Environment", "Wiz01DataFile", fi.Name)
                    'Call DumpWiz01(dPath & "\" & txtFile.Text)
                    'Load frmWiz01
                    'frmWiz01.DataFile = dPath & "\" & txtFile.Text
                    'frmWiz01.Caption = picWiz01.ToolTipText
                    'frmWiz01.Icon = imgIcons32.ListImages("Wiz01").ExtractIcon
                    'frmWiz01.picWiz01.Visible = True
                    'Me.Hide()
                    'frmWiz01.Show
                Case "02"
                    'If Not Wiz02ValidateScenario(dPath & "\" & txtFile.Text) Then GoTo ExitSub
                    'Load frmWiz02
                    'frmWiz02.DataFile = dPath & "\" & txtFile.Text
                    'frmWiz02.Caption = picWiz02.ToolTipText
                    'frmWiz02.Icon = imgIcons32.ListImages("Wiz02").ExtractIcon
                    'frmWiz02.picWiz02.Visible = True
                    'Me.Hide()
                    'frmWiz02.Show
                Case "03"
                    'If Not Wiz03ValidateScenario(dPath & "\" & txtFile.Text) Then GoTo ExitSub
                    'Load frmWiz03
                    'frmWiz03.DataFile = dPath & "\" & txtFile.Text
                    'frmWiz03.Caption = picWiz03.ToolTipText
                    'frmWiz03.Icon = imgIcons32.ListImages("Wiz03").ExtractIcon
                    'frmWiz03.picWiz03.Visible = True
                    'Me.Hide()
                    'frmWiz03.Show
                Case "04"
                    'If Not Wiz04ValidateScenario(dPath & "\" & txtFile.Text) Then GoTo ExitSub
                    'Load frmWiz04
                    'frmWiz04.DataFile = dPath & "\" & txtFile.Text
                    'frmWiz04.Caption = picWiz04.ToolTipText
                    'frmWiz04.Icon = imgIcons32.ListImages("Wiz04").ExtractIcon
                    'frmWiz04.picWiz04.Visible = True
                    'Me.Hide()
                    'frmWiz04.Show
                Case "05"
                    'If Not Wiz05ValidateScenario(dPath & "\" & txtFile.Text) Then GoTo ExitSub
                    'Load frmWiz05
                    'frmWiz05.DataFile = dPath & "\" & txtFile.Text
                    'frmWiz05.Caption = picWiz05.ToolTipText
                    'frmWiz05.Icon = imgIcons32.ListImages("Wiz05").ExtractIcon
                    'frmWiz05.picWiz05.Visible = True
                    'Me.Hide()
                    'frmWiz05.Show
                Case "07"
                    ''If Not Wiz07ValidateScenario(dPath & "\" & txtFile.Text) Then GoTo ExitSub
                    ''Call DumpWiz07(dPath & "\" & txtFile.Text)
                    'Load frmWiz07
                    'frmWiz07.DataFile = dPath & "\" & txtFile.Text
                    'frmWiz07.Caption = picWiz07.ToolTipText
                    'frmWiz07.Icon = imgIcons32.ListImages("Wiz07").ExtractIcon
                    'frmWiz07.picWiz07.Visible = True
                    ''frmWiz07.picWiz07.Picture = picWiz07.Picture
                    'Me.Hide()
                    'frmWiz07.Show
                Case "07G"
                    ''If Not Wiz07ValidateScenario(dPath & "\" & txtFile.Text) Then GoTo ExitSub
                    ''Call DumpWiz07(dPath & "\" & txtFile.Text)
                    'Load frmWiz07
                    'frmWiz07.DataFile = dPath & "\" & txtFile.Text
                    'frmWiz07.Caption = Left(picWiz07g.ToolTipText, 11) & Mid(picWiz07g.ToolTipText, 13)
                    'frmWiz07.Icon = imgIcons32.ListImages("Wiz07g").ExtractIcon
                    'frmWiz07.picWiz07Gold.Visible = True
                    ''frmWiz07.picWiz07Gold.Picture = picWiz07g.Picture
                    'Me.Hide()
                    'frmWiz07.Show
                Case Else
                    'MsgBox "Sorry, I haven't implemented this scenario yet...", vbExclamation, Me.Caption
            End Select
        Catch ex As FileNotFoundException
            epMain.SetError(sender, ex.Message)
            cmdBrowse_Click(sender, e)
        Catch ex As Exception : epMain.SetError(sender, ex.Message)
        End Try
    End Sub
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

            fScenarioSelected = False
            With pbWiz01 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz02 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz03 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz04 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz05 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz06 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz07 : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With
            With pbWiz07g : AddHandler .Click, AddressOf pb_Click : AddHandler .GotFocus, AddressOf pb_GotFocus : AddHandler .LostFocus, AddressOf pb_LostFocus : End With

            tsslMessage.Text = ""
        Catch ex As Exception : MessageBox.Show(Me, ex.Message, ex.GetType.Name)
        End Try
    End Sub
    Private Sub pb_Click(sender As Object, e As EventArgs)
        EnableFields(ttMain.GetToolTip(CType(sender, PictureBox)))
    End Sub
    Private Sub pb_GotFocus(sender As Object, e As EventArgs)
        tsslMessage.Text = ttMain.GetToolTip(CType(sender, PictureBox))
    End Sub
    Private Sub pb_LostFocus(sender As Object, e As EventArgs)
        tsslMessage.Text = ""
    End Sub
End Class
