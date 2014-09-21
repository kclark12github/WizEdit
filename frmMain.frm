VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WizEdit 2000© for The Ultimate Wizardry Archives®"
   ClientHeight    =   6420
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   9288
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":1CCA
   ScaleHeight     =   6420
   ScaleWidth      =   9288
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdgMain 
      Left            =   240
      Top             =   5460
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Height          =   432
      Left            =   7560
      Picture         =   "frmMain.frx":43810
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton cmdBrowse 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   6108
      Picture         =   "frmMain.frx":44D8B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4380
      Width           =   1332
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Left            =   1848
      TabIndex        =   9
      Text            =   "Save Game File"
      Top             =   4380
      Width           =   4152
   End
   Begin VB.PictureBox picWiz07g 
      AutoSize        =   -1  'True
      Height          =   1524
      Left            =   7782
      Picture         =   "frmMain.frx":462FB
      ScaleHeight     =   1476
      ScaleWidth      =   1224
      TabIndex        =   8
      ToolTipText     =   "Wizardry 07G - Wizardry Gold"
      Top             =   1680
      Width           =   1272
   End
   Begin VB.PictureBox picWizardryLogo 
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   240
      Picture         =   "frmMain.frx":46CEE
      ScaleHeight     =   372
      ScaleWidth      =   1920
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1968
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   17
      Top             =   6168
      Width           =   9288
      _ExtentX        =   16383
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10986
            Key             =   "Message"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   900
            TextSave        =   "SCRL"
            Key             =   "SCRL"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   910
            MinWidth        =   900
            TextSave        =   "CAPS"
            Key             =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   889
            MinWidth        =   891
            TextSave        =   "NUM"
            Key             =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "7:33 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picWiz07 
      AutoSize        =   -1  'True
      Height          =   1536
      Left            =   6702
      Picture         =   "frmMain.frx":477CB
      ScaleHeight     =   1488
      ScaleWidth      =   984
      TabIndex        =   7
      ToolTipText     =   "Wizardry 07 - Crusaders of the Dark Savant"
      Top             =   1680
      Width           =   1032
   End
   Begin VB.PictureBox picWiz06 
      AutoSize        =   -1  'True
      Height          =   1536
      Left            =   5622
      Picture         =   "frmMain.frx":47F2E
      ScaleHeight     =   1488
      ScaleWidth      =   984
      TabIndex        =   6
      ToolTipText     =   "Wizardry 06 - Bane of the Cosmic Forge"
      Top             =   1680
      Width           =   1032
   End
   Begin VB.PictureBox picWiz05 
      AutoSize        =   -1  'True
      Height          =   1548
      Left            =   4566
      Picture         =   "frmMain.frx":486EE
      ScaleHeight     =   1500
      ScaleWidth      =   960
      TabIndex        =   5
      ToolTipText     =   "Wizardry 05 - Heart of the Maelstrom"
      Top             =   1680
      Width           =   1008
   End
   Begin VB.PictureBox picWiz04 
      AutoSize        =   -1  'True
      Height          =   1536
      Left            =   3474
      Picture         =   "frmMain.frx":48E6A
      ScaleHeight     =   1488
      ScaleWidth      =   996
      TabIndex        =   4
      ToolTipText     =   "Wizardry 04 - Return of Werdna"
      Top             =   1680
      Width           =   1044
   End
   Begin VB.PictureBox picWiz03 
      AutoSize        =   -1  'True
      Height          =   1536
      Left            =   2394
      Picture         =   "frmMain.frx":494F8
      ScaleHeight     =   1488
      ScaleWidth      =   984
      TabIndex        =   3
      ToolTipText     =   "Wizardry 03 - Legacy of Llylgamn"
      Top             =   1680
      Width           =   1032
   End
   Begin VB.PictureBox picWiz02 
      AutoSize        =   -1  'True
      Height          =   1536
      Left            =   1302
      Picture         =   "frmMain.frx":49AA1
      ScaleHeight     =   1488
      ScaleWidth      =   996
      TabIndex        =   2
      ToolTipText     =   "Wizardry 02 - Knight of Diamonds"
      Top             =   1680
      Width           =   1044
   End
   Begin VB.PictureBox picWiz01 
      AutoSize        =   -1  'True
      Height          =   1524
      Left            =   234
      Picture         =   "frmMain.frx":4A0FB
      ScaleHeight     =   1476
      ScaleWidth      =   972
      TabIndex        =   1
      ToolTipText     =   "Wizardry 01 - Proving Grounds of the Mad Overlord"
      Top             =   1680
      Width           =   1020
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   432
      Left            =   7560
      Picture         =   "frmMain.frx":4A711
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Height          =   432
      Left            =   6240
      Picture         =   "frmMain.frx":4BC5D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1212
   End
   Begin VB.PictureBox picSirTech 
      AutoSize        =   -1  'True
      Height          =   468
      Left            =   8220
      Picture         =   "frmMain.frx":4D1BA
      ScaleHeight     =   420
      ScaleWidth      =   780
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Width           =   828
   End
   Begin VB.PictureBox picUWA 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1236
      Left            =   2820
      Picture         =   "frmMain.frx":4D664
      ScaleHeight     =   1188
      ScaleWidth      =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   3648
   End
   Begin MSComctlLib.ImageList imgIcons16 
      Left            =   600
      Top             =   5460
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E20A
            Key             =   "Wiz01"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FEE6
            Key             =   "Wiz02"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51BC2
            Key             =   "Wiz03"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5389E
            Key             =   "Wiz04"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53BBA
            Key             =   "Wiz05"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55896
            Key             =   "Wiz06"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57572
            Key             =   "Wiz07"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5924E
            Key             =   "Wiz07g"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons32 
      Left            =   1020
      Top             =   5460
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5956A
            Key             =   "Wiz01"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B246
            Key             =   "Wiz02"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5CF22
            Key             =   "Wiz03"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EBFE
            Key             =   "Wiz04"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EF1A
            Key             =   "Wiz05"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60BF6
            Key             =   "Wiz06"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":628D2
            Key             =   "Wiz07"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":645AE
            Key             =   "Wiz07g"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Save Game File..."
      BeginProperty Font 
         Name            =   "Heidelberg"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   252
      Left            =   1848
      TabIndex        =   19
      Top             =   4080
      Width           =   2112
   End
   Begin VB.Label lblSelectedScenario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Scenario..."
      BeginProperty Font 
         Name            =   "Heidelberg"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   288
      Left            =   480
      TabIndex        =   16
      Top             =   3660
      Width           =   1956
   End
   Begin VB.Label lblScenario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Scenario..."
      BeginProperty Font 
         Name            =   "Heidelberg"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   288
      Left            =   240
      TabIndex        =   15
      Top             =   3300
      Width           =   1956
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmMain - frmMain.frm
'   Main form for the WizEdit Application...
'   Copyright © 2000, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   08/26/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit
Dim fScenarioSelected As Boolean
Dim Scenario As String
Dim dFilter As String
Dim dPath As String
Private Sub cmdBrowse_Click()
    On Error GoTo ErrorHandler

    If Dir(dPath, vbDirectory) <> vbNullString Then ChDir dPath
    initCommonDialog
    With frmMain.cdgMain
        '.InitDir = dPath
        .FileName = txtFile.Text
        .Filter = dFilter
        .FilterIndex = 1
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist + _
            cdlOFNNoChangeDir + cdlOFNNoReadOnlyReturn
        .CancelError = True
        .ShowOpen    ' Call the open file procedure.
        txtFile.Text = .FileName
    End With
    
    'Parse the directory from the rest of the path and save it in the registry...
    dPath = txtFile.Text
    If dPath <> vbNullString Then
        'If (GetAttr(dPath) And vbDirectory) <> vbDirectory Then dPath = Mid(dPath, 1, InStrRev(dPath, "\") - 1)
        dPath = ParsePath(txtFile.Text, DrvDirNoSlash)
        txtFile.Text = ParsePath(txtFile.Text, FileNameBaseExt)
        cmdOK.SetFocus
    End If
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case cdlCancel
            sbStatus.Panels("Status").Text = "Cancelled"
            MsgBox "Operation cancelled at User's request.", vbInformation, Me.Caption
            Call cmdCancel_Click
            Exit Sub
        Case 52, 76
            Resume Next
        Case Else
            MsgBox Err.Description & " (Error #" & Err.Number & ")", Me.Caption
            Exit Sub
    End Select
End Sub
Private Sub cmdCancel_Click()
    cmdExit.Visible = True
    
    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdExit.Enabled = True
    
    lblScenario.Caption = "Please select a scenario by clicking on the appropriate game box above..."
    lblSelectedScenario.Visible = False
    
    lblFile.Visible = False
    txtFile.Visible = False
    cmdBrowse.Visible = False
    
    fScenarioSelected = False
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    If Dir(dPath & "\" & txtFile.Text, vbNormal) = vbNullString Then
        'Before we bail, try the  current directory...
        If Dir(CurDir & "\" & txtFile.Text, vbNormal) = vbNullString Then
            Call Beep
            Call MsgBox("Save Game file not found!" & vbCrLf & vbCrLf & dPath & "\" & txtFile.Text, vbExclamation, Me.Caption)
            Call cmdBrowse_Click
            Exit Sub
        Else
            dPath = CurDir
        End If
    End If
    
    Select Case Scenario
        Case "01"
            If Not Wiz01ValidateScenario(dPath & "\" & txtFile.Text) Then GoTo ExitSub
            'Call DumpWiz01(dPath & "\" & txtFile.Text)
            Load frmWiz01
            frmWiz01.DataFile = dPath & "\" & txtFile.Text
            frmWiz01.Caption = picWiz01.ToolTipText
            frmWiz01.Icon = imgIcons32.ListImages("Wiz01").ExtractIcon
            frmWiz01.picWiz01.Visible = True
            Me.Hide
            frmWiz01.Show
        Case "07"
            'If Not Wiz07ValidateScenario(dPath & "\" & txtFile.Text) Then GoTo ExitSub
            'Call DumpWiz07(dPath & "\" & txtFile.Text)
            Load frmWiz07
            frmWiz07.DataFile = dPath & "\" & txtFile.Text
            frmWiz07.Caption = picWiz07.ToolTipText
            frmWiz07.Icon = imgIcons32.ListImages("Wiz07").ExtractIcon
            frmWiz07.picWiz07.Visible = True
            'frmWiz07.picWiz07.Picture = picWiz07.Picture
            Me.Hide
            frmWiz07.Show
        Case "07G"
            'If Not Wiz07ValidateScenario(dPath & "\" & txtFile.Text) Then GoTo ExitSub
            'Call DumpWiz07(dPath & "\" & txtFile.Text)
            Load frmWiz07
            frmWiz07.DataFile = dPath & "\" & txtFile.Text
            frmWiz07.Caption = Left(picWiz07g.ToolTipText, 11) & Mid(picWiz07g.ToolTipText, 13)
            frmWiz07.Icon = imgIcons32.ListImages("Wiz07g").ExtractIcon
            frmWiz07.picWiz07Gold.Visible = True
            'frmWiz07.picWiz07Gold.Picture = picWiz07g.Picture
            Me.Hide
            frmWiz07.Show
        Case Else
            MsgBox "Sorry, I haven't implemented this scenario yet...", vbExclamation, Me.Caption
    End Select
    
ExitSub:
    Exit Sub
End Sub
Private Sub EnableFields(ByVal strCaption As String)
    Dim dFileName As String
    
    If fScenarioSelected Then
        Call Beep
        Exit Sub
    End If
    
    Scenario = Trim(UCase(Mid(strCaption, 10, 3)))
    
    lblScenario.Caption = "Scenario Selected..."
    If Scenario = "07G" Then    'Drop off the "G" for display purposes...
        lblSelectedScenario.Caption = Left(strCaption, 11) & Mid(strCaption, 13)
    Else
        lblSelectedScenario.Caption = strCaption
    End If
    lblSelectedScenario.Visible = True
    
    cmdOK.Visible = True
    cmdExit.Visible = False
    cmdCancel.Visible = True
    lblFile.Visible = True
    txtFile.Visible = True
    cmdBrowse.Visible = True
    
    fScenarioSelected = True

    dFileName = GetRegSetting("Environment", "Wiz" & Scenario & "DataFile", vbNullString)
    If dFileName = vbNullString Then
        Select Case Scenario
            Case "01", "02", "03", "04", "05"
                dFileName = "SAVE" & Right(Scenario, 1) & ".DSK"
            Case "06", "07"
                dFileName = "SCENARIO.DBS"
            Case "07G"
                dFileName = "SCENARIO.GLD"
        End Select
    End If
    txtFile.Text = dFileName

    'See if we know where UWA is installed... If not, use defaults from UWA...
    Select Case Scenario
        Case "01", "02", "03", "04", "05"
            dPath = GetRegSetting("Environment", "UWAPath" & Scenario, "C:\WIZARD15")
            dFilter = "Saved Games (SAVE?.DSK)|SAVE?.DSK|All Files (*.*)|*.*"
        Case "06"
            dPath = GetRegSetting("Environment", "UWAPath" & Scenario, "C:\BANE")
            dFilter = "Saved Games (*.DBS)|*.DBS|All Files (*.*)|*.*"
        Case "07"
            dPath = GetRegSetting("Environment", "UWAPath" & Scenario, "C:\DSAVANT")
            dFilter = "Saved Games (*.DBS)|*.DBS|All Files (*.*)|*.*"
        Case "07G"
            dPath = GetRegSetting("Environment", "UWAPath" & Scenario, "C:\Sirtech\WizGold")
            dFilter = "Saved Games (*.GLD)|*.GLD|All Files (*.*)|*.*"
    End Select
End Sub
Private Sub Form_Load()
    'Handle Command Line Arguments (if any)...
    CommandLineArgs = GetCommandLine()
    If UBound(CommandLineArgs) > 0 Then
    End If
    
    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdExit.Enabled = True
    
    lblScenario.Caption = "Please select a scenario by clicking on the appropriate game box above..."
    lblSelectedScenario.Visible = False
    
    lblFile.Visible = False
    txtFile.Visible = False
    cmdBrowse.Visible = False
    
    fScenarioSelected = False
End Sub
Public Sub MainCancel()
    cmdCancel_Click
End Sub
Private Sub picWiz01_Click()
    Call EnableFields(picWiz01.ToolTipText)
End Sub
Private Sub picWiz01_GotFocus()
    sbStatus.Panels("Message").Text = picWiz01.ToolTipText
End Sub
Private Sub picWiz01_LostFocus()
    sbStatus.Panels("Message").Text = vbNullString
End Sub
Private Sub picWiz02_Click()
    Call EnableFields(picWiz02.ToolTipText)
End Sub
Private Sub picWiz02_GotFocus()
    sbStatus.Panels("Message").Text = picWiz02.ToolTipText
End Sub
Private Sub picWiz02_LostFocus()
    sbStatus.Panels("Message").Text = vbNullString
End Sub
Private Sub picWiz03_Click()
    Call EnableFields(picWiz03.ToolTipText)
End Sub
Private Sub picWiz03_GotFocus()
    sbStatus.Panels("Message").Text = picWiz03.ToolTipText
End Sub
Private Sub picWiz03_LostFocus()
    sbStatus.Panels("Message").Text = vbNullString
End Sub
Private Sub picWiz04_Click()
    Call EnableFields(picWiz04.ToolTipText)
End Sub
Private Sub picWiz04_GotFocus()
    sbStatus.Panels("Message").Text = picWiz04.ToolTipText
End Sub
Private Sub picWiz04_LostFocus()
    sbStatus.Panels("Message").Text = vbNullString
End Sub
Private Sub picWiz05_Click()
    Call EnableFields(picWiz05.ToolTipText)
End Sub
Private Sub picWiz05_GotFocus()
    sbStatus.Panels("Message").Text = picWiz05.ToolTipText
End Sub
Private Sub picWiz05_LostFocus()
    sbStatus.Panels("Message").Text = vbNullString
End Sub
Private Sub picWiz06_Click()
    Call EnableFields(picWiz06.ToolTipText)
End Sub
Private Sub picWiz06_GotFocus()
    sbStatus.Panels("Message").Text = picWiz06.ToolTipText
End Sub
Private Sub picWiz06_LostFocus()
    sbStatus.Panels("Message").Text = vbNullString
End Sub
Private Sub picWiz07_Click()
    Call EnableFields(picWiz07.ToolTipText)
End Sub
Private Sub picWiz07_GotFocus()
    sbStatus.Panels("Message").Text = picWiz07.ToolTipText
End Sub
Private Sub picWiz07_LostFocus()
    sbStatus.Panels("Message").Text = vbNullString
End Sub
Private Sub picWiz07g_Click()
    Call EnableFields(picWiz07g.ToolTipText)
End Sub
Private Sub picWiz07g_GotFocus()
    sbStatus.Panels("Message").Text = Left(picWiz07g.ToolTipText, 11) & Mid(picWiz07g.ToolTipText, 13)
End Sub
Private Sub picWiz07g_LostFocus()
    sbStatus.Panels("Message").Text = vbNullString
End Sub

