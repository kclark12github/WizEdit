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
      Picture         =   "frmMain.frx":1B4EC
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   372
      Left            =   6108
      Picture         =   "frmMain.frx":1BAEA
      Style           =   1  'Graphical
      TabIndex        =   18
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
      TabIndex        =   16
      Text            =   "Save Game File"
      Top             =   4380
      Width           =   4152
   End
   Begin VB.PictureBox picWiz07g 
      AutoSize        =   -1  'True
      Height          =   1524
      Left            =   7782
      Picture         =   "frmMain.frx":1C665
      ScaleHeight     =   1476
      ScaleWidth      =   1224
      TabIndex        =   15
      ToolTipText     =   "Wizardry 07G - Wizardry Gold"
      Top             =   1680
      Width           =   1272
   End
   Begin VB.PictureBox picWizardryLogo 
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   240
      Picture         =   "frmMain.frx":1D058
      ScaleHeight     =   372
      ScaleWidth      =   1920
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   1968
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   13
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
            TextSave        =   "1:40 AM"
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
      Picture         =   "frmMain.frx":1DB35
      ScaleHeight     =   1488
      ScaleWidth      =   984
      TabIndex        =   10
      ToolTipText     =   "Wizardry 07 - Crusaders of the Dark Savant"
      Top             =   1680
      Width           =   1032
   End
   Begin VB.PictureBox picWiz06 
      AutoSize        =   -1  'True
      Height          =   1536
      Left            =   5622
      Picture         =   "frmMain.frx":1E298
      ScaleHeight     =   1488
      ScaleWidth      =   984
      TabIndex        =   9
      ToolTipText     =   "Wizardry 06 - Bane of the Cosmic Forge"
      Top             =   1680
      Width           =   1032
   End
   Begin VB.PictureBox picWiz05 
      AutoSize        =   -1  'True
      Height          =   1548
      Left            =   4566
      Picture         =   "frmMain.frx":1EA58
      ScaleHeight     =   1500
      ScaleWidth      =   960
      TabIndex        =   8
      ToolTipText     =   "Wizardry 05 - Heart of the Maelstrom"
      Top             =   1680
      Width           =   1008
   End
   Begin VB.PictureBox picWiz04 
      AutoSize        =   -1  'True
      Height          =   1536
      Left            =   3474
      Picture         =   "frmMain.frx":1F1D4
      ScaleHeight     =   1488
      ScaleWidth      =   996
      TabIndex        =   7
      ToolTipText     =   "Wizardry 04 - Return of Werdna"
      Top             =   1680
      Width           =   1044
   End
   Begin VB.PictureBox picWiz03 
      AutoSize        =   -1  'True
      Height          =   1536
      Left            =   2394
      Picture         =   "frmMain.frx":1F862
      ScaleHeight     =   1488
      ScaleWidth      =   984
      TabIndex        =   6
      ToolTipText     =   "Wizardry 03 - Legacy of Llylgamn"
      Top             =   1680
      Width           =   1032
   End
   Begin VB.PictureBox picWiz02 
      AutoSize        =   -1  'True
      Height          =   1536
      Left            =   1302
      Picture         =   "frmMain.frx":1FE0B
      ScaleHeight     =   1488
      ScaleWidth      =   996
      TabIndex        =   5
      ToolTipText     =   "Wizardry 02 - Knight of Diamonds"
      Top             =   1680
      Width           =   1044
   End
   Begin VB.PictureBox picWiz01 
      AutoSize        =   -1  'True
      Height          =   1524
      Left            =   234
      Picture         =   "frmMain.frx":20465
      ScaleHeight     =   1476
      ScaleWidth      =   972
      TabIndex        =   4
      ToolTipText     =   "Wizardry 01 - Proving Grounds of the Mad Overlord"
      Top             =   1680
      Width           =   1020
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   432
      Left            =   7560
      Picture         =   "frmMain.frx":20A7B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Height          =   432
      Left            =   6240
      Picture         =   "frmMain.frx":21064
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1212
   End
   Begin VB.PictureBox picSirTech 
      AutoSize        =   -1  'True
      Height          =   468
      Left            =   8220
      Picture         =   "frmMain.frx":2163E
      ScaleHeight     =   420
      ScaleWidth      =   780
      TabIndex        =   1
      Top             =   240
      Width           =   828
   End
   Begin VB.PictureBox picUWA 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1236
      Left            =   2820
      Picture         =   "frmMain.frx":21AE8
      ScaleHeight     =   1188
      ScaleWidth      =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   3648
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
      TabIndex        =   17
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
      TabIndex        =   12
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
      TabIndex        =   11
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
        .flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist + _
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
        Call SaveWizEditSetting("Environment", "UWAPath" & Scenario, dPath)
        
        txtFile.Text = ParsePath(txtFile.Text, FileNameBaseExt)
        Call SaveWizEditSetting("Environment", "Wiz" & Scenario & "DataFile", txtFile.Text)
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
Private Sub EnableFields(ByVal strCaption As String)
    Dim dFileName As String
    
    If fScenarioSelected Then
        Call Beep
        Exit Sub
    End If
    
    lblScenario.Caption = "Scenario Selected..."
    lblSelectedScenario.Caption = strCaption
    lblSelectedScenario.Visible = True
    
    Scenario = Trim(UCase(Mid(strCaption, 10, 3)))
    
    cmdOK.Visible = True
    cmdExit.Visible = False
    cmdCancel.Visible = True
    lblFile.Visible = True
    txtFile.Visible = True
    cmdBrowse.Visible = True
    
    fScenarioSelected = True

    dFileName = GetWizEditSetting("Environment", "Wiz" & Scenario & "DataFile", vbNullString)
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
            dPath = GetWizEditSetting("Environment", "UWAPath" & Scenario, "C:\WIZARD15")
            dFilter = "Saved Games (SAVE?.DSK)|SAVE?.DSK|All Files (*.*)|*.*"
        Case "06"
            dPath = GetWizEditSetting("Environment", "UWAPath" & Scenario, "C:\BANE")
            dFilter = "Saved Games (*.DBS)|*.DBS|All Files (*.*)|*.*"
        Case "07"
            dPath = GetWizEditSetting("Environment", "UWAPath" & Scenario, "C:\DSAVANT")
            dFilter = "Saved Games (*.DBS)|*.DBS|All Files (*.*)|*.*"
        Case "07G"
            dPath = GetWizEditSetting("Environment", "UWAPath" & Scenario, "C:\Sirtech\WizGold")
            dFilter = "Saved Games (*.GLD)|*.GLD|All Files (*.*)|*.*"
    End Select
End Sub
Private Sub cmdOK_Click()
    If Dir(dPath & "\" & txtFile.Text, vbNormal) = vbNullString Then
        'Before we bail, try the  current directory...
        If Dir(CurDir & "\" & txtFile.Text, vbNormal) = vbNullString Then
            Call Beep
            MsgBox "Save Game file not found!" & vbCrLf & vbCrLf & dPath & "\" & txtFile.Text, vbExclamation, Me.Caption
            Call cmdBrowse_Click
            Exit Sub
        Else
            dPath = CurDir
            Call SaveWizEditSetting("Environment", "UWAPath" & Scenario, dPath)
        End If
    End If
    
    Select Case Scenario
        Case "07"
            Call DumpWiz07(dPath & "\" & txtFile.Text)
        Case Else
            MsgBox "Sorry, I haven't implemented this scenario yet...", vbExclamation, Me.Caption
    End Select
    
    cmdCancel_Click
    Exit Sub
End Sub
Private Sub Form_Load()
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
Private Sub picWiz01_Click()
    Call EnableFields(picWiz01.ToolTipText)
End Sub
Private Sub picWiz02_Click()
    Call EnableFields(picWiz02.ToolTipText)
End Sub
Private Sub picWiz03_Click()
    Call EnableFields(picWiz03.ToolTipText)
End Sub
Private Sub picWiz04_Click()
    Call EnableFields(picWiz04.ToolTipText)
End Sub
Private Sub picWiz05_Click()
    Call EnableFields(picWiz05.ToolTipText)
End Sub
Private Sub picWiz06_Click()
    Call EnableFields(picWiz06.ToolTipText)
End Sub
Private Sub picWiz07_Click()
    Call EnableFields(picWiz07.ToolTipText)
End Sub
Private Sub picWiz07g_Click()
    Call EnableFields(picWiz07g.ToolTipText)
End Sub

