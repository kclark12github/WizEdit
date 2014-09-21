VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WizEdit 2000© for The Ultimate Wizardry Archives®"
   ClientHeight    =   6420
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   7896
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":1CCA
   ScaleHeight     =   6420
   ScaleWidth      =   7896
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Height          =   312
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   1332
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H00808080&
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
      Height          =   312
      Left            =   300
      TabIndex        =   16
      Text            =   "Save Game File"
      Top             =   4800
      Width           =   4152
   End
   Begin VB.PictureBox picWiz07g 
      AutoSize        =   -1  'True
      Height          =   1524
      Left            =   6420
      Picture         =   "frmMain.frx":EABB
      ScaleHeight     =   1476
      ScaleWidth      =   1224
      TabIndex        =   15
      ToolTipText     =   "Wizardry 07G - Wizardry Gold"
      Top             =   3300
      Width           =   1272
   End
   Begin VB.PictureBox picWizardryLogo 
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   120
      Picture         =   "frmMain.frx":F4AE
      ScaleHeight     =   372
      ScaleWidth      =   1920
      TabIndex        =   14
      Top             =   300
      Visible         =   0   'False
      Width           =   1968
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   13
      Top             =   6168
      Width           =   7896
      _ExtentX        =   13928
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
            Object.Width           =   8530
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
            TextSave        =   "10:04 PM"
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
      Left            =   6672
      Picture         =   "frmMain.frx":FF8B
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
      Left            =   5590
      Picture         =   "frmMain.frx":106EE
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
      Left            =   4532
      Picture         =   "frmMain.frx":10EAE
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
      Left            =   3438
      Picture         =   "frmMain.frx":1162A
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
      Left            =   2356
      Picture         =   "frmMain.frx":11CB8
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
      Left            =   1262
      Picture         =   "frmMain.frx":12261
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
      Left            =   192
      Picture         =   "frmMain.frx":128BB
      ScaleHeight     =   1476
      ScaleWidth      =   972
      TabIndex        =   4
      ToolTipText     =   "Wizardry 01 - Proving Grounds of the Mad Overlord"
      Top             =   1680
      Width           =   1020
   End
   Begin VB.CommandButton cmdExit 
      DisabledPicture =   "frmMain.frx":12ED1
      DownPicture     =   "frmMain.frx":134A7
      Enabled         =   0   'False
      Height          =   432
      Left            =   5160
      Picture         =   "frmMain.frx":13A4F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1572
   End
   Begin VB.CommandButton cmdOK 
      DisabledPicture =   "frmMain.frx":14038
      DownPicture     =   "frmMain.frx":1460E
      Height          =   432
      Left            =   1200
      Picture         =   "frmMain.frx":14C1B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5460
      Width           =   1572
   End
   Begin VB.PictureBox picSirTech 
      AutoSize        =   -1  'True
      Height          =   468
      Left            =   6000
      Picture         =   "frmMain.frx":151F5
      ScaleHeight     =   420
      ScaleWidth      =   780
      TabIndex        =   1
      Top             =   660
      Width           =   828
   End
   Begin VB.PictureBox picUWA 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1236
      Left            =   2130
      Picture         =   "frmMain.frx":1569F
      ScaleHeight     =   1188
      ScaleWidth      =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   3648
   End
   Begin VB.Label Label1 
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
      Left            =   300
      TabIndex        =   17
      Top             =   4500
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
      Left            =   240
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
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    'cmdOK.Enabled = False
    cmdOK.Visible = False
    cmdExit.Enabled = True
    lblScenario.Caption = "Please select a scenario by clicking on the appropriate game box above..."
    lblSelectedScenario.Visible = False
    
    'See if we know where UWA is installed... If not, use defaults from UWA...
    gstrUWAPath1 = GetWizEditSetting("Environment", "UWAPath1", vbNullString)
    If gstrUWAPath1 = vbNullString Then gstrUWAPath1 = "C:\WIZARD15"
    gstrUWAPath2 = GetWizEditSetting("Environment", "UWAPath2", vbNullString)
    If gstrUWAPath2 = vbNullString Then gstrUWAPath2 = "C:\WIZARD15"
    gstrUWAPath3 = GetWizEditSetting("Environment", "UWAPath3", vbNullString)
    If gstrUWAPath3 = vbNullString Then gstrUWAPath3 = "C:\WIZARD15"
    gstrUWAPath4 = GetWizEditSetting("Environment", "UWAPath4", vbNullString)
    If gstrUWAPath4 = vbNullString Then gstrUWAPath4 = "C:\WIZARD15"
    gstrUWAPath5 = GetWizEditSetting("Environment", "UWAPath5", vbNullString)
    If gstrUWAPath5 = vbNullString Then gstrUWAPath5 = "C:\WIZARD15"
    gstrUWAPath6 = GetWizEditSetting("Environment", "UWAPath6", vbNullString)
    If gstrUWAPath6 = vbNullString Then gstrUWAPath6 = "C:\BANE"
    gstrUWAPath7 = GetWizEditSetting("Environment", "UWAPath7", vbNullString)
    If gstrUWAPath7 = vbNullString Then gstrUWAPath7 = "C:\DSAVANT"
    gstrUWAPath7g = GetWizEditSetting("Environment", "UWAPath7g", vbNullString)
    If gstrUWAPath7g = vbNullString Then gstrUWAPath7g = "C:\Program Files\Sir-Tech\WizGold"
    
    '
End Sub
Private Sub picWiz01_Click()
    lblScenario.Caption = "Scenario Selected..."
    lblSelectedScenario.Caption = picWiz01.ToolTipText
    lblSelectedScenario.Visible = True
    cmdOK.Visible = True
End Sub
Private Sub picWiz02_Click()
    lblScenario.Caption = "Scenario Selected..."
    lblSelectedScenario.Caption = picWiz02.ToolTipText
    lblSelectedScenario.Visible = True
    cmdOK.Visible = True
End Sub
Private Sub picWiz03_Click()
    lblScenario.Caption = "Scenario Selected..."
    lblSelectedScenario.Caption = picWiz03.ToolTipText
    lblSelectedScenario.Visible = True
    cmdOK.Visible = True
End Sub
Private Sub picWiz04_Click()
    lblScenario.Caption = "Scenario Selected..."
    lblSelectedScenario.Caption = picWiz04.ToolTipText
    lblSelectedScenario.Visible = True
    cmdOK.Visible = True
End Sub
Private Sub picWiz05_Click()
    lblScenario.Caption = "Scenario Selected..."
    lblSelectedScenario.Caption = picWiz05.ToolTipText
    lblSelectedScenario.Visible = True
    cmdOK.Visible = True
End Sub
Private Sub picWiz06_Click()
    lblScenario.Caption = "Scenario Selected..."
    lblSelectedScenario.Caption = picWiz06.ToolTipText
    lblSelectedScenario.Visible = True
    cmdOK.Visible = True
End Sub
Private Sub picWiz07_Click()
    lblScenario.Caption = "Scenario Selected..."
    lblSelectedScenario.Caption = picWiz07.ToolTipText
    lblSelectedScenario.Visible = True
    cmdOK.Visible = True
End Sub
Private Sub picWiz07g_Click()
    lblScenario.Caption = "Scenario Selected..."
    lblSelectedScenario.Caption = picWiz07g.ToolTipText
    lblSelectedScenario.Visible = True
    cmdOK.Visible = True
End Sub

