VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWiz07 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wizardry 07 - Crusaders of the Dark Savant"
   ClientHeight    =   6444
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   11256
   Icon            =   "frmWiz07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6444
   ScaleWidth      =   11256
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picWizardryLogo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   420
      Left            =   4260
      ScaleHeight     =   372
      ScaleWidth      =   1920
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   5700
      Width           =   1968
   End
   Begin VB.PictureBox picFrames 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5112
      Index           =   2
      Left            =   120
      ScaleHeight     =   5088
      ScaleWidth      =   9108
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   1020
      Width           =   9132
      Begin VB.TextBox txtSkillsAcademiaDesc 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   552
         Left            =   1380
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   3960
         Width           =   7272
      End
      Begin VB.TextBox txtSkillsPersonalDesc 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   552
         Left            =   1380
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   2880
         Width           =   7272
      End
      Begin VB.TextBox txtSkillsPhysicalDesc 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   552
         Left            =   1380
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   1800
         Width           =   7272
      End
      Begin VB.TextBox txtSkillsWeaponryDesc 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   552
         Left            =   1380
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   720
         Width           =   7272
      End
      Begin VB.ComboBox cboSkillsAcademia 
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
         Height          =   336
         ItemData        =   "frmWiz07.frx":1CCA
         Left            =   1380
         List            =   "frmWiz07.frx":1CCC
         Style           =   2  'Dropdown List
         TabIndex        =   118
         ToolTipText     =   "Character's Academia Skills..."
         Top             =   3552
         Width           =   2712
      End
      Begin VB.TextBox txtSkillsAcademia 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4140
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   116
         Text            =   "frmWiz07.frx":1CCE
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3540
         Width           =   576
      End
      Begin VB.ComboBox cboSkillsPersonal 
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
         Height          =   336
         ItemData        =   "frmWiz07.frx":1CD4
         Left            =   1380
         List            =   "frmWiz07.frx":1CD6
         Style           =   2  'Dropdown List
         TabIndex        =   115
         ToolTipText     =   "Character's Personal Skills..."
         Top             =   2472
         Width           =   2712
      End
      Begin VB.TextBox txtSkillsPersonal 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4140
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   113
         Text            =   "frmWiz07.frx":1CD8
         ToolTipText     =   "Wand && Dagger..."
         Top             =   2460
         Width           =   576
      End
      Begin VB.ComboBox cboSkillsPhysical 
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
         Height          =   336
         ItemData        =   "frmWiz07.frx":1CDE
         Left            =   1380
         List            =   "frmWiz07.frx":1CE0
         Style           =   2  'Dropdown List
         TabIndex        =   112
         ToolTipText     =   "Character's Physical Skills..."
         Top             =   1392
         Width           =   2712
      End
      Begin VB.TextBox txtSkillsPhysical 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4140
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   110
         Text            =   "frmWiz07.frx":1CE2
         ToolTipText     =   "Wand && Dagger..."
         Top             =   1380
         Width           =   576
      End
      Begin VB.ComboBox cboSkillsWeaponry 
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
         Height          =   336
         ItemData        =   "frmWiz07.frx":1CE8
         Left            =   1380
         List            =   "frmWiz07.frx":1CEA
         Style           =   2  'Dropdown List
         TabIndex        =   109
         ToolTipText     =   "Character's Weaponry Skills..."
         Top             =   336
         Width           =   2712
      End
      Begin VB.TextBox txtSkillsWeaponry 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4140
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   107
         Text            =   "frmWiz07.frx":1CEC
         ToolTipText     =   "Wand && Dagger..."
         Top             =   324
         Width           =   576
      End
      Begin MSComCtl2.UpDown udSkillsWeaponry 
         Height          =   360
         Left            =   4680
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   324
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSkillsWeaponry"
         BuddyDispid     =   196710
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udSkillsPhysical 
         Height          =   360
         Left            =   4680
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   1380
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSkillsPhysical"
         BuddyDispid     =   196713
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown usSkillsPersonal 
         Height          =   360
         Left            =   4680
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   2460
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSkillsPersonal"
         BuddyDispid     =   196718
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown usSkillsAcademia 
         Height          =   360
         Left            =   4680
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   3540
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSkillsAcademia"
         BuddyDispid     =   196722
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblSkillsAcademia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Academia:"
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
         Left            =   216
         TabIndex        =   117
         Top             =   3576
         Width           =   1032
      End
      Begin VB.Label lblSkillsPersonal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal:"
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
         Left            =   252
         TabIndex        =   114
         Top             =   2496
         Width           =   996
      End
      Begin VB.Label lblSkillsPhysical 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Physical:"
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
         Left            =   300
         TabIndex        =   111
         Top             =   1416
         Width           =   948
      End
      Begin VB.Label lblSkillsWeaponry 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weaponry:"
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
         Left            =   120
         TabIndex        =   108
         Top             =   360
         Width           =   1128
      End
   End
   Begin VB.PictureBox picFrames 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5112
      Index           =   1
      Left            =   120
      ScaleHeight     =   5088
      ScaleWidth      =   9108
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1020
      Width           =   9132
      Begin VB.TextBox txtDivineSP 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   7620
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   27
         Text            =   "frmWiz07.frx":1CF2
         ToolTipText     =   "Divine Spell Points (automatically ""topped-off"")..."
         Top             =   3780
         Width           =   996
      End
      Begin VB.TextBox txtMentalSP 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   7620
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "frmWiz07.frx":1CF5
         ToolTipText     =   "Mental Spell Points (automatically ""topped-off"")..."
         Top             =   3360
         Width           =   996
      End
      Begin VB.TextBox txtEarthSP 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   7620
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "frmWiz07.frx":1CF8
         ToolTipText     =   "Earth Spell Points (automatically ""topped-off"")..."
         Top             =   2940
         Width           =   996
      End
      Begin VB.TextBox txtAirSP 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   7620
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "frmWiz07.frx":1CFB
         ToolTipText     =   "Air Spell Points (automatically ""topped-off"")..."
         Top             =   2520
         Width           =   996
      End
      Begin VB.TextBox txtWaterSP 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   7620
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frmWiz07.frx":1CFE
         ToolTipText     =   "Water Spell Points (automatically ""topped-off"")..."
         Top             =   2100
         Width           =   996
      End
      Begin VB.TextBox txtFireSP 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   7620
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "frmWiz07.frx":1D01
         ToolTipText     =   "Fire Spell Points (automatically ""topped-off"")..."
         Top             =   1680
         Width           =   996
      End
      Begin VB.TextBox txtAge 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4680
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmWiz07.frx":1D04
         ToolTipText     =   "Age...? (Still not sure of this one)..."
         Top             =   840
         Width           =   996
      End
      Begin VB.TextBox txtLives 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4680
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmWiz07.frx":1D07
         ToolTipText     =   "Lives..."
         Top             =   420
         Width           =   996
      End
      Begin VB.TextBox txtLVL 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4680
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmWiz07.frx":1D0A
         ToolTipText     =   "Level..."
         Top             =   0
         Width           =   996
      End
      Begin VB.TextBox txtGP 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   7380
         MaxLength       =   13
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmWiz07.frx":1D0D
         ToolTipText     =   "Gold Pieces..."
         Top             =   840
         Width           =   1476
      End
      Begin VB.TextBox txtMKS 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   7380
         MaxLength       =   13
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "frmWiz07.frx":1D10
         ToolTipText     =   "Monster Kills..."
         Top             =   420
         Width           =   1476
      End
      Begin VB.TextBox txtEXP 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   7380
         MaxLength       =   13
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmWiz07.frx":1D13
         ToolTipText     =   "Experience Points..."
         Top             =   0
         Width           =   1476
      End
      Begin VB.TextBox txtCC 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4320
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "frmWiz07.frx":1D16
         ToolTipText     =   "Carrying Capacity (automatically ""topped-off"")..."
         Top             =   2520
         Width           =   996
      End
      Begin VB.TextBox txtSTM 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4320
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "frmWiz07.frx":1D19
         ToolTipText     =   "Stamina (automatically ""topped-off"")..."
         Top             =   2100
         Width           =   996
      End
      Begin VB.TextBox txtHP 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4320
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "frmWiz07.frx":1D1C
         ToolTipText     =   "Hit Points (automatically ""topped-off"")..."
         Top             =   1680
         Width           =   996
      End
      Begin VB.ComboBox cboCondition 
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
         Height          =   336
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Character's Condition (i.e. OK, Afraid, Poisoned, etc.)..."
         Top             =   2940
         Width           =   1632
      End
      Begin VB.ComboBox cboProfession 
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
         Height          =   336
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Character's Profession (i.e. Fighter, Mage, etc.)..."
         Top             =   780
         Width           =   1692
      End
      Begin VB.ComboBox cboRace 
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
         Height          =   336
         ItemData        =   "frmWiz07.frx":1D1F
         Left            =   1440
         List            =   "frmWiz07.frx":1D26
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Character's Race (i.e. Human, Elf, etc.)..."
         Top             =   60
         Width           =   1692
      End
      Begin MSComCtl2.UpDown udPIE 
         Height          =   360
         Left            =   2400
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPIE"
         BuddyDispid     =   196641
         OrigLeft        =   120
         OrigTop         =   1560
         OrigRight       =   360
         OrigBottom      =   2172
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtPIE 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   2040
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "frmWiz07.frx":1D31
         ToolTipText     =   "Piety..."
         Top             =   2520
         Width           =   396
      End
      Begin VB.TextBox txtKAR 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   2040
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "frmWiz07.frx":1D34
         ToolTipText     =   "Karma..."
         Top             =   4620
         Width           =   396
      End
      Begin VB.TextBox txtPER 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   2040
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "frmWiz07.frx":1D37
         ToolTipText     =   "Personality..."
         Top             =   4200
         Width           =   396
      End
      Begin VB.TextBox txtSPD 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   2040
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "frmWiz07.frx":1D3A
         ToolTipText     =   "Speed..."
         Top             =   3780
         Width           =   396
      End
      Begin VB.TextBox txtDEX 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   2040
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "frmWiz07.frx":1D3D
         ToolTipText     =   "Dexterity..."
         Top             =   3360
         Width           =   396
      End
      Begin VB.TextBox txtVIT 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   2040
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmWiz07.frx":1D40
         ToolTipText     =   "Vitality..."
         Top             =   2940
         Width           =   396
      End
      Begin VB.TextBox txtINT 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   2040
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmWiz07.frx":1D43
         ToolTipText     =   "Intelligence..."
         Top             =   2100
         Width           =   396
      End
      Begin VB.ComboBox cboGender 
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
         Height          =   336
         ItemData        =   "frmWiz07.frx":1D46
         Left            =   1440
         List            =   "frmWiz07.frx":1D50
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Character's Gender (Male/Female)..."
         Top             =   420
         Width           =   1692
      End
      Begin VB.TextBox txtSTR 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   2040
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmWiz07.frx":1D62
         ToolTipText     =   "Strength..."
         Top             =   1680
         Width           =   396
      End
      Begin MSComCtl2.UpDown udSTR 
         Height          =   360
         Left            =   2400
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSTR"
         BuddyDispid     =   196649
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udINT 
         Height          =   360
         Left            =   2400
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2100
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtINT"
         BuddyDispid     =   196647
         OrigLeft        =   2280
         OrigTop         =   720
         OrigRight       =   2520
         OrigBottom      =   1032
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udVIT 
         Height          =   360
         Left            =   2400
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2940
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtVIT"
         BuddyDispid     =   196646
         OrigRight       =   240
         OrigBottom      =   612
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udDEX 
         Height          =   360
         Left            =   2400
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtDEX"
         BuddyDispid     =   196645
         OrigRight       =   240
         OrigBottom      =   612
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udSPD 
         Height          =   360
         Left            =   2400
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   3780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSPD"
         BuddyDispid     =   196644
         OrigRight       =   240
         OrigBottom      =   612
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPER 
         Height          =   360
         Left            =   2400
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   4200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPER"
         BuddyDispid     =   196643
         OrigRight       =   240
         OrigBottom      =   612
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udKAR 
         Height          =   360
         Left            =   2400
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   4620
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtKAR"
         BuddyDispid     =   196642
         OrigRight       =   240
         OrigBottom      =   612
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udHP 
         Height          =   360
         Left            =   5280
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtHP"
         BuddyDispid     =   196637
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udSTM 
         Height          =   360
         Left            =   5280
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2100
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSTM"
         BuddyDispid     =   196636
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udCC 
         Height          =   360
         Left            =   5280
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtCC"
         BuddyDispid     =   196635
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udEXP 
         Height          =   360
         Left            =   8820
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtEXP"
         BuddyDispid     =   196634
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMKS 
         Height          =   360
         Left            =   8820
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMKS"
         BuddyDispid     =   196633
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udGP 
         Height          =   360
         Left            =   8820
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtGP"
         BuddyDispid     =   196632
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udLVL 
         Height          =   360
         Left            =   5640
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLVL"
         BuddyDispid     =   196631
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udLives 
         Height          =   360
         Left            =   5640
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLives"
         BuddyDispid     =   196630
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udAge 
         Height          =   360
         Left            =   5640
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtAge"
         BuddyDispid     =   196629
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udFireSP 
         Height          =   360
         Left            =   8580
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtFireSP"
         BuddyDispid     =   196628
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udWaterSP 
         Height          =   360
         Left            =   8580
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   2100
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtWaterSP"
         BuddyDispid     =   196627
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udAirSP 
         Height          =   360
         Left            =   8580
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtAirSP"
         BuddyDispid     =   196626
         OrigLeft        =   8460
         OrigTop         =   2940
         OrigRight       =   8700
         OrigBottom      =   3552
         Max             =   65535
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udEarthSP 
         Height          =   360
         Left            =   8580
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   2940
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtEarthSP"
         BuddyDispid     =   196625
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMentalSP 
         Height          =   360
         Left            =   8580
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   3360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMentalSP"
         BuddyDispid     =   196624
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udDivineSP 
         Height          =   360
         Left            =   8580
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   3780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtDivineSP"
         BuddyDispid     =   196623
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl2Stats 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Secondary Statistics:"
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
         Left            =   3012
         TabIndex        =   67
         Top             =   1380
         Width           =   2148
      End
      Begin VB.Label lbl1Stats 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Statistics:"
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
         Left            =   180
         TabIndex        =   66
         Top             =   1380
         Width           =   1608
      End
      Begin VB.Label lblGender 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
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
         Left            =   60
         TabIndex        =   63
         Top             =   444
         Width           =   780
      End
      Begin VB.Label lblKAR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Karma:"
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
         Left            =   1104
         TabIndex        =   61
         Top             =   4656
         Width           =   756
      End
      Begin VB.Label lblPER 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personality:"
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
         Left            =   576
         TabIndex        =   60
         Top             =   4236
         Width           =   1284
      End
      Begin VB.Label lblSPD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
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
         Left            =   1212
         TabIndex        =   59
         Top             =   3816
         Width           =   648
      End
      Begin VB.Label lblDEX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dexterity:"
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
         Left            =   804
         TabIndex        =   58
         Top             =   3396
         Width           =   1056
      End
      Begin VB.Label lblVIT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vitality:"
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
         Left            =   960
         TabIndex        =   57
         Top             =   2976
         Width           =   900
      End
      Begin VB.Label lblPIE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Piety:"
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
         Left            =   1224
         TabIndex        =   56
         Top             =   2556
         Width           =   636
      End
      Begin VB.Label lblINT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intelligence:"
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
         Left            =   600
         TabIndex        =   55
         Top             =   2136
         Width           =   1260
      End
      Begin VB.Label lblSTR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strength:"
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
         Left            =   936
         TabIndex        =   54
         Top             =   1716
         Width           =   924
      End
      Begin VB.Label lblSPDivine 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Divine:"
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
         Left            =   6684
         TabIndex        =   53
         Top             =   3816
         Width           =   744
      End
      Begin VB.Label lblSPMental 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mental:"
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
         Left            =   6612
         TabIndex        =   52
         Top             =   3396
         Width           =   816
      End
      Begin VB.Label lblSPAir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Air:"
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
         Left            =   7020
         TabIndex        =   51
         Top             =   2556
         Width           =   408
      End
      Begin VB.Label lblSPEarth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Earth:"
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
         Left            =   6792
         TabIndex        =   50
         Top             =   2976
         Width           =   636
      End
      Begin VB.Label lblSPWater 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Water:"
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
         Left            =   6720
         TabIndex        =   49
         Top             =   2136
         Width           =   708
      End
      Begin VB.Label lblSPFire 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fire:"
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
         Left            =   6912
         TabIndex        =   48
         Top             =   1716
         Width           =   516
      End
      Begin VB.Label lblSP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spell Points:"
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
         Left            =   6480
         TabIndex        =   47
         Top             =   1380
         Width           =   1332
      End
      Begin VB.Label lblCC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity:"
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
         Left            =   3276
         TabIndex        =   46
         Top             =   2556
         Width           =   936
      End
      Begin VB.Label lblSTM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stamina:"
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
         Left            =   3312
         TabIndex        =   45
         Top             =   2136
         Width           =   900
      End
      Begin VB.Label lblHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hit Points:"
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
         Left            =   3048
         TabIndex        =   44
         Top             =   1716
         Width           =   1164
      End
      Begin VB.Label lblGP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gold:"
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
         Left            =   6744
         TabIndex        =   43
         Top             =   876
         Width           =   528
      End
      Begin VB.Label lblMKS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kills:"
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
         Left            =   6696
         TabIndex        =   42
         Top             =   456
         Width           =   576
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Experience:"
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
         Left            =   6072
         TabIndex        =   41
         Top             =   36
         Width           =   1200
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age?:"
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
         Left            =   4020
         TabIndex        =   40
         Top             =   876
         Width           =   552
      End
      Begin VB.Label lblLife 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life:"
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
         Left            =   4080
         TabIndex        =   39
         Top             =   456
         Width           =   480
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
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
         Left            =   3960
         TabIndex        =   38
         Top             =   36
         Width           =   624
      End
      Begin VB.Label lblCondition 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condition:"
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
         Left            =   3180
         TabIndex        =   37
         Top             =   2940
         Width           =   1032
      End
      Begin VB.Label lblProfession 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Profession:"
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
         Left            =   60
         TabIndex        =   36
         Top             =   804
         Width           =   1188
      End
      Begin VB.Label lblRace 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Race:"
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
         Left            =   60
         TabIndex        =   35
         Top             =   84
         Width           =   576
      End
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Heidelberg"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   9720
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmWiz07.frx":1D65
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2460
      UseMaskColor    =   -1  'True
      Width           =   1212
   End
   Begin VB.PictureBox picPortrait 
      BackColor       =   &H00000000&
      Height          =   2076
      Left            =   9420
      ScaleHeight     =   2028
      ScaleWidth      =   1668
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   120
      Width           =   1716
      Begin VB.Label lblPortrait 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Character Portrait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Left            =   60
         TabIndex        =   94
         Top             =   900
         Width           =   1608
      End
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Height          =   432
      Left            =   9720
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2460
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Height          =   432
      Left            =   9720
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2940
      Width           =   1212
   End
   Begin VB.PictureBox picWiz07 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2064
      Left            =   9600
      Picture         =   "frmWiz07.frx":32B4
      ScaleHeight     =   2016
      ScaleWidth      =   1344
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1392
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   1
      Left            =   120
      Picture         =   "frmWiz07.frx":3CFF
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   660
      Width           =   1152
      Begin VB.Label lblStats 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Statistics:"
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
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Width           =   1032
      End
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   2
      Left            =   1272
      Picture         =   "frmWiz07.frx":1D521
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   660
      Width           =   1152
      Begin VB.Label lblSkills 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skills:"
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
         Left            =   60
         TabIndex        =   96
         Top             =   0
         Width           =   1032
      End
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   3
      Left            =   2424
      Picture         =   "frmWiz07.frx":36D43
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   660
      Width           =   1152
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Items:"
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
         Left            =   60
         TabIndex        =   98
         Top             =   0
         Width           =   1032
      End
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   4
      Left            =   3576
      Picture         =   "frmWiz07.frx":50565
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   660
      Width           =   1152
      Begin VB.Label lblSwagBag 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Swag Bag:"
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
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Width           =   1032
      End
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   5
      Left            =   4728
      Picture         =   "frmWiz07.frx":69D87
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   660
      Width           =   1152
      Begin VB.Label lblSpells 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spells:"
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
         Left            =   60
         TabIndex        =   102
         Top             =   0
         Width           =   1032
      End
   End
   Begin VB.PictureBox picFrames 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5112
      Index           =   3
      Left            =   120
      ScaleHeight     =   5088
      ScaleWidth      =   9108
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   1020
      Width           =   9132
   End
   Begin VB.PictureBox picFrames 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5112
      Index           =   4
      Left            =   120
      ScaleHeight     =   5088
      ScaleWidth      =   9108
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   1020
      Width           =   9132
   End
   Begin VB.PictureBox picFrames 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5112
      Index           =   5
      Left            =   120
      ScaleHeight     =   5088
      ScaleWidth      =   9108
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   1020
      Width           =   9132
   End
   Begin VB.ComboBox cboCharacter 
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
      Height          =   336
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4212
   End
   Begin VB.PictureBox picWiz07Gold 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2076
      Left            =   9420
      Picture         =   "frmWiz07.frx":835A9
      ScaleHeight     =   2028
      ScaleWidth      =   1668
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1716
   End
   Begin MSComctlLib.ImageList imgIcons32 
      Left            =   10140
      Top             =   4500
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWiz07.frx":8447C
            Key             =   "Wiz07"
            Object.Tag             =   "Wiz07"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWiz07.frx":86158
            Key             =   "Wiz07g"
            Object.Tag             =   "Wiz07g"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons16 
      Left            =   10140
      Top             =   5040
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWiz07.frx":86474
            Key             =   "Wiz07"
            Object.Tag             =   "Wiz07"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWiz07.frx":88150
            Key             =   "Wiz07g"
            Object.Tag             =   "Wiz07g"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   62
      Top             =   6192
      Width           =   11256
      _ExtentX        =   19854
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
            Object.Width           =   14457
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
            TextSave        =   "10:45 PM"
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
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Heidelberg"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   9720
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2940
      Width           =   1212
   End
   Begin VB.Label lblCharacter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Character:"
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
      Left            =   120
      TabIndex        =   33
      Top             =   264
      Width           =   1020
   End
End
Attribute VB_Name = "frmWiz07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const EnabledColor = &HFFFF&
Const DisabledColor = &HC0C0&
Public DataFile As String
Private SaveMessage As String
Private SelectedCharacter As Integer
Private Characters(1 To 6) As Wiz07Character
Private Sub cboCharacter_Click()
    sbStatus.Panels("Message").Text = vbNullString
    sbStatus.Panels("Status").Text = vbNullString
    
    If cboCharacter.ListIndex = -1 Then
        sbStatus.Panels("Message").Text = "Select Character from the List..."
        cmdEdit.Visible = False
        Exit Sub
    End If
    cmdEdit.Visible = True
    SelectedCharacter = cboCharacter.ListIndex + 1
    With Characters(SelectedCharacter)
        cboCondition.ListIndex = .ConditionCode
        cboGender.ListIndex = .Gender
        cboProfession.ListIndex = .Profession
        cboRace.ListIndex = .Race
        
        txtSTR.Text = .STR
        txtINT.Text = .INT
        txtPIE.Text = .PIE
        txtVIT.Text = .VIT
        txtDEX.Text = .DEX
        txtSPD.Text = .SPD
        txtPER.Text = .PER
        txtKAR.Text = .KAR
        
        txtLVL.Text = Format(.Level, "#,##0")
        txtLives.Text = Format(.Lives, "#,##0")
        txtAge.Text = Format(.AgeInWeeks \ 52, "#,##0")
        
        txtHP.Text = Format(.HP.Maximum, "#,##0")
        txtSTM.Text = Format(.STM.Maximum, "#,##0")
        txtCC.Text = Format(.CC.Maximum / 10, "#,##0.0")
        txtEXP.Text = Format(.EXP, "#,##0")
        txtMKS.Text = Format(.MKS, "#,##0")
        txtGP.Text = Format(.GP, "#,##0")
    
        txtFireSP.Text = Format(.FireSpellPoints.Maximum, "#,##0")
        txtWaterSP.Text = Format(.WaterSpellPoints.Maximum, "#,##0")
        txtAirSP.Text = Format(.AirSpellPoints.Maximum, "#,##0")
        txtEarthSP.Text = Format(.EarthSpellPoints.Maximum, "#,##0")
        txtMentalSP.Text = Format(.MentalSpellPoints.Maximum, "#,##0")
        txtDivineSP.Text = Format(.DivineSpellPoints.Maximum, "#,##0")
                
'            Debug.Print vbCrLf & "List of Items (not stowed)..."
'            For i = 1 To 10
'                Debug.Print strItem(.ItemList(i))
'            Next i
'
'            Debug.Print vbCrLf & "List of Stowed items..."
'            For i = 1 To 10
'                Debug.Print strItem(.SwagBag(i))
'            Next i
'
        cboSkillsAcademia.ListIndex = 0
        cboSkillsPersonal.ListIndex = 0
        cboSkillsPhysical.ListIndex = 0
        cboSkillsWeaponry.ListIndex = 0

'            Debug.Print "'Natural' Armor Class:" & vbTab & .NaturalArmorClass
'
'            Debug.Print vbCrLf & "SpellBooks..."
'            For j = 1 To 12
'                For i = 1 To 8
'                    Debug.Print vbTab & strSpell(((j - 1) * 8) + i, .SpellBooks(j), i - 1)
'                Next i
'            Next j
'
'            Debug.Print "PictureCode:       " & vbTab & .PictureCode
        
        sbStatus.Panels("Status").Text = .Name
        sbStatus.Panels("Message").Text = "Level " & .Level & " " & cboProfession.List(cboProfession.ListIndex) & "..."
    End With
End Sub
Private Sub cboSkillsAcademia_Click()
    With Characters(SelectedCharacter)
        Select Case cboSkillsAcademia.ListIndex
            Case 0
                txtSkillsAcademia.Text = .Artifacts
                txtSkillsAcademia.ToolTipText = cboSkillsAcademia.Text & _
                    ": The ability to effectively use and invoke magical items depends on this skill. " & _
                    "Without a developed Artifact skill, there's a chance the item's power will fizzle or backfire. " & _
                    "This skill also affects a character's ability to successfully assay an item to determine its intricacies."
            Case 1
                txtSkillsAcademia.Text = .Mythology
                txtSkillsAcademia.ToolTipText = cboSkillsAcademia.Text & _
                    ": The ability to recognize, while in combat, the true identities of monsters."
            Case 2
                txtSkillsAcademia.Text = .Mapping
                txtSkillsAcademia.ToolTipText = cboSkillsAcademia.Text & _
                    ": The ability to transcribe an accurate record of the party's adventure. " & _
                    "The higher ths skill, the more detail (doors, stairs, trees, gates, etc.) included. " & _
                    "This skill requires a mapping kit to be effective."
            Case 3
                txtSkillsAcademia.Text = .Scribe
                txtSkillsAcademia.ToolTipText = cboSkillsAcademia.Text & _
                    ": The ability to successfully invoke the magical power of a scroll during combat."
            Case 4
                txtSkillsAcademia.Text = .Diplomacy
                txtSkillsAcademia.ToolTipText = cboSkillsAcademia.Text & _
                    ": The art of negotiation and creation of mutual pacts and trust between the party and another group. " & _
                    "Allows the negotiator to truce well and form alliances with NPCs."
            Case 5
                txtSkillsAcademia.Text = .Alchemy
                txtSkillsAcademia.ToolTipText = cboSkillsAcademia.Text & _
                    ": The art of learning, practicing, and exercising Alchemist spells."
            Case 6
                txtSkillsAcademia.Text = .Theology
                txtSkillsAcademia.ToolTipText = cboSkillsAcademia.Text & _
                    ": The pursuit of the divine interests leading to the study of Priest spells."
            Case 7
                txtSkillsAcademia.Text = .Theosophy
                txtSkillsAcademia.ToolTipText = cboSkillsAcademia.Text & _
                    ": The possession of mental and spiritual insight that allows its possessor to study Psionic spells."
            Case 8
                txtSkillsAcademia.Text = .Thaumaturgy
                txtSkillsAcademia.ToolTipText = cboSkillsAcademia.Text & _
                    ": The path of study followed by the Mage and those who follow him to learn Mage spells."
            Case 9
                txtSkillsAcademia.Text = .Kirijutsu
                txtSkillsAcademia.ToolTipText = cboSkillsAcademia.Text & _
                    ": The deadly skill and knowledge of the body which allows its possessor to strike a vital blow or critical area, hopefully killing an opponent with a single blow."
            Case Else
        End Select
    End With
    txtSkillsAcademiaDesc.Text = txtSkillsAcademia.ToolTipText
End Sub
Private Sub cboSkillsPersonal_Click()
    With Characters(SelectedCharacter)
        Select Case cboSkillsPersonal.ListIndex
            Case 0
                txtSkillsPersonal.Text = .Firearms
                txtSkillsPersonal.ToolTipText = cboSkillsPersonal.Text & _
                    ": The use of small firearms like muskets. " & _
                    "This skill determines your character's ability to load and accurately fire such a weapon."
            Case 1
                txtSkillsPersonal.Text = .Reflextion
                txtSkillsPersonal.ToolTipText = cboSkillsPersonal.Text & _
                    ": The ability to take small jumps so quickly - faster than the eye - that a double image is created. " & _
                    "This skill's most effective use is in a character's ability to avoid an attack. " & _
                    "The ""After Image"" is usually the target."
            Case 2
                txtSkillsPersonal.Text = .SnakeSpeed
                txtSkillsPersonal.ToolTipText = cboSkillsPersonal.Text & _
                    ": Allows characters to move with lightning reflexes, increasing speed in all aspects where it is a factor."
            Case 3
                txtSkillsPersonal.Text = .EagleEye
                txtSkillsPersonal.ToolTipText = cboSkillsPersonal.Text & _
                    ": The ability to target a creature with a weapon or spell and strike with an amazing degree of accuracy."
            Case 4
                txtSkillsPersonal.Text = .PowerStrike
                txtSkillsPersonal.ToolTipText = cboSkillsPersonal.Text & _
                    ": An ability to strike a blow that does maximum damage - and sometimes yields more than that!"
            Case 5
                txtSkillsPersonal.Text = .MindControl
                txtSkillsPersonal.ToolTipText = cboSkillsPersonal.Text & _
                    ": Those adept in this skill can master control of their own minds. " & _
                    "This extra willpower helps them fend off sleep or Psionic spells and to retain control of their own mind."
            Case Else
        End Select
    End With
    txtSkillsPersonalDesc.Text = txtSkillsPersonal.ToolTipText
End Sub
Private Sub cboSkillsPhysical_Click()
    With Characters(SelectedCharacter)
        Select Case cboSkillsPhysical.ListIndex
            Case 0
                txtSkillsPhysical.Text = .Swimming
                txtSkillsPhysical.ToolTipText = cboSkillsPhysical.Text & _
                    ": A measurement of your character's ability to swim across deep water. " & _
                    "Characters with fewer than 10 skill points may drown from entering deep water."
            Case 1
                txtSkillsPhysical.Text = .Climbing
                txtSkillsPhysical.ToolTipText = cboSkillsPhysical.Text & _
                    ": The knack of taking falls, climbing into pits, and scaling the sides of walls without taking damage."
            Case 2
                txtSkillsPhysical.Text = .Scouting
                txtSkillsPhysical.ToolTipText = cboSkillsPhysical.Text & _
                    ": The Knack of seeing and finding things such as secret doors, hidden entrances or buried items others seem to pass by. " & _
                    "You must add points manually to ""Scout"" to increase your character's proficiency."
            Case 3
                txtSkillsPhysical.Text = .Music
                txtSkillsPhysical.ToolTipText = cboSkillsPhysical.Text & _
                    ": The art of playing enchanted musical instruments and bringing forth from them different magical spells."
            Case 4
                txtSkillsPhysical.Text = .Oratory
                txtSkillsPhysical.ToolTipText = cboSkillsPhysical.Text & _
                    ": The vocal discipline required to properly recite a magical spell in combat. " & _
                    "Without good oratory, spells meant for monsters may fizzle or backfire on the party."
            Case 5
                txtSkillsPhysical.Text = .Legerdemain
                txtSkillsPhysical.ToolTipText = cboSkillsPhysical.Text & _
                    ": The ability to pickpocket (steal) items or gold from NPCs without their knowledge or permission."
            Case 6
                txtSkillsPhysical.Text = .Skulduggery
                txtSkillsPhysical.ToolTipText = cboSkillsPhysical.Text & _
                    ": The delicate skill of inspecting and disarming traps on chects and picking locks on doors."
            Case 7
                txtSkillsPhysical.Text = .Ninjutsu
                txtSkillsPhysical.ToolTipText = cboSkillsPhysical.Text & _
                    ": THe legendary art that allows characters to hide themselves from their opponents." & _
                    " For the Ninja and Monk, proficiency in Ninjitsu helps to lower their armor class rating."
            Case Else
        End Select
    End With
    txtSkillsPhysicalDesc.Text = txtSkillsPhysical.ToolTipText
End Sub
Private Sub cboSkillsWeaponry_Click()
    With Characters(SelectedCharacter)
        Select Case cboSkillsWeaponry.ListIndex
            Case 0
                txtSkillsWeaponry.Text = .Wand
                txtSkillsWeaponry.ToolTipText = cboSkillsWeaponry.Text & _
                    ": The talent of weilding daggers, wands, and other small items used as weapons in combat."
            Case 1
                txtSkillsWeaponry.Text = .Sword
                txtSkillsWeaponry.ToolTipText = cboSkillsWeaponry.Text & _
                    ": Any sword, including the katana, used as a weapon in combat is covered under this skill."
            Case 2
                txtSkillsWeaponry.Text = .Axe
                txtSkillsWeaponry.ToolTipText = cboSkillsWeaponry.Text & _
                    ": This ability covers any axe, such as the battle axe or hand axe, used as a weapon in combat."
            Case 3
                txtSkillsWeaponry.Text = .Mace
                txtSkillsWeaponry.ToolTipText = cboSkillsWeaponry.Text & _
                    ": The talent needed to use any mace-like item, including the flail or hammer, as a weapon in combat."
            Case 4
                txtSkillsWeaponry.Text = .Pole
                txtSkillsWeaponry.ToolTipText = cboSkillsWeaponry.Text & _
                    ": The mastery of any pole & staff, such as the halberd, bo, or staff, used as a weapon in combat."
            Case 5
                txtSkillsWeaponry.Text = .Throwing
                txtSkillsWeaponry.ToolTipText = cboSkillsWeaponry.Text & _
                    ": The demonstrated ability to be on target when any weapon is thrown. " & _
                    "This includes such things as shurikens, darts, potions, and weapons that are thrown accidentally."
            Case 6
                txtSkillsWeaponry.Text = .Sling
                txtSkillsWeaponry.ToolTipText = cboSkillsWeaponry.Text & _
                    ": The ability to use any weapon which consists of a leather strap and two cords which, when whirled and released, hurls stones and other like objects at an opponent."
            Case 7
                txtSkillsWeaponry.Text = .Bow
                txtSkillsWeaponry.ToolTipText = cboSkillsWeaponry.Text & _
                    ": The flair of handling any bow which fires arrows and is used as a weapon in combat."
            Case 8
                txtSkillsWeaponry.Text = .Shield
                txtSkillsWeaponry.ToolTipText = cboSkillsWeaponry.Text & _
                    ": The art of using a shield effectively to block an opponent's blow while fighting or parrying."
            Case 9
                txtSkillsWeaponry.Text = .HandToHand
                txtSkillsWeaponry.ToolTipText = cboSkillsWeaponry.Text & _
                    ": The talent of using one's hands and feet as lethal weapons to strike and hopefully kill an opponent."
            Case Else
        End Select
    End With
    txtSkillsWeaponryDesc.Text = txtSkillsWeaponry.ToolTipText
End Sub
Private Sub ClearFields()
    Dim ctl As Control
    
    If cboCharacter.ListIndex = -1 Then
        cmdEdit.Visible = False
    Else
        cmdEdit.Visible = True
    End If
    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdExit.Visible = True
    
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.Text = vbNullString
            Case "ComboBox"
                If ctl Is cboCharacter Then
                Else
                    ctl.ListIndex = -1
                End If
            Case Else
        End Select
    Next ctl
End Sub

'char *SpellMap[] = {
'     /* Fire */     "EnergyBlast","BlindingFlash","PsionicFire","Fireball",
'                    "FireShield","DazzlingLights","FireBomb","Lightning",
'                    "PrismicMissile","Firestorm","NuclearBlast",
'     /* Water */    "ChillingTouch","Stamina","Terror","Weaken","Slow",
'                    "Haste","CureParalysis","IceShield","Restfull","IceBall",
'                    "Paralyze","Superman","Deepfreeze","DrainingCloud",
'                    "CureDisease",
'     /* Air */      "Poison","MissileShield","ShrillSound","StinkBomb",
'                    "AirPocket","Silence","PoisonGas","CurePoison",
'                    "Whirlwind","PurifyAir","DeadlyPoison","Levitate",
'                    "ToxicVapors","NoxiousFumes","Asphyxiation","DeadlyAir",
'                    "DeathCloud",
'     /* Earth */    "AcidSplash","ItchingSkin","ArmorShield","Direction",
'                    "KnockKnock","Blades","Armorplate","Web","WhippingRocks",
'                    "AcidBomb","Armormelt","Crush","CreateLife","CureStone",
'     /* Mental */   "MentalAttack","Sleep","Bless","Charm","CureLesserCnd",
'                    "DivineTrap","DetectSecret","Identify","Confusion",
'                    "Watchbells","HoldMonsters","Mindread","SaneMind",
'                    "PsionicBlast","Illusion","WizardsEye","Spooks","Death",
'                    "LocateObject","MindFlay","FindPerson",
'     /* Divine */   "HealWounds","MakeWounds","MagicMissile","DispellUndead",
'                    "EnchantedBlade","Blink","MagicScreen","Conjuration",
'                    "AntiMagic","RemoveCurse","Healfull","Lifesteal",
'                    "AstralGate","ZapUndead","Recharge","WordOfDeath",
'                    "Resurrection","DeathWish"};
'
Private Sub cmdCancel_Click()
    Call DisableFields
    Call picTabs_Click(1)
    cboCharacter.SetFocus
End Sub
Private Sub cmdEdit_Click()
    Call EnableFields
    Call picTabs_Click(1)
    cboRace.SetFocus
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    'Move data from screen controls back into Characters array...
    'Write Data back to DataFile...
    Call DisableFields
    Call picTabs_Click(1)
End Sub
Private Sub DisableFields()
    Dim ctl As Control
    
    If cboCharacter.ListIndex = -1 Then
        cmdEdit.Visible = False
    Else
        cmdEdit.Visible = True
    End If
    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdExit.Visible = True
    
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "UpDown"
                ctl.Enabled = False
            Case "TextBox"
                Select Case ctl.Name
                    Case "txtSkillsAcademiaDesc", "txtSkillsPersonalDesc", _
                        "txtSkillsPhysicalDesc", "txtSkillsWeaponryDesc"
                    Case Else
                        ctl.Enabled = False
                        ctl.ForeColor = DisabledColor
                End Select
            Case "Label"
                Select Case ctl.Name
                    Case "lblPortrait", "lblStats", "lblSkills", "lblItems", "lblSwagBag", "lblSpells"
                    Case "lblCharacter"
                        ctl.ForeColor = EnabledColor
                    Case Else
                        ctl.ForeColor = DisabledColor
                End Select
            Case "ComboBox"
                If ctl Is cboCharacter Then
                    ctl.Enabled = True
                    ctl.ForeColor = EnabledColor
                Else
                    ctl.Enabled = False
                    ctl.ForeColor = DisabledColor
                End If
            Case Else
        End Select
    Next ctl
End Sub
Private Sub EnableFields()
    Dim ctl As Control
    
    cmdEdit.Visible = False
    cmdOK.Visible = True
    cmdCancel.Visible = True
    cmdExit.Visible = False
    
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "UpDown"
                ctl.Enabled = True
            Case "TextBox"
                Select Case ctl.Name
                    Case "txtSkillsAcademiaDesc", "txtSkillsPersonalDesc", _
                        "txtSkillsPhysicalDesc", "txtSkillsWeaponryDesc"
                    Case Else
                        ctl.Enabled = True
                        ctl.ForeColor = EnabledColor
                End Select
            Case "Label"
                Select Case ctl.Name
                    Case "lblPortrait", "lblStats", "lblSkills", "lblItems", "lblSwagBag", "lblSpells"
                    Case "lblCharacter"
                        ctl.ForeColor = DisabledColor
                    Case Else
                        ctl.ForeColor = EnabledColor
                End Select
            Case "ComboBox"
                If ctl Is cboCharacter Then
                    ctl.Enabled = False
                    ctl.ForeColor = DisabledColor
                Else
                    ctl.Enabled = True
                    ctl.ForeColor = EnabledColor
                End If
            Case Else
        End Select
    Next ctl
End Sub
Private Sub Form_Activate()
    Dim i As Integer
    
    'Populate Form with data from disk...
    Call ReadWiz07(DataFile, Characters)
    For i = 1 To 6
        cboCharacter.AddItem Characters(i).Name, i - 1
    Next i
    
    cmdEdit.Visible = False
    cboCharacter.ListIndex = -1
    Call DisableFields
    Call ClearFields
    Call picTabs_Click(1)
End Sub
Private Sub Form_Load()
    Dim i As Integer
    
    Me.Picture = frmMain.Picture
    For i = 1 To 5
        picTabs(i).Picture = frmMain.Picture
        picFrames(i).Picture = frmMain.Picture
    Next i
    picWizardryLogo.Picture = frmMain.picWizardryLogo.Picture
    cmdOK.Picture = frmMain.cmdOK.Picture
    cmdCancel.Picture = frmMain.cmdCancel.Picture
    cmdExit.Picture = frmMain.cmdExit.Picture
    
    Call PopulateWiz07Condition(cboCondition)
    Call PopulateWiz07Gender(cboGender)
    Call PopulateWiz07Profession(cboProfession)
    Call PopulateWiz07Race(cboRace)
    Call PopulateWiz07SkillsAcademia(cboSkillsAcademia)
    Call PopulateWiz07SkillsPersonal(cboSkillsPersonal)
    Call PopulateWiz07SkillsPhysical(cboSkillsPhysical)
    Call PopulateWiz07SkillsWeaponry(cboSkillsWeaponry)
End Sub
Private Sub lblItems_Click()
    Call picTabs_Click(3)
End Sub
Private Sub lblSkills_Click()
    Call picTabs_Click(2)
End Sub
Private Sub lblSpells_Click()
    Call picTabs_Click(5)
End Sub
Private Sub lblStats_Click()
    Call picTabs_Click(1)
End Sub
Private Sub lblSwagBag_Click()
    Call picTabs_Click(4)
End Sub
Private Sub picTabs_Click(Index As Integer)
    Dim i As Integer
    
    lblStats.ForeColor = DisabledColor
    lblSkills.ForeColor = DisabledColor
    lblItems.ForeColor = DisabledColor
    lblSwagBag.ForeColor = DisabledColor
    lblSpells.ForeColor = DisabledColor
    
    For i = 1 To 5
        If i <> Index Then
            Call picFrames(i).ZOrder(1) '1 = SendToBack;
        Else
            'Don't want to bring the frame to the very front 'cause it'll preclude
            'the picWizardryLogo image...
            'Call picFrames(i).ZOrder(0) '0 = BringToFront;
        End If
        
        Select Case Index
            Case 1  'Basic Statistics...
                lblStats.ForeColor = EnabledColor
            Case 2  'Skills...
                lblSkills.ForeColor = EnabledColor
            Case 3  'Items...
                lblItems.ForeColor = EnabledColor
            Case 4  'Swag Bag...
                lblSwagBag.ForeColor = EnabledColor
            Case 5  'Spells...
                lblSpells.ForeColor = EnabledColor
            Case Else
        End Select
    Next i
    
    'Move the picWizardryLogo to the appropriate position, per frame...
    Select Case Index
        Case 1  'Basic Statistics...
            picWizardryLogo.Move 4272, 4920
        Case 2  'Skills...
            picWizardryLogo.Move 4260, 5700
        Case 3  'Items...
        Case 4  'Swag Bag...
        Case 5  'Spells...
        Case Else
    End Select
    
End Sub
Private Sub picTabs_GotFocus(Index As Integer)
    SaveMessage = sbStatus.Panels("Message").Text
    Select Case Index
        Case 1  'Basic Statistics...
            sbStatus.Panels("Message").Text = "Basic Character Statistics..."
        Case 2  'Skills...
            sbStatus.Panels("Message").Text = "Character Skills..."
        Case 3  'Items...
            sbStatus.Panels("Message").Text = "Carried Items..."
        Case 4  'Swag Bag...
            sbStatus.Panels("Message").Text = "Stowed Items..."
        Case 5  'Spells...
            sbStatus.Panels("Message").Text = "Spells..."
        Case Else
    End Select
End Sub
Private Sub picTabs_LostFocus(Index As Integer)
    Select Case Index
        Case 1  'Basic Statistics...
        Case 2  'Skills...
        Case 3  'Items...
        Case 4  'Swag Bag...
        Case 5  'Spells...
        Case Else
    End Select
    sbStatus.Panels("Message").Text = SaveMessage
End Sub
Private Sub txtAge_GotFocus()
    TextSelected
End Sub
Private Sub txtAge_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtAge_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtAirSP_GotFocus()
    TextSelected
End Sub
Private Sub txtAirSP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtAirSP_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtCC_GotFocus()
    TextSelected
End Sub
Private Sub txtCC_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtCC_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    txtCC.Text = Format(txtCC.Text, "#,##0.0")
End Sub
Private Sub txtDEX_GotFocus()
    TextSelected
End Sub
Private Sub txtDEX_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtDEX_Validate(Cancel As Boolean)
    Cancel = ValidateByte()
End Sub
Private Sub txtDivineSP_GotFocus()
    TextSelected
End Sub
Private Sub txtDivineSP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtDivineSP_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtEarthSP_GotFocus()
    TextSelected
End Sub
Private Sub txtEarthSP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtEarthSP_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtEXP_GotFocus()
    TextSelected
End Sub
Private Sub txtEXP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtEXP_Validate(Cancel As Boolean)
    Cancel = ValidateI4()
End Sub
Private Sub txtFireSP_GotFocus()
    TextSelected
End Sub
Private Sub txtFireSP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtFireSP_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtGP_GotFocus()
    TextSelected
End Sub
Private Sub txtGP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtGP_Validate(Cancel As Boolean)
    Cancel = ValidateI4()
End Sub
Private Sub txtHP_GotFocus()
    TextSelected
End Sub
Private Sub txtHP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtHP_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtINT_GotFocus()
    TextSelected
End Sub
Private Sub txtINT_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtINT_Validate(Cancel As Boolean)
    Cancel = ValidateByte()
End Sub
Private Sub txtKAR_GotFocus()
    TextSelected
End Sub
Private Sub txtKAR_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtKAR_Validate(Cancel As Boolean)
    Cancel = ValidateByte()
End Sub
Private Sub txtLives_GotFocus()
    TextSelected
End Sub
Private Sub txtLives_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtLives_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtLVL_GotFocus()
    TextSelected
End Sub
Private Sub txtLVL_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtLVL_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtMentalSP_GotFocus()
    TextSelected
End Sub
Private Sub txtMentalSP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMentalSP_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtMKS_GotFocus()
    TextSelected
End Sub
Private Sub txtMKS_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMKS_Validate(Cancel As Boolean)
    Cancel = ValidateI4()
End Sub
Private Sub txtPER_GotFocus()
    TextSelected
End Sub
Private Sub txtPER_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPER_Validate(Cancel As Boolean)
    Cancel = ValidateByte()
End Sub
Private Sub txtPIE_GotFocus()
    TextSelected
End Sub
Private Sub txtPIE_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPIE_Validate(Cancel As Boolean)
    Cancel = ValidateByte()
End Sub
Private Sub txtSPD_GotFocus()
    TextSelected
End Sub
Private Sub txtSPD_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtSPD_Validate(Cancel As Boolean)
    Cancel = ValidateByte()
End Sub
Private Sub txtSTM_GotFocus()
    TextSelected
End Sub
Private Sub txtSTM_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtSTM_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtSTR_GotFocus()
    TextSelected
End Sub
Private Sub txtSTR_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtSTR_Validate(Cancel As Boolean)
    Cancel = ValidateByte()
End Sub
Private Sub txtVIT_GotFocus()
    TextSelected
End Sub
Private Sub txtVIT_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtVIT_Validate(Cancel As Boolean)
    Cancel = ValidateByte()
End Sub
Private Sub txtWaterSP_GotFocus()
    TextSelected
End Sub
Private Sub txtWaterSP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtWaterSP_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub udAge_Change()
    Call ValidateI2(txtAge)
End Sub
Private Sub udAirSP_Change()
    Call ValidateI2(txtAirSP)
End Sub
Private Sub udCC_Change()
    Call ValidateI2(txtCC)
    txtCC.Text = Format(txtCC.Text, "#,##0.0")
End Sub
Private Sub udDEX_Change()
    Call ValidateByte(txtDEX)
End Sub
Private Sub udDivineSP_Change()
    Call ValidateI2(txtDivineSP)
End Sub
Private Sub udEarthSP_Change()
    Call ValidateI2(txtEarthSP)
End Sub
Private Sub udEXP_Change()
    Call ValidateI4(txtEXP)
End Sub
Private Sub udFireSP_Change()
    Call ValidateI2(txtFireSP)
End Sub
Private Sub udGP_Change()
    Call ValidateI4(txtGP)
End Sub
Private Sub udHP_Change()
    Call ValidateI2(txtHP)
End Sub
Private Sub udINT_Change()
    Call ValidateByte(txtINT)
End Sub
Private Sub udKAR_Change()
    Call ValidateByte(txtKAR)
End Sub
Private Sub udLives_Change()
    Call ValidateI2(txtLives)
End Sub
Private Sub udLVL_Change()
    Call ValidateI2(txtLVL)
End Sub
Private Sub udMentalSP_Change()
    Call ValidateI2(txtMentalSP)
End Sub
Private Sub udMKS_Change()
    Call ValidateI4(txtMKS)
End Sub
Private Sub udPER_Change()
    Call ValidateByte(txtPER)
End Sub
Private Sub udPIE_Change()
    Call ValidateByte(txtPIE)
End Sub
Private Sub udSPD_Change()
    Call ValidateByte(txtSPD)
End Sub
Private Sub udSTM_Change()
    Call ValidateI2(txtSTM)
End Sub
Private Sub udSTR_Change()
    Call ValidateByte(txtSTR)
End Sub
Private Sub udVIT_Change()
    Call ValidateByte(txtVIT)
End Sub
Private Sub udWaterSP_Change()
    Call ValidateI2(txtWaterSP)
End Sub

