VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWiz07 
   BackColor       =   &H00000000&
   Caption         =   "Wizardry 07 - Crusaders of the Dark Savant"
   ClientHeight    =   6444
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   11256
   Icon            =   "frmWiz07.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6444
   ScaleWidth      =   11256
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Height          =   432
      Left            =   9720
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   90
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
      TabIndex        =   89
      Top             =   2940
      Width           =   1212
   End
   Begin VB.PictureBox picWiz07 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2064
      Left            =   9600
      Picture         =   "frmWiz07.frx":1CCA
      ScaleHeight     =   2016
      ScaleWidth      =   1344
      TabIndex        =   57
      Top             =   120
      Width           =   1392
   End
   Begin VB.PictureBox picTabBasic 
      BackColor       =   &H00000000&
      Height          =   372
      Left            =   120
      Picture         =   "frmWiz07.frx":2715
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   35
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
         TabIndex        =   36
         Top             =   0
         Width           =   1032
      End
   End
   Begin VB.PictureBox picBasic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5112
      Left            =   120
      ScaleHeight     =   5088
      ScaleWidth      =   9108
      TabIndex        =   3
      Top             =   1020
      Width           =   9132
      Begin VB.PictureBox picWizardryLogo 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Height          =   420
         Left            =   4152
         ScaleHeight     =   372
         ScaleWidth      =   1920
         TabIndex        =   88
         Top             =   3900
         Width           =   1968
      End
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
         TabIndex        =   86
         Text            =   "frmWiz07.frx":1BF37
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
         TabIndex        =   84
         Text            =   "frmWiz07.frx":1BF3A
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
         TabIndex        =   82
         Text            =   "frmWiz07.frx":1BF3D
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
         TabIndex        =   80
         Text            =   "frmWiz07.frx":1BF40
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
         TabIndex        =   78
         Text            =   "frmWiz07.frx":1BF43
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
         TabIndex        =   76
         Text            =   "frmWiz07.frx":1BF46
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
         TabIndex        =   74
         Text            =   "frmWiz07.frx":1BF49
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
         TabIndex        =   72
         Text            =   "frmWiz07.frx":1BF4C
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
         TabIndex        =   70
         Text            =   "frmWiz07.frx":1BF4F
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
         TabIndex        =   68
         Text            =   "frmWiz07.frx":1BF52
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
         TabIndex        =   66
         Text            =   "frmWiz07.frx":1BF55
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
         TabIndex        =   64
         Text            =   "frmWiz07.frx":1BF58
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
         TabIndex        =   62
         Text            =   "frmWiz07.frx":1BF5B
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
         TabIndex        =   60
         Text            =   "frmWiz07.frx":1BF5E
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
         TabIndex        =   58
         Text            =   "frmWiz07.frx":1BF61
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         ItemData        =   "frmWiz07.frx":1BF64
         Left            =   1440
         List            =   "frmWiz07.frx":1BF6B
         Style           =   2  'Dropdown List
         TabIndex        =   54
         ToolTipText     =   "Character's Race (i.e. Human, Elf, etc.)..."
         Top             =   60
         Width           =   1692
      End
      Begin MSComCtl2.UpDown udPIE 
         Height          =   360
         Left            =   2400
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPIE"
         BuddyDispid     =   196617
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
         TabIndex        =   45
         Text            =   "frmWiz07.frx":1BF76
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
         TabIndex        =   44
         Text            =   "frmWiz07.frx":1BF79
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
         TabIndex        =   43
         Text            =   "frmWiz07.frx":1BF7C
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
         TabIndex        =   42
         Text            =   "frmWiz07.frx":1BF7F
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
         TabIndex        =   41
         Text            =   "frmWiz07.frx":1BF82
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
         TabIndex        =   40
         Text            =   "frmWiz07.frx":1BF85
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
         TabIndex        =   39
         Text            =   "frmWiz07.frx":1BF88
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
         ItemData        =   "frmWiz07.frx":1BF8B
         Left            =   1440
         List            =   "frmWiz07.frx":1BF95
         Style           =   2  'Dropdown List
         TabIndex        =   34
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
         TabIndex        =   31
         Text            =   "frmWiz07.frx":1BFA7
         ToolTipText     =   "Strength..."
         Top             =   1680
         Width           =   396
      End
      Begin MSComCtl2.UpDown udSTR 
         Height          =   360
         Left            =   2400
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSTR"
         BuddyDispid     =   196626
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
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2100
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtINT"
         BuddyDispid     =   196623
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
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2940
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtVIT"
         BuddyDispid     =   196622
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
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtDEX"
         BuddyDispid     =   196621
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
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSPD"
         BuddyDispid     =   196620
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   4200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPER"
         BuddyDispid     =   196619
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
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   4620
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtKAR"
         BuddyDispid     =   196618
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
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtHP"
         BuddyDispid     =   196613
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
      Begin MSComCtl2.UpDown upSTM 
         Height          =   360
         Left            =   5280
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2100
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSTM"
         BuddyDispid     =   196672
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
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtCC"
         BuddyDispid     =   196674
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
      Begin MSComCtl2.UpDown upEXP 
         Height          =   360
         Left            =   8820
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtEXP"
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
      Begin MSComCtl2.UpDown upMKS 
         Height          =   360
         Left            =   8820
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMKS"
         BuddyDispid     =   196677
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
      Begin MSComCtl2.UpDown upGP 
         Height          =   360
         Left            =   8820
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtGP"
         BuddyDispid     =   196679
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
      Begin MSComCtl2.UpDown upLVL 
         Height          =   360
         Left            =   5640
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLVL"
         BuddyDispid     =   196681
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
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLives"
         BuddyDispid     =   196683
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
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtAge"
         BuddyDispid     =   196685
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
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtFireSP"
         BuddyDispid     =   196687
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
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   2100
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtWaterSP"
         BuddyDispid     =   196689
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
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtAirSP"
         BuddyDispid     =   196691
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
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   2940
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtEarthSP"
         BuddyDispid     =   196693
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
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   3360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMentalSP"
         BuddyDispid     =   196695
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
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   3780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtDivineSP"
         BuddyDispid     =   196697
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   33
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   84
         Width           =   576
      End
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
      TabIndex        =   2
      Top             =   240
      Width           =   4212
   End
   Begin VB.PictureBox picWiz07Gold 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2076
      Left            =   9420
      Picture         =   "frmWiz07.frx":1BFAA
      ScaleHeight     =   2028
      ScaleWidth      =   1668
      TabIndex        =   0
      Top             =   120
      Width           =   1716
   End
   Begin MSComctlLib.ImageList imgIcons32 
      Left            =   10140
      Top             =   600
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
            Picture         =   "frmWiz07.frx":1CE7D
            Key             =   "Wiz07"
            Object.Tag             =   "Wiz07"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWiz07.frx":1EB59
            Key             =   "Wiz07g"
            Object.Tag             =   "Wiz07g"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons16 
      Left            =   10140
      Top             =   1140
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
            Picture         =   "frmWiz07.frx":1EE75
            Key             =   "Wiz07"
            Object.Tag             =   "Wiz07"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWiz07.frx":20B51
            Key             =   "Wiz07g"
            Object.Tag             =   "Wiz07g"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   32
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
            Object.Width           =   14012
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
            TextSave        =   "2:16 AM"
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
      Enabled         =   0   'False
      Height          =   432
      Left            =   9720
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   91
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
      TabIndex        =   1
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
Public DataFile As String
Private Characters(1 To 6) As Character
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
    cmdExit.Visible = True
    
    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdExit.Enabled = True
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Dim i As Integer
    
    'Populate Form with data from disk...
    Call ReadWiz07(DataFile, Characters)

    For i = 1 To 1  '6
        With Characters(i)
            cboCharacter.AddItem .Name, i - 1
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
            txtAge.Text = Format(.Age, "#,##0")
            
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
'            Debug.Print vbCrLf & "Unknown Region #2 (64 bytes):"
'            Debug.Print strHex(.Unknown2, 64) & vbCrLf
'
'            Debug.Print vbCrLf & "Weaponry Skills..."
'            Debug.Print "Wand:              " & vbTab & .Wand
'            Debug.Print "Sword:             " & vbTab & .Sword
'            Debug.Print "Axe:               " & vbTab & .Axe
'            Debug.Print "Mace:              " & vbTab & .Mace
'            Debug.Print "Pole:              " & vbTab & .Pole
'            Debug.Print "Throwing:          " & vbTab & .Throwing
'            Debug.Print "Sling:             " & vbTab & .Sling
'            Debug.Print "Bow:               " & vbTab & .Bow
'            Debug.Print "Shield:            " & vbTab & .Shield
'            Debug.Print "HandToHand:        " & vbTab & .HandToHand
'
'            Debug.Print vbCrLf & "Physical Skills..."
'            Debug.Print "Swimming:          " & vbTab & .Swimming
'            Debug.Print "Climbing:          " & vbTab & .Climbing
'            Debug.Print "Scouting:          " & vbTab & .Scouting
'            Debug.Print "Music:             " & vbTab & .Music
'            Debug.Print "Oratory:           " & vbTab & .Oratory
'            Debug.Print "Legerdemain:       " & vbTab & .Legerdemain
'            Debug.Print "Skulduggery:       " & vbTab & .Skulduggery
'            Debug.Print "Ninjutsu:          " & vbTab & .Ninjutsu
'
'            Debug.Print vbCrLf & "Personal Skills..."
'            Debug.Print "Firearms:          " & vbTab & .Firearms
'            Debug.Print "Reflextion:        " & vbTab & .Reflextion
'            Debug.Print "SnakeSpeed:        " & vbTab & .SnakeSpeed
'            Debug.Print "EagleEye:          " & vbTab & .EagleEye
'            Debug.Print "PowerStrike:       " & vbTab & .PowerStrike
'            Debug.Print "MindControl:       " & vbTab & .MindControl
'
'            Debug.Print vbCrLf & "Academia Skills..."
'            Debug.Print "Artifacts:         " & vbTab & .Artifacts
'            Debug.Print "Mythology:         " & vbTab & .Mythology
'            Debug.Print "Mapping:           " & vbTab & .Mapping
'            Debug.Print "Scribe:            " & vbTab & .Scribe
'            Debug.Print "Diplomacy:         " & vbTab & .Diplomacy
'            Debug.Print "Alchemy:           " & vbTab & .Alchemy
'            Debug.Print "Theology:          " & vbTab & .Theology
'            Debug.Print "Theosophy:         " & vbTab & .Theosophy
'            Debug.Print "Thaumaturgy:       " & vbTab & .Thaumaturgy
'            Debug.Print "Kirijutsu:         " & vbTab & .Kirijutsu
'
'            Debug.Print vbCrLf & "Unknown Region #3a (28 bytes):"
'            Debug.Print strHex(.Unknown3a, 28) & vbCrLf
'
'            Debug.Print "'Natural' Armor Class:" & vbTab & .NaturalArmorClass
'
'            Debug.Print vbCrLf & "Unknown Region #3b (7 bytes):"
'            Debug.Print strHex(.Unknown3b, 7) & vbCrLf
'
'            Debug.Print vbCrLf & "Aptitute (96 bytes):"
'            Debug.Print strHex(.Aptitude, 96) & vbCrLf   'Aptitude - I don't remember how I determined this...
'
'            Debug.Print vbCrLf & "SpellBooks..."
'            For j = 1 To 12
'                For i = 1 To 8
'                    Debug.Print vbTab & strSpell(((j - 1) * 8) + i, .SpellBooks(j), i - 1)
'                Next i
'            Next j
'
'            Debug.Print vbCrLf & "Unknown Region #4 (12 bytes):"
'            Debug.Print strHex(.Unknown4, 12) & vbCrLf
'
'            Debug.Print "PictureCode:       " & vbTab & .PictureCode
'            Debug.Print "Race:              " & vbTab & strRace(.Race)
'            Debug.Print "Gender:            " & vbTab & strGender(.Gender)
'            Debug.Print "Profession:        " & vbTab & strProfession(.Profession)
'            Debug.Print "?Alive?:           " & vbTab & .Alive
'            Debug.Print "ConditionCode:     " & vbTab & strCondition(.ConditionCode)
'
'            Debug.Print vbCrLf & "Unknown Region #5 (12 bytes):"
'            Debug.Print strHex(.Unknown5, 12) & vbCrLf
'
'            Debug.Print "Unknown Recap:"
'            Debug.Print "Unknown Region #1 (4 bytes): "
'            Debug.Print strHex(.Unknown1, 4) & vbCrLf
'            Debug.Print vbCrLf & "Unknown Region #2 (64 bytes):"
'            Debug.Print strHex(.Unknown2, 64) & vbCrLf
'            Debug.Print vbCrLf & "Unknown Region #3a (28 bytes):"
'            Debug.Print strHex(.Unknown3a, 28) & vbCrLf
'            Debug.Print vbCrLf & "Unknown Region #3b (7 bytes):"
'            Debug.Print strHex(.Unknown3b, 7) & vbCrLf
'            Debug.Print vbCrLf & "Unknown Region #4 (12 bytes):"
'            Debug.Print strHex(.Unknown4, 12) & vbCrLf
'            Debug.Print vbCrLf & "Unknown Region #5 (12 bytes):"
'            Debug.Print strHex(.Unknown5, 12) & vbCrLf
'            Debug.Print vbCrLf & "Aptitute (96 bytes):"
'            Debug.Print strHex(.Aptitude, 96) & vbCrLf   'Aptitude - I don't remember how I determined this...
            
        End With
    Next i
    cboCharacter.ListIndex = 0
End Sub
Private Sub Form_Load()
    Me.Picture = frmMain.Picture
    picBasic.Picture = frmMain.Picture
    picWizardryLogo.Picture = frmMain.picWizardryLogo.Picture
    cmdOK.Picture = frmMain.cmdOK.Picture
    cmdCancel.Picture = frmMain.cmdCancel.Picture
    cmdExit.Picture = frmMain.cmdExit.Picture
    
    Call PopulateCondition(cboCondition)
    Call PopulateGender(cboGender)
    Call PopulateProfession(cboProfession)
    Call PopulateRace(cboRace)
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

