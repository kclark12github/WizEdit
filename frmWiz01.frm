VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWiz01 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wizardry 01 - Proving Grounds of the Mad Overlord"
   ClientHeight    =   6420
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   9288
   Icon            =   "frmWiz01.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9288
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picFrames 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   4452
      Index           =   1
      Left            =   120
      ScaleHeight     =   4428
      ScaleWidth      =   7008
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   1020
      Width           =   7032
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
         Left            =   4320
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "frmWiz01.frx":1CCA
         ToolTipText     =   "Age...? (Still not sure of this one)..."
         Top             =   2940
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
         Left            =   4320
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "frmWiz01.frx":1CCD
         ToolTipText     =   "Level..."
         Top             =   2100
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
         Left            =   4320
         MaxLength       =   13
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "frmWiz01.frx":1CD0
         ToolTipText     =   "Gold Pieces..."
         Top             =   3720
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
         Left            =   4320
         MaxLength       =   13
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmWiz01.frx":1CD3
         ToolTipText     =   "Experience Points..."
         Top             =   1680
         Width           =   1476
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
         TabIndex        =   15
         Text            =   "frmWiz01.frx":1CD6
         ToolTipText     =   "Hit Points (automatically ""topped-off"")..."
         Top             =   2520
         Width           =   996
      End
      Begin VB.ComboBox cboStatus 
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
         TabIndex        =   17
         ToolTipText     =   "Character's Status (i.e. OK, Afraid, Poisoned, etc.)..."
         Top             =   3360
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
         TabIndex        =   6
         Tag             =   "TabStop"
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
         ItemData        =   "frmWiz01.frx":1CD9
         Left            =   1440
         List            =   "frmWiz01.frx":1CE0
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Character's Race (i.e. Human, Elf, etc.)..."
         Top             =   60
         Width           =   1692
      End
      Begin MSComCtl2.UpDown udPIE 
         Height          =   360
         Left            =   2400
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPIE"
         BuddyDispid     =   196618
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
         TabIndex        =   9
         Text            =   "frmWiz01.frx":1CEB
         ToolTipText     =   "Piety..."
         Top             =   2520
         Width           =   396
      End
      Begin VB.TextBox txtLCK 
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
         Text            =   "frmWiz01.frx":1CEE
         ToolTipText     =   "Speed..."
         Top             =   3780
         Width           =   396
      End
      Begin VB.TextBox txtAGL 
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
         Text            =   "frmWiz01.frx":1CF1
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
         TabIndex        =   10
         Text            =   "frmWiz01.frx":1CF4
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
         TabIndex        =   8
         Text            =   "frmWiz01.frx":1CF7
         ToolTipText     =   "Intelligence..."
         Top             =   2100
         Width           =   396
      End
      Begin VB.ComboBox cboGender 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
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
         ItemData        =   "frmWiz01.frx":1CFA
         Left            =   1440
         List            =   "frmWiz01.frx":1D04
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "TabStop"
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
         TabIndex        =   7
         Text            =   "frmWiz01.frx":1D16
         ToolTipText     =   "Strength..."
         Top             =   1680
         Width           =   396
      End
      Begin MSComCtl2.UpDown udSTR 
         Height          =   360
         Left            =   2400
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSTR"
         BuddyDispid     =   196624
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
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   2100
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtINT"
         BuddyDispid     =   196622
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
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   2940
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtVIT"
         BuddyDispid     =   196621
         OrigRight       =   240
         OrigBottom      =   612
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udAGL 
         Height          =   360
         Left            =   2400
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   3360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtAGL"
         BuddyDispid     =   196620
         OrigLeft        =   2400
         OrigTop         =   3360
         OrigRight       =   2640
         OrigBottom      =   3720
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udLCK 
         Height          =   360
         Left            =   2400
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   3780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLCK"
         BuddyDispid     =   196619
         OrigLeft        =   2400
         OrigTop         =   3780
         OrigRight       =   2640
         OrigBottom      =   4140
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udHP 
         Height          =   360
         Left            =   5280
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtHP"
         BuddyDispid     =   196614
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udEXP 
         Height          =   360
         Left            =   5760
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtEXP"
         BuddyDispid     =   196613
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udGP 
         Height          =   360
         Left            =   5760
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   3720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtGP"
         BuddyDispid     =   196612
         OrigLeft        =   2280
         OrigTop         =   360
         OrigRight       =   2520
         OrigBottom      =   972
         Max             =   65535
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udLVL 
         Height          =   360
         Left            =   5280
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   2100
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLVL"
         BuddyDispid     =   196611
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
         Left            =   5280
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   2940
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtAge"
         BuddyDispid     =   196610
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   88
         Top             =   444
         Width           =   780
      End
      Begin VB.Label lblLCK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luck:"
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
         Left            =   1320
         TabIndex        =   86
         Top             =   3816
         Width           =   540
      End
      Begin VB.Label lblAGL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agility:"
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
         Left            =   1080
         TabIndex        =   85
         Top             =   3396
         Width           =   780
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
         Top             =   1716
         Width           =   924
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
         TabIndex        =   80
         Top             =   2556
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
         Left            =   3684
         TabIndex        =   79
         Top             =   3756
         Width           =   528
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
         Left            =   3012
         TabIndex        =   78
         Top             =   1716
         Width           =   1200
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
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
         Left            =   3660
         TabIndex        =   77
         Top             =   2976
         Width           =   456
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
         Left            =   3600
         TabIndex        =   76
         Top             =   2136
         Width           =   624
      End
      Begin VB.Label lblCondition 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   3492
         TabIndex        =   75
         Top             =   3360
         Width           =   720
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
         TabIndex        =   74
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
         TabIndex        =   73
         Top             =   84
         Width           =   576
      End
   End
   Begin VB.PictureBox picFrames 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   4452
      Index           =   2
      Left            =   120
      ScaleHeight     =   4428
      ScaleWidth      =   7008
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   1020
      Width           =   7032
      Begin VB.CheckBox chkIdentified 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   8
         Left            =   2100
         TabIndex        =   49
         ToolTipText     =   "Identified?"
         Top             =   3792
         Width           =   204
      End
      Begin VB.CheckBox chkIdentified 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   7
         Left            =   2100
         TabIndex        =   45
         ToolTipText     =   "Identified?"
         Top             =   3372
         Width           =   204
      End
      Begin VB.CheckBox chkIdentified 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   6
         Left            =   2100
         TabIndex        =   41
         ToolTipText     =   "Identified?"
         Top             =   2952
         Width           =   204
      End
      Begin VB.CheckBox chkIdentified 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   5
         Left            =   2100
         TabIndex        =   37
         ToolTipText     =   "Identified?"
         Top             =   2532
         Width           =   204
      End
      Begin VB.CheckBox chkIdentified 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   4
         Left            =   2100
         TabIndex        =   33
         ToolTipText     =   "Identified?"
         Top             =   2112
         Width           =   204
      End
      Begin VB.CheckBox chkIdentified 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   3
         Left            =   2100
         TabIndex        =   29
         ToolTipText     =   "Identified?"
         Top             =   1692
         Width           =   204
      End
      Begin VB.CheckBox chkIdentified 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   2
         Left            =   2100
         TabIndex        =   25
         ToolTipText     =   "Identified?"
         Top             =   1272
         Width           =   204
      End
      Begin VB.CheckBox chkCursed 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   8
         Left            =   1800
         TabIndex        =   48
         ToolTipText     =   "Cursed?"
         Top             =   3792
         Width           =   204
      End
      Begin VB.CheckBox chkCursed 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   7
         Left            =   1800
         TabIndex        =   44
         ToolTipText     =   "Cursed?"
         Top             =   3372
         Width           =   204
      End
      Begin VB.CheckBox chkCursed 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   6
         Left            =   1800
         TabIndex        =   40
         ToolTipText     =   "Cursed?"
         Top             =   2952
         Width           =   204
      End
      Begin VB.CheckBox chkCursed 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   5
         Left            =   1800
         TabIndex        =   36
         ToolTipText     =   "Cursed?"
         Top             =   2532
         Width           =   204
      End
      Begin VB.CheckBox chkCursed 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   4
         Left            =   1800
         TabIndex        =   32
         ToolTipText     =   "Cursed?"
         Top             =   2112
         Width           =   204
      End
      Begin VB.CheckBox chkCursed 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   3
         Left            =   1800
         TabIndex        =   28
         ToolTipText     =   "Cursed?"
         Top             =   1692
         Width           =   204
      End
      Begin VB.CheckBox chkCursed 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   2
         Left            =   1800
         TabIndex        =   24
         ToolTipText     =   "Cursed?"
         Top             =   1272
         Width           =   204
      End
      Begin VB.CheckBox chkEquipped 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   8
         Left            =   1500
         TabIndex        =   47
         ToolTipText     =   "Equipped?"
         Top             =   3792
         Width           =   204
      End
      Begin VB.CheckBox chkEquipped 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   7
         Left            =   1500
         TabIndex        =   43
         ToolTipText     =   "Equipped?"
         Top             =   3372
         Width           =   204
      End
      Begin VB.CheckBox chkEquipped 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   6
         Left            =   1500
         TabIndex        =   39
         ToolTipText     =   "Equipped?"
         Top             =   2952
         Width           =   204
      End
      Begin VB.CheckBox chkEquipped 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   5
         Left            =   1500
         TabIndex        =   35
         ToolTipText     =   "Equipped?"
         Top             =   2532
         Width           =   204
      End
      Begin VB.CheckBox chkEquipped 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   4
         Left            =   1500
         TabIndex        =   31
         ToolTipText     =   "Equipped?"
         Top             =   2112
         Width           =   204
      End
      Begin VB.CheckBox chkEquipped 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   3
         Left            =   1500
         TabIndex        =   27
         ToolTipText     =   "Equipped?"
         Top             =   1692
         Width           =   204
      End
      Begin VB.CheckBox chkEquipped 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   2
         Left            =   1500
         TabIndex        =   23
         ToolTipText     =   "Equipped?"
         Top             =   1272
         Width           =   204
      End
      Begin VB.CheckBox chkEquipped 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   1
         Left            =   1500
         TabIndex        =   19
         ToolTipText     =   "Equipped?"
         Top             =   840
         Width           =   204
      End
      Begin VB.CheckBox chkCursed 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   1
         Left            =   1800
         TabIndex        =   20
         ToolTipText     =   "Cursed?"
         Top             =   840
         Width           =   204
      End
      Begin VB.CheckBox chkIdentified 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   1
         Left            =   2100
         TabIndex        =   21
         ToolTipText     =   "Identified?"
         Top             =   840
         Width           =   204
      End
      Begin VB.ComboBox cboItem 
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
         Index           =   1
         ItemData        =   "frmWiz01.frx":1D19
         Left            =   2520
         List            =   "frmWiz01.frx":1D1B
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #1..."
         Top             =   828
         Width           =   2712
      End
      Begin VB.ComboBox cboItem 
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
         Index           =   8
         ItemData        =   "frmWiz01.frx":1D1D
         Left            =   2520
         List            =   "frmWiz01.frx":1D1F
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #8..."
         Top             =   3780
         Width           =   2712
      End
      Begin VB.ComboBox cboItem 
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
         Index           =   7
         ItemData        =   "frmWiz01.frx":1D21
         Left            =   2520
         List            =   "frmWiz01.frx":1D23
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #7..."
         Top             =   3360
         Width           =   2712
      End
      Begin VB.ComboBox cboItem 
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
         Index           =   6
         ItemData        =   "frmWiz01.frx":1D25
         Left            =   2520
         List            =   "frmWiz01.frx":1D27
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #6..."
         Top             =   2940
         Width           =   2712
      End
      Begin VB.ComboBox cboItem 
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
         Index           =   5
         ItemData        =   "frmWiz01.frx":1D29
         Left            =   2520
         List            =   "frmWiz01.frx":1D2B
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #5..."
         Top             =   2520
         Width           =   2712
      End
      Begin VB.ComboBox cboItem 
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
         Index           =   4
         ItemData        =   "frmWiz01.frx":1D2D
         Left            =   2520
         List            =   "frmWiz01.frx":1D2F
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #4..."
         Top             =   2100
         Width           =   2712
      End
      Begin VB.ComboBox cboItem 
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
         Index           =   3
         ItemData        =   "frmWiz01.frx":1D31
         Left            =   2520
         List            =   "frmWiz01.frx":1D33
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #3..."
         Top             =   1680
         Width           =   2712
      End
      Begin VB.ComboBox cboItem 
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
         Index           =   2
         ItemData        =   "frmWiz01.frx":1D35
         Left            =   2520
         List            =   "frmWiz01.frx":1D37
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #2..."
         Top             =   1260
         Width           =   2712
      End
      Begin VB.Label lbl8i 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8"
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
         Left            =   1200
         TabIndex        =   149
         Top             =   3804
         Width           =   144
      End
      Begin VB.Label lbl7i 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7"
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
         Left            =   1200
         TabIndex        =   148
         Top             =   3384
         Width           =   144
      End
      Begin VB.Label lbl6i 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
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
         Left            =   1200
         TabIndex        =   147
         Top             =   2964
         Width           =   144
      End
      Begin VB.Label lbl5i 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
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
         Left            =   1200
         TabIndex        =   146
         Top             =   2544
         Width           =   144
      End
      Begin VB.Label lbl4i 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
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
         Left            =   1200
         TabIndex        =   145
         Top             =   2124
         Width           =   144
      End
      Begin VB.Label lbl3i 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Left            =   1200
         TabIndex        =   144
         Top             =   1704
         Width           =   144
      End
      Begin VB.Label lbl2i 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
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
         Left            =   1200
         TabIndex        =   143
         Top             =   1284
         Width           =   144
      End
      Begin VB.Label lblItemNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Heidelberg"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   384
         Left            =   1200
         TabIndex        =   142
         Top             =   420
         Width           =   156
      End
      Begin VB.Label lbl1i 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Left            =   1200
         TabIndex        =   141
         Top             =   852
         Width           =   144
      End
      Begin VB.Label lblIdentified 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "?"
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
         Left            =   2100
         TabIndex        =   120
         ToolTipText     =   "Identified?"
         Top             =   480
         Width           =   108
      End
      Begin VB.Label lblCursed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Heidelberg"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   432
         Left            =   1800
         TabIndex        =   119
         ToolTipText     =   "Cursed?"
         Top             =   360
         Width           =   168
      End
      Begin VB.Label lblEquipped 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Heidelberg"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   432
         Left            =   1500
         TabIndex        =   118
         ToolTipText     =   "Equipped?"
         Top             =   480
         Width           =   144
      End
      Begin VB.Label lblItemList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item List:"
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
         Left            =   360
         TabIndex        =   117
         Top             =   120
         Width           =   1008
      End
   End
   Begin VB.PictureBox picFrames 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   4452
      Index           =   3
      Left            =   120
      ScaleHeight     =   4428
      ScaleWidth      =   7008
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   1020
      Width           =   7032
      Begin VB.TextBox txtPriest7 
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
         Left            =   6000
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   66
         Text            =   "frmWiz01.frx":1D39
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3900
         Width           =   576
      End
      Begin VB.TextBox txtPriest6 
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
         Left            =   5160
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   65
         Text            =   "frmWiz01.frx":1D3F
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3900
         Width           =   576
      End
      Begin VB.TextBox txtPriest5 
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
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   64
         Text            =   "frmWiz01.frx":1D45
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3900
         Width           =   576
      End
      Begin VB.TextBox txtMage7 
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
         Left            =   6000
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   59
         Text            =   "frmWiz01.frx":1D4B
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3420
         Width           =   576
      End
      Begin VB.TextBox txtMage6 
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
         Left            =   5160
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   58
         Text            =   "frmWiz01.frx":1D51
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3420
         Width           =   576
      End
      Begin VB.TextBox txtMage5 
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
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   57
         Text            =   "frmWiz01.frx":1D57
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3420
         Width           =   576
      End
      Begin VB.TextBox txtPriest4 
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
         Left            =   3480
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   63
         Text            =   "frmWiz01.frx":1D5D
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3900
         Width           =   576
      End
      Begin VB.TextBox txtMage4 
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
         Left            =   3480
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   56
         Text            =   "frmWiz01.frx":1D63
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3420
         Width           =   576
      End
      Begin VB.TextBox txtPriest3 
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
         Left            =   2640
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   62
         Text            =   "frmWiz01.frx":1D69
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3900
         Width           =   576
      End
      Begin VB.TextBox txtMage3 
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
         Left            =   2640
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   55
         Text            =   "frmWiz01.frx":1D6F
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3420
         Width           =   576
      End
      Begin VB.ListBox lstPriestSpells 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Heidelberg"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1776
         Left            =   4140
         Style           =   1  'Checkbox
         TabIndex        =   52
         Top             =   840
         Width           =   2472
      End
      Begin VB.ListBox lstMageSpells 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Heidelberg"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1776
         Left            =   480
         Style           =   1  'Checkbox
         TabIndex        =   51
         Top             =   840
         Width           =   2472
      End
      Begin VB.TextBox txtPriest2 
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
         Left            =   1800
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   61
         Text            =   "frmWiz01.frx":1D75
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3900
         Width           =   576
      End
      Begin VB.TextBox txtMage2 
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
         Left            =   1800
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   54
         Text            =   "frmWiz01.frx":1D7B
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3420
         Width           =   576
      End
      Begin VB.TextBox txtPriest1 
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
         Left            =   960
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   60
         Text            =   "frmWiz01.frx":1D81
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3900
         Width           =   576
      End
      Begin VB.TextBox txtMage1 
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
         Left            =   960
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   53
         Text            =   "frmWiz01.frx":1D87
         ToolTipText     =   "Wand && Dagger..."
         Top             =   3444
         Width           =   576
      End
      Begin MSComCtl2.UpDown udMage1 
         Height          =   360
         Left            =   1500
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   3444
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage1"
         BuddyDispid     =   196674
         OrigLeft        =   2400
         OrigTop         =   3444
         OrigRight       =   2640
         OrigBottom      =   3804
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest1 
         Height          =   360
         Left            =   1500
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest1"
         BuddyDispid     =   196673
         OrigLeft        =   2400
         OrigTop         =   3900
         OrigRight       =   2640
         OrigBottom      =   4260
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage2 
         Height          =   360
         Left            =   2340
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage2"
         BuddyDispid     =   196672
         OrigLeft        =   3300
         OrigTop         =   3420
         OrigRight       =   3540
         OrigBottom      =   3780
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest2 
         Height          =   360
         Left            =   2340
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest2"
         BuddyDispid     =   196671
         OrigLeft        =   3240
         OrigTop         =   3900
         OrigRight       =   3480
         OrigBottom      =   4260
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage3 
         Height          =   360
         Left            =   3180
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage3"
         BuddyDispid     =   196668
         OrigLeft        =   4140
         OrigTop         =   3420
         OrigRight       =   4380
         OrigBottom      =   3780
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest3 
         Height          =   360
         Left            =   3180
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest3"
         BuddyDispid     =   196667
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage4 
         Height          =   360
         Left            =   4020
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage4"
         BuddyDispid     =   196666
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest4 
         Height          =   360
         Left            =   4020
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest4"
         BuddyDispid     =   196665
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage5 
         Height          =   360
         Left            =   4860
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage5"
         BuddyDispid     =   196664
         OrigLeft        =   5760
         OrigTop         =   3420
         OrigRight       =   6000
         OrigBottom      =   3780
         Max             =   999
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage6 
         Height          =   360
         Left            =   5700
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage6"
         BuddyDispid     =   196663
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage7 
         Height          =   360
         Left            =   6540
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage7"
         BuddyDispid     =   196662
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest5 
         Height          =   360
         Left            =   4860
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest5"
         BuddyDispid     =   196661
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest6 
         Height          =   360
         Left            =   5700
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest6"
         BuddyDispid     =   196660
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest7 
         Height          =   360
         Left            =   6540
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest7"
         BuddyDispid     =   196659
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7"
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
         Left            =   6300
         TabIndex        =   130
         Top             =   3060
         Width           =   144
      End
      Begin VB.Label lbl6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
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
         Left            =   5460
         TabIndex        =   129
         Top             =   3060
         Width           =   144
      End
      Begin VB.Label lbl5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
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
         Left            =   4620
         TabIndex        =   128
         Top             =   3060
         Width           =   144
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
         Left            =   60
         TabIndex        =   127
         Top             =   2760
         Width           =   1332
      End
      Begin VB.Label lblLevel1 
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
         Left            =   192
         TabIndex        =   126
         Top             =   3120
         Width           =   624
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Left            =   1260
         TabIndex        =   125
         Top             =   3060
         Width           =   144
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Left            =   2940
         TabIndex        =   124
         Top             =   3060
         Width           =   144
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
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
         Left            =   2100
         TabIndex        =   123
         Top             =   3060
         Width           =   144
      End
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
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
         Left            =   3780
         TabIndex        =   122
         Top             =   3060
         Width           =   144
      End
      Begin VB.Label lblPriestPoints 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Priest:"
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
         Left            =   108
         TabIndex        =   121
         Top             =   3960
         Width           =   708
      End
      Begin VB.Label lblSpellBooks 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spell Books:"
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
         TabIndex        =   112
         Top             =   120
         Width           =   1272
      End
      Begin VB.Label lblPriestSpells 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Priest Spells:"
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
         Left            =   4140
         TabIndex        =   111
         Top             =   420
         Width           =   1392
      End
      Begin VB.Label lblMagePoints 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mage:"
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
         TabIndex        =   110
         Top             =   3480
         Width           =   636
      End
      Begin VB.Label lblMageSpells 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mage Spells:"
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
         TabIndex        =   109
         Top             =   420
         Width           =   1320
      End
   End
   Begin VB.PictureBox picWizardryLogo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   420
      Left            =   3660
      ScaleHeight     =   372
      ScaleWidth      =   1920
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1968
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
      Left            =   7680
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmWiz01.frx":1D8D
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2460
      UseMaskColor    =   -1  'True
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Height          =   432
      Left            =   7680
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   2460
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   432
      Left            =   7680
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   2940
      Width           =   1212
   End
   Begin VB.PictureBox picWiz01 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2052
      Left            =   7560
      Picture         =   "frmWiz01.frx":32DC
      ScaleHeight     =   2004
      ScaleWidth      =   1332
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   120
      Width           =   1380
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   1
      Left            =   120
      Picture         =   "frmWiz01.frx":3AE6
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   89
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
         TabIndex        =   1
         Top             =   0
         Width           =   1032
      End
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   2
      Left            =   1260
      Picture         =   "frmWiz01.frx":1D308
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   105
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
         TabIndex        =   2
         Top             =   0
         Width           =   1032
      End
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   3
      Left            =   2400
      Picture         =   "frmWiz01.frx":36B2A
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   106
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
         TabIndex        =   3
         Top             =   0
         Width           =   1032
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
      TabIndex        =   0
      Top             =   240
      Width           =   4212
   End
   Begin MSComctlLib.ImageList imgIcons32 
      Left            =   7980
      Top             =   660
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWiz01.frx":5034C
            Key             =   "Wiz01"
            Object.Tag             =   "Wiz01"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons16 
      Left            =   7980
      Top             =   1200
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWiz01.frx":52028
            Key             =   "Wiz01"
            Object.Tag             =   "Wiz01"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   87
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
            TextSave        =   "10:36 PM"
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
      Left            =   7680
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   69
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
      TabIndex        =   71
      Top             =   264
      Width           =   1020
   End
End
Attribute VB_Name = "frmWiz01"
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
Private ActiveTab As Integer
Private Characters(1 To Wiz01CharactersMax) As Wiz01Character
Private Sub cboCharacter_Click()
    Dim i As Integer
    Dim j As Integer
    Dim SpellNumber As Integer
    Dim bString As String
    Dim ctl As Control
    
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
        cboStatus.ListIndex = .Status
        cboGender.ListIndex = 0
        cboProfession.ListIndex = .Profession
        cboRace.ListIndex = .Race
        
        txtSTR.Text = icvtStatistic(.Statistics, 1)
        txtINT.Text = icvtStatistic(.Statistics, 2)
        txtPIE.Text = icvtStatistic(.Statistics, 3)
        txtVIT.Text = icvtStatistic(.Statistics, 4)
        txtAGL.Text = icvtStatistic(.Statistics, 5)
        txtLCK.Text = icvtStatistic(.Statistics, 6)
        
        txtLVL.Text = Format(.LVL.Maximum, "#,##0")
        txtAge.Text = Format(.AgeInWeeks \ 52, "#,##0")
        
        txtHP.Text = Format(.HP.Maximum, "#,##0")
        txtEXP.Text = Format(.EXP, "#,##0")
        txtGP.Text = Format(.GP, "#,##0")
    
        'SpellBooks...
        bString = icvtSpellsToBin(.SpellBooks)
        For i = 1 To Wiz01SpellMapMax
'            If Mid(bString, i + 1, 1) = "1" Then
'                Debug.Print "[X] " & GetSpell(i)
'            Else
'                Debug.Print "[ ] " & GetSpell(i)
'            End If
            If i <= 21 Then
                If Mid(bString, i + 1, 1) = "1" Then
                    lstMageSpells.Selected(i - 1) = True
                Else
                    lstMageSpells.Selected(i - 1) = False
                End If
            Else
                If Mid(bString, i + 1, 1) = "1" Then
                    lstPriestSpells.Selected(i - 1 - 21) = True
                Else
                    lstPriestSpells.Selected(i - 1 - 21) = False
                End If
            End If
        Next i
        
        txtMage1.Text = .MageSpellPoints(1)
        txtPriest1.Text = .PriestSpellPoints(1)
        txtMage2.Text = .MageSpellPoints(2)
        txtPriest2.Text = .PriestSpellPoints(2)
        txtMage3.Text = .MageSpellPoints(3)
        txtPriest3.Text = .PriestSpellPoints(3)
        txtMage4.Text = .MageSpellPoints(4)
        txtPriest4.Text = .PriestSpellPoints(4)
        txtMage5.Text = .MageSpellPoints(5)
        txtPriest5.Text = .PriestSpellPoints(5)
        txtMage6.Text = .MageSpellPoints(6)
        txtPriest6.Text = .PriestSpellPoints(6)
        txtMage7.Text = .MageSpellPoints(7)
        txtPriest7.Text = .PriestSpellPoints(7)
        
        For i = 1 To Wiz01ItemListMax
            cboItem(i).ListIndex = .ItemList(i).ItemCode
            If .ItemList(i).Identified = 1 Then chkIdentified(i).Value = vbChecked Else chkIdentified(i).Value = vbUnchecked
            If .ItemList(i).Equipped = 1 Then chkEquipped(i).Value = vbChecked Else chkEquipped(i).Value = vbUnchecked
            If .ItemList(i).Cursed = -1 Then chkCursed(i).Value = vbChecked Else chkCursed(i).Value = vbUnchecked
        Next i
        
        sbStatus.Panels("Status").Text = .Name
        sbStatus.Panels("Message").Text = "Level " & .LVL.Maximum & " " & cboProfession.List(cboProfession.ListIndex) & "..."
        lstMageSpells.ListIndex = -1
        lstPriestSpells.ListIndex = -1
    End With
End Sub
Private Sub ClearFields()
    Dim ctl As Control
    Dim i As Integer
    
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
            Case "ListBox"
                For i = 0 To ctl.ListCount - 1
                    ctl.Selected(i) = False
                Next i
            Case "CheckBox"
                ctl.Value = vbUnchecked
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
Private Sub cboGender_GotFocus()
    TextSelected
End Sub
Private Sub cboItem_Change(Index As Integer)
    TextSelected
End Sub
Private Sub cboProfession_GotFocus()
    TextSelected
End Sub
Private Sub cboRace_GotFocus()
    TextSelected
End Sub
Private Sub cboStatus_GotFocus()
    TextSelected
End Sub
Private Sub chkCursed_GotFocus(Index As Integer)
    TextSelected
End Sub
Private Sub chkEquipped_GotFocus(Index As Integer)
    TextSelected
End Sub
Private Sub chkIdentified_GotFocus(Index As Integer)
    TextSelected
End Sub
Private Sub cmdCancel_Click()
    Call DisableFields
    Call picTabs_Click(1)
    cboCharacter.SetFocus
End Sub
Private Sub cmdCancel_GotFocus()
    TextSelected
End Sub
Private Sub cmdEdit_Click()
    Call EnableFields
    Select Case ActiveTab
        Case 1
            cboRace.SetFocus
        Case 2
            chkEquipped(1).SetFocus
        Case 3
            lstMageSpells.SetFocus
    End Select
End Sub
Private Sub cmdEdit_GotFocus()
    TextSelected
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdExit_GotFocus()
    TextSelected
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
            Case "ListBox"
                ctl.Enabled = False
            Case "CheckBox"
                ctl.Enabled = False
            Case "TextBox"
                Select Case ctl.Name
                    Case ""
                    Case Else
                        ctl.Enabled = False
                        ctl.ForeColor = DisabledColor
                End Select
            Case "Label"
                Select Case ctl.Name
                    Case "lblStats", "lblItems", "lblSpells"
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
            Case "ListBox"
                ctl.Enabled = True
            Case "CheckBox"
                ctl.Enabled = True
            Case "TextBox"
                Select Case ctl.Name
                    Case ""
                    Case Else
                        ctl.Enabled = True
                        ctl.ForeColor = EnabledColor
                End Select
            Case "Label"
                Select Case ctl.Name
                    Case "lblStats", "lblItems", "lblSpells"
                    Case "lblCharacter"
                        ctl.ForeColor = DisabledColor
                    Case Else
                        ctl.ForeColor = EnabledColor
                End Select
            Case "ComboBox"
                Select Case ctl.Name
                    Case "cboGender"
                    Case Else
                        If ctl Is cboCharacter Then
                            ctl.Enabled = False
                            ctl.ForeColor = DisabledColor
                        Else
                            ctl.Enabled = True
                            ctl.ForeColor = EnabledColor
                        End If
                    End Select
            Case Else
        End Select
    Next ctl
End Sub
Private Sub cmdOK_GotFocus()
    TextSelected
End Sub
Private Sub Form_Activate()
    Dim i As Integer
    
    'Populate Form with data from disk...
    Call ReadWiz01(DataFile, Characters)
    For i = 1 To Wiz01CharactersMax
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
    For i = 1 To 3
        picTabs(i).Picture = frmMain.Picture
        picFrames(i).Picture = frmMain.Picture
    Next i
    'picWiz01.Picture = frmMain.picWiz01.Picture
    picWizardryLogo.Picture = frmMain.picWizardryLogo.Picture
    cmdOK.Picture = frmMain.cmdOK.Picture
    cmdCancel.Picture = frmMain.cmdCancel.Picture
    cmdExit.Picture = frmMain.cmdExit.Picture
    
    Call InitializeWiz01ItemList
    Call InitializeWiz01Spells
    
    Call PopulateWiz01Status(cboStatus)
    Call PopulateWiz01Profession(cboProfession)
    Call PopulateWiz01Race(cboRace)
    For i = 1 To Wiz01ItemListMax
        Call PopulateWiz01Item(cboItem(i))
    Next i
    Call PopulateWiz01SpellBooks(lstMageSpells, lstPriestSpells)
End Sub
Private Sub lblItems_Click()
    Call picTabs_Click(2)
End Sub
Private Sub lblSpells_Click()
    Call picTabs_Click(3)
End Sub
Private Sub lblStats_Click()
    Call picTabs_Click(1)
End Sub
Private Sub lstMageSpells_GotFocus()
    TextSelected
    lstMageSpells.ListIndex = 0
End Sub
Private Sub lstMageSpells_LostFocus()
    lstMageSpells.ListIndex = -1
End Sub
Private Sub lstPriestSpells_GotFocus()
    TextSelected
    lstPriestSpells.ListIndex = 0
End Sub
Private Sub lstPriestSpells_LostFocus()
    lstPriestSpells.ListIndex = -1
End Sub
Private Sub picFrames_GotFocus(Index As Integer)
    TextSelected
End Sub
Private Sub picTabs_Click(Index As Integer)
    Dim i As Integer
    Dim ctl As Control
    
    ActiveTab = Index
    lblStats.ForeColor = DisabledColor
    lblItems.ForeColor = DisabledColor
    lblSpells.ForeColor = DisabledColor
    
    For i = 1 To 3
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
            Case 2  'Items...
                lblItems.ForeColor = EnabledColor
            Case 3  'Spells...
                lblSpells.ForeColor = EnabledColor
            Case Else
        End Select
    Next i
    
    'Debug.Print "Testing..."
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "Label", "ImageList", "CommandButton", "UpDown"
            Case Else
                If Not ctl.Container Is Me And Not IsNull(ctl.Container) Then
                    If ctl.Container Is picFrames(Index) Then
                        ctl.TabStop = True
                        'Debug.Print ctl.Name & ".TabStop = True"
                    Else
                        ctl.TabStop = False
                        'Debug.Print ctl.Name & ".TabStop = False"
                    End If
                End If
        End Select
    Next ctl
    
    'Move the picWizardryLogo to the appropriate position, per frame...
    Select Case Index
        Case 1  'Basic Statistics...
            picWizardryLogo.Move (picFrames(1).Width + picFrames(1).Left) / 2, 5520
        Case 2  'Items...
            picWizardryLogo.Move (picFrames(1).Width + picFrames(1).Left) / 2, 5520
        Case 3  'Spells...
            picWizardryLogo.Move (picFrames(1).Width + picFrames(1).Left) / 2, 5520
        Case Else
    End Select
    
    On Error Resume Next
    Select Case Index
        Case 1
            cboRace.SetFocus
        Case 2
            chkEquipped(1).SetFocus
        Case 3
            lstMageSpells.SetFocus
    End Select
End Sub
Private Sub picTabs_GotFocus(Index As Integer)
    TextSelected
    SaveMessage = sbStatus.Panels("Message").Text
    Select Case Index
        Case 1  'Basic Statistics...
            sbStatus.Panels("Message").Text = "Basic Character Statistics..."
        Case 2  'Items...
            sbStatus.Panels("Message").Text = "Carried Items..."
        Case 3  'Spells...
            sbStatus.Panels("Message").Text = "Spells..."
        Case Else
    End Select
End Sub
Private Sub picTabs_LostFocus(Index As Integer)
    Select Case Index
        Case 1  'Basic Statistics...
        Case 2  'Items...
        Case 3  'Spells...
        Case Else
    End Select
    sbStatus.Panels("Message").Text = SaveMessage
End Sub
Private Sub picWiz01_GotFocus()
    TextSelected
End Sub
Private Sub picWizardryLogo_GotFocus()
    TextSelected
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
Private Sub txtAGL_GotFocus()
    TextSelected
End Sub
Private Sub txtAGL_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtAGL_Validate(Cancel As Boolean)
    Cancel = ValidateByte()
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
Private Sub txtLCK_GotFocus()
    TextSelected
End Sub
Private Sub txtLCK_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtLCK_Validate(Cancel As Boolean)
    Cancel = ValidateByte()
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
Private Sub txtMage1_GotFocus()
    TextSelected
End Sub
Private Sub txtMage1_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage1_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtMage2_GotFocus()
    TextSelected
End Sub
Private Sub txtMage2_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage2_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtMage3_GotFocus()
    TextSelected
End Sub
Private Sub txtMage3_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage3_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtMage4_GotFocus()
    TextSelected
End Sub
Private Sub txtMage4_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage4_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtMage5_GotFocus()
    TextSelected
End Sub
Private Sub txtMage5_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage5_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtMage6_GotFocus()
    TextSelected
End Sub
Private Sub txtMage6_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage6_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtMage7_GotFocus()
    TextSelected
End Sub
Private Sub txtMage7_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage7_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
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
Private Sub txtPriest1_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest1_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest1_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtPriest2_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest2_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest2_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtPriest3_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest3_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest3_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtPriest4_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest4_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest4_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtPriest5_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest5_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest5_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtPriest6_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest6_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest6_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
End Sub
Private Sub txtPriest7_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest7_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest7_Validate(Cancel As Boolean)
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
Private Sub udAge_Change()
    Call ValidateI2(txtAge)
End Sub
Private Sub udAGL_Change()
    Call ValidateByte(txtAGL)
End Sub
Private Sub udMage1_Change()
    Call ValidateI2(txtMage1)
End Sub
Private Sub udMage2_Change()
    Call ValidateI2(txtMage2)
End Sub
Private Sub udMage3_Change()
    Call ValidateI2(txtMage3)
End Sub
Private Sub udMage4_Change()
    Call ValidateI2(txtMage4)
End Sub
Private Sub udMage5_Change()
    Call ValidateI2(txtMage5)
End Sub
Private Sub udMage6_Change()
    Call ValidateI2(txtMage6)
End Sub
Private Sub udMage7_Change()
    Call ValidateI2(txtMage7)
End Sub
Private Sub udEXP_Change()
    Call ValidateI4(txtEXP)
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
Private Sub udLCK_Change()
    Call ValidateByte(txtLCK)
End Sub
Private Sub udLVL_Change()
    Call ValidateI2(txtLVL)
End Sub
Private Sub udPIE_Change()
    Call ValidateByte(txtPIE)
End Sub
Private Sub udPriest1_Change()
    Call ValidateI2(txtPriest1)
End Sub
Private Sub udPriest2_Change()
    Call ValidateI2(txtPriest2)
End Sub
Private Sub udPriest3_Change()
    Call ValidateI2(txtPriest3)
End Sub
Private Sub udPriest4_Change()
    Call ValidateI2(txtPriest4)
End Sub
Private Sub udPriest5_Change()
    Call ValidateI2(txtPriest5)
End Sub
Private Sub udPriest6_Change()
    Call ValidateI2(txtPriest6)
End Sub
Private Sub udPriest7_Change()
    Call ValidateI2(txtPriest7)
End Sub
Private Sub udSTR_Change()
    Call ValidateByte(txtSTR)
End Sub
Private Sub udVIT_Change()
    Call ValidateByte(txtVIT)
End Sub

