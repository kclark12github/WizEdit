VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWiz02 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wizardry 02 - The Knight of Diamonds"
   ClientHeight    =   6420
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   9288
   Icon            =   "frmWiz02.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9288
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
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   1020
      Width           =   7032
      Begin VB.TextBox txtDown 
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
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "frmWiz02.frx":1CCA
         ToolTipText     =   "Down..."
         Top             =   1140
         Width           =   396
      End
      Begin VB.TextBox txtNorth 
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
         Left            =   5700
         Locked          =   -1  'True
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "frmWiz02.frx":1CCD
         ToolTipText     =   "North..."
         Top             =   1140
         Width           =   396
      End
      Begin VB.TextBox txtEast 
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
         Left            =   4980
         Locked          =   -1  'True
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "frmWiz02.frx":1CD0
         ToolTipText     =   "East..."
         Top             =   1140
         Width           =   396
      End
      Begin VB.TextBox txtAC 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H0000C0C0&
         Height          =   360
         Left            =   2700
         Locked          =   -1  'True
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   164
         TabStop         =   0   'False
         Text            =   "frmWiz02.frx":1CD3
         ToolTipText     =   "Age...? (Still not sure of this one)..."
         Top             =   900
         Width           =   576
      End
      Begin VB.ComboBox cboAlignment 
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
         ItemData        =   "frmWiz02.frx":1CD6
         Left            =   4980
         List            =   "frmWiz02.frx":1CE0
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "TabStop"
         ToolTipText     =   "Character's Alignment (Good, Neutral, or Evil)..."
         Top             =   780
         Width           =   1932
      End
      Begin VB.CheckBox chkOut 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Left            =   1560
         TabIndex        =   6
         ToolTipText     =   "Left out in the Maze?"
         Top             =   924
         Width           =   204
      End
      Begin VB.TextBox txtYears 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H0000C0C0&
         Height          =   360
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   160
         Text            =   "frmWiz02.frx":1CF2
         ToolTipText     =   "Age...? (Still not sure of this one)..."
         Top             =   3240
         Width           =   396
      End
      Begin VB.TextBox txtPassword 
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "Password"
         ToolTipText     =   "Character Password..."
         Top             =   480
         Width           =   2076
      End
      Begin VB.TextBox txtName 
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "Name"
         ToolTipText     =   "Character Name..."
         Top             =   60
         Width           =   2076
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
         Left            =   4320
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frmWiz02.frx":1CF5
         ToolTipText     =   "Age...? (Still not sure of this one)..."
         Top             =   3240
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
         TabIndex        =   21
         Text            =   "frmWiz02.frx":1CF8
         ToolTipText     =   "Level..."
         Top             =   2400
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
         MaxLength       =   16
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "frmWiz02.frx":1CFB
         ToolTipText     =   "Gold Pieces..."
         Top             =   4020
         Width           =   1836
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
         MaxLength       =   16
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "frmWiz02.frx":1D0D
         ToolTipText     =   "Experience Points..."
         Top             =   1980
         Width           =   1836
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
         TabIndex        =   22
         Text            =   "frmWiz02.frx":1D1F
         ToolTipText     =   "Hit Points (automatically ""topped-off"")..."
         Top             =   2820
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
         TabIndex        =   24
         ToolTipText     =   "Character's Status (i.e. OK, Afraid, Poisoned, etc.)..."
         Top             =   3660
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
         Left            =   4980
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "TabStop"
         ToolTipText     =   "Character's Profession (i.e. Fighter, Mage, etc.)..."
         Top             =   420
         Width           =   1932
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
         ItemData        =   "frmWiz02.frx":1D22
         Left            =   4980
         List            =   "frmWiz02.frx":1D29
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Character's Race (i.e. Human, Elf, etc.)..."
         Top             =   60
         Width           =   1932
      End
      Begin MSComCtl2.UpDown udPIE 
         Height          =   360
         Left            =   2400
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   2820
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPIE"
         BuddyDispid     =   196625
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
         TabIndex        =   16
         Text            =   "frmWiz02.frx":1D34
         ToolTipText     =   "Piety..."
         Top             =   2820
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
         TabIndex        =   19
         Text            =   "frmWiz02.frx":1D37
         ToolTipText     =   "Speed..."
         Top             =   4080
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
         TabIndex        =   18
         Text            =   "frmWiz02.frx":1D3A
         ToolTipText     =   "Dexterity..."
         Top             =   3660
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
         TabIndex        =   17
         Text            =   "frmWiz02.frx":1D3D
         ToolTipText     =   "Vitality..."
         Top             =   3240
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
         TabIndex        =   15
         Text            =   "frmWiz02.frx":1D40
         ToolTipText     =   "Intelligence..."
         Top             =   2400
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
         ItemData        =   "frmWiz02.frx":1D43
         Left            =   4980
         List            =   "frmWiz02.frx":1D4D
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
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
         TabIndex        =   14
         Text            =   "frmWiz02.frx":1D5F
         ToolTipText     =   "Strength..."
         Top             =   1980
         Width           =   396
      End
      Begin MSComCtl2.UpDown udSTR 
         Height          =   360
         Left            =   2400
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   1980
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSTR"
         BuddyDispid     =   196631
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
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   2400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtINT"
         BuddyDispid     =   196629
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
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   3240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtVIT"
         BuddyDispid     =   196628
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
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   3660
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtAGL"
         BuddyDispid     =   196627
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
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   4080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLCK"
         BuddyDispid     =   196626
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
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   2820
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtHP"
         BuddyDispid     =   196620
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
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   2400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLVL"
         BuddyDispid     =   196617
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
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   2940
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtAge"
         BuddyDispid     =   196616
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
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D"
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
         Left            =   6240
         TabIndex        =   169
         Top             =   1176
         Width           =   180
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N"
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
         TabIndex        =   168
         Top             =   1176
         Width           =   204
      End
      Begin VB.Label lblE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E"
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
         Left            =   4740
         TabIndex        =   167
         Top             =   1176
         Width           =   168
      End
      Begin VB.Label lblLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location:"
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
         Left            =   3480
         TabIndex        =   166
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblAC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AC:"
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
         Left            =   2220
         TabIndex        =   165
         Top             =   936
         Width           =   360
      End
      Begin VB.Label lblAlignment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alignment:"
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
         TabIndex        =   163
         Top             =   804
         Width           =   1116
      End
      Begin VB.Label lblOut 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Out (Reset) ?"
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
         TabIndex        =   162
         Top             =   936
         Width           =   1344
      End
      Begin VB.Label lblYears 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Years)"
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
         Left            =   5580
         TabIndex        =   161
         Top             =   3276
         Width           =   732
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         TabIndex        =   159
         Top             =   516
         Width           =   1068
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   468
         TabIndex        =   158
         Top             =   96
         Width           =   660
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
         TabIndex        =   98
         Top             =   1680
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
         TabIndex        =   97
         Top             =   1680
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
         Left            =   3600
         TabIndex        =   95
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
         TabIndex        =   93
         Top             =   4116
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
         TabIndex        =   92
         Top             =   3696
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
         TabIndex        =   91
         Top             =   3276
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
         TabIndex        =   90
         Top             =   2856
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
         TabIndex        =   89
         Top             =   2436
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
         TabIndex        =   88
         Top             =   2016
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
         TabIndex        =   87
         Top             =   2856
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
         TabIndex        =   86
         Top             =   4056
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
         TabIndex        =   85
         Top             =   2016
         Width           =   1200
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age (Weeks):"
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
         Left            =   2880
         TabIndex        =   84
         Top             =   3276
         Width           =   1332
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
         TabIndex        =   83
         Top             =   2436
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
         TabIndex        =   82
         Top             =   3660
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
         Left            =   3600
         TabIndex        =   81
         Top             =   444
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
         Left            =   3600
         TabIndex        =   80
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
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   1020
      Width           =   7032
      Begin VB.CheckBox chkIdentified 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   312
         Index           =   8
         Left            =   2100
         TabIndex        =   56
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
         TabIndex        =   52
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
         TabIndex        =   48
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
         TabIndex        =   44
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
         TabIndex        =   40
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
         TabIndex        =   36
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
         TabIndex        =   32
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
         TabIndex        =   55
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
         TabIndex        =   51
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
         TabIndex        =   47
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
         TabIndex        =   43
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
         TabIndex        =   39
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
         TabIndex        =   35
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
         TabIndex        =   31
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
         TabIndex        =   54
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
         TabIndex        =   50
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
         TabIndex        =   46
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
         TabIndex        =   42
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
         TabIndex        =   38
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
         TabIndex        =   34
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
         TabIndex        =   30
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
         TabIndex        =   26
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
         TabIndex        =   27
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
         TabIndex        =   28
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
         ItemData        =   "frmWiz02.frx":1D62
         Left            =   2520
         List            =   "frmWiz02.frx":1D64
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #1..."
         Top             =   828
         Width           =   4092
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
         ItemData        =   "frmWiz02.frx":1D66
         Left            =   2520
         List            =   "frmWiz02.frx":1D68
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #8..."
         Top             =   3780
         Width           =   4092
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
         ItemData        =   "frmWiz02.frx":1D6A
         Left            =   2520
         List            =   "frmWiz02.frx":1D6C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #7..."
         Top             =   3360
         Width           =   4092
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
         ItemData        =   "frmWiz02.frx":1D6E
         Left            =   2520
         List            =   "frmWiz02.frx":1D70
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #6..."
         Top             =   2940
         Width           =   4092
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
         ItemData        =   "frmWiz02.frx":1D72
         Left            =   2520
         List            =   "frmWiz02.frx":1D74
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #5..."
         Top             =   2520
         Width           =   4092
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
         ItemData        =   "frmWiz02.frx":1D76
         Left            =   2520
         List            =   "frmWiz02.frx":1D78
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #4..."
         Top             =   2100
         Width           =   4092
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
         ItemData        =   "frmWiz02.frx":1D7A
         Left            =   2520
         List            =   "frmWiz02.frx":1D7C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #3..."
         Top             =   1680
         Width           =   4092
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
         ItemData        =   "frmWiz02.frx":1D7E
         Left            =   2520
         List            =   "frmWiz02.frx":1D80
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Tag             =   "TabStop"
         ToolTipText     =   "Item Carried in slot #2..."
         Top             =   1260
         Width           =   4092
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
         TabIndex        =   153
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
         TabIndex        =   152
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
         TabIndex        =   151
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
         TabIndex        =   150
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
         TabIndex        =   149
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
         TabIndex        =   148
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
         TabIndex        =   147
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
         TabIndex        =   146
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
         TabIndex        =   145
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   1020
      Width           =   7032
      Begin VB.CommandButton cmdPriestNone 
         Caption         =   "None"
         Height          =   252
         Left            =   5400
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   2400
         Width           =   552
      End
      Begin VB.CommandButton cmdPriestAll 
         Caption         =   "All"
         Height          =   252
         Left            =   4800
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   2400
         Width           =   552
      End
      Begin VB.CommandButton cmdMageNone 
         Caption         =   "None"
         Height          =   252
         Left            =   1680
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   2400
         Width           =   552
      End
      Begin VB.CommandButton cmdMageAll 
         Caption         =   "All"
         Height          =   252
         Left            =   1080
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   2400
         Width           =   552
      End
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   73
         Text            =   "frmWiz02.frx":1D82
         ToolTipText     =   "Level 7 Priest Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   72
         Text            =   "frmWiz02.frx":1D88
         ToolTipText     =   "Level 6 Priest Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   71
         Text            =   "frmWiz02.frx":1D8E
         ToolTipText     =   "Level 5 Priest Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   66
         Text            =   "frmWiz02.frx":1D94
         ToolTipText     =   "Level 7 Mage Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   65
         Text            =   "frmWiz02.frx":1D9A
         ToolTipText     =   "Level 6 Mage Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   64
         Text            =   "frmWiz02.frx":1DA0
         ToolTipText     =   "Level 5 Mage Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   70
         Text            =   "frmWiz02.frx":1DA6
         ToolTipText     =   "Level 4 Priest Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   63
         Text            =   "frmWiz02.frx":1DAC
         ToolTipText     =   "Level 4 Mage Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   69
         Text            =   "frmWiz02.frx":1DB2
         ToolTipText     =   "Level 3 Priest Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   62
         Text            =   "frmWiz02.frx":1DB8
         ToolTipText     =   "Level 3 Mage Spell Points..."
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
         TabIndex        =   59
         Top             =   540
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
         TabIndex        =   58
         Top             =   540
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   68
         Text            =   "frmWiz02.frx":1DBE
         ToolTipText     =   "Level 2 Priest Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   61
         Text            =   "frmWiz02.frx":1DC4
         ToolTipText     =   "Level 2 Mage Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   67
         Text            =   "frmWiz02.frx":1DCA
         ToolTipText     =   "Level 1 Priest Spell Points..."
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
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   60
         Text            =   "frmWiz02.frx":1DD0
         ToolTipText     =   "Level 1 Mage Spell Points..."
         Top             =   3444
         Width           =   576
      End
      Begin MSComCtl2.UpDown udMage1 
         Height          =   360
         Left            =   1500
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   3444
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage1"
         BuddyDispid     =   196699
         OrigLeft        =   2400
         OrigTop         =   3444
         OrigRight       =   2640
         OrigBottom      =   3804
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest1 
         Height          =   360
         Left            =   1500
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest1"
         BuddyDispid     =   196698
         OrigLeft        =   2400
         OrigTop         =   3900
         OrigRight       =   2640
         OrigBottom      =   4260
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage2 
         Height          =   360
         Left            =   2340
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage2"
         BuddyDispid     =   196697
         OrigLeft        =   3300
         OrigTop         =   3420
         OrigRight       =   3540
         OrigBottom      =   3780
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest2 
         Height          =   360
         Left            =   2340
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest2"
         BuddyDispid     =   196696
         OrigLeft        =   3240
         OrigTop         =   3900
         OrigRight       =   3480
         OrigBottom      =   4260
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage3 
         Height          =   360
         Left            =   3180
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage3"
         BuddyDispid     =   196693
         OrigLeft        =   4140
         OrigTop         =   3420
         OrigRight       =   4380
         OrigBottom      =   3780
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest3 
         Height          =   360
         Left            =   3180
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest3"
         BuddyDispid     =   196692
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage4 
         Height          =   360
         Left            =   4020
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage4"
         BuddyDispid     =   196691
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest4 
         Height          =   360
         Left            =   4020
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest4"
         BuddyDispid     =   196690
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage5 
         Height          =   360
         Left            =   4860
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage5"
         BuddyDispid     =   196689
         OrigLeft        =   5760
         OrigTop         =   3420
         OrigRight       =   6000
         OrigBottom      =   3780
         Max             =   9
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage6 
         Height          =   360
         Left            =   5700
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage6"
         BuddyDispid     =   196688
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMage7 
         Height          =   360
         Left            =   6540
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   3420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMage7"
         BuddyDispid     =   196687
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest5 
         Height          =   360
         Left            =   4860
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest5"
         BuddyDispid     =   196686
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest6 
         Height          =   360
         Left            =   5700
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest6"
         BuddyDispid     =   196685
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPriest7 
         Height          =   360
         Left            =   6540
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   3900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPriest7"
         BuddyDispid     =   196684
         OrigLeft        =   5340
         OrigTop         =   324
         OrigRight       =   5580
         OrigBottom      =   684
         Max             =   9
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
         TabIndex        =   134
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
         TabIndex        =   133
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   127
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
         TabIndex        =   126
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
         TabIndex        =   125
         Top             =   3960
         Width           =   708
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
         TabIndex        =   116
         Top             =   120
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
         TabIndex        =   115
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
         TabIndex        =   114
         Top             =   120
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
      TabIndex        =   109
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
      Picture         =   "frmWiz02.frx":1DD6
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   2460
      UseMaskColor    =   -1  'True
      Width           =   1212
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Height          =   432
      Left            =   7680
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmWiz02.frx":3325
      Style           =   1  'Graphical
      TabIndex        =   74
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
      TabIndex        =   77
      Top             =   2940
      Width           =   1212
   End
   Begin VB.PictureBox picWiz02 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2052
      Left            =   7560
      Picture         =   "frmWiz02.frx":4882
      ScaleHeight     =   2004
      ScaleWidth      =   1332
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   120
      Width           =   1380
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H00000000&
      Height          =   372
      Index           =   1
      Left            =   120
      Picture         =   "frmWiz02.frx":508C
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   96
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
      Picture         =   "frmWiz02.frx":1E8AE
      ScaleHeight     =   324
      ScaleWidth      =   1104
      TabIndex        =   110
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
      Picture         =   "frmWiz02.frx":380D0
      ScaleHeight     =   324
      ScaleWidth      =   1464
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   660
      Width           =   1512
      Begin VB.Label lblSpells 
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
         TabIndex        =   3
         Top             =   0
         Width           =   1272
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
            Picture         =   "frmWiz02.frx":518F2
            Key             =   "Wiz02"
            Object.Tag             =   "Wiz02"
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
            Picture         =   "frmWiz02.frx":535CE
            Key             =   "Wiz02"
            Object.Tag             =   "Wiz02"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   94
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
            TextSave        =   "10:32 PM"
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
      TabIndex        =   76
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
      TabIndex        =   78
      Top             =   264
      Width           =   1020
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuOptionsHexDump 
         Caption         =   "&Hex Dump..."
      End
      Begin VB.Menu mnuOptionsPrint 
         Caption         =   "&Print Characters..."
      End
      Begin VB.Menu mnuOptionsSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About &WizEdit 2000..."
      End
   End
End
Attribute VB_Name = "frmWiz02"
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
Private Characters(1 To Wiz02CharactersMax) As Wiz02Character
Private SaveCharacter As Wiz02Character
Private Sub cboAlignment_Validate(Cancel As Boolean)
    Cancel = False
    Select Case cboAlignment.Text
        Case "Good"
            Select Case cboProfession.Text
                Case "Thief", "Ninja"
                    Cancel = True
            End Select
        Case "Neutral"
            Select Case cboProfession.Text
                Case "Priest", "Bishop", "Lord", "Ninja"
                    Cancel = True
            End Select
        Case "Evil"
            Select Case cboProfession.Text
                Case "Samurai", "Lord"
                    Cancel = True
            End Select
    End Select
    If Cancel Then Call MsgBox(cboProfession.Text & " may not be of " & cboAlignment.Text & " alignment.", vbExclamation, Me.Caption)
End Sub
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
    Call LoadCharacter(SelectedCharacter)
End Sub
Private Sub ClearFields()
    Dim ctl As Control
    Dim i As Integer
    
    If cboCharacter.ListIndex = -1 Then
        cmdEdit.Visible = False
    Else
        cmdEdit.Visible = True
    End If
    cmdSave.Visible = False
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
Private Sub cboProfession_Validate(Cancel As Boolean)
    ResetSpellPointMax
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
    'Verify Cancel if fields have changed...
    
    Characters(SelectedCharacter) = SaveCharacter
    Call LoadCharacter(SelectedCharacter)
    Call DisableFields
    Call picTabs_Click(1)
    cboCharacter.SetFocus
End Sub
Private Sub cmdCancel_GotFocus()
    TextSelected
End Sub
Private Sub cmdEdit_Click()
    SaveCharacter = Characters(SelectedCharacter)
    Call EnableFields
    Select Case ActiveTab
        Case 1
            txtName.SetFocus
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
Private Sub cmdMageAll_Click()
    Dim i As Integer
    For i = 0 To lstMageSpells.ListCount - 1
        lstMageSpells.Selected(i) = True
    Next i
    lstMageSpells.ListIndex = -1
End Sub
Private Sub cmdMageNone_Click()
    Dim i As Integer
    For i = 0 To lstMageSpells.ListCount - 1
        lstMageSpells.Selected(i) = False
    Next i
    lstMageSpells.ListIndex = -1
End Sub
Private Sub cmdPriestAll_Click()
    Dim i As Integer
    For i = 0 To lstPriestSpells.ListCount - 1
        lstPriestSpells.Selected(i) = True
    Next i
    lstPriestSpells.ListIndex = -1
End Sub
Private Sub cmdPriestNone_Click()
    Dim i As Integer
    For i = 0 To lstPriestSpells.ListCount - 1
        lstPriestSpells.Selected(i) = False
    Next i
    lstPriestSpells.ListIndex = -1
End Sub
Private Sub cmdSave_Click()
    'Move data from screen controls back into Characters array...
    'Debug.Print String(80, "=") & vbCrLf & "Before:"
    'Call DumpWiz02(Characters(SelectedCharacter))
    Call UnloadCharacter(SelectedCharacter)
    'Debug.Print String(80, "=") & vbCrLf & "After:"
    'Call Wiz02DumpCharacter(Characters(SelectedCharacter))
    
    'Write Data back to DataFile...
    Call Wiz02Write(DataFile, Characters)
    
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
    cmdSave.Visible = False
    cmdCancel.Visible = False
    cmdExit.Visible = True
    mnuOptions.Enabled = True
    mnuHelp.Enabled = True
    
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "CommandButton"
                Select Case ctl.Name
                    Case "cmdMageAll", "cmdMageNone", "cmdPriestAll", "cmdPriestNone"
                        ctl.Enabled = False
                    Case Else
                End Select
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
    cmdSave.Visible = True
    cmdCancel.Visible = True
    cmdExit.Visible = False
    mnuOptions.Enabled = False
    mnuHelp.Enabled = False
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "CommandButton"
                Select Case ctl.Name
                    Case "cmdMageAll", "cmdMageNone", "cmdPriestAll", "cmdPriestNone"
                        ctl.Enabled = True
                    Case Else
                End Select
            Case "UpDown"
                ctl.Enabled = True
            Case "ListBox"
                ctl.Enabled = True
            Case "CheckBox"
                ctl.Enabled = True
            Case "TextBox"
                Select Case ctl.Name
                    Case "txtAC", "txtYears"
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
Private Sub cmdSave_GotFocus()
    TextSelected
End Sub
Private Sub Form_Activate()
    Dim i As Integer
    
    'Populate Form with data from disk...
    Call Wiz02Read(DataFile, Characters)
    For i = 1 To Wiz02CharactersMax
        If Trim(Characters(i).Name) <> vbNullString Then cboCharacter.AddItem Trim(Characters(i).Name), i - 1
    Next i
    
    cmdEdit.Visible = False
    Call DisableFields
    Call ClearFields
    Call picTabs_Click(1)
    If cboCharacter.ListCount > 0 Then cboCharacter.ListIndex = 0
End Sub
Private Sub Form_Load()
    Dim i As Integer
    
    Me.Picture = frmMain.Picture
    For i = 1 To 3
        picTabs(i).Picture = frmMain.Picture
        picFrames(i).Picture = frmMain.Picture
    Next i
    'picWiz02.Picture = frmMain.picWiz02.Picture
    picWizardryLogo.Picture = frmMain.picWizardryLogo.Picture
    cmdCancel.Picture = frmMain.cmdCancel.Picture
    cmdExit.Picture = frmMain.cmdExit.Picture
    
    cboGender.Visible = False
    lblGender.Visible = False
    
    Call InitializeWiz02ItemList
    Call InitializeWiz02Spells
    
    Call PopulateWiz02Status(cboStatus)
    Call PopulateWiz02Profession(cboProfession)
    Call PopulateWiz02Alignment(cboAlignment)
    Call PopulateWiz02Race(cboRace)
    For i = 1 To Wiz02ItemListMax
        Call PopulateWiz02Item(cboItem(i))
    Next i
    Call PopulateWiz02SpellBooks(lstMageSpells, lstPriestSpells)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.MainCancel
    frmMain.Show
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
Private Sub LoadCharacter(iCharacter As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim SpellNumber As Integer
    Dim bString As String
    Dim ctl As Control
    Dim dTemp As Double
    
    With Characters(iCharacter)
        txtName.Text = Trim(Left(.Name, .NameLength))
        txtPassword.Text = Trim(Left(.Password, .PasswordLength))
        If .Out = 1 Then chkOut.Value = vbChecked Else chkOut.Value = vbUnchecked
        
        cboStatus.ListIndex = .Status
        cboAlignment.ListIndex = .Alignment
        cboGender.ListIndex = -1
        cboProfession.ListIndex = .Profession
        
        cboRace.ListIndex = .Race
        
        txtSTR.Text = cvtStatisticToInt(.Statistics, 1)
        txtINT.Text = cvtStatisticToInt(.Statistics, 2)
        txtPIE.Text = cvtStatisticToInt(.Statistics, 3)
        txtVIT.Text = cvtStatisticToInt(.Statistics, 4)
        txtAGL.Text = cvtStatisticToInt(.Statistics, 5)
        txtLCK.Text = cvtStatisticToInt(.Statistics, 6)
        
        txtLVL.Text = Format(.LVL.Maximum, "#,##0")
        txtAge.Text = Format(.AgeInWeeks, "#,##0")
        txtYears.Text = Format(.AgeInWeeks \ 52, "#,##0")
        txtAC.Text = .AC
        txtEast.Text = .Location \ 100
        txtNorth.Text = .Location Mod 100
        txtDown.Text = .Down
                
        txtHP.Text = Format(.HP.Maximum, "#,##0")
        txtEXP.Text = Format(I6toD(.EXP), "#,##0")
        txtGP.Text = Format(I6toD(.GP), "#,##0")
    
        'SpellBooks...
        bString = cvtSpellsToBstr(.SpellBooks)
        For i = 1 To Wiz02SpellMapMax
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
        ResetSpellPointMax
        
        For i = 1 To Wiz02ItemListMax
            cboItem(i).ListIndex = -1
        Next i
        For i = 1 To .ItemCount
            cboItem(i).ListIndex = lkupWiz02ItemByCbo(.ItemList(i).ItemCode, cboItem(i))
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
Private Sub mnuOptionsExit_Click()
    cmdExit_Click
End Sub
Private Sub mnuOptionsHexDump_Click()
    Dim xOutput As String
    Dim sStatus As String
    On Error GoTo ErrorHandler

    initCommonDialog
    With frmMain.cdgMain
        .FileName = ParsePath(DataFile, DrvDirFileNameBase) & " " & _
            Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00") & "-" & _
            Format(Hour(Time), "00") & Format(Minute(Time), "00") & Format(Second(Time), "00") & _
            ".dmp"
        .Filter = "Hex Dump Files (*.DMP)|*.dmp|Text Files (*.TXT)|*.txt|All Files (*.*)|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist + _
            cdlOFNNoChangeDir + cdlOFNNoReadOnlyReturn
        .CancelError = True
        .ShowSave    ' Call the open file procedure.
        xOutput = .FileName
    End With
    
    cboCharacter.Enabled = False
    cmdExit.Enabled = False
    cmdEdit.Enabled = False
    Me.MousePointer = vbHourglass
    sStatus = sbStatus.Panels("Message").Text
    sbStatus.Panels("Message").Text = "Creating " & xOutput & "..."
    Call HexDump(DataFile, xOutput)
    Call MsgBox("HexDump complete.", vbInformation, Me.Caption)
    sbStatus.Panels("Message").Text = sStatus
    Me.MousePointer = vbDefault
    cboCharacter.Enabled = True
    cmdExit.Enabled = True
    cmdEdit.Enabled = True
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case cdlCancel
            sbStatus.Panels("Status").Text = "Cancelled"
            MsgBox "Operation cancelled at User's request.", vbInformation, Me.Caption
            Exit Sub
        Case 52, 76
            Resume Next
        Case Else
            MsgBox Err.Description & " (Error #" & Err.Number & ")", Me.Caption
            Exit Sub
    End Select
End Sub
Private Sub mnuOptionsPrint_Click()
    Dim i As Long
    Dim oUnit As Integer
    Dim xOutput As String
    Dim sStatus As String
    Dim errorCode As Long
    
    On Error GoTo ErrorHandler

    initCommonDialog
    With frmMain.cdgMain
        .FileName = ParsePath(DataFile, DrvDirFileNameBase) & " Characters " & _
            Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00") & "-" & _
            Format(Hour(Time), "00") & Format(Minute(Time), "00") & Format(Second(Time), "00") & _
            ".txt"
        .Filter = "Text Files (*.TXT)|*.txt|All Files (*.*)|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist + _
            cdlOFNNoChangeDir + cdlOFNNoReadOnlyReturn
        .CancelError = True
        .ShowSave    ' Call the open file procedure.
        xOutput = .FileName
    End With
    
    cboCharacter.Enabled = False
    cmdExit.Enabled = False
    cmdEdit.Enabled = False
    Me.MousePointer = vbHourglass
    sStatus = sbStatus.Panels("Message").Text
    sbStatus.Panels("Message").Text = "Creating " & xOutput & "..."
    
    oUnit = FreeFile
    Open xOutput For Output Access Write Lock Read Write As #oUnit
    For i = 1 To UBound(Characters)
        If Trim(Characters(i).Name) <> vbNullString Then Call Wiz02PrintCharacter(oUnit, Characters(i))
    Next i
    Close #oUnit
    
    Call MsgBox("Print complete.", vbInformation, Me.Caption)
    sbStatus.Panels("Message").Text = sStatus
    Me.MousePointer = vbDefault
    cboCharacter.Enabled = True
    cmdExit.Enabled = True
    cmdEdit.Enabled = True
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case cdlCancel
            sbStatus.Panels("Status").Text = "Cancelled"
            MsgBox "Operation cancelled at User's request.", vbInformation, Me.Caption
            Exit Sub
        Case 52, 76
            Resume Next
        Case Else
            MsgBox Err.Description & " (Error #" & Err.Number & ")", Me.Caption
            Exit Sub
    End Select
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
            Case "Label", "ImageList", "CommandButton", "UpDown", "Menu"
            Case Else
                Select Case ctl.Name
                    Case "txtAC", "txtYears"
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
            txtName.SetFocus
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
Private Sub picWiz02_GotFocus()
    TextSelected
End Sub
Private Sub picWizardryLogo_GotFocus()
    TextSelected
End Sub
Private Sub ResetSpellPointMax()
    Dim cClass As Integer
    
    cClass = cboProfession.ListIndex
    udMage1.Max = 0
    udMage2.Max = 0
    udMage3.Max = 0
    udMage4.Max = 0
    udMage5.Max = 0
    udMage6.Max = 0
    udMage7.Max = 0
    udPriest1.Max = 0
    udPriest2.Max = 0
    udPriest3.Max = 0
    udPriest4.Max = 0
    udPriest5.Max = 0
    udPriest6.Max = 0
    udPriest7.Max = 0
    
    If IsMage(cClass) Then
        udMage1.Max = 9
        udMage2.Max = 9
        udMage3.Max = 9
        udMage4.Max = 9
        udMage5.Max = 9
        udMage6.Max = 9
        udMage7.Max = 9
    End If
    If IsPriest(cClass) Then
        udPriest1.Max = 9
        udPriest2.Max = 9
        udPriest3.Max = 9
        udPriest4.Max = 9
        udPriest5.Max = 9
        udPriest6.Max = 9
        udPriest7.Max = 9
    End If
    If IsSamurai(cClass) Then  '?
        udMage1.Max = 4
        udMage2.Max = 2
        udMage3.Max = 2
        udMage4.Max = 3
        udMage5.Max = 3
        udMage6.Max = 4
        udMage7.Max = 3
    ElseIf IsLord(cClass) Then  'Usually, you get to be a Lord after being a Bishop, so I'll give you Mage points too...!
        udMage1.Max = 4
        udMage2.Max = 2
        udMage3.Max = 2
        udMage4.Max = 3
        udMage5.Max = 3
        udMage6.Max = 4
        udMage7.Max = 3
        
        udPriest1.Max = 4
        udPriest2.Max = 2
        udPriest3.Max = 2
        udPriest4.Max = 3
        udPriest5.Max = 3
        udPriest6.Max = 4
        udPriest7.Max = 3
    End If
    
    txtMage1.Text = udMage1.Max
    txtMage2.Text = udMage2.Max
    txtMage3.Text = udMage3.Max
    txtMage4.Text = udMage4.Max
    txtMage5.Text = udMage5.Max
    txtMage6.Text = udMage6.Max
    txtMage7.Text = udMage7.Max
    txtPriest1.Text = udPriest1.Max
    txtPriest2.Text = udPriest2.Max
    txtPriest3.Text = udPriest3.Max
    txtPriest4.Text = udPriest4.Max
    txtPriest5.Text = udPriest5.Max
    txtPriest6.Text = udPriest6.Max
    txtPriest7.Text = udPriest7.Max
End Sub
Private Sub txtAge_Change()
    If Trim(txtAge.Text) = vbNullString Then Exit Sub
    txtYears.Text = Format(CInt(txtAge.Text) \ 52, "#,##0")
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
Private Sub txtDown_GotFocus()
    TextSelected
End Sub
Private Sub txtDown_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtDown_Validate(Cancel As Boolean)
    If Trim(txtDown.Text) = vbNullString Or Val(txtDown.Text) < 0 Or Val(txtDown.Text) > 10 Then
        Call MsgBox("Levels must be between 0 (Castle) and 10.", vbExclamation, Me.Caption)
        Cancel = True
    End If
End Sub
Private Sub txtEast_GotFocus()
    TextSelected
End Sub
Private Sub txtEast_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtEast_Validate(Cancel As Boolean)
    If Trim(txtEast.Text) = vbNullString Or Val(txtEast.Text) < 0 Or Val(txtEast.Text) > 19 Then
        Call MsgBox("Coordinates must be between 0 and 19.", vbExclamation, Me.Caption)
        Cancel = True
    End If
End Sub
Private Sub txtEXP_GotFocus()
    TextSelected
End Sub
Private Sub txtEXP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtEXP_Validate(Cancel As Boolean)
    Cancel = ValidateI4()
    If CDbl(txtEXP.Text) > 999999999999# Then
        Call MsgBox("Hey, Wizardry's only going to display 12 digits... Lighten-Up...!", vbExclamation, Me.Caption)
        Cancel = True
    End If
End Sub
Private Sub txtGP_GotFocus()
    TextSelected
End Sub
Private Sub txtGP_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtGP_Validate(Cancel As Boolean)
    Cancel = ValidateI4()
    If CDbl(txtGP.Text) > 999999999999# Then
        Call MsgBox("Hey, Wizardry's only going to display 12 digits... Lighten-Up...!", vbExclamation, Me.Caption)
        Cancel = True
    End If
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
    If CInt(txtMage1.Text) > udMage1.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtMage2_GotFocus()
    TextSelected
End Sub
Private Sub txtMage2_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage2_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtMage2.Text) > udMage2.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtMage3_GotFocus()
    TextSelected
End Sub
Private Sub txtMage3_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage3_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtMage3.Text) > udMage3.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtMage4_GotFocus()
    TextSelected
End Sub
Private Sub txtMage4_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage4_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtMage4.Text) > udMage4.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtMage5_GotFocus()
    TextSelected
End Sub
Private Sub txtMage5_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage5_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtMage5.Text) > udMage5.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtMage6_GotFocus()
    TextSelected
End Sub
Private Sub txtMage6_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage6_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtMage6.Text) > udMage6.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtMage7_GotFocus()
    TextSelected
End Sub
Private Sub txtMage7_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtMage7_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtMage7.Text) > udMage7.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtName_GotFocus()
    TextSelected
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    txtName.Text = UCase(txtName.Text)
End Sub
Private Sub txtNorth_GotFocus()
    TextSelected
End Sub
Private Sub txtNorth_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtNorth_Validate(Cancel As Boolean)
    If Trim(txtNorth.Text) = vbNullString Or Val(txtNorth.Text) < 0 Or Val(txtNorth.Text) > 19 Then
        Call MsgBox("Coordinates must be between 0 and 19.", vbExclamation, Me.Caption)
        Cancel = True
    End If
End Sub
Private Sub txtPassword_GotFocus()
    TextSelected
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub
Private Sub txtPassword_Validate(Cancel As Boolean)
    txtPassword.Text = UCase(txtPassword.Text)
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
    If CInt(txtPriest1.Text) > udPriest1.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtPriest2_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest2_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest2_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtPriest2.Text) > udPriest2.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtPriest3_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest3_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest3_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtPriest3.Text) > udPriest3.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtPriest4_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest4_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest4_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtPriest4.Text) > udPriest4.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtPriest5_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest5_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest5_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtPriest5.Text) > udPriest5.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtPriest6_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest6_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest6_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtPriest6.Text) > udPriest6.Max Then Cancel = True
    If Cancel Then Beep
End Sub
Private Sub txtPriest7_GotFocus()
    TextSelected
End Sub
Private Sub txtPriest7_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPriest7_Validate(Cancel As Boolean)
    Cancel = ValidateI2()
    If CInt(txtPriest7.Text) > udPriest7.Max Then Cancel = True
    If Cancel Then Beep
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
Private Sub UnloadCharacter(iCharacter As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim SpellNumber As Integer
    Dim bString As String
    Dim ctl As Control
    
    With Characters(iCharacter)
        .Name = Trim(txtName.Text)
        .NameLength = Len(Trim(txtName.Text))
        If cboCharacter.Text <> .Name Then cboCharacter.List(iCharacter - 1) = .Name
        .Password = Trim(txtPassword.Text)
        .PasswordLength = Len(Trim(txtPassword.Text))
        If chkOut.Value = vbChecked Then .Out = 1 Else .Out = 0

        .Status = cboStatus.ListIndex
        .Profession = cboProfession.ListIndex
        .Alignment = cboAlignment.ListIndex
        .Race = cboRace.ListIndex
        
        .Statistics = cvtStatisticsToLong(txtSTR.Text, txtINT.Text, txtPIE.Text, txtVIT.Text, txtAGL.Text, txtLCK.Text)
        
        .LVL.Maximum = CInt(txtLVL.Text)
        .LVL.Current = .LVL.Maximum
        .AgeInWeeks = CInt(txtAge.Text)
        .HP.Maximum = CInt(txtHP.Text)
        .HP.Current = .HP.Maximum
        Call DtoI6(CDbl(txtEXP.Text), .EXP)
        Call DtoI6(CDbl(txtGP.Text), .GP)
        txtEast.Text = .Location \ 100
        txtNorth.Text = .Location Mod 100
        .Down = txtDown.Text
    
        'SpellBooks...
        bString = String(UBound(.SpellBooks) * 8, "0")
        For i = 1 To Wiz02SpellMapMax
            If i <= 21 Then
                If lstMageSpells.Selected(i - 1) Then
                    Mid(bString, i + 1, 1) = "1"
                Else
                    Mid(bString, i + 1, 1) = "0"
                End If
            Else
                If lstPriestSpells.Selected(i - 1 - 21) Then
                    Mid(bString, i + 1, 1) = "1"
                Else
                    Mid(bString, i + 1, 1) = "0"
                End If
            End If
        Next i
        Call cvtBstrToSpells(bString, .SpellBooks)
        
        .MageSpellPoints(1) = CInt(txtMage1.Text)
        .MageSpellPoints(2) = CInt(txtMage2.Text)
        .MageSpellPoints(3) = CInt(txtMage3.Text)
        .MageSpellPoints(4) = CInt(txtMage4.Text)
        .MageSpellPoints(5) = CInt(txtMage5.Text)
        .MageSpellPoints(6) = CInt(txtMage6.Text)
        .MageSpellPoints(7) = CInt(txtMage7.Text)
        .PriestSpellPoints(1) = CInt(txtPriest1.Text)
        .PriestSpellPoints(2) = CInt(txtPriest2.Text)
        .PriestSpellPoints(3) = CInt(txtPriest3.Text)
        .PriestSpellPoints(4) = CInt(txtPriest4.Text)
        .PriestSpellPoints(5) = CInt(txtPriest5.Text)
        .PriestSpellPoints(6) = CInt(txtPriest6.Text)
        .PriestSpellPoints(7) = CInt(txtPriest7.Text)
        
        .ItemCount = 0
        For i = 1 To Wiz02ItemListMax
            .ItemList(i).ItemCode = lkupWiz02ItemByName(cboItem(i).Text)
            If .ItemList(i).ItemCode > 0 Then .ItemCount = .ItemCount + 1
            If chkIdentified(i).Value = vbChecked Then .ItemList(i).Identified = 1 Else .ItemList(i).Identified = 0
            If chkEquipped(i).Value = vbChecked Then .ItemList(i).Equipped = 1 Else .ItemList(i).Equipped = 0
            If chkCursed(i).Value = vbChecked Then .ItemList(i).Cursed = -1 Else .ItemList(i).Cursed = 0
        Next i
    End With
End Sub
