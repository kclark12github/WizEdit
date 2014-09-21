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
   Begin VB.PictureBox picWiz07 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2064
      Left            =   9600
      Picture         =   "frmWiz07.frx":1CCA
      ScaleHeight     =   2016
      ScaleWidth      =   1344
      TabIndex        =   58
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
      TabIndex        =   36
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
         TabIndex        =   37
         Top             =   0
         Width           =   1032
      End
   End
   Begin VB.PictureBox picBasic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4992
      Left            =   120
      ScaleHeight     =   4968
      ScaleWidth      =   9108
      TabIndex        =   3
      Top             =   1020
      Width           =   9132
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
         Left            =   4740
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   59
         Text            =   "frmWiz07.frx":1BF37
         ToolTipText     =   "Hit Points..."
         Top             =   1560
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
         Left            =   4020
         Style           =   2  'Dropdown List
         TabIndex        =   57
         ToolTipText     =   "Character's Condition (i.e. OK, Afraid, Poisoned, etc.)..."
         Top             =   3960
         Width           =   1692
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
         TabIndex        =   56
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
         ItemData        =   "frmWiz07.frx":1BF3A
         Left            =   1440
         List            =   "frmWiz07.frx":1BF41
         Style           =   2  'Dropdown List
         TabIndex        =   55
         ToolTipText     =   "Character's Race (i.e. Human, Elf, etc.)..."
         Top             =   60
         Width           =   1692
      End
      Begin MSComCtl2.UpDown udPIE 
         Height          =   360
         Left            =   2280
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPIE"
         BuddyDispid     =   196658
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
         Left            =   1920
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   46
         Text            =   "frmWiz07.frx":1BF4C
         ToolTipText     =   "Piety..."
         Top             =   2400
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
         Left            =   1920
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   45
         Text            =   "frmWiz07.frx":1BF4F
         ToolTipText     =   "Karma..."
         Top             =   4500
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
         Left            =   1920
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   44
         Text            =   "frmWiz07.frx":1BF52
         ToolTipText     =   "Personality..."
         Top             =   4080
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
         Left            =   1920
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   43
         Text            =   "frmWiz07.frx":1BF55
         ToolTipText     =   "Speed..."
         Top             =   3660
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
         Left            =   1920
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   42
         Text            =   "frmWiz07.frx":1BF58
         ToolTipText     =   "Dexterity..."
         Top             =   3240
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
         Left            =   1920
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   41
         Text            =   "frmWiz07.frx":1BF5B
         ToolTipText     =   "Vitality..."
         Top             =   2820
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
         Left            =   1920
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   40
         Text            =   "frmWiz07.frx":1BF5E
         ToolTipText     =   "Intelligence..."
         Top             =   1980
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
         ItemData        =   "frmWiz07.frx":1BF61
         Left            =   1440
         List            =   "frmWiz07.frx":1BF6B
         Style           =   2  'Dropdown List
         TabIndex        =   35
         ToolTipText     =   "Character's Gender (Male/Female)..."
         Top             =   420
         Width           =   1692
      End
      Begin VB.TextBox Text1 
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
         Height          =   312
         Left            =   6300
         TabIndex        =   32
         ToolTipText     =   "Character Name"
         Top             =   4380
         Width           =   4152
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
         Left            =   1920
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "frmWiz07.frx":1BF7D
         ToolTipText     =   "Strength..."
         Top             =   1560
         Width           =   396
      End
      Begin MSComCtl2.UpDown udSTR 
         Height          =   360
         Left            =   2280
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1560
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSTR"
         BuddyDispid     =   196642
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
         Left            =   2280
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1980
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtINT"
         BuddyDispid     =   196652
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
         Left            =   2280
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2820
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtVIT"
         BuddyDispid     =   196653
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
         Left            =   2280
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtDEX"
         BuddyDispid     =   196654
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
         Left            =   2280
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3660
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSPD"
         BuddyDispid     =   196655
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
         Left            =   2280
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   4080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPER"
         BuddyDispid     =   196656
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
         Left            =   2280
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   4500
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtKAR"
         BuddyDispid     =   196657
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
         Left            =   5700
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1560
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtHP"
         BuddyDispid     =   196670
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
         Left            =   3000
         TabIndex        =   39
         Top             =   1260
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
         Left            =   60
         TabIndex        =   38
         Top             =   1260
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
         TabIndex        =   34
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
         Left            =   480
         TabIndex        =   30
         Top             =   4536
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
         Left            =   480
         TabIndex        =   29
         Top             =   4116
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
         Left            =   480
         TabIndex        =   28
         Top             =   3696
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
         Left            =   480
         TabIndex        =   27
         Top             =   3276
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
         Left            =   480
         TabIndex        =   26
         Top             =   2856
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
         Left            =   480
         TabIndex        =   25
         Top             =   2436
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
         Left            =   480
         TabIndex        =   24
         Top             =   2016
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
         Left            =   480
         TabIndex        =   23
         Top             =   1596
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
         Left            =   7200
         TabIndex        =   22
         Top             =   3780
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
         Left            =   7140
         TabIndex        =   21
         Top             =   3480
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
         Left            =   5640
         TabIndex        =   20
         Top             =   3720
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
         Left            =   7140
         TabIndex        =   19
         Top             =   3180
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
         Left            =   5640
         TabIndex        =   18
         Top             =   3480
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
         Left            =   5640
         TabIndex        =   17
         Top             =   3180
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
         Left            =   5400
         TabIndex        =   16
         Top             =   2820
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
         Left            =   2880
         TabIndex        =   15
         Top             =   3660
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
         Left            =   2880
         TabIndex        =   14
         Top             =   2016
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
         Left            =   3300
         TabIndex        =   13
         Top             =   1596
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
         Left            =   2880
         TabIndex        =   12
         Top             =   3060
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
         Left            =   2880
         TabIndex        =   11
         Top             =   2760
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
         Left            =   2880
         TabIndex        =   10
         Top             =   2436
         Width           =   1200
      End
      Begin VB.Label lblAlive 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "?Alive?:"
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
         TabIndex        =   9
         Top             =   4260
         Width           =   804
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
         Left            =   2880
         TabIndex        =   8
         Top             =   3360
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
         Left            =   2880
         TabIndex        =   7
         Top             =   2460
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
         Left            =   2880
         TabIndex        =   6
         Top             =   3960
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
      Picture         =   "frmWiz07.frx":1BF80
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
            Picture         =   "frmWiz07.frx":1CE53
            Key             =   "Wiz07"
            Object.Tag             =   "Wiz07"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWiz07.frx":1EB2F
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
            Picture         =   "frmWiz07.frx":1EE4B
            Key             =   "Wiz07"
            Object.Tag             =   "Wiz07"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWiz07.frx":20B27
            Key             =   "Wiz07g"
            Object.Tag             =   "Wiz07g"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   33
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
            TextSave        =   "12:52 AM"
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

'char *RaceMap[] = {"Human","Elf","Dwarf","Gnome","Hobbit","Faerie","Lizardman","Dracon","Rawulf","Felpurr","Mook"};
'char *ProfessionMap[] = {"Fighter","Mage","Priest","Thief","Ranger","Alchemist","Bard","Psionic","Valkyrie","Bishop","Lord","Samurai","Monk","Ninja"};
'char *ConditionMap[] = {"OK","Asleep","Blinded","Dead","Poisoned","Stoned","Insane","Afraid","Nauseated","Paralyzed","Irritated","Diseased"};
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
'/* The WIZEDIT program */
'
'#include <stdio.h>
'#include <stdlib.h>
'#include <string.h>
'#include <alloc.h>
'#include <io.h>
'#include <ctype.h>
'#include <conio.h>
'
'#include <input.h>
'#include <wizardry.h>
'
'#define MAX_NAME_LENGTH 7
'#define MAX_I4 4294967295
'#define MAX_I2 65535
'#define MAX_I1 255
'
'extern int  titlecolor;
'extern int  headercolor;
'extern int  promptcolor;
'extern int  errorcolor;
'extern int  warningcolor;
'extern int  selectcolor;
'extern int  menucolor;
'extern int  datecolor;
'extern int  timecolor;
'extern int  screencolor;
'extern int  inputcolor;
'extern int  displaycolor;
'
'long    FileSize = 0;
'char    *Buffer = NULL,
'        FileName[133],
'        CharacterName[8];
'struct ProgramInfo Program;
'struct Character   *Party[6];
'
'#define WAND_XY     22,7
'#define SWOR_XY     22,8
'#define AXE_XY      22,9
'#define MACE_XY     22,10
'#define POLE_XY     22,11
'#define THRO_XY     22,12
'#define SLIN_XY     22,13
'#define BOWS_XY     22,14
'#define SHIE_XY     22,15
'#define HAND_XY     22,16
'#define SWIM_XY     39,7
'#define CLIM_XY     39,8
'#define SCOU_XY     39,9
'#define MUSI_XY     39,10
'#define ORAT_XY     39,11
'#define LEGE_XY     39,12
'#define SKUL_XY     39,13
'#define NINJ_XY     39,14
'#define FIREA_XY    56,7
'#define REFL_XY     56,8
'#define SNAK_XY     56,9
'#define EAGL_XY     56,10
'#define POWE_XY     56,11
'#define MIND_XY     56,12
'#define ARTI_XY     73,7
'#define MYTH_XY     73,8
'#define MAPP_XY     73,9
'#define SCRI_XY     73,10
'#define DIPL_XY     73,11
'#define ALCH_XY     73,12
'#define THEOL_XY    73,13
'#define THEOS_XY    73,14
'#define THAU_XY     73,15
'#define KIRI_XY     73,16
'
'void EditSkills(void * CharacterBuffer)
'{
'   char           Temp[81];
'   struct Character *cp = (struct Character *)CharacterBuffer;
'
'AGAIN:
'    clrscr();
'   header("Edit Skills");
'
'   gotoxy(1, 3);  clreol();
'    textattr(displaycolor + (BLACK << 4));  lowvideo();
'    cprintf("Character: ");                highvideo();
'    cprintf("%s", cp->Name);
'
'    textattr(screencolor + (BLACK << 4));
'    highvideo();
'
'    gotoxy(4, 5);  cprintf("Skills");
'    gotoxy(7, 6);  cprintf("Weaponry         Physical         Personal         Academia");
'    lowvideo();
'    gotoxy(10, 7); cprintf("Wand&Dagger      Swimming         Firearms         Artifacts      ");
'    gotoxy(10, 8); cprintf("Sword            Climbing         Reflextion       Mythology      ");
'    gotoxy(10, 9); cprintf("Axe              Scouting         SnakeSpeed       Mapping        ");
'    gotoxy(10,10); cprintf("Mace&Flail       Music            EagleEye         Scribe         ");
'    gotoxy(10,11); cprintf("Pole&Staff       Oratory          PowerStrike      Diplomacy      ");
'    gotoxy(10,12); cprintf("Throwing         Legerdemain      MindControl      Alchemy        ");
'    gotoxy(10,13); cprintf("Sling            Skulduggery                       Theology       ");
'    gotoxy(10,14); cprintf("Bows             Ninjutsu                          Theosophy      ");
'    gotoxy(10,15); cprintf("Shield                                             Thaumaturgy    ");
'    gotoxy(10,16); cprintf("Hands&Feet                                         Kirijutsu      ");
'
'   gotoxy(WAND_XY); cprintf("%u", cp->Skills.Weaponry.Wand);
'   gotoxy(SWOR_XY); cprintf("%u", cp->Skills.Weaponry.Sword);
'   gotoxy(AXE_XY); cprintf("%u", cp->Skills.Weaponry.Axe);
'   gotoxy(MACE_XY); cprintf("%u", cp->Skills.Weaponry.Mace);
'   gotoxy(POLE_XY); cprintf("%u", cp->Skills.Weaponry.Pole);
'   gotoxy(THRO_XY); cprintf("%u", cp->Skills.Weaponry.Throwing);
'   gotoxy(SLIN_XY); cprintf("%u", cp->Skills.Weaponry.Sling);
'   gotoxy(BOWS_XY); cprintf("%u", cp->Skills.Weaponry.Bow);
'   gotoxy(SHIE_XY); cprintf("%u", cp->Skills.Weaponry.Shield);
'   gotoxy(HAND_XY); cprintf("%u", cp->Skills.Weaponry.HandToHand);
'
'   gotoxy(SWIM_XY); cprintf("%u", cp->Skills.Physical.Swimming);
'   gotoxy(CLIM_XY); cprintf("%u", cp->Skills.Physical.Climbing);
'   gotoxy(SCOU_XY); cprintf("%u", cp->Skills.Physical.Scouting);
'   gotoxy(MUSI_XY); cprintf("%u", cp->Skills.Physical.Music);
'   gotoxy(ORAT_XY); cprintf("%u", cp->Skills.Physical.Oratory);
'   gotoxy(LEGE_XY); cprintf("%u", cp->Skills.Physical.Legerdemain);
'   gotoxy(SKUL_XY); cprintf("%u", cp->Skills.Physical.Skulduggery);
'   gotoxy(NINJ_XY); cprintf("%u", cp->Skills.Physical.Ninjutsu);
'
'   gotoxy(FIREA_XY); cprintf("%u", cp->Skills.Personal.Firearms);
'   gotoxy(REFL_XY); cprintf("%u", cp->Skills.Personal.Reflextion);
'   gotoxy(SNAK_XY); cprintf("%u", cp->Skills.Personal.SnakeSpeed);
'   gotoxy(EAGL_XY); cprintf("%u", cp->Skills.Personal.EagleEye);
'   gotoxy(POWE_XY); cprintf("%u", cp->Skills.Personal.PowerStrike);
'   gotoxy(MIND_XY); cprintf("%u", cp->Skills.Personal.MindControl);
'
'   gotoxy(ARTI_XY); cprintf("%u", cp->Skills.Academia.Artifacts);
'   gotoxy(MYTH_XY); cprintf("%u", cp->Skills.Academia.Mythology);
'   gotoxy(MAPP_XY); cprintf("%u", cp->Skills.Academia.Mapping);
'   gotoxy(SCRI_XY); cprintf("%u", cp->Skills.Academia.Scribe);
'   gotoxy(DIPL_XY); cprintf("%u", cp->Skills.Academia.Diplomacy);
'   gotoxy(ALCH_XY); cprintf("%u", cp->Skills.Academia.Alchemy);
'   gotoxy(THEOL_XY); cprintf("%u", cp->Skills.Academia.Theology);
'   gotoxy(THEOS_XY); cprintf("%u", cp->Skills.Academia.Theosophy);
'   gotoxy(THAU_XY); cprintf("%u", cp->Skills.Academia.Thaumaturgy);
'   gotoxy(KIRI_XY); cprintf("%u", cp->Skills.Academia.Kirijutsu);
'
'   cp->Skills.Weaponry.Wand = GetNewI1(WAND_XY, 3, cp->Skills.Weaponry.Wand, 100);
'   cp->Skills.Weaponry.Sword = GetNewI1(SWOR_XY, 3, cp->Skills.Weaponry.Sword, 100);
'   cp->Skills.Weaponry.Axe = GetNewI1(AXE_XY, 3, cp->Skills.Weaponry.Axe, 100);
'   cp->Skills.Weaponry.Mace = GetNewI1(MACE_XY, 3, cp->Skills.Weaponry.Mace, 100);
'   cp->Skills.Weaponry.Pole = GetNewI1(POLE_XY, 3, cp->Skills.Weaponry.Pole, 100);
'   cp->Skills.Weaponry.Throwing = GetNewI1(THRO_XY, 3, cp->Skills.Weaponry.Throwing, 100);
'   cp->Skills.Weaponry.Sling = GetNewI1(SLIN_XY, 3, cp->Skills.Weaponry.Sling, 100);
'   cp->Skills.Weaponry.Bow = GetNewI1(BOWS_XY, 3, cp->Skills.Weaponry.Bow, 100);
'   cp->Skills.Weaponry.Shield = GetNewI1(SHIE_XY, 3, cp->Skills.Weaponry.Shield, 100);
'   cp->Skills.Weaponry.HandToHand = GetNewI1(HAND_XY, 3, cp->Skills.Weaponry.HandToHand, 100);
'
'   cp->Skills.Physical.Swimming = GetNewI1(SWIM_XY, 3, cp->Skills.Physical.Swimming, 100);
'   cp->Skills.Physical.Climbing = GetNewI1(CLIM_XY, 3, cp->Skills.Physical.Climbing, 100);
'   cp->Skills.Physical.Scouting = GetNewI1(SCOU_XY, 3, cp->Skills.Physical.Scouting, 100);
'   cp->Skills.Physical.Music = GetNewI1(MUSI_XY, 3, cp->Skills.Physical.Music, 100);
'   cp->Skills.Physical.Oratory = GetNewI1(ORAT_XY, 3, cp->Skills.Physical.Oratory, 100);
'   cp->Skills.Physical.Legerdemain = GetNewI1(LEGE_XY, 3, cp->Skills.Physical.Legerdemain, 100);
'   cp->Skills.Physical.Skulduggery = GetNewI1(SKUL_XY, 3, cp->Skills.Physical.Skulduggery, 100);
'   cp->Skills.Physical.Ninjutsu = GetNewI1(NINJ_XY, 3, cp->Skills.Physical.Ninjutsu, 100);
'
'   cp->Skills.Personal.Firearms = GetNewI1(FIREA_XY, 3, cp->Skills.Personal.Firearms, 100);
'   cp->Skills.Personal.Reflextion = GetNewI1(REFL_XY, 3, cp->Skills.Personal.Reflextion, 100);
'   cp->Skills.Personal.SnakeSpeed = GetNewI1(SNAK_XY, 3, cp->Skills.Personal.SnakeSpeed, 100);
'   cp->Skills.Personal.EagleEye = GetNewI1(EAGL_XY,3 , cp->Skills.Personal.EagleEye, 100);
'   cp->Skills.Personal.PowerStrike = GetNewI1(POWE_XY, 3, cp->Skills.Personal.PowerStrike, 100);
'   cp->Skills.Personal.MindControl = GetNewI1(MIND_XY, 3, cp->Skills.Personal.MindControl, 100);
'
'   cp->Skills.Academia.Artifacts = GetNewI1(ARTI_XY, 3, cp->Skills.Academia.Artifacts, 100);
'   cp->Skills.Academia.Mythology = GetNewI1(MYTH_XY, 3, cp->Skills.Academia.Mythology, 100);
'   cp->Skills.Academia.Mapping = GetNewI1(MAPP_XY, 3, cp->Skills.Academia.Mapping, 100);
'   cp->Skills.Academia.Scribe = GetNewI1(SCRI_XY, 3, cp->Skills.Academia.Scribe, 100);
'   cp->Skills.Academia.Diplomacy = GetNewI1(DIPL_XY, 3, cp->Skills.Academia.Diplomacy, 100);
'   cp->Skills.Academia.Alchemy = GetNewI1(ALCH_XY, 3, cp->Skills.Academia.Alchemy, 100);
'   cp->Skills.Academia.Theology = GetNewI1(THEOL_XY, 3, cp->Skills.Academia.Theology, 100);
'   cp->Skills.Academia.Theosophy = GetNewI1(THEOS_XY, 3, cp->Skills.Academia.Theosophy, 100);
'   cp->Skills.Academia.Thaumaturgy = GetNewI1(THAU_XY, 3, cp->Skills.Academia.Thaumaturgy, 100);
'   cp->Skills.Academia.Kirijutsu = GetNewI1(KIRI_XY, 3, cp->Skills.Academia.Kirijutsu, 100);
'
'   if (!Query("continue", ""))
'      goto AGAIN;
'
'EXIT_LABEL:
'   printf("\n");
'   return;
'}
'
'#define ITEM_START_ROW 7
'#define ITEMCODE_X     1
'#define ITEMNAME_X     7
'#define ITEMPIC_X      25
'#define ITEMWEIGHT_X   33
'#define ITEMCOUNT_X    40
'#define ITEMSTATUS_X   46
'#define ITEMSTATUSx_X  52
'#define ITEMAC_X       58
'#define ITEMUNKNOWN_X  61
'#define ITEMUNKNOWNx_X 69
'
'void EditItems(void * CharacterBuffer)
'{
'   int            i = 0;
'   char           Temp[81];
'   struct Item *ip = NULL;
'   struct Character *cp = (struct Character *)CharacterBuffer;
'
'AGAIN:
'    clrscr();
'   header("Edit Items");
'
'   gotoxy(1, 3);  clreol();
'    textattr(displaycolor + (BLACK << 4));  lowvideo();
'    cprintf("Character: ");                highvideo();
'    cprintf("%s", cp->Name);
'
'    textattr(screencolor + (BLACK << 4));
'    highvideo();
'
'    gotoxy(1, 5);  cprintf("Items");
'    lowvideo();
'
'    gotoxy(1, 6);  cprintf("Code  Name             Picture  Weight Count Status      AC Unknown");
'   textattr(displaycolor + (BLACK << 4));
'    lowvideo();
'   for (i = 0; i < 10; i++)
'   {
'      ip = (struct Item *)&cp->ItemList[i];
'      gotoxy(ITEMCODE_X, ITEM_START_ROW+i);      cprintf("%d", ip->ItemCode);
'
'      memset(Temp, NULL, 81);
'      strncpy(Temp, MapItemCode(ip->ItemCode), 17);
'      gotoxy(ITEMNAME_X, ITEM_START_ROW+i);      cprintf("%s", Temp);
'
'      gotoxy(ITEMPIC_X, ITEM_START_ROW+i);       cprintf("%u", ip->PictureCode);
'      gotoxy(ITEMWEIGHT_X, ITEM_START_ROW+i);    cprintf("%u", ip->Weight);
'      gotoxy(ITEMCOUNT_X, ITEM_START_ROW+i);     cprintf("%u", ip->Count);
'      gotoxy(ITEMSTATUS_X, ITEM_START_ROW+i);    cprintf("%u", ip->Status);
'      gotoxy(ITEMSTATUSx_X, ITEM_START_ROW+i);   cprintf("(x%02X)", ip->Status);
'      gotoxy(ITEMAC_X, ITEM_START_ROW+i);        cprintf("%u", ip->AC);
'      gotoxy(ITEMUNKNOWN_X, ITEM_START_ROW+i);   cprintf("%u", ip->Unknown);
'      gotoxy(ITEMUNKNOWNx_X, ITEM_START_ROW+i);  cprintf("(x%04X)", ip->Unknown);
'   }
'   textattr(warningcolor + (BLACK << 4));
'    lowvideo();
'   gotoxy(1, 19); cprintf("Note: Care should be taken in changing these values.");
'   gotoxy(1, 20); cprintf("      This layout is not yet fully understood.");
'   textattr(displaycolor + (BLACK << 4));
'
'   for (i = 0; i < 10; i++)
'   {
'      ip = (struct Item *)&cp->ItemList[i];
'      ip->ItemCode = GetNewI2(ITEMCODE_X, ITEM_START_ROW+i, 5, ip->ItemCode, ItemMapMax);
'      memset(Temp, NULL, 81);
'      strncpy(Temp, MapItemCode(ip->ItemCode), 17);
'      gotoxy(ITEMNAME_X, ITEM_START_ROW+i);      cprintf("%s", Temp);
'
'      ip->PictureCode = GetNewI2(ITEMPIC_X, ITEM_START_ROW+i, 5, ip->PictureCode, MAX_I2);
'      ip->Weight = GetNewI2(ITEMWEIGHT_X, ITEM_START_ROW+i, 5, ip->Weight, 3000);
'      ip->Count = GetNewI1(ITEMCOUNT_X, ITEM_START_ROW+i, 4, ip->Count, MAX_I1);
'      ip->Status = GetNewI2(ITEMSTATUS_X, ITEM_START_ROW+i, 3, ip->Status, MAX_I1);
'      gotoxy(ITEMSTATUSx_X, ITEM_START_ROW+i);   cprintf("(x%02X)", ip->Status);
'      ip->AC = GetNewI2(ITEMAC_X, ITEM_START_ROW+i, 2, ip->AC, MAX_I1);
'      ip->Unknown = GetNewI2(ITEMUNKNOWN_X, ITEM_START_ROW+i, 5, ip->Unknown, MAX_I2);
'      gotoxy(ITEMUNKNOWNx_X, ITEM_START_ROW+i);  cprintf("(x%04X)", ip->Unknown);
'   }
'
'   if (!Query("continue", ""))
'      goto AGAIN;
'
'EXIT_LABEL:
'   printf("\n");
'   return;
'}
'
'void EditSwagBag(void * CharacterBuffer)
'{
'   int            i = 0;
'   char           Temp[81];
'   struct Item *ip = NULL;
'   struct Character *cp = (struct Character *)CharacterBuffer;
'
'AGAIN:
'    clrscr();
'   header("Edit SwagBag");
'
'   gotoxy(1, 3);  clreol();
'    textattr(displaycolor + (BLACK << 4));  lowvideo();
'    cprintf("Character: ");                highvideo();
'    cprintf("%s", cp->Name);
'
'    textattr(screencolor + (BLACK << 4));
'    highvideo();
'
'    gotoxy(1, 5);  cprintf("SwagBag");
'    lowvideo();
'
'    gotoxy(1, 6);  cprintf("Code  Name             Picture  Weight Count Status      AC Unknown");
'   textattr(displaycolor + (BLACK << 4));
'    lowvideo();
'   for (i = 0; i < 10; i++)
'   {
'      ip = (struct Item *)&cp->SwagBag[i];
'      gotoxy(ITEMCODE_X, ITEM_START_ROW+i);      cprintf("%d", ip->ItemCode);
'
'      memset(Temp, NULL, 81);
'      strncpy(Temp, MapItemCode(ip->ItemCode), 17);
'      gotoxy(ITEMNAME_X, ITEM_START_ROW+i);      cprintf("%s", Temp);
'
'      gotoxy(ITEMPIC_X, ITEM_START_ROW+i);       cprintf("%u", ip->PictureCode);
'      gotoxy(ITEMWEIGHT_X, ITEM_START_ROW+i);    cprintf("%u", ip->Weight);
'      gotoxy(ITEMCOUNT_X, ITEM_START_ROW+i);     cprintf("%u", ip->Count);
'      gotoxy(ITEMSTATUS_X, ITEM_START_ROW+i);    cprintf("%u", ip->Status);
'      gotoxy(ITEMSTATUSx_X, ITEM_START_ROW+i);   cprintf("(x%02X)", ip->Status);
'      gotoxy(ITEMAC_X, ITEM_START_ROW+i);        cprintf("%u", ip->AC);
'      gotoxy(ITEMUNKNOWN_X, ITEM_START_ROW+i);   cprintf("%u", ip->Unknown);
'      gotoxy(ITEMUNKNOWNx_X, ITEM_START_ROW+i);  cprintf("(x%04X)", ip->Unknown);
'   }
'   textattr(warningcolor + (BLACK << 4));
'    lowvideo();
'   gotoxy(1, 19); cprintf("Note: Care should be taken in changing these values.");
'   gotoxy(1, 20); cprintf("      This layout is not yet fully understood.");
'   textattr(displaycolor + (BLACK << 4));
'
'   for (i = 0; i < 10; i++)
'   {
'      ip = (struct Item *)&cp->SwagBag[i];
'      ip->ItemCode = GetNewI2(ITEMCODE_X, ITEM_START_ROW+i, 5, ip->ItemCode, ItemMapMax);
'      memset(Temp, NULL, 81);
'      strncpy(Temp, MapItemCode(ip->ItemCode), 17);
'      gotoxy(ITEMNAME_X, ITEM_START_ROW+i);      cprintf("%s", Temp);
'
'      ip->PictureCode = GetNewI2(ITEMPIC_X, ITEM_START_ROW+i, 5, ip->PictureCode, MAX_I2);
'      ip->Weight = GetNewI2(ITEMWEIGHT_X, ITEM_START_ROW+i, 5, ip->Weight, 3000);
'      ip->Count = GetNewI1(ITEMCOUNT_X, ITEM_START_ROW+i, 4, ip->Count, 99);
'      ip->Status = GetNewI2(ITEMSTATUS_X, ITEM_START_ROW+i, 3, ip->Status, MAX_I1);
'      gotoxy(ITEMSTATUSx_X, ITEM_START_ROW+i);   cprintf("(x%02X)", ip->Status);
'      ip->AC = GetNewI2(ITEMAC_X, ITEM_START_ROW+i, 2, ip->AC, MAX_I1);
'      ip->Unknown = GetNewI2(ITEMUNKNOWN_X, ITEM_START_ROW+i, 5, ip->Unknown, MAX_I2);
'      gotoxy(ITEMUNKNOWNx_X, ITEM_START_ROW+i);  cprintf("(x%04X)", ip->Unknown);
'   }
'
'   if (!Query("continue", ""))
'      goto AGAIN;
'
'EXIT_LABEL:
'   printf("\n");
'   return;
'}
'
'#define B_YN 'Y' : 'N'
'
'void EditSpells1(void * CharacterBuffer)
'{
'   char           Temp[81];
'   struct Character *cp = (struct Character *)CharacterBuffer;
'
'AGAIN:
'    clrscr();
'   header("Edit Spells 1");
'
'   gotoxy(1, 3);  clreol();
'    textattr(displaycolor + (BLACK << 4));  lowvideo();
'    cprintf("Character: ");                highvideo();
'    cprintf("%s", cp->Name);
'
'    textattr(screencolor + (BLACK << 4));
'    highvideo();
'
'   gotoxy(7, 5);cprintf("Fire               Water              Air\n");
'   lowvideo();
'
'   gotoxy(7, 6);cprintf("EnergyBlast:       ChillingTouch:     Poison:          ");
'   gotoxy(7, 7);cprintf("BlindingFlash:     Stamina:           MissileShield:   ");
'   gotoxy(7, 8);cprintf("PsionicFire:       Terror:            ShrillSound:     ");
'   gotoxy(7, 9);cprintf("Fireball:          Weaken:            StinkBomb:       ");
'   gotoxy(7,10);cprintf("FireShield:        Slow:              AirPocket:       ");
'   gotoxy(7,11);cprintf("DazzlingLights:    Haste:             Silence:         ");
'   gotoxy(7,12);cprintf("FireBomb:          CureParalysis:     PoisonGas:       ");
'   gotoxy(7,13);cprintf("Lightning:         IceShield:         CurePoison:      ");
'   gotoxy(7,14);cprintf("PrismicMissile:    Restfull:          Whirlwind:       ");
'   gotoxy(7,15);cprintf("Firestorm:         IceBall:           PurifyAir:       ");
'   gotoxy(7,16);cprintf("NuclearBlast:      Paralyze:          DeadlyPoison:    ");
'   gotoxy(7,17);cprintf("                   Superman:          Levitate:        ");
'   gotoxy(7,18);cprintf("                   Deepfreeze:        ToxicVapors:     ");
'   gotoxy(7,19);cprintf("                   DrainingCloud:     NoxiousFumes:    ");
'   gotoxy(7,20);cprintf("                   CureDisease:       Asphyxiation:    ");
'   gotoxy(7,21);cprintf("                                      DeadlyAir:       ");
'   gotoxy(7,22);cprintf("                                      DeathCloud:      ");
'
'   textattr(displaycolor + (BLACK << 4));
'    lowvideo();
'
'   gotoxy(23, 6);cprintf("%c", cp->Spells.EnergyBlast ? B_YN);
'   gotoxy(23, 7);cprintf("%c", cp->Spells.BlindingFlash ? B_YN);
'   gotoxy(23, 8);cprintf("%c", cp->Spells.PsionicFire ? B_YN);
'   gotoxy(23, 9);cprintf("%c", cp->Spells.Fireball ? B_YN);
'   gotoxy(23,10);cprintf("%c", cp->Spells.FireShield ? B_YN);
'   gotoxy(23,11);cprintf("%c", cp->Spells.DazzlingLights ? B_YN);
'   gotoxy(23,12);cprintf("%c", cp->Spells.FireBomb ? B_YN);
'   gotoxy(23,13);cprintf("%c", cp->Spells.Lightning ? B_YN);
'   gotoxy(23,14);cprintf("%c", cp->Spells.PrismicMissile ? B_YN);
'   gotoxy(23,15);cprintf("%c", cp->Spells.Firestorm ? B_YN);
'   gotoxy(23,16);cprintf("%c", cp->Spells.NuclearBlast ? B_YN);
'
'   gotoxy(42, 6);cprintf("%c", cp->Spells.ChillingTouch ? B_YN);
'   gotoxy(42, 7);cprintf("%c", cp->Spells.Stamina ? B_YN);
'   gotoxy(42, 8);cprintf("%c", cp->Spells.Terror ? B_YN);
'   gotoxy(42, 9);cprintf("%c", cp->Spells.Weaken ? B_YN);
'   gotoxy(42,10);cprintf("%c", cp->Spells.Slow ? B_YN);
'   gotoxy(42,11);cprintf("%c", cp->Spells.Haste ? B_YN);
'   gotoxy(42,12);cprintf("%c", cp->Spells.CureParalysis ? B_YN);
'   gotoxy(42,13);cprintf("%c", cp->Spells.IceShield ? B_YN);
'   gotoxy(42,14);cprintf("%c", cp->Spells.Restfull ? B_YN);
'   gotoxy(42,15);cprintf("%c", cp->Spells.IceBall ? B_YN);
'   gotoxy(42,16);cprintf("%c", cp->Spells.Paralyze ? B_YN);
'   gotoxy(42,17);cprintf("%c", cp->Spells.Superman ? B_YN);
'   gotoxy(42,18);cprintf("%c", cp->Spells.Deepfreeze ? B_YN);
'   gotoxy(42,19);cprintf("%c", cp->Spells.DrainingCloud ? B_YN);
'   gotoxy(42,20);cprintf("%c", cp->Spells.CureDisease ? B_YN);
'
'   gotoxy(61, 6);cprintf("%c", cp->Spells.Poison ? B_YN);
'   gotoxy(61, 7);cprintf("%c", cp->Spells.MissileShield ? B_YN);
'   gotoxy(61, 8);cprintf("%c", cp->Spells.ShrillSound ? B_YN);
'   gotoxy(61, 9);cprintf("%c", cp->Spells.StinkBomb ? B_YN);
'   gotoxy(61,10);cprintf("%c", cp->Spells.AirPocket ? B_YN);
'   gotoxy(61,11);cprintf("%c", cp->Spells.Silence ? B_YN);
'   gotoxy(61,12);cprintf("%c", cp->Spells.PoisonGas ? B_YN);
'   gotoxy(61,13);cprintf("%c", cp->Spells.CurePoison ? B_YN);
'   gotoxy(61,14);cprintf("%c", cp->Spells.Whirlwind ? B_YN);
'   gotoxy(61,15);cprintf("%c", cp->Spells.PurifyAir ? B_YN);
'   gotoxy(61,16);cprintf("%c", cp->Spells.DeadlyPoison ? B_YN);
'   gotoxy(61,17);cprintf("%c", cp->Spells.Levitate ? B_YN);
'   gotoxy(61,18);cprintf("%c", cp->Spells.ToxicVapors ? B_YN);
'   gotoxy(61,19);cprintf("%c", cp->Spells.NoxiousFumes ? B_YN);
'   gotoxy(61,20);cprintf("%c", cp->Spells.Asphyxiation ? B_YN);
'   gotoxy(61,21);cprintf("%c", cp->Spells.DeadlyAir ? B_YN);
'   gotoxy(61,22);cprintf("%c", cp->Spells.DeathCloud ? B_YN);
'
'   cp->Spells.EnergyBlast = YesNo(23, 6, cp->Spells.EnergyBlast ? B_YN);
'   cp->Spells.BlindingFlash = YesNo(23, 7, cp->Spells.BlindingFlash ? B_YN);
'   cp->Spells.PsionicFire = YesNo(23, 8, cp->Spells.PsionicFire ? B_YN);
'   cp->Spells.Fireball = YesNo(23, 9, cp->Spells.Fireball ? B_YN);
'   cp->Spells.FireShield = YesNo(23,10, cp->Spells.FireShield ? B_YN);
'   cp->Spells.DazzlingLights = YesNo(23,11, cp->Spells.DazzlingLights ? B_YN);
'   cp->Spells.FireBomb = YesNo(23,12, cp->Spells.FireBomb ? B_YN);
'   cp->Spells.Lightning = YesNo(23,13, cp->Spells.Lightning ? B_YN);
'   cp->Spells.PrismicMissile = YesNo(23,14, cp->Spells.PrismicMissile ? B_YN);
'   cp->Spells.Firestorm = YesNo(23,15, cp->Spells.Firestorm ? B_YN);
'   cp->Spells.NuclearBlast = YesNo(23,16, cp->Spells.NuclearBlast ? B_YN);
'
'   cp->Spells.ChillingTouch = YesNo(42, 6, cp->Spells.ChillingTouch ? B_YN);
'   cp->Spells.Stamina = YesNo(42, 7, cp->Spells.Stamina ? B_YN);
'   cp->Spells.Terror = YesNo(42, 8, cp->Spells.Terror ? B_YN);
'   cp->Spells.Weaken = YesNo(42, 9, cp->Spells.Weaken ? B_YN);
'   cp->Spells.Slow = YesNo(42,10, cp->Spells.Slow ? B_YN);
'   cp->Spells.Haste = YesNo(42,11, cp->Spells.Haste ? B_YN);
'   cp->Spells.CureParalysis = YesNo(42,12, cp->Spells.CureParalysis ? B_YN);
'   cp->Spells.IceShield = YesNo(42,13, cp->Spells.IceShield ? B_YN);
'   cp->Spells.Restfull = YesNo(42,14, cp->Spells.Restfull ? B_YN);
'   cp->Spells.IceBall = YesNo(42,15, cp->Spells.IceBall ? B_YN);
'   cp->Spells.Paralyze = YesNo(42,16, cp->Spells.Paralyze ? B_YN);
'   cp->Spells.Superman = YesNo(42,17, cp->Spells.Superman ? B_YN);
'   cp->Spells.Deepfreeze = YesNo(42,18, cp->Spells.Deepfreeze ? B_YN);
'   cp->Spells.DrainingCloud = YesNo(42,19, cp->Spells.DrainingCloud ? B_YN);
'   cp->Spells.CureDisease = YesNo(42,20, cp->Spells.CureDisease ? B_YN);
'
'   cp->Spells.Poison = YesNo(61, 6, cp->Spells.Poison ? B_YN);
'   cp->Spells.MissileShield = YesNo(61, 7, cp->Spells.MissileShield ? B_YN);
'   cp->Spells.ShrillSound = YesNo(61, 8, cp->Spells.ShrillSound ? B_YN);
'   cp->Spells.StinkBomb = YesNo(61, 9, cp->Spells.StinkBomb ? B_YN);
'   cp->Spells.AirPocket = YesNo(61,10, cp->Spells.AirPocket ? B_YN);
'   cp->Spells.Silence = YesNo(61,11, cp->Spells.Silence ? B_YN);
'   cp->Spells.PoisonGas = YesNo(61,12, cp->Spells.PoisonGas ? B_YN);
'   cp->Spells.CurePoison = YesNo(61,13, cp->Spells.CurePoison ? B_YN);
'   cp->Spells.Whirlwind = YesNo(61,14, cp->Spells.Whirlwind ? B_YN);
'   cp->Spells.PurifyAir = YesNo(61,15, cp->Spells.PurifyAir ? B_YN);
'   cp->Spells.DeadlyPoison = YesNo(61,16, cp->Spells.DeadlyPoison ? B_YN);
'   cp->Spells.Levitate = YesNo(61,17, cp->Spells.Levitate ? B_YN);
'   cp->Spells.ToxicVapors = YesNo(61,18, cp->Spells.ToxicVapors ? B_YN);
'   cp->Spells.NoxiousFumes = YesNo(61,19, cp->Spells.NoxiousFumes ? B_YN);
'   cp->Spells.Asphyxiation = YesNo(61,20, cp->Spells.Asphyxiation ? B_YN);
'   cp->Spells.DeadlyAir = YesNo(61,21, cp->Spells.DeadlyAir ? B_YN);
'   cp->Spells.DeathCloud = YesNo(61,22, cp->Spells.DeathCloud ? B_YN);
'
'   if (!Query("continue", ""))
'      goto AGAIN;
'
'EXIT_LABEL:
'   printf("\n");
'   return;
'}
'
'void EditSpells2(void * CharacterBuffer)
'{
'   char           Temp[81];
'   struct Character *cp = (struct Character *)CharacterBuffer;
'
'AGAIN:
'    clrscr();
'   header("Edit Spells 2");
'
'   gotoxy(1, 3);  clreol();
'    textattr(displaycolor + (BLACK << 4));  lowvideo();
'    cprintf("Character: ");                highvideo();
'    cprintf("%s", cp->Name);
'
'    textattr(screencolor + (BLACK << 4));
'    highvideo();
'
'   gotoxy(7, 5);cprintf("Earth              Mental             Divine");
'   lowvideo();
'
'   gotoxy(7, 6);cprintf("AcidSplash:        MentalAttack:   _  HealWounds:     _");
'   gotoxy(7, 7);cprintf("ItchingSkin:       Sleep:          _  MakeWounds:     _");
'   gotoxy(7, 8);cprintf("ArmorShield:       Bless:          _  MagicMissile:   _");
'   gotoxy(7, 9);cprintf("Direction:         Charm:          _  DispellUndead:  _");
'   gotoxy(7,10);cprintf("KnockKnock:        CureLesserCnd:  _  EnchantedBlade: _");
'   gotoxy(7,11);cprintf("Blades:            DivineTrap:     _  Blink:          _");
'   gotoxy(7,12);cprintf("Armorplate:        DetectSecret:   _  MagicScreen:    _");
'   gotoxy(7,13);cprintf("Web:               Identify:       _  Conjuration:    _");
'   gotoxy(7,14);cprintf("WhippingRocks:     Confusion:      _  AntiMagic:      _");
'   gotoxy(7,15);cprintf("AcidBomb:          Watchbells:     _  RemoveCurse:    _");
'   gotoxy(7,16);cprintf("Armormelt:         HoldMonsters:   _  Healfull:       _");
'   gotoxy(7,17);cprintf("Crush:             Mindread:       _  Lifesteal:      _");
'   gotoxy(7,18);cprintf("CreateLife:        SaneMind:       _  AstralGate:     _");
'   gotoxy(7,19);cprintf("CureStone:         PsionicBlast:   _  ZapUndead:      _");
'   gotoxy(7,20);cprintf("                   Illusion:       _  Recharge:       _");
'   gotoxy(7,21);cprintf("LocateObject:      WizardsEye:     _  WordOfDeath:    _");
'   gotoxy(7,22);cprintf("MindFlay:          Spooks:         _  Resurrection:   _");
'   gotoxy(7,23);cprintf("FindPerson:        Death:          _  DeathWish:      _");
'
'   textattr(displaycolor + (BLACK << 4));
'    lowvideo();
'
'   gotoxy(23, 6);cprintf("%c", cp->Spells.AcidSplash ? B_YN);
'   gotoxy(23, 7);cprintf("%c", cp->Spells.ItchingSkin ? B_YN);
'   gotoxy(23, 8);cprintf("%c", cp->Spells.ArmorShield ? B_YN);
'   gotoxy(23, 9);cprintf("%c", cp->Spells.Direction ? B_YN);
'   gotoxy(23,10);cprintf("%c", cp->Spells.KnockKnock ? B_YN);
'   gotoxy(23,11);cprintf("%c", cp->Spells.Blades ? B_YN);
'   gotoxy(23,12);cprintf("%c", cp->Spells.Armorplate ? B_YN);
'   gotoxy(23,13);cprintf("%c", cp->Spells.Web ? B_YN);
'   gotoxy(23,14);cprintf("%c", cp->Spells.WhippingRocks ? B_YN);
'   gotoxy(23,15);cprintf("%c", cp->Spells.AcidBomb ? B_YN);
'   gotoxy(23,16);cprintf("%c", cp->Spells.Armormelt ? B_YN);
'   gotoxy(23,17);cprintf("%c", cp->Spells.Crush ? B_YN);
'   gotoxy(23,18);cprintf("%c", cp->Spells.CreateLife ? B_YN);
'   gotoxy(23,19);cprintf("%c", cp->Spells.CureStone ? B_YN);
'   gotoxy(23,21);cprintf("%c", cp->Spells.LocateObject ? B_YN);
'   gotoxy(23,22);cprintf("%c", cp->Spells.MindFlay ? B_YN);
'   gotoxy(23,23);cprintf("%c", cp->Spells.FindPerson ? B_YN);
'
'   gotoxy(42, 6);cprintf("%c", cp->Spells.MentalAttack ? B_YN);
'   gotoxy(42, 7);cprintf("%c", cp->Spells.Sleep ? B_YN);
'   gotoxy(42, 8);cprintf("%c", cp->Spells.Bless ? B_YN);
'   gotoxy(42, 9);cprintf("%c", cp->Spells.Charm ? B_YN);
'   gotoxy(42,10);cprintf("%c", cp->Spells.CureLesserCnd ? B_YN);
'   gotoxy(42,11);cprintf("%c", cp->Spells.DivineTrap ? B_YN);
'   gotoxy(42,12);cprintf("%c", cp->Spells.DetectSecret ? B_YN);
'   gotoxy(42,13);cprintf("%c", cp->Spells.Identify ? B_YN);
'   gotoxy(42,14);cprintf("%c", cp->Spells.Confusion ? B_YN);
'   gotoxy(42,15);cprintf("%c", cp->Spells.Watchbells ? B_YN);
'   gotoxy(42,16);cprintf("%c", cp->Spells.HoldMonsters ? B_YN);
'   gotoxy(42,17);cprintf("%c", cp->Spells.Mindread ? B_YN);
'   gotoxy(42,18);cprintf("%c", cp->Spells.SaneMind ? B_YN);
'   gotoxy(42,19);cprintf("%c", cp->Spells.PsionicBlast ? B_YN);
'   gotoxy(42,20);cprintf("%c", cp->Spells.Illusion ? B_YN);
'   gotoxy(42,21);cprintf("%c", cp->Spells.WizardsEye ? B_YN);
'   gotoxy(42,22);cprintf("%c", cp->Spells.Spooks ? B_YN);
'   gotoxy(42,23);cprintf("%c", cp->Spells.Death ? B_YN);
'
'   gotoxy(61, 6);cprintf("%c", cp->Spells.HealWounds ? B_YN);
'   gotoxy(61, 7);cprintf("%c", cp->Spells.MakeWounds ? B_YN);
'   gotoxy(61, 8);cprintf("%c", cp->Spells.MagicMissile ? B_YN);
'   gotoxy(61, 9);cprintf("%c", cp->Spells.DispellUndead ? B_YN);
'   gotoxy(61,10);cprintf("%c", cp->Spells.EnchantedBlade ? B_YN);
'   gotoxy(61,11);cprintf("%c", cp->Spells.Blink ? B_YN);
'   gotoxy(61,12);cprintf("%c", cp->Spells.MagicScreen ? B_YN);
'   gotoxy(61,13);cprintf("%c", cp->Spells.Conjuration ? B_YN);
'   gotoxy(61,14);cprintf("%c", cp->Spells.AntiMagic ? B_YN);
'   gotoxy(61,15);cprintf("%c", cp->Spells.RemoveCurse ? B_YN);
'   gotoxy(61,16);cprintf("%c", cp->Spells.Healfull ? B_YN);
'   gotoxy(61,17);cprintf("%c", cp->Spells.Lifesteal ? B_YN);
'   gotoxy(61,18);cprintf("%c", cp->Spells.AstralGate ? B_YN);
'   gotoxy(61,19);cprintf("%c", cp->Spells.ZapUndead ? B_YN);
'   gotoxy(61,20);cprintf("%c", cp->Spells.Recharge ? B_YN);
'   gotoxy(61,21);cprintf("%c", cp->Spells.WordOfDeath ? B_YN);
'   gotoxy(61,22);cprintf("%c", cp->Spells.Resurrection ? B_YN);
'   gotoxy(61,23);cprintf("%c", cp->Spells.DeathWish ? B_YN);
'
'   cp->Spells.AcidSplash = YesNo(23, 6, cp->Spells.AcidSplash ? B_YN);
'   cp->Spells.ItchingSkin = YesNo(23, 7, cp->Spells.ItchingSkin ? B_YN);
'   cp->Spells.ArmorShield = YesNo(23, 8, cp->Spells.ArmorShield ? B_YN);
'   cp->Spells.Direction = YesNo(23, 9, cp->Spells.Direction ? B_YN);
'   cp->Spells.KnockKnock = YesNo(23,10, cp->Spells.KnockKnock ? B_YN);
'   cp->Spells.Blades = YesNo(23,11, cp->Spells.Blades ? B_YN);
'   cp->Spells.Armorplate = YesNo(23,12, cp->Spells.Armorplate ? B_YN);
'   cp->Spells.Web = YesNo(23,13, cp->Spells.Web ? B_YN);
'   cp->Spells.WhippingRocks = YesNo(23,14, cp->Spells.WhippingRocks ? B_YN);
'   cp->Spells.AcidBomb = YesNo(23,15, cp->Spells.AcidBomb ? B_YN);
'   cp->Spells.Armormelt = YesNo(23,16, cp->Spells.Armormelt ? B_YN);
'   cp->Spells.Crush = YesNo(23,17, cp->Spells.Crush ? B_YN);
'   cp->Spells.CreateLife = YesNo(23,18, cp->Spells.CreateLife ? B_YN);
'   cp->Spells.CureStone = YesNo(23,19, cp->Spells.CureStone ? B_YN);
'   cp->Spells.LocateObject = YesNo(23,21, cp->Spells.LocateObject ? B_YN);
'   cp->Spells.MindFlay = YesNo(23,22, cp->Spells.MindFlay ? B_YN);
'   cp->Spells.FindPerson = YesNo(23,23, cp->Spells.FindPerson ? B_YN);
'
'   cp->Spells.MentalAttack = YesNo(42, 6, cp->Spells.MentalAttack ? B_YN);
'   cp->Spells.Sleep = YesNo(42, 7, cp->Spells.Sleep ? B_YN);
'   cp->Spells.Bless = YesNo(42, 8, cp->Spells.Bless ? B_YN);
'   cp->Spells.Charm = YesNo(42, 9, cp->Spells.Charm ? B_YN);
'   cp->Spells.CureLesserCnd = YesNo(42,10, cp->Spells.CureLesserCnd ? B_YN);
'   cp->Spells.DivineTrap = YesNo(42,11, cp->Spells.DivineTrap ? B_YN);
'   cp->Spells.DetectSecret = YesNo(42,12, cp->Spells.DetectSecret ? B_YN);
'   cp->Spells.Identify = YesNo(42,13, cp->Spells.Identify ? B_YN);
'   cp->Spells.Confusion = YesNo(42,14, cp->Spells.Confusion ? B_YN);
'   cp->Spells.Watchbells = YesNo(42,15, cp->Spells.Watchbells ? B_YN);
'   cp->Spells.HoldMonsters = YesNo(42,16, cp->Spells.HoldMonsters ? B_YN);
'   cp->Spells.Mindread = YesNo(42,17, cp->Spells.Mindread ? B_YN);
'   cp->Spells.SaneMind = YesNo(42,18, cp->Spells.SaneMind ? B_YN);
'   cp->Spells.PsionicBlast = YesNo(42,19, cp->Spells.PsionicBlast ? B_YN);
'   cp->Spells.Illusion = YesNo(42,20, cp->Spells.Illusion ? B_YN);
'   cp->Spells.WizardsEye = YesNo(42,21, cp->Spells.WizardsEye ? B_YN);
'   cp->Spells.Spooks = YesNo(42,22, cp->Spells.Spooks ? B_YN);
'   cp->Spells.Death = YesNo(42,23, cp->Spells.Death ? B_YN);
'
'   cp->Spells.HealWounds = YesNo(61, 6, cp->Spells.HealWounds ? B_YN);
'   cp->Spells.MakeWounds = YesNo(61, 7, cp->Spells.MakeWounds ? B_YN);
'   cp->Spells.MagicMissile = YesNo(61, 8, cp->Spells.MagicMissile ? B_YN);
'   cp->Spells.DispellUndead = YesNo(61, 9, cp->Spells.DispellUndead ? B_YN);
'   cp->Spells.EnchantedBlade = YesNo(61,10, cp->Spells.EnchantedBlade ? B_YN);
'   cp->Spells.Blink = YesNo(61,11, cp->Spells.Blink ? B_YN);
'   cp->Spells.MagicScreen = YesNo(61,12, cp->Spells.MagicScreen ? B_YN);
'   cp->Spells.Conjuration = YesNo(61,13, cp->Spells.Conjuration ? B_YN);
'   cp->Spells.AntiMagic = YesNo(61,14, cp->Spells.AntiMagic ? B_YN);
'   cp->Spells.RemoveCurse = YesNo(61,15, cp->Spells.RemoveCurse ? B_YN);
'   cp->Spells.Healfull = YesNo(61,16, cp->Spells.Healfull ? B_YN);
'   cp->Spells.Lifesteal = YesNo(61,17, cp->Spells.Lifesteal ? B_YN);
'   cp->Spells.AstralGate = YesNo(61,18, cp->Spells.AstralGate ? B_YN);
'   cp->Spells.ZapUndead = YesNo(61,19, cp->Spells.ZapUndead ? B_YN);
'   cp->Spells.Recharge = YesNo(61,20, cp->Spells.Recharge ? B_YN);
'   cp->Spells.WordOfDeath = YesNo(61,21, cp->Spells.WordOfDeath ? B_YN);
'   cp->Spells.Resurrection = YesNo(61,22, cp->Spells.Resurrection ? B_YN);
'   cp->Spells.DeathWish = YesNo(61,23, cp->Spells.DeathWish ? B_YN);
'
'   if (!Query("continue", ""))
'      goto AGAIN;
'
'EXIT_LABEL:
'   printf("\n");
'   return;
'}
'
'#define GENDd_XY    13,5
'#define GEND_XY     17,5
'#define RACEd_XY    13,6
'#define RACE_XY     17,6
'#define PROFd_XY    13,7
'#define PROF_XY     17,7
'#define LEVEL_XY    36,7
'#define CONDd_XY    13,8
'#define COND_XY     17,8
'#define LIVES_XY    36,8
'#define ALIVE_XY    52,8
'
'#define STR_XY      11,11
'#define INT_XY      11,12
'#define PIE_XY      11,13
'#define VIT_XY      11,14
'#define DEX_XY      11,15
'#define SPD_XY      11,16
'#define PER_XY      11,17
'#define KAR_XY      11,18
'#define EXP_XY      31,11
'#define MKS_XY      31,12
'#define GP_XY       31,13
'#define HP_XY       31,15
'#define STA_XY      31,16
'#define CC_XY       31,17
'#define FIRE_XY     52,11
'#define WATER_XY    52,12
'#define AIR_XY      52,13
'#define EARTH_XY    64,11
'#define MENTAL_XY   64,12
'#define DIVINE_XY   64,13
'#define SKILLS_XY   60,16
'#define ITEMS_XY    60,17
'#define SWAG_XY     60,18
'#define SPELLS_XY   60,19
'
'void EditCharacter(void * CharacterBuffer)
'{
'   short int      WriteChanges = FALSE,
'                  DoSkills, DoItems, DoSwagBag, DoSpells;
'   char           Temp[81];
'   struct Character *cp = (struct Character *)CharacterBuffer;
'    SCREEN_BUFFER   MainEditMenuPage;
'
'AGAIN:
'    clrscr();
'   header("Edit Character");
'
'   gotoxy(1, 3);  clreol();
'    textattr(displaycolor + (BLACK << 4));  lowvideo();
'    cprintf(" Character: ");                highvideo();
'    cprintf("%s", cp->Name);
'    gettext(1, 1, 80, 24, MainEditMenuPage);
'
'    textattr(screencolor + (BLACK << 4));
'   lowvideo();
'   gotoxy(1, 5); cprintf("    Gender:                 ");
'   gotoxy(1, 6); cprintf("      Race:                 ");
'   gotoxy(1, 7); cprintf("Profession:                 Level:      ");
'   gotoxy(1, 8); cprintf(" Condition:                 Life:         ?Alive?:   ");
'
'    highvideo();
'    gotoxy(4, 10);  cprintf("Attributes                                   Spell Points");
'    gotoxy(49, 15); cprintf("More Edit Stuff...");
'    lowvideo();
'
'    gotoxy(7, 11);  cprintf("STR       Experience                   Fire       Earth     ");
'    gotoxy(7, 12);  cprintf("INT       Monster Kills                Water      Mental    ");
'    gotoxy(7, 13);  cprintf("PIE       Gold Pieces                  Air        Divine    ");
'    gotoxy(7, 14);  cprintf("VIT       ");
'    gotoxy(7, 15); cprintf("DEX          Hit Points");
'    gotoxy(7, 16); cprintf("SPD          Stamina                         Skills   ");
'    gotoxy(7, 17); cprintf("PER          Capacity                        Items    ");
'    gotoxy(7, 18); cprintf("KAR                                          SwagBag  ");
'    gotoxy(7, 19); cprintf("                                             Spells   ");
'
'   textattr(displaycolor + (BLACK << 4));
'    lowvideo();
'
'   gotoxy(GENDd_XY); cprintf("%u", cp->Gender);
'   gotoxy(GEND_XY);  cprintf("%-6.6s", cp->Gender ? "Female" : "Male");
'   gotoxy(RACEd_XY); cprintf("%u", cp->Race);
'   gotoxy(RACE_XY);  cprintf("%-9.9s", RaceMap[cp->Race]);
'   gotoxy(PROFd_XY); cprintf("%u", cp->Profession);
'   gotoxy(PROF_XY);  cprintf("%-9.9s", ProfessionMap[cp->Profession]);
'   gotoxy(LEVEL_XY); cprintf("%d", cp->Level);
'   gotoxy(CONDd_XY); cprintf("%u", cp->ConditionCode);
'   gotoxy(COND_XY);  cprintf("%-9.9s", ConditionMap[cp->ConditionCode]);
'   gotoxy(LIVES_XY); cprintf("%d", cp->Lives);
'   gotoxy(ALIVE_XY); cprintf("%u", cp->Alive);
'
'   gotoxy(STR_XY);   cprintf("%u", cp->Attributes.STR);
'   gotoxy(INT_XY);   cprintf("%u", cp->Attributes.INT);
'   gotoxy(PIE_XY);   cprintf("%u", cp->Attributes.PIE);
'   gotoxy(VIT_XY);   cprintf("%u", cp->Attributes.VIT);
'   gotoxy(DEX_XY);   cprintf("%u", cp->Attributes.DEX);
'   gotoxy(SPD_XY);   cprintf("%u", cp->Attributes.SPD);
'   gotoxy(PER_XY);   cprintf("%u", cp->Attributes.PER);
'   gotoxy(KAR_XY);   cprintf("%u", cp->Attributes.KAR);
'   gotoxy(EXP_XY);   cprintf("%ld", cp->EXP);
'   gotoxy(MKS_XY);   cprintf("%ld", cp->MKS);
'   gotoxy(GP_XY);    cprintf("%ld", cp->GP);
'   gotoxy(HP_XY);    cprintf("%d", cp->HP.Maximum);
'   gotoxy(STA_XY);   cprintf("%d", cp->STA.Maximum);
'   gotoxy(CC_XY);    cprintf("%d", cp->CC.Maximum);
'   gotoxy(FIRE_XY);  cprintf("%d", cp->SpellPoints.Fire.Maximum);
'   gotoxy(WATER_XY); cprintf("%d", cp->SpellPoints.Water.Maximum);
'   gotoxy(AIR_XY);   cprintf("%d", cp->SpellPoints.Air.Maximum);
'   gotoxy(EARTH_XY); cprintf("%d", cp->SpellPoints.Earth.Maximum);
'   gotoxy(MENTAL_XY);cprintf("%d", cp->SpellPoints.Mental.Maximum);
'   gotoxy(DIVINE_XY);cprintf("%d", cp->SpellPoints.Divine.Maximum);
'   gotoxy(SKILLS_XY);  cprintf("N");
'   gotoxy(ITEMS_XY);   cprintf("N");
'   gotoxy(SWAG_XY);    cprintf("N");
'   gotoxy(SPELLS_XY);  cprintf("N");
'
'   cp->Gender = GetNewI1(GENDd_XY, 1, cp->Gender, 1);
'   gotoxy(GEND_XY);  cprintf("%-6.6s", cp->Gender ? "Female" : "Male");
'   cp->Race = GetNewI1(RACEd_XY, 2, cp->Race, RaceMapMax);
'   gotoxy(RACE_XY);  cprintf("%-9.9s", RaceMap[cp->Race]);
'   cp->Profession = GetNewI1(PROFd_XY, 2, cp->Profession, ProfessionMapMax);
'   gotoxy(PROF_XY);  cprintf("%-9.9s", ProfessionMap[cp->Profession]);
'   cp->ConditionCode = GetNewI1(CONDd_XY, 2, cp->ConditionCode, ConditionMapMax);
'   gotoxy(COND_XY);  cprintf("%-9.9s", ConditionMap[cp->ConditionCode]);
'   cp->Level = GetNewI2(LEVEL_XY, 3, cp->Level, 999);
'   cp->Lives = GetNewI2(LIVES_XY, 3, cp->Lives, 999);
'   cp->Alive = GetNewI1(ALIVE_XY, 2, cp->Alive, MAX_I1);
'
'   cp->Attributes.STR = GetNewI1(STR_XY, 2, cp->Attributes.STR, 99);
'   cp->Attributes.INT = GetNewI1(INT_XY, 2, cp->Attributes.INT, 99);
'   cp->Attributes.PIE = GetNewI1(PIE_XY, 2, cp->Attributes.PIE, 99);
'   cp->Attributes.VIT = GetNewI1(VIT_XY, 2, cp->Attributes.VIT, 99);
'   cp->Attributes.DEX = GetNewI1(DEX_XY, 2, cp->Attributes.DEX, 99);
'   cp->Attributes.SPD = GetNewI1(SPD_XY, 2, cp->Attributes.SPD, 99);
'   cp->Attributes.PER = GetNewI1(PER_XY, 2, cp->Attributes.PER, 99);
'   cp->Attributes.KAR = GetNewI1(KAR_XY, 2, cp->Attributes.KAR, 99);
'   cp->EXP = GetNewI4(EXP_XY, 10, cp->EXP, MAX_I4);
'   cp->MKS = GetNewI4(MKS_XY, 10, cp->MKS, MAX_I4);
'   cp->GP =  GetNewI4(GP_XY, 10, cp->GP, (long int)(MAX_I4/6));
'   cp->HP.Maximum =  GetNewI2(HP_XY, 4, cp->HP.Maximum, 9999);
'   cp->STA.Maximum = GetNewI2(STA_XY, 4, cp->STA.Maximum, 9999);
'   cp->CC.Maximum =  GetNewI2(CC_XY, 4, cp->CC.Maximum, 9999);
'   cp->SpellPoints.Fire.Maximum = GetNewI2(FIRE_XY, 3,
'                                cp->SpellPoints.Fire.Maximum, 999);
'   cp->SpellPoints.Water.Maximum = GetNewI2(WATER_XY, 3,
'                                 cp->SpellPoints.Water.Maximum, 999);
'   cp->SpellPoints.Air.Maximum = GetNewI2(AIR_XY, 3,
'                               cp->SpellPoints.Air.Maximum, 999);
'   cp->SpellPoints.Mental.Maximum = GetNewI2(EARTH_XY, 3,
'                                  cp->SpellPoints.Earth.Maximum, 999);
'   cp->SpellPoints.Mental.Maximum = GetNewI2(MENTAL_XY, 3,
'                                  cp->SpellPoints.Mental.Maximum, 999);
'   cp->SpellPoints.Divine.Maximum = GetNewI2(DIVINE_XY, 3,
'                                  cp->SpellPoints.Divine.Maximum, 999);
'    gettext(1, 1, 80, 24, MainEditMenuPage);
'
'   DoSkills = YesNo(SKILLS_XY, 'N');
'   DoItems  = YesNo(ITEMS_XY, 'N');
'   DoSwagBag= YesNo(SWAG_XY, 'N');
'   DoSpells = YesNo(SPELLS_XY, 'N');
'
'   if (!Query("continue", ""))
'      goto AGAIN;
'
'   if (DoSkills)
'   {
'      EditSkills(cp);
'      puttext(1, 1, 80, 24, MainEditMenuPage);
'   }
'   if (DoItems)
'   {
'      EditItems(cp);
'      puttext(1, 1, 80, 24, MainEditMenuPage);
'   }
'   if (DoSwagBag)
'   {
'      EditSwagBag(cp);
'      puttext(1, 1, 80, 24, MainEditMenuPage);
'   }
'   if (DoSpells)
'   {
'      EditSpells1(cp);
'      EditSpells2(cp);
'      puttext(1, 1, 80, 24, MainEditMenuPage);
'   }
'
'   WriteChanges = Query("write", "changes to disk");
'   if (WriteChanges)
'   {
'      if (!WriteDBS(FileName, Buffer, FileSize))
'         SystemError("Unable to write saved game to disk");
'    }
'
'EXIT_LABEL:
'   printf("\n");
'   return;
'}
'
'void CharacterMenu(void)
'{
'    int     choice = -1,
'                i = 0,
'            menu = 0;
'    char        c   =   ' ',
'            Temp[11];
'    SCREEN_BUFFER   CharMenuPage;
'
'   memset(Temp, NULL, 11);
'
'    /* Format the Main Menu Screen. */
'
'    clrscr();
'   header("Character Menu");
'
'    for (menu = 0, i = 1; menu < 6 && strcmp(Party[menu]->Name, ""); menu++, i++)
'   {
'        gotoxy(35, i + 10); clreol();
'        textattr(selectcolor + (BLACK << 4));
'        highvideo();    cprintf("%d)", i);
'        textattr(menucolor + (BLACK << 4));
'        lowvideo();     cprintf(" %s", Party[menu]->Name);
'    }
'
'    gettext(1, 1, 80, 24, CharMenuPage);
'
'    while (choice != 99)
'   {
'        gotoxy(1, Y_PROMPT);
'        textattr(promptcolor + (BLACK << 4));   highvideo();
'
'        cprintf("Edit Character: ");    lowvideo();
'        cprintf(" (ESC to quit ");
'
'        textattr(headercolor + (BLACK << 4));   highvideo();
'
'        cprintf("%s", Program.Name);    lowvideo();
'        cprintf("): ");               textattr(promptcolor + (BLACK << 4));
'
'        c = wgetch();
'      Temp[0] = c;
'        Switch ((choice = atoi(Temp)))
'      {
'            Case 0:
'              Switch (C)
'              {
'                  case RETURN:
'                  Case Escape:
'                  case 'Q':
'                  case 'q':
'                       choice = 99;
'                       break;
'default:
'                             cprintf("%c", BELL);
'                       continue;
'                       break;
'              };
'                  break;
'         Case 1:
'         Case 2:
'         Case 3:
'         Case 4:
'         Case 5:
'         Case 6:
'               gettext(1, 1, 80, 24, CharMenuPage);
'              EditCharacter(Party[choice-1]);
'               puttext(1, 1, 80, 24, CharMenuPage);
'              break;
'default:
'              cprintf("%c", BELL);
'              continue;
'              break;
'        }
'   }
'
'EXIT_LABEL:
'   free(Buffer);
'   return;
'};
'
'void MainMenu(Char * FileName, Char * CharacterName)
'{
'    int     i = 0;
'    char        *p = NULL;
'    SCREEN_BUFFER   MainMenuPage;
'
'    /* Format the Main Menu Screen. */
'
'    clrscr();
'   header("Saved Game");
'
'    gettext(1, 1, 80, 24, MainMenuPage);
'
'GetFileName:
'    While (!strcmp(FileName, ""))
'   {
'       puttext(1, 1, 80, 24, MainMenuPage);
'        gotoxy(1, Y_PROMPT-1);  clreol();
'        textattr(promptcolor + (BLACK << 4));   highvideo();
'
'        cprintf("Enter Wizardry Saved Game File Name"); lowvideo();
'
'        gotoxy(1, Y_PROMPT);    clreol();
'        cprintf("(ESC to quit ");
'
'        textattr(headercolor + (BLACK << 4));   highvideo();
'
'        cprintf("%s", Program.Name);    lowvideo();
'        cprintf("): ");
'
'        textattr(promptcolor + (BLACK << 4));
'
'        gotoxy(24, Y_PROMPT);
'        i = input(FileName, 40);
'      strupr(FileName);
'        Switch (i)
'      {
'            case RETURN:
'         case TAB:
'              if (!strcmp(FileName, ""))
'                 goto EXIT_LABEL;
'                  break;
'         Case Escape:
'              goto EXIT_LABEL;
'              break;
'default:
'              cprintf("%c", BELL);
'              continue;
'              break;
'        };
'
'        gotoxy(1, Y_PROMPT-1);  clreol();
'        gotoxy(1, Y_PROMPT);    clreol();
'    };
'
'   if (!ReadDBS(FileName, &Buffer, &FileSize))
'   {
'      strcpy(FileName, "");
'      goto GetFileName;
'   }
'   gotoxy(1, 3);  clreol();
'    textattr(displaycolor + (BLACK << 4));  lowvideo();
'    cprintf("Saved Game: ");                highvideo();
'    cprintf("%s", FileName);
'    gettext(1, 1, 80, 24, MainMenuPage);
'
'GetCharacterName:
'    While (!strcmp(CharacterName, ""))
'   {
'       puttext(1, 1, 80, 24, MainMenuPage);
'        gotoxy(1, Y_PROMPT);  clreol();
'        textattr(promptcolor + (BLACK << 4));   highvideo();
'        cprintf("Enter First Character Name");  lowvideo();
'        cprintf(" (ESC to quit ");
'        textattr(headercolor + (BLACK << 4));   highvideo();
'        cprintf("%s", Program.Name);    lowvideo();
'        textattr(promptcolor + (BLACK << 4));
'        cprintf("): ");
'
'        gotoxy(51, Y_PROMPT);
'        i = input(CharacterName, 7);
'      strupr(CharacterName);
'        Switch (i)
'      {
'            case RETURN:
'         case TAB:
'              if (!strcmp(CharacterName, ""))
'                 goto EXIT_LABEL;
'                  break;
'         Case Escape:
'              goto EXIT_LABEL;
'              break;
'default:
'              cprintf("%c", BELL);
'              continue;
'              break;
'        };
'
'        gotoxy(1, Y_PROMPT);    clreol();
'    };
'
'    gotoxy(1, Y_PROMPT);  clreol();
'    textattr(promptcolor + (BLACK << 4));   lowvideo();
'    cprintf("   Searching for \"");
'   highvideo();          cprintf("%s", CharacterName);
'   lowvideo();           cprintf("\"...");
'
'   for (p = Buffer; p < Buffer+FileSize && strcmp(p, CharacterName); p++);
'   if (!p || strcmp(p, CharacterName))
'   {
'      PrintError("Could not find \"%s\"", CharacterName);
'      strcpy(CharacterName, "");
'      goto GetCharacterName;
'   }
'    gotoxy(1, Y_PROMPT);  clreol();
'    textattr(promptcolor + (BLACK << 4));   highvideo();
'   cprintf("Found \"%s\" @ byte position %d (%d bytes)...", CharacterName, p-Buffer+1, sizeof(struct Character));
'   for (i = 0; i < 6 && p < Buffer+FileSize; i++, p += sizeof(struct Character))
'       Party[i] = (struct Character *)p;
'    lowvideo(); gotoxy(1, Y_PROMPT);    clreol();
'
'   While (True)
'   {
'       gettext(1, 1, 80, 24, MainMenuPage);
'      CharacterMenu();
'       puttext(1, 1, 80, 24, MainMenuPage);
'      if (!Query("edit", "more characters"))
'         break;
'      strcpy(FileName, "");
'      strcpy(CharacterName, "");
'      goto GetFileName;
'   }
'
'EXIT_LABEL:
'   return;
'};
'
'int main(int argc, char **argv)
'{
'   short i;
'   char Args[513],
'        Temp[33];
'
'   strcpy(FileName, "");
'   strcpy(CharacterName, "");
'   for (i = 0; i < 6; i++) Party[i] = NULL;
'
'   strcpy(Program.Name, "WIZEDIT");
'   strcpy(Program.Title, "Wizardry VII: Crusaders of the Dark Savant");
'   Program.Version = 1;
'   Program.Revision = 1;
'   Program.ScreenMode = TRUE;
'
'    textmode(C80);
'    window(1, 1, 80, 24);
'    clrscr();
'
'   strcpy(Args, strupr(argv[0]));
'   for (i = 1; i <= argc; i++)
'   {
'      sprintf(Temp, " %s", strupr(argv[i]));
'      strcat(Args, Temp);
'   }
'
'   Switch (argc)
'   {
'      Case 2:
'           strcpy(FileName, strupr(argv[1]));
'           break;
'      Case 3:
'           strcpy(FileName, strupr(argv[1]));
'           strcpy(CharacterName, strupr(argv[2]));
'           break;
'default:
'              break;
'   }
'
'   MainMenu(FileName, CharacterName);
'    normvideo();
'   clrscr();
'   exit(0);
'}

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
            
            txtHP.Text = Format(.HP.Maximum, "#,##0")
            'txtexp.Text = .EXP
            'txtmks.Text = .MKS
'            Debug.Print "Name:              " & vbTab & .Name
'            Debug.Print "Unknown Region #1 (4 bytes): "
'            Debug.Print strHex(.Unknown1, 4) & vbCrLf
'
'            Debug.Print vbCrLf & "Secondary Statistics..."
'            Debug.Print "Experience Points: " & vbTab & .EXP & vbTab & "0x" & Hex(.EXP)
'            Debug.Print "Monster Kills:     " & vbTab & .MKS & vbTab & "0x" & Hex(.MKS)
'            Debug.Print "Gold Pieces:       " & vbTab & .GP & vbTab & "0x" & Hex(.GP)
'            Debug.Print "Hit Points:        " & vbTab & strPoints(.HP)
'            Debug.Print "Stamina:           " & vbTab & strPoints(.STA)
'            Debug.Print "Carrying Capacity: " & vbTab & strPoints(.CC)
'            Debug.Print "Level:             " & vbTab & .Level
'            Debug.Print "Lives:             " & vbTab & .Lives
'
'            Debug.Print vbCrLf & "Spell Points..."
'            Debug.Print "Fire:              " & vbTab & strPoints(.FireSpellPoints)
'            Debug.Print "Water:             " & vbTab & strPoints(.WaterSpellPoints)
'            Debug.Print "Air:               " & vbTab & strPoints(.AirSpellPoints)
'            Debug.Print "Earth:             " & vbTab & strPoints(.EarthSpellPoints)
'            Debug.Print "Mental:            " & vbTab & strPoints(.MentalSpellPoints)
'            Debug.Print "Divine:            " & vbTab & strPoints(.DivineSpellPoints)
'
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
'            Debug.Print vbCrLf & "Basic Statistics..."
'            Debug.Print "Strength:          " & vbTab & .STR       '
'            Debug.Print "Intellegence:      " & vbTab & .INT
'            Debug.Print "Piety:             " & vbTab & .PIE
'            Debug.Print "Vitality:          " & vbTab & .VIT
'            Debug.Print "Dexterity:         " & vbTab & .DEX
'            Debug.Print "Speed:             " & vbTab & .SPD
'            Debug.Print "Personality:       " & vbTab & .PER
'            Debug.Print "Karma:             " & vbTab & .KAR
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
    
    Call PopulateCondition(cboCondition)
    Call PopulateGender(cboGender)
    Call PopulateProfession(cboProfession)
    Call PopulateRace(cboRace)
End Sub
Private Sub txtDEX_GotFocus()
    TextSelected
End Sub
Private Sub txtDEX_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtDEX_Validate(Cancel As Boolean)
    Cancel = ValidateAttribute()
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
    Cancel = ValidateAttribute()
End Sub
Private Sub txtKAR_GotFocus()
    TextSelected
End Sub
Private Sub txtKAR_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtKAR_Validate(Cancel As Boolean)
    Cancel = ValidateAttribute()
End Sub
Private Sub txtPER_GotFocus()
    TextSelected
End Sub
Private Sub txtPER_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPER_Validate(Cancel As Boolean)
    Cancel = ValidateAttribute()
End Sub
Private Sub txtPIE_GotFocus()
    TextSelected
End Sub
Private Sub txtPIE_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtPIE_Validate(Cancel As Boolean)
    Cancel = ValidateAttribute()
End Sub
Private Sub txtSPD_GotFocus()
    TextSelected
End Sub
Private Sub txtSPD_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtSPD_Validate(Cancel As Boolean)
    Cancel = ValidateAttribute()
End Sub
Private Sub txtSTR_GotFocus()
    TextSelected
End Sub
Private Sub txtSTR_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtSTR_Validate(Cancel As Boolean)
    Cancel = ValidateAttribute()
End Sub
Private Sub txtVIT_GotFocus()
    TextSelected
End Sub
Private Sub txtVIT_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkINumber(KeyAscii, True)
End Sub
Private Sub txtVIT_Validate(Cancel As Boolean)
    Cancel = ValidateAttribute()
End Sub
Private Sub udDEX_Change()
    Call ValidateAttribute(txtDEX)
End Sub
Private Sub udHP_Change()
    Call ValidateI2(txtHP)
End Sub
Private Sub udINT_Change()
    Call ValidateAttribute(txtINT)
End Sub
Private Sub udKAR_Change()
    Call ValidateAttribute(txtKAR)
End Sub
Private Sub udPER_Change()
    Call ValidateAttribute(txtPER)
End Sub
Private Sub udPIE_Change()
    Call ValidateAttribute(txtPIE)
End Sub
Private Sub udSPD_Change()
    Call ValidateAttribute(txtSPD)
End Sub
Private Sub udSTR_Change()
    Call ValidateAttribute(txtSTR)
End Sub
Private Sub udVIT_Change()
    Call ValidateAttribute(txtVIT)
End Sub
