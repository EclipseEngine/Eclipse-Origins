VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Panel"
   ClientHeight    =   8955
   ClientLeft      =   12540
   ClientTop       =   1650
   ClientWidth     =   5490
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   5490
   Begin VB.PictureBox picEditPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8970
      Left            =   0
      ScaleHeight     =   596
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   5505
      Begin VB.Frame fraEditPlayer 
         BackColor       =   &H00000000&
         Caption         =   "Edit Player"
         ForeColor       =   &H00FFFFFF&
         Height          =   8295
         Left            =   120
         TabIndex        =   81
         Top             =   480
         Visible         =   0   'False
         Width           =   5160
         Begin VB.ComboBox cmbAccess 
            Height          =   315
            ItemData        =   "frmAdmin.frx":0000
            Left            =   1200
            List            =   "frmAdmin.frx":0013
            TabIndex        =   118
            Text            =   "cmbAccess"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.ComboBox cmbDir 
            Height          =   315
            ItemData        =   "frmAdmin.frx":004F
            Left            =   3720
            List            =   "frmAdmin.frx":005F
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelEditPlayer 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   2760
            TabIndex        =   116
            Top             =   7920
            Width           =   2295
         End
         Begin VB.CommandButton cmdEditPlayerOk 
            Caption         =   "Save and Close"
            Height          =   255
            Left            =   2760
            TabIndex        =   115
            Top             =   7560
            Width           =   2295
         End
         Begin VB.TextBox txtY 
            Height          =   285
            Left            =   3720
            TabIndex        =   114
            Text            =   "0"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtX 
            Height          =   285
            Left            =   3720
            TabIndex        =   113
            Text            =   "0"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtMap 
            Height          =   285
            Left            =   3720
            TabIndex        =   112
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00000000&
            Caption         =   "Spell Editing"
            ForeColor       =   &H00FFFFFF&
            Height          =   2175
            Left            =   2760
            TabIndex        =   109
            Top             =   3960
            Width           =   2295
            Begin VB.ListBox lstSpells 
               Height          =   1230
               Left            =   120
               TabIndex        =   144
               Top             =   240
               Width           =   2055
            End
            Begin VB.ComboBox cmbSpells 
               Height          =   315
               ItemData        =   "frmAdmin.frx":007A
               Left            =   120
               List            =   "frmAdmin.frx":007C
               Style           =   2  'Dropdown List
               TabIndex        =   110
               Top             =   1800
               Width           =   2055
            End
            Begin VB.Label lblRandom 
               BackStyle       =   0  'Transparent
               Caption         =   "Set Spell:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   26
               Left            =   120
               TabIndex        =   111
               Top             =   1560
               Width           =   2055
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00000000&
            Caption         =   "Inventory Editing"
            ForeColor       =   &H00FFFFFF&
            Height          =   2175
            Left            =   2760
            TabIndex        =   102
            Top             =   1680
            Width           =   2295
            Begin VB.TextBox txtItemQuantity 
               Height          =   285
               Left            =   120
               TabIndex        =   105
               Text            =   "0"
               Top             =   1680
               Width           =   2055
            End
            Begin VB.ComboBox cmbItems 
               Height          =   315
               ItemData        =   "frmAdmin.frx":007E
               Left            =   120
               List            =   "frmAdmin.frx":0080
               Style           =   2  'Dropdown List
               TabIndex        =   104
               Top             =   1080
               Width           =   2055
            End
            Begin VB.ComboBox cmbInvSlot 
               Height          =   315
               ItemData        =   "frmAdmin.frx":0082
               Left            =   120
               List            =   "frmAdmin.frx":0084
               Style           =   2  'Dropdown List
               TabIndex        =   103
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label lblRandom 
               BackStyle       =   0  'Transparent
               Caption         =   "Set Quantity:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   24
               Left            =   120
               TabIndex        =   108
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label lblRandom 
               BackStyle       =   0  'Transparent
               Caption         =   "Set Item:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   23
               Left            =   120
               TabIndex        =   107
               Top             =   840
               Width           =   975
            End
            Begin VB.Label lblRandom 
               BackStyle       =   0  'Transparent
               Caption         =   "Slot:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   22
               Left            =   120
               TabIndex        =   106
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.ComboBox cmbShield 
            Height          =   315
            ItemData        =   "frmAdmin.frx":0086
            Left            =   1200
            List            =   "frmAdmin.frx":0088
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   7440
            Width           =   1335
         End
         Begin VB.ComboBox cmbHelmet 
            Height          =   315
            ItemData        =   "frmAdmin.frx":008A
            Left            =   1200
            List            =   "frmAdmin.frx":008C
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   7080
            Width           =   1335
         End
         Begin VB.ComboBox cmbArmor 
            Height          =   315
            ItemData        =   "frmAdmin.frx":008E
            Left            =   1200
            List            =   "frmAdmin.frx":0090
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   6720
            Width           =   1335
         End
         Begin VB.ComboBox cmbWeapon 
            Height          =   315
            ItemData        =   "frmAdmin.frx":0092
            Left            =   1200
            List            =   "frmAdmin.frx":0094
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   6360
            Width           =   1335
         End
         Begin VB.TextBox txtPoints 
            Height          =   285
            Left            =   1200
            TabIndex        =   97
            Text            =   "0"
            Top             =   6000
            Width           =   1335
         End
         Begin VB.TextBox txtWillPower 
            Height          =   285
            Left            =   1200
            TabIndex        =   96
            Text            =   "0"
            Top             =   5640
            Width           =   1335
         End
         Begin VB.TextBox txtAgility 
            Height          =   285
            Left            =   1200
            TabIndex        =   95
            Text            =   "0"
            Top             =   5280
            Width           =   1335
         End
         Begin VB.TextBox txtIntelligence 
            Height          =   285
            Left            =   1200
            TabIndex        =   94
            Text            =   "0"
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox txtEndurance 
            Height          =   285
            Left            =   1200
            TabIndex        =   93
            Text            =   "0"
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox txtStrength 
            Height          =   285
            Left            =   1200
            TabIndex        =   92
            Text            =   "0"
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox txtMP 
            Height          =   285
            Left            =   1200
            TabIndex        =   91
            Text            =   "0"
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox txtHP 
            Height          =   285
            Left            =   1200
            TabIndex        =   90
            Text            =   "0"
            Top             =   3480
            Width           =   1335
         End
         Begin VB.ComboBox cmbPK 
            Height          =   315
            ItemData        =   "frmAdmin.frx":0096
            Left            =   1200
            List            =   "frmAdmin.frx":00A0
            TabIndex        =   89
            Text            =   "cmbPK"
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txtExp 
            Height          =   285
            Left            =   1200
            TabIndex        =   88
            Text            =   "0"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtLevel 
            Height          =   285
            Left            =   1200
            TabIndex        =   87
            Text            =   "0"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.ComboBox cmbClass 
            Height          =   315
            ItemData        =   "frmAdmin.frx":00AD
            Left            =   1200
            List            =   "frmAdmin.frx":00AF
            TabIndex        =   86
            Text            =   "cmbClass"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.ComboBox cmbSex 
            Height          =   315
            ItemData        =   "frmAdmin.frx":00B1
            Left            =   1200
            List            =   "frmAdmin.frx":00BB
            TabIndex        =   85
            Text            =   "cmbSex"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtCharName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   84
            Text            =   "Character Name"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtPassword 
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   83
            Text            =   "Password"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtLogin 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   82
            Text            =   "Login Name"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Dir:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   30
            Left            =   2760
            TabIndex        =   143
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Y:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   29
            Left            =   2760
            TabIndex        =   142
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "X:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   28
            Left            =   2760
            TabIndex        =   141
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Map:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   25
            Left            =   2760
            TabIndex        =   140
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Shield:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   139
            Top             =   7440
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Helmet:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   138
            Top             =   7080
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Armor:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   137
            Top             =   6720
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Weapon:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   136
            Top             =   6360
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Stat Points:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   135
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Willpower:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   134
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Agility:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   133
            Top             =   5280
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Intelligence:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   132
            Top             =   4920
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Endurance:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   131
            Top             =   4560
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Strength:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   130
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "MP:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   129
            Top             =   3840
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "HP:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   128
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "PKer:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   127
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Access:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   126
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Exp:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   125
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Level:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   124
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Class:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   123
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   122
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Char Name:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   121
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   120
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraCharList 
         BackColor       =   &H00000000&
         Caption         =   "Choose a Character to Edit."
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   120
         TabIndex        =   76
         Top             =   480
         Visible         =   0   'False
         Width           =   5175
         Begin VB.CommandButton cmdSelChar 
            Caption         =   "Select Character"
            Height          =   255
            Left            =   360
            TabIndex        =   80
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CommandButton cmdCloseChars 
            Caption         =   "Close Character List"
            Height          =   255
            Left            =   2760
            TabIndex        =   79
            Top             =   2400
            Width           =   2055
         End
         Begin VB.ListBox lstChars 
            Height          =   2010
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Player Editor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   77
         Top             =   120
         Width           =   2145
      End
   End
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8970
      Left            =   0
      ScaleHeight     =   596
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   0
      Top             =   0
      Width           =   5505
      Begin VB.Frame Frame18 
         BackColor       =   &H00000000&
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   163
         Top             =   6120
         Width           =   2535
         Begin VB.CommandButton cmdAGameOpts 
            Caption         =   "Game Options"
            Height          =   255
            Left            =   120
            TabIndex        =   165
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton cmdAServer 
            Caption         =   "Server Options"
            Height          =   255
            Left            =   120
            TabIndex        =   164
            Top             =   720
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdCloseAPanel 
         Caption         =   "Close Admin Panel"
         Height          =   255
         Left            =   2760
         TabIndex        =   74
         Top             =   8640
         Width           =   2535
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00000000&
         Caption         =   "Extras"
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   2760
         TabIndex        =   67
         Top             =   6360
         Width           =   2535
         Begin VB.CommandButton cmdLevel 
            Caption         =   "Level Me Up"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   1800
            Width           =   2175
         End
         Begin VB.HScrollBar scrlAItem 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   70
            Top             =   480
            Value           =   1
            Width           =   2175
         End
         Begin VB.HScrollBar scrlAAmount 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   69
            Top             =   960
            Value           =   1
            Width           =   2175
         End
         Begin VB.CommandButton cmdASpawn 
            Caption         =   "Spawn Item"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00FFFFFF&
            X1              =   120
            X2              =   2280
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label lblAItem 
            BackStyle       =   0  'Transparent
            Caption         =   "Spawn Item: None"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblAAmount 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount: 1"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   720
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Show/Hide"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2760
         TabIndex        =   56
         Top             =   5640
         Width           =   2535
         Begin VB.CheckBox chkShowLoc 
            BackColor       =   &H00000000&
            Caption         =   "Location"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   59
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkShowFPS 
            BackColor       =   &H00000000&
            Caption         =   "FPS"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   58
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox chkShowPing 
            BackColor       =   &H00000000&
            Caption         =   "Ping"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Player List"
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   5175
         Begin MSComctlLib.ListView lstPlayers 
            Height          =   1935
            Index           =   3
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Visible         =   0   'False
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Index"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "IP Banned"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Account Banned"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Character (When Banned)"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Reason"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lstPlayers 
            Height          =   1935
            Index           =   2
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Visible         =   0   'False
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Account"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Character(s)"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Banned?"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lstPlayers 
            Height          =   1935
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Index"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Account"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Class"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Map"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Level"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.TabStrip tabLists 
            Height          =   2415
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   4260
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Online Players"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "All Accounts"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Bans"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Map Report (Dbl Click to Warp)"
         ForeColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   2760
         TabIndex        =   13
         Top             =   3240
         Width           =   2535
         Begin VB.ListBox lstMaps 
            Height          =   2010
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame fraEditors 
         BackColor       =   &H00000000&
         Caption         =   "Content Editors"
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   2535
         Begin VB.CommandButton cmdProjectiles 
            Caption         =   "Projectile"
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdDungeon 
            Caption         =   "Dungeon"
            Height          =   255
            Left            =   1320
            TabIndex        =   155
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton cmdAPets 
            Caption         =   "Pet"
            Height          =   255
            Left            =   120
            TabIndex        =   151
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton cmdAHouses 
            Caption         =   "Player House"
            Height          =   255
            Left            =   1320
            TabIndex        =   12
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton cmdAAnim 
            Caption         =   "Animation"
            Height          =   255
            Left            =   1320
            TabIndex        =   11
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton cmdASpell 
            Caption         =   "Spell"
            Height          =   255
            Left            =   1320
            TabIndex        =   10
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton cmdAShop 
            Caption         =   "Shop"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton cmdAResource 
            Caption         =   "Resource"
            Height          =   255
            Left            =   1320
            TabIndex        =   8
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmdANpc 
            Caption         =   "NPC"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmdAMap 
            Caption         =   "Map"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdAItem 
            Caption         =   "Item"
            Height          =   255
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdAQuest 
            Caption         =   "Quest"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton cmdAZones 
            Caption         =   "Zones"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   1680
            Width           =   1095
         End
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Panel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   120
         Width           =   2145
      End
   End
   Begin VB.PictureBox picGAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8970
      Left            =   0
      ScaleHeight     =   596
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   145
      Top             =   0
      Visible         =   0   'False
      Width           =   5505
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "Death Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   160
         Top             =   2880
         Width           =   5175
         Begin VB.CheckBox chkDisableExpLoss 
            BackColor       =   &H80000007&
            Caption         =   "Disable Exp Loss"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2160
            TabIndex        =   162
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox chkDisableItemLoss 
            BackColor       =   &H80000007&
            Caption         =   "Disable Item Loss"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   161
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00000000&
         Caption         =   "Main Menu Music"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   157
         Top             =   2160
         Width           =   5175
         Begin VB.TextBox txtMainMenuMusic 
            Height          =   285
            Left            =   2160
            TabIndex        =   158
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Filename (IE: Sound.mp3)"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00000000&
         Caption         =   "Maxes:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   152
         Top             =   1320
         Width           =   5175
         Begin VB.HScrollBar scrlMaxLevel 
            Height          =   255
            Left            =   1320
            Max             =   30000
            Min             =   1
            TabIndex        =   154
            Top             =   240
            Value           =   1
            Width           =   3735
         End
         Begin VB.Label lblMaxLevel 
            BackStyle       =   0  'Transparent
            Caption         =   "Max Level: 1"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   153
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdCloseGameOpts 
         Caption         =   "Close"
         Height          =   255
         Left            =   4200
         TabIndex        =   150
         Top             =   8640
         Width           =   1215
      End
      Begin VB.Frame Frame21 
         BackColor       =   &H00000000&
         Caption         =   "Backwards Compatability Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   147
         Top             =   600
         Width           =   5175
         Begin VB.CheckBox chkNewBatForumlas 
            BackColor       =   &H00000000&
            Caption         =   "Use New Combat Formulas."
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   149
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.CommandButton cmdSaveGameOpts 
         Caption         =   "Save and Close"
         Height          =   255
         Left            =   2640
         TabIndex        =   146
         Top             =   8640
         Width           =   1455
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Game Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   148
         Top             =   120
         Width           =   5505
      End
   End
   Begin VB.PictureBox picSAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8970
      Left            =   0
      ScaleHeight     =   596
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   5505
      Begin VB.Frame Frame12 
         BackColor       =   &H00000000&
         Caption         =   "Extra Option(s)"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   65
         Top             =   6600
         Width           =   5175
         Begin VB.CheckBox chkStaffOnly 
            BackColor       =   &H00000000&
            Caption         =   "Staff Only?"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   225
            Width           =   2655
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00000000&
         Caption         =   "Server Stats"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   60
         Top             =   7080
         Width           =   5175
         Begin VB.Label lblVersion 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Version: 4.0.1"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2280
            TabIndex        =   64
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblUpTime 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Up Time: 0s"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2280
            TabIndex        =   63
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label lblPlayersOnline 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Players Online: 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblAccounts 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Accounts: 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdHideServerPanel 
         Caption         =   "Hide Server Options"
         Height          =   255
         Left            =   3360
         TabIndex        =   55
         Top             =   8640
         Width           =   2055
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00000000&
         Caption         =   "Restart/Update Server (30 second shutdown timer.)"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   7920
         Width           =   5175
         Begin VB.CommandButton cmdRestartServer 
            Caption         =   "Restart/Update Server"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00000000&
         Caption         =   "Edit Update Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   2760
         TabIndex        =   46
         Top             =   4080
         Width           =   2535
         Begin VB.CommandButton cmdUpdateHelp 
            Caption         =   "Help!"
            Height          =   205
            Left            =   1800
            TabIndex        =   52
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtDataFolder 
            Height          =   285
            Left            =   120
            TabIndex        =   49
            Text            =   "txtDataFolder"
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox txtUpdateURL 
            Height          =   285
            Left            =   120
            TabIndex        =   48
            Text            =   "txtUpdateURL"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.CommandButton cmdSaveUpdateInfo 
            Caption         =   "Save Update Info"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label5 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Game's Data Folder:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Update.ini URL"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00000000&
         Caption         =   "Edit Message of the Day (Displayed in chatbox when players login.)"
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   42
         Top             =   2880
         Width           =   5175
         Begin VB.TextBox txtMOTD 
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Text            =   "txtMOTD"
            Top             =   480
            Width           =   4935
         End
         Begin VB.CommandButton cmdSaveMOTD 
            Caption         =   "Save MOTD"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   4935
         End
         Begin VB.Label Label4 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Message of the day:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00000000&
         Caption         =   "News and Credits"
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   5175
         Begin VB.CommandButton cmdSaveCredits 
            Caption         =   "Save Credits"
            Height          =   195
            Left            =   2700
            TabIndex        =   36
            Top             =   1920
            Width           =   2415
         End
         Begin VB.TextBox txtCredits 
            Height          =   1575
            Left            =   2700
            MultiLine       =   -1  'True
            TabIndex        =   35
            Text            =   "frmAdmin.frx":00CD
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtNews 
            Height          =   1575
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   34
            Text            =   "frmAdmin.frx":00D5
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton cmdSaveNews 
            Caption         =   "Save News"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   2600
            X2              =   2600
            Y1              =   120
            Y2              =   2160
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         Caption         =   "Server Reload"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   22
         Top             =   5760
         Width           =   5175
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Height          =   255
            Left            =   3720
            TabIndex        =   32
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Height          =   255
            Left            =   3720
            TabIndex        =   31
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Height          =   255
            Left            =   2520
            TabIndex        =   30
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            Height          =   255
            Left            =   2520
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Height          =   255
            Left            =   1320
            TabIndex        =   28
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Height          =   255
            Left            =   1320
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Classes"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Edit Server Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   4080
         Width           =   2535
         Begin VB.CommandButton cmdSaveOpts 
            Caption         =   "Save Options"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox txtGameWebsite 
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Text            =   "txtGameWebsite"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txtGameName 
            Height          =   285
            Left            =   120
            TabIndex        =   38
            Text            =   "txtGameName"
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Game Website:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Game Name:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Server Admin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   24
         Top             =   120
         Width           =   5505
      End
   End
   Begin VB.Menu menuBan 
      Caption         =   "Ban Menu"
      Visible         =   0   'False
      Begin VB.Menu mUnbanIPAccount 
         Caption         =   "Unban Account/IP"
      End
   End
   Begin VB.Menu AccountPopup 
      Caption         =   "Account Editor"
      Visible         =   0   'False
      Begin VB.Menu mEditAccountChar 
         Caption         =   "Edit Characters"
      End
      Begin VB.Menu mBanAccount 
         Caption         =   "Ban Account/IP"
      End
      Begin VB.Menu mDeleteAccount 
         Caption         =   "Delete Account"
      End
   End
   Begin VB.Menu MenuPopup 
      Caption         =   "Player Editor"
      Visible         =   0   'False
      Begin VB.Menu MEditPlayer 
         Caption         =   "Edit Player"
         Index           =   1
      End
      Begin VB.Menu MKickPlayer 
         Caption         =   "Kick Player"
      End
      Begin VB.Menu MBanPlayer 
         Caption         =   "Ban Player"
      End
      Begin VB.Menu MWarpMe2 
         Caption         =   "Warp Me To"
      End
      Begin VB.Menu MWarp2Me 
         Caption         =   "Warp To Me"
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ****************
' ** Admin Menu **
' ****************
Public EditingPlayer As String
Public EditingAccount As String
Public EditingBan As Long

Private Sub chkShowFPS_Click()


   On Error GoTo errorhandler
    If chkShowFPS.Value = 1 Then
        BFPS = True
    Else
        BFPS = False
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkShowFPS_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkShowLoc_Click()


   On Error GoTo errorhandler
    If chkShowLoc.Value = 1 Then
        BLoc = True
    Else
        BLoc = False
    End If
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkShowLoc_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkShowPing_Click()


   On Error GoTo errorhandler
    If chkShowPing.Value = 1 Then
        BPing = True
    Else
        BPing = False
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkShowPing_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkStaffOnly_Click()
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 14
    If chkStaffOnly.Value = 1 Then
        buffer.WriteLong 1
    Else
        buffer.WriteLong 0
    End If
    SendData buffer.ToArray
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkStaffOnly_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdALoc_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
            Exit Sub
    End If
    BLoc = Not BLoc




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdALoc_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmbInvSlot_Click()
    

   On Error GoTo errorhandler
    
    cmbItems.ListIndex = TempPlayerInv(cmbInvSlot.ListIndex + 1).Num
    txtItemQuantity.Text = CStr(TempPlayerInv(cmbInvSlot.ListIndex + 1).Value)

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbInvSlot_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbItems_Click()


   On Error GoTo errorhandler
    If cmbInvSlot.ListIndex > -1 Then
        If cmbItems.ListIndex > -1 Then
            TempPlayerInv(cmbInvSlot.ListIndex + 1).Num = cmbItems.ListIndex
            InitPlayerItemEditor
        End If
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbItems_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSpells_Click()


   On Error GoTo errorhandler
    If lstSpells.ListIndex > -1 Then
        If cmbSpells.ListIndex > -1 Then
            TempPlayerSpells(lstSpells.ListIndex + 1) = cmbSpells.ListIndex
            InitPlayerSpellEditor
        End If
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSpells_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdAGameOpts_Click()
Dim buffer As clsBuffer

   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CGameOpts
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAGameOpts_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdAHouses_Click()


   On Error GoTo errorhandler
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
            Exit Sub
    End If
    SendRequestEditHouse
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAHouses_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdAMap_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
            Exit Sub
    End If
    SendRequestEditMap




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAMap_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAPets_Click()


   On Error GoTo errorhandler
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then

    Exit Sub
    End If

    SendRequestEditPet
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAPets_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdAQuest_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    SendRequestEditQuest




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAQuest_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAWarp2Me_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
            Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.Text)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAWarp2Me_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAWarpMe2_Click()


   On Error GoTo errorhandler





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAWarpMe2_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAServer_Click()
Dim buffer As clsBuffer

   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CServerOpts
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAServer_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long



   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
            Exit Sub
    End If

    If Len(Trim$(txtAMap.Text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.Text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.Text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAWarp_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdASprite_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
            Exit Sub
    End If

    If Len(Trim$(txtASprite.Text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtASprite.Text)) Then
        Exit Sub
    End If

    SendSetSprite CLng(Trim$(txtASprite.Text))




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdASprite_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAMapReport_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
            Exit Sub
    End If

    SendMapReport




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAMapReport_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdARespawn_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
            Exit Sub
    End If
    SendMapRespawn




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdARespawn_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAItem_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
            Exit Sub
    End If

    SendRequestEditItem




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAItem_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdANpc_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
            Exit Sub
    End If

    SendRequestEditNpc




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdANpc_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAResource_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
            Exit Sub
    End If

    SendRequestEditResource




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAResource_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAShop_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
            Exit Sub
    End If

    SendRequestEditShop




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAShop_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdASpell_Click()

   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
            Exit Sub
    End If

    SendRequestEditSpell


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdASpell_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdAAccess_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
            Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Or Not IsNumeric(Trim$(txtAAccess.Text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.Text), CLng(Trim$(txtAAccess.Text))




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAAccess_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdASpawn_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
            Exit Sub
    End If
    SendSpawnItem scrlAItem.Value, scrlAAmount.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdASpawn_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAAnim_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
            Exit Sub
    End If

    SendRequestEditAnimation




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAAnim_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAZones_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
            Exit Sub
    End If
    SendRequestEditZone




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAZones_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Private Sub cmdCancelEditPlayer_Click()


   On Error GoTo errorhandler

    frmAdmin.fraEditPlayer.Visible = False
    frmAdmin.fraCharList.Visible = False
    frmAdmin.picEditPlayer.Visible = False


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancelEditPlayer_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseAPanel_Click()


   On Error GoTo errorhandler
    frmAdmin.Visible = False
    frmAdmin.picSAdmin.Visible = False
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseAPanel_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseChars_Click()


   On Error GoTo errorhandler
    picEditPlayer.Visible = False
    fraCharList.Visible = False
    lstChars.Clear
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseChars_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdDungeon_Click()

   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    'removed


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdDungeon_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdEditPlayerOk_Click()
Dim buffer As clsBuffer, i As Long

   On Error GoTo errorhandler
        If IsNumeric(txtLevel.Text) = False Then MsgBox ("Level must be a number!"): Exit Sub
        If IsNumeric(txtExp.Text) = False Then MsgBox ("Exp must be a number!"): Exit Sub
        If IsNumeric(txtHP.Text) = False Then MsgBox ("HP must be a number!"): Exit Sub
        If IsNumeric(txtMP.Text) = False Then MsgBox ("MP must be a number!"): Exit Sub
        If IsNumeric(txtLevel.Text) = False Then MsgBox ("Level must be a number!"): Exit Sub
        If IsNumeric(txtStrength.Text) = False Then MsgBox ("Strength must be a number!"): Exit Sub
        If IsNumeric(txtEndurance.Text) = False Then MsgBox ("Endurance must be a number!"): Exit Sub
        If IsNumeric(txtIntelligence.Text) = False Then MsgBox ("Intelligence must be a number!"): Exit Sub
        If IsNumeric(txtAgility.Text) = False Then MsgBox ("Agility must be a number!"): Exit Sub
        If IsNumeric(txtWillPower.Text) = False Then MsgBox ("Willpower must be a number!"): Exit Sub
        If IsNumeric(txtPoints.Text) = False Then MsgBox ("Stat Points must be a number!"): Exit Sub
        If IsNumeric(txtMap.Text) = False Then MsgBox ("Map must be a number!"): Exit Sub
        If IsNumeric(txtX.Text) = False Then MsgBox ("X must be a number!"): Exit Sub
        If IsNumeric(txtY.Text) = False Then MsgBox ("Y must be a number!"): Exit Sub
            'Check Everything First.
        If Val(txtLevel.Text) > MAX_LEVELS Or Val(txtLevel.Text) < 1 Then
            lblNotifications.Caption = "Player Saving Failed: Level is invalid."
            Exit Sub
        End If
        If Val(txtStrength.Text) > 255 Or Val(txtStrength.Text) < 0 Then
            lblNotifications.Caption = "Strength is invalid."
            Exit Sub
        End If
        If Val(txtEndurance.Text) > 255 Or Val(txtEndurance.Text) < 0 Then
            lblNotifications.Caption = "Endurance is invalid."
            Exit Sub
        End If
        If Val(txtIntelligence.Text) > 255 Or Val(txtIntelligence.Text) < 0 Then
            lblNotifications.Caption = "Intelligence is invalid."
            Exit Sub
        End If
        If Val(txtAgility.Text) > 255 Or Val(txtAgility.Text) < 0 Then
            MsgBox "Error, Agility is invalid."
            Exit Sub
        End If
        If Val(txtWillPower.Text) > 255 Or Val(txtWillPower.Text) < 0 Then
            MsgBox "Error, Willpower is invalid."
            Exit Sub
        End If
        If Val(txtMap.Text) <= 0 Or Val(txtMap.Text) > MAX_MAPS Then
            MsgBox "Map number is invalid."
            Exit Sub
        End If
        If Val(txtX.Text) < 0 Or Val(txtX.Text) > 255 Then
            MsgBox "X position is invalid."
            Exit Sub
        End If
        If Val(txtY.Text) < 0 Or Val(txtY.Text) > 255 Then
            MsgBox "Y position is invalid."
            Exit Sub
        End If
        
        Set buffer = New clsBuffer
        buffer.WriteLong CSavePlayer
        buffer.WriteString EditingAccount
        buffer.WriteString EditingPlayer
        
        'Changable Player Info Here
        buffer.WriteLong cmbSex.ListIndex
        buffer.WriteLong cmbClass.ListIndex + 1
        buffer.WriteLong Val(txtLevel.Text)
        buffer.WriteLong Val(txtExp.Text)
        buffer.WriteLong cmbAccess.ListIndex
        buffer.WriteLong cmbPK.ListIndex
        buffer.WriteLong Val(txtHP.Text)
        buffer.WriteLong Val(txtMP.Text)
        buffer.WriteLong Val(txtStrength.Text)
        buffer.WriteLong Val(txtEndurance.Text)
        buffer.WriteLong Val(txtIntelligence.Text)
        buffer.WriteLong Val(txtAgility.Text)
        buffer.WriteLong Val(txtWillPower.Text)
        buffer.WriteLong Val(txtPoints.Text)
        buffer.WriteLong cmbWeapon.ListIndex
        buffer.WriteLong cmbArmor.ListIndex
        buffer.WriteLong cmbHelmet.ListIndex
        buffer.WriteLong cmbShield.ListIndex
        buffer.WriteLong Val(txtMap.Text)
        buffer.WriteLong Val(txtX.Text)
        buffer.WriteLong Val(txtY.Text)
        buffer.WriteLong cmbDir.ListIndex
        
        For i = 1 To MAX_PLAYER_SPELLS
            buffer.WriteLong TempPlayerSpells(i)
        Next
        
        For i = 1 To MAX_INV
            buffer.WriteLong TempPlayerInv(i).Num
            buffer.WriteLong TempPlayerInv(i).Value
        Next
        
        'End Changable Player Info
        SendData buffer.ToArray
        Set buffer = Nothing
        
        frmAdmin.fraEditPlayer.Visible = False
        frmAdmin.fraCharList.Visible = False
        frmAdmin.picEditPlayer.Visible = False
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdEditPlayerOk_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdHideServerPanel_Click()


   On Error GoTo errorhandler
    picSAdmin.Visible = False
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdHideServerPanel_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdProjectiles_Click()
   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    SendRequestEditProjectiles

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdProjectiles_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdReloadAnimations_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 13
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadAnimations_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdReloadClasses_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 6
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadClasses_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdReloadItems_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 12
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadItems_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdReloadMaps_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 10
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadMaps_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdReloadNPCs_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 8
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadNPCs_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdReloadResources_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 9
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadResources_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdReloadShops_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 11
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadShops_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub CmdReloadSpells_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 7
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadSpells_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdRestartServer_Click()
    Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRestartServer
    SendData buffer.ToArray
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdRestartServer_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Private Sub cmdLevel_Click()


   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
            Exit Sub
    End If

    SendRequestLevelUp




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdSaveCredits_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 2
    buffer.WriteString txtCredits.Text
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveCredits_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveGameOpts_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveGameOpt
    If frmAdmin.chkNewBatForumlas.Value = 1 Then
        buffer.WriteLong 1
    Else
        buffer.WriteLong 2
    End If
    buffer.WriteLong frmAdmin.scrlMaxLevel.Value
    buffer.WriteString frmAdmin.txtMainMenuMusic.Text
    If frmAdmin.chkDisableItemLoss.Value = 1 Then
        buffer.WriteLong 1
    Else
        buffer.WriteLong 0
    End If
    If frmAdmin.chkDisableExpLoss.Value = 1 Then
        buffer.WriteLong 1
    Else
        buffer.WriteLong 0
    End If
    SendData buffer.ToArray
    Set buffer = Nothing
    
    frmAdmin.picGAdmin.Visible = False
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveGameOpts_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveMOTD_Click()
Dim buffer As clsBuffer

   On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 3
    buffer.WriteString txtMOTD.Text
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveMOTD_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveNews_Click()
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 1
    buffer.WriteString txtNews.Text
    SendData buffer.ToArray
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveNews_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveOpts_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
   
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 4
    buffer.WriteString txtGameName.Text
    buffer.WriteString txtGameWebsite.Text
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveOpts_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveUpdateInfo_Click()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveServerOpt
    buffer.WriteLong 5
    buffer.WriteString txtDataFolder.Text
    buffer.WriteString txtUpdateURL.Text
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveUpdateInfo_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSelChar_Click()
Dim buffer As clsBuffer

   On Error GoTo errorhandler
    If lstChars.ListIndex > -1 Then
        Set buffer = New clsBuffer
        buffer.WriteLong CEditPlayer
        buffer.WriteLong 2
        buffer.WriteString EditingAccount
        buffer.WriteLong lstChars.ListIndex + 1
        SendData buffer.ToArray
        frmAdmin.fraCharList.Visible = False
        fraEditPlayer.Visible = False
        picEditPlayer.Visible = False
        Set buffer = Nothing
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSelChar_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseGameOpts_Click()


   On Error GoTo errorhandler
    frmAdmin.picGAdmin.Visible = False
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseGameOpts_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim buffer As clsBuffer
   On Error GoTo errorhandler

Select Case KeyCode
        Case vbKeyInsert
            Me.Visible = False
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_Load()

   On Error GoTo errorhandler

    scrlAItem.max = MAX_ITEMS
    scrlAItem.Value = 1
    scrlAItem.min = 1
    
    If BFPS = True Then
        frmAdmin.chkShowFPS.Value = 1
    Else
        frmAdmin.chkShowFPS.Value = 0
    End If
    
    If BLoc = True Then
        frmAdmin.chkShowLoc.Value = 1
    Else
        frmAdmin.chkShowLoc.Value = 0
    End If
    
    If BPing = True Then
        frmAdmin.chkShowPing.Value = 1
    Else
        frmAdmin.chkShowPing.Value = 0
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lstMaps_DblClick()
   On Error GoTo errorhandler
    If lstMaps.ListIndex > -1 Then
        Call WarpTo(lstMaps.ListIndex + 1)
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstMaps_DblClick", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lstPlayers_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo errorhandler

    If Button = 2 Then
        If Index = 1 Then
            If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then Exit Sub
            If Trim$(lstPlayers(1).SelectedItem.SubItems(2)) <> "" Then
                EditingPlayer = Trim$(lstPlayers(1).SelectedItem.SubItems(2))
                Me.PopupMenu MenuPopup
            End If
        ElseIf Index = 2 Then
            If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
            If Trim$(lstPlayers(2).SelectedItem.SubItems(1)) <> "" Then
                EditingAccount = Trim$(lstPlayers(2).SelectedItem.SubItems(1))
                EditingPlayer = Trim$(lstPlayers(2).SelectedItem.SubItems(2))
                Me.PopupMenu AccountPopup
            End If
        ElseIf Index = 3 Then
            If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
            If Trim$(lstPlayers(3).SelectedItem.Text) <> "" Then
                EditingBan = Val(Trim$(lstPlayers(3).SelectedItem.Text))
                Me.PopupMenu menuBan
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstPlayers_MouseDown", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lstSpells_Click()


   On Error GoTo errorhandler
    If lstSpells.ListIndex > -1 Then
        cmbSpells.ListIndex = TempPlayerSpells(lstSpells.ListIndex + 1)
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstSpells_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub mBanAccount_Click()

   On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
            Exit Sub
    End If

    If Len(Trim$(EditingAccount)) < 1 Then
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to ban " & Trim$(EditingAccount) & "?", vbYesNo) = vbYes Then
        SendBan Trim$(EditingAccount), InputBox("Briefly describe why you are banning this account.", "Ban " & Trim$(EditingAccount)), True
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mBanAccount_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub MBanPlayer_Click()


   On Error GoTo errorhandler
   
    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
            Exit Sub
    End If

    If Len(Trim$(EditingPlayer)) < 1 Then
        Exit Sub
    End If
    If MsgBox("Are you sure you want to ban " & Trim$(EditingPlayer) & "?", vbYesNo) = vbYes Then
        SendBan Trim$(EditingPlayer), InputBox("Briefly describe why you are banning this player.", "Ban " & Trim$(EditingPlayer)), False
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MBanPlayer_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub mDeleteAccount_Click()


   On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
            Exit Sub
    End If

    If Len(Trim$(EditingAccount)) < 1 Then
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete " & Trim$(EditingAccount) & "?", vbYesNo) = vbYes Then
        SendDelAccount Trim$(EditingAccount)
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mDeleteAccount_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub mEditAccountChar_Click()
Dim buffer As clsBuffer

   On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CEditPlayer
    buffer.WriteLong 1
    buffer.WriteString EditingAccount
    SendData buffer.ToArray
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mEditAccountChar_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub MEditPlayer_Click(Index As Integer)
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CEditPlayer
    buffer.WriteLong 0
    buffer.WriteString EditingPlayer
    SendData buffer.ToArray
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MEditPlayer_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub MKickPlayer_Click()


   On Error GoTo errorhandler
    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
            Exit Sub
    End If

    If Len(Trim$(EditingPlayer)) < 1 Then
        Exit Sub
    End If

    SendKick Trim$(EditingPlayer)
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MKickPlayer_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub mUnbanAccount_Click()

End Sub

Private Sub mUnbanIPAccount_Click()


   On Error GoTo errorhandler
    If MsgBox("Are you sure you want to unban this IP and/or Account?", vbYesNo) = vbYes Then
        SendBanDestroy EditingBan
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mUnbanIPAccount_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub MWarp2Me_Click()


   On Error GoTo errorhandler
    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
            Exit Sub
    End If

    If Len(Trim$(EditingPlayer)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(EditingPlayer)) Then
        Exit Sub
    End If

    WarpToMe Trim$(EditingPlayer)
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MWarp2Me_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub MWarpMe2_Click()


   On Error GoTo errorhandler
    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
            Exit Sub
    End If

    If Len(Trim$(EditingPlayer)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(EditingPlayer)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(EditingPlayer)

    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MWarpMe2_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlAAmount_Change()


   On Error GoTo errorhandler

    lblAAmount.Caption = "Amount: " & scrlAAmount.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAAmount_Change", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlAItem_Change()

   On Error GoTo errorhandler

    If scrlAItem.Value > MAX_ITEMS Then Exit Sub
    lblAItem.Caption = "Item: " & scrlAItem.Value & ". " & Trim$(Item(scrlAItem.Value).Name)
    If Item(scrlAItem.Value).Stackable = 1 Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    scrlAAmount.Enabled = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMaxLevel_Change()


   On Error GoTo errorhandler
    lblmaxlevel.Caption = "Max Level: " & scrlMaxLevel.Value
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMaxLevel_Change", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub tabLists_Click()


   On Error GoTo errorhandler
   
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        If tabLists.Tabs(1).Selected = False Then
            tabLists.Tabs(1).Selected = True
            lstPlayers(1).Visible = True
            Exit Sub
        End If
    End If

    For i = 1 To 3
        lstPlayers(i).Visible = False
    Next
    lstPlayers(tabLists.SelectedItem.Index).Visible = True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "tabLists_Click", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtItemQuantity_Validate(Cancel As Boolean)


   On Error GoTo errorhandler

    If cmbInvSlot.ListIndex > -1 Then
        If IsNumeric(txtItemQuantity.Text) Then
            If Val(txtItemQuantity.Text) >= 0 Then
                TempPlayerInv(cmbInvSlot.ListIndex + 1).Value = Val(txtItemQuantity.Text)
                InitPlayerItemEditor
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtItemQuantity_Validate", "frmAdmin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

