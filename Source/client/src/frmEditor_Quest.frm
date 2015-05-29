VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest Editor"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9105
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   607
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdPrevStep 
      Caption         =   "Go To Previous Page"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   193
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdNextStep 
      Caption         =   "Go To Next Page"
      Height          =   255
      Left            =   6960
      TabIndex        =   192
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quest List"
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ListBox lstIndex 
         Height          =   6720
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraStep1 
      Caption         =   "Step 1 - Define General Quest Information"
      Height          =   7095
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   6375
      Begin VB.Frame Frame4 
         Caption         =   "Quest Requirements:"
         Height          =   3375
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   6135
         Begin VB.TextBox txtVariableReq 
            Height          =   285
            Left            =   5160
            TabIndex        =   189
            Top             =   2400
            Width           =   495
         End
         Begin VB.ComboBox cmbPlayerVarReq 
            Height          =   300
            ItemData        =   "frmEditor_Quest.frx":0000
            Left            =   1920
            List            =   "frmEditor_Quest.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   188
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CheckBox chkVariableReq 
            Caption         =   "Player Variable"
            Height          =   255
            Left            =   120
            TabIndex        =   187
            Top             =   2400
            Width           =   1695
         End
         Begin VB.ComboBox cmbPlayerSwitchReq 
            Height          =   300
            ItemData        =   "frmEditor_Quest.frx":0004
            Left            =   1920
            List            =   "frmEditor_Quest.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   186
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CheckBox chkSwitchReq 
            Caption         =   "Player Switch"
            Height          =   255
            Left            =   120
            TabIndex        =   185
            Top             =   2880
            Width           =   1695
         End
         Begin VB.ComboBox cmbSwitchReqCompare 
            Height          =   300
            ItemData        =   "frmEditor_Quest.frx":0008
            Left            =   3720
            List            =   "frmEditor_Quest.frx":0012
            Style           =   2  'Dropdown List
            TabIndex        =   184
            Top             =   2880
            Width           =   1095
         End
         Begin VB.ComboBox cmbVariableReqCompare 
            Height          =   300
            ItemData        =   "frmEditor_Quest.frx":0023
            Left            =   3720
            List            =   "frmEditor_Quest.frx":0039
            Style           =   2  'Dropdown List
            TabIndex        =   183
            Top             =   2400
            Width           =   1335
         End
         Begin VB.HScrollBar scrlItemReqVal 
            Height          =   255
            Left            =   4080
            Max             =   32000
            TabIndex        =   29
            Top             =   1680
            Width           =   1815
         End
         Begin VB.HScrollBar scrlQuestCompleteCount 
            Height          =   255
            Left            =   3360
            Max             =   100
            TabIndex        =   24
            Top             =   2040
            Width           =   2535
         End
         Begin VB.ComboBox cmbItemReq 
            Height          =   300
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1680
            Width           =   1935
         End
         Begin VB.ComboBox cmbQuestReq 
            Height          =   300
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1320
            Width           =   3015
         End
         Begin VB.ComboBox cmbClassReq 
            Height          =   300
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   600
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevelReq 
            Height          =   255
            Left            =   2880
            Max             =   100
            TabIndex        =   19
            Top             =   960
            Width           =   3015
         End
         Begin VB.CheckBox chkNumQuestReq 
            Caption         =   "# of Quests Completed Req?"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2040
            Width           =   2535
         End
         Begin VB.CheckBox chkItemReq 
            Caption         =   "Item Req?"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CheckBox chkQuestReq 
            Caption         =   "Quest Completed Req?"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CheckBox chkLevelReq 
            Caption         =   "Level Req?"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1455
         End
         Begin VB.CheckBox classReq 
            Caption         =   "Class Req?"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "is"
            Height          =   255
            Index           =   5
            Left            =   3480
            TabIndex        =   191
            Top             =   2505
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "is"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   190
            Top             =   2955
            Width           =   255
         End
         Begin VB.Label lblItemReqVal 
            Caption         =   "x1"
            Height          =   255
            Left            =   3360
            TabIndex        =   28
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label lblQuestComplete 
            Caption         =   "1"
            Height          =   255
            Left            =   2880
            TabIndex        =   23
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label lblLevelReq 
            Caption         =   "Level: 1"
            Height          =   255
            Left            =   1920
            TabIndex        =   18
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label 
            Caption         =   "These are different requirements that you can set on being able to start this specific quest. "
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   5895
         End
      End
      Begin VB.Frame fraQuestOffer 
         Caption         =   "Quest Offer:"
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   6135
         Begin VB.TextBox txtQuestOffer 
            Height          =   615
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   5895
         End
         Begin VB.Label Label1 
            Caption         =   $"frmEditor_Quest.frx":009F
            Height          =   735
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   5895
         End
      End
      Begin VB.Frame fraQuestName 
         Caption         =   "Quest Name:"
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6135
         Begin VB.TextBox txtQuestName 
            Height          =   270
            Left            =   120
            MaxLength       =   30
            TabIndex        =   6
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label lblQuestName 
            Caption         =   "Come up with a creative name for your quest. This is what will appear in the quest log while the player is on this quest."
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   5895
         End
      End
   End
   Begin VB.Frame fraStep4 
      Caption         =   "Step 4 - Final Quest Options"
      Height          =   7095
      Left            =   2640
      TabIndex        =   198
      Top             =   120
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame 
         Caption         =   "Extra Quest Specific Options"
         Height          =   2655
         Index           =   3
         Left            =   120
         TabIndex        =   223
         Top             =   3960
         Width           =   6135
         Begin VB.CheckBox chkAbandonable 
            Caption         =   "Abandonable?"
            Height          =   255
            Left            =   120
            TabIndex        =   247
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox chkRepeat 
            Caption         =   "Repeatable?"
            Height          =   255
            Left            =   120
            TabIndex        =   246
            Top             =   240
            Width           =   1575
         End
         Begin VB.Frame fraBeforeStarted 
            Caption         =   "Quest Log Options (Will show in log during quest.)"
            Height          =   2295
            Left            =   1800
            TabIndex        =   239
            Top             =   240
            Width           =   4215
            Begin VB.TextBox txtAfterQuest 
               Height          =   495
               Left            =   120
               MaxLength       =   200
               TabIndex        =   244
               Top             =   1680
               Width           =   3975
            End
            Begin VB.CheckBox chkShowAfterQuest 
               Caption         =   "Show In Quest Log After Quest is Completed."
               Height          =   255
               Left            =   120
               TabIndex        =   243
               Top             =   1200
               Width           =   3975
            End
            Begin VB.TextBox txtBeforeQuest 
               Height          =   495
               Left            =   120
               MaxLength       =   200
               TabIndex        =   241
               Top             =   720
               Width           =   3975
            End
            Begin VB.CheckBox chkShowBeforeQuest 
               Caption         =   "Show In Quest Log Before Quest is Accepted."
               Height          =   255
               Left            =   120
               TabIndex        =   240
               Top             =   240
               Width           =   3975
            End
            Begin VB.Label Label 
               BackStyle       =   0  'Transparent
               Caption         =   "Quest Log Description after quest is completed."
               Height          =   255
               Index           =   60
               Left            =   120
               TabIndex        =   245
               Top             =   1440
               Width           =   3975
            End
            Begin VB.Label Label 
               BackStyle       =   0  'Transparent
               Caption         =   "Quest Log Description before quest is started."
               Height          =   255
               Index           =   59
               Left            =   120
               TabIndex        =   242
               Top             =   480
               Width           =   3975
            End
         End
      End
      Begin VB.Frame fraRewards 
         Caption         =   "Rewards"
         Height          =   3735
         Left            =   120
         TabIndex        =   199
         Top             =   240
         Width           =   6135
         Begin VB.ComboBox cmbSetVariableCompare 
            Height          =   300
            ItemData        =   "frmEditor_Quest.frx":01AA
            Left            =   4200
            List            =   "frmEditor_Quest.frx":01B7
            Style           =   2  'Dropdown List
            TabIndex        =   237
            Top             =   2760
            Width           =   1095
         End
         Begin VB.ComboBox cmbSetSwitchCompare 
            Height          =   300
            ItemData        =   "frmEditor_Quest.frx":01CF
            Left            =   4200
            List            =   "frmEditor_Quest.frx":01D9
            Style           =   2  'Dropdown List
            TabIndex        =   236
            Top             =   3240
            Width           =   1095
         End
         Begin VB.ComboBox cmbSetPlayerSwitch 
            Height          =   300
            ItemData        =   "frmEditor_Quest.frx":01EA
            Left            =   2400
            List            =   "frmEditor_Quest.frx":01EC
            Style           =   2  'Dropdown List
            TabIndex        =   235
            Top             =   3240
            Width           =   1455
         End
         Begin VB.ComboBox cmbSetPlayerVar 
            Height          =   300
            ItemData        =   "frmEditor_Quest.frx":01EE
            Left            =   2400
            List            =   "frmEditor_Quest.frx":01F0
            Style           =   2  'Dropdown List
            TabIndex        =   234
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox txtSetPlayerVarValue 
            Height          =   285
            Left            =   5400
            MaxLength       =   5
            TabIndex        =   233
            Top             =   2760
            Width           =   495
         End
         Begin VB.Frame Frame 
            Caption         =   "Teleport Player"
            Height          =   2295
            Index           =   2
            Left            =   4080
            TabIndex        =   224
            Top             =   120
            Width           =   1815
            Begin VB.CheckBox chkTeleAfter 
               Caption         =   "Teleport Player when complete?"
               Height          =   375
               Left            =   120
               TabIndex        =   228
               Top             =   220
               Width           =   1575
            End
            Begin VB.HScrollBar scrlAfterY 
               Height          =   255
               Left            =   120
               Max             =   255
               TabIndex        =   227
               Top             =   1800
               Value           =   1
               Width           =   1575
            End
            Begin VB.HScrollBar scrlAfterX 
               Height          =   255
               Left            =   120
               Max             =   255
               TabIndex        =   226
               Top             =   1320
               Value           =   1
               Width           =   1575
            End
            Begin VB.HScrollBar scrlAfterMap 
               Height          =   255
               Left            =   120
               Min             =   1
               TabIndex        =   225
               Top             =   840
               Value           =   1
               Width           =   1575
            End
            Begin VB.Label lblAfterY 
               Caption         =   "Y: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   231
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label lblAfterX 
               Caption         =   "X: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   230
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label lblAfterMap 
               Caption         =   "Map: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   229
               Top             =   600
               Width           =   975
            End
         End
         Begin VB.CheckBox chkRestoreMana 
            Caption         =   "Restore Mana"
            Height          =   255
            Left            =   120
            TabIndex        =   222
            Top             =   3360
            Width           =   1935
         End
         Begin VB.CheckBox chkRestoreHealth 
            Caption         =   "Restore Health"
            Height          =   255
            Left            =   120
            TabIndex        =   221
            Top             =   3120
            Width           =   1935
         End
         Begin VB.HScrollBar scrlGiveExp 
            Height          =   255
            Left            =   120
            TabIndex        =   220
            Top             =   2760
            Width           =   1935
         End
         Begin VB.ComboBox cmbSpellReward 
            Height          =   300
            Index           =   1
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   215
            Top             =   1680
            Width           =   3375
         End
         Begin VB.ComboBox cmbSpellReward 
            Height          =   300
            Index           =   2
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   214
            Top             =   2040
            Width           =   3375
         End
         Begin VB.HScrollBar scrlRewardItemVal 
            Height          =   255
            Index           =   1
            Left            =   2760
            Max             =   32000
            Min             =   1
            TabIndex        =   205
            Top             =   360
            Value           =   1
            Width           =   1095
         End
         Begin VB.ComboBox cmbRewardItem 
            Height          =   300
            Index           =   1
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   360
            Width           =   1575
         End
         Begin VB.HScrollBar scrlRewardItemVal 
            Height          =   255
            Index           =   2
            Left            =   2760
            Max             =   32000
            Min             =   1
            TabIndex        =   203
            Top             =   720
            Value           =   1
            Width           =   1095
         End
         Begin VB.ComboBox cmbRewardItem 
            Height          =   300
            Index           =   2
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   202
            Top             =   720
            Width           =   1575
         End
         Begin VB.HScrollBar scrlRewardItemVal 
            Height          =   255
            Index           =   3
            Left            =   2760
            Max             =   32000
            Min             =   1
            TabIndex        =   201
            Top             =   1080
            Value           =   1
            Width           =   1095
         End
         Begin VB.ComboBox cmbRewardItem 
            Height          =   300
            Index           =   3
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   200
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "to"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   238
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label Label 
            Caption         =   "Set Switch/Variable"
            Height          =   255
            Index           =   58
            Left            =   3360
            TabIndex        =   232
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000A&
            X1              =   2280
            X2              =   2280
            Y1              =   3600
            Y2              =   2520
         End
         Begin VB.Label lblGiveExp 
            Caption         =   "Give Exp: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   219
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label 
            Caption         =   "2."
            Height          =   255
            Index           =   57
            Left            =   120
            TabIndex        =   218
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label 
            Caption         =   "1."
            Height          =   255
            Index           =   56
            Left            =   120
            TabIndex        =   217
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label 
            Caption         =   "Teach Spells"
            Height          =   255
            Index           =   55
            Left            =   1680
            TabIndex        =   216
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Line Line 
            BorderColor     =   &H8000000A&
            X1              =   120
            X2              =   3960
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label Label 
            Caption         =   "3."
            Height          =   255
            Index           =   54
            Left            =   120
            TabIndex        =   213
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label 
            Caption         =   "2."
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   212
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label 
            Caption         =   "1."
            Height          =   255
            Index           =   52
            Left            =   120
            TabIndex        =   211
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label 
            Caption         =   "Amount"
            Height          =   255
            Index           =   51
            Left            =   3000
            TabIndex        =   210
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label 
            Caption         =   "Item"
            Height          =   255
            Index           =   50
            Left            =   1080
            TabIndex        =   209
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblRewardItemVal 
            Alignment       =   1  'Right Justify
            Caption         =   "x1"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   208
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblRewardItemVal 
            Alignment       =   1  'Right Justify
            Caption         =   "x1"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   207
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblRewardItemVal 
            Alignment       =   1  'Right Justify
            Caption         =   "x1"
            Height          =   255
            Index           =   3
            Left            =   2040
            TabIndex        =   206
            Top             =   1080
            Width           =   615
         End
      End
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Step 3 - Setup Quest Tasks"
      Height          =   7095
      Left            =   2640
      TabIndex        =   76
      Top             =   120
      Width           =   6375
      Begin VB.CheckBox chkEndQuestOnCompletion 
         Caption         =   "End quest immediately when this task is finished?"
         Height          =   255
         Left            =   1920
         TabIndex        =   104
         Top             =   6160
         Width           =   4095
      End
      Begin VB.Frame fraCurrentTask 
         Caption         =   "Task 1"
         Height          =   5295
         Left            =   120
         TabIndex        =   78
         Top             =   1320
         Width           =   6135
         Begin VB.CommandButton cmdHelp2 
            Caption         =   "Explain ->"
            Height          =   255
            Left            =   120
            TabIndex        =   197
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtTaskDesc 
            Height          =   735
            Left            =   1200
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   196
            Text            =   "frmEditor_Quest.frx":01F2
            Top             =   480
            Width           =   4815
         End
         Begin VB.ComboBox cmbTaskType 
            Height          =   300
            ItemData        =   "frmEditor_Quest.frx":01FE
            Left            =   1200
            List            =   "frmEditor_Quest.frx":021A
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   120
            Width           =   2775
         End
         Begin VB.Frame fraKillNpcs 
            Caption         =   "Task Type: Kill NPC(s)"
            Height          =   3855
            Left            =   120
            TabIndex        =   85
            Top             =   1320
            Width           =   5895
            Begin VB.HScrollBar scrlKillNPCCount 
               Height          =   255
               Index           =   3
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   103
               Top             =   1200
               Value           =   1
               Width           =   1815
            End
            Begin VB.HScrollBar scrlKillNPCCount 
               Height          =   255
               Index           =   2
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   102
               Top             =   840
               Value           =   1
               Width           =   1815
            End
            Begin VB.HScrollBar scrlKillNPCCount 
               Height          =   255
               Index           =   4
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   101
               Top             =   1560
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbKillNPC 
               Height          =   300
               Index           =   3
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   1200
               Width           =   1935
            End
            Begin VB.ComboBox cmbKillNPC 
               Height          =   300
               Index           =   2
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   840
               Width           =   1935
            End
            Begin VB.ComboBox cmbKillNPC 
               Height          =   300
               Index           =   4
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   95
               Top             =   1560
               Width           =   1935
            End
            Begin VB.HScrollBar scrlKillNPCCount 
               Height          =   255
               Index           =   1
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   87
               Top             =   480
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbKillNPC 
               Height          =   300
               Index           =   1
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Top             =   480
               Width           =   1935
            End
            Begin VB.Label lblKillNPCValue 
               Caption         =   "x1"
               Height          =   255
               Index           =   3
               Left            =   3240
               TabIndex        =   100
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label lblKillNPCValue 
               Caption         =   "x1"
               Height          =   255
               Index           =   2
               Left            =   3240
               TabIndex        =   99
               Top             =   840
               Width           =   615
            End
            Begin VB.Label lblKillNPCValue 
               Caption         =   "x1"
               Height          =   255
               Index           =   4
               Left            =   3240
               TabIndex        =   98
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "4."
               Height          =   255
               Index           =   19
               Left            =   120
               TabIndex        =   94
               Top             =   1560
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "3."
               Height          =   255
               Index           =   18
               Left            =   120
               TabIndex        =   93
               Top             =   1200
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "2."
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   92
               Top             =   840
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "1."
               Height          =   255
               Index           =   16
               Left            =   120
               TabIndex        =   91
               Top             =   480
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "Count"
               Height          =   255
               Index           =   15
               Left            =   4440
               TabIndex        =   90
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "NPC"
               Height          =   255
               Index           =   14
               Left            =   1080
               TabIndex        =   89
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lblKillNPCValue 
               Caption         =   "x1"
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   88
               Top             =   480
               Width           =   615
            End
         End
         Begin VB.Frame fraGotoMap 
            Caption         =   "Task Type: Reach Map"
            Height          =   3855
            Left            =   120
            TabIndex        =   157
            Top             =   1320
            Width           =   5895
            Begin VB.TextBox txtGoToMap 
               Height          =   375
               Left            =   120
               TabIndex        =   163
               Text            =   "txtGoToMap"
               Top             =   2280
               Width           =   5655
            End
            Begin VB.HScrollBar scrlGotoMap 
               Height          =   255
               Left            =   600
               Max             =   100
               Min             =   1
               TabIndex        =   158
               Top             =   960
               Value           =   1
               Width           =   5055
            End
            Begin VB.Label Label 
               Caption         =   $"frmEditor_Quest.frx":029C
               ForeColor       =   &H000000FF&
               Height          =   975
               Index           =   42
               Left            =   120
               TabIndex        =   162
               Top             =   1320
               Width           =   5655
            End
            Begin VB.Label lblGotoMap 
               Caption         =   "x1"
               Height          =   255
               Left            =   240
               TabIndex        =   161
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "Map #:"
               Height          =   255
               Index           =   40
               Left            =   600
               TabIndex        =   160
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Label 
               Caption         =   "Instead of simply using the map name you can define what you want the player to do although it is just reaching a certain map."
               ForeColor       =   &H000000FF&
               Height          =   615
               Index           =   38
               Left            =   120
               TabIndex        =   159
               Top             =   240
               Width           =   5655
            End
         End
         Begin VB.Frame fraKillPlayers 
            Caption         =   "Task Type: Kill Players"
            Height          =   3855
            Left            =   120
            TabIndex        =   152
            Top             =   1320
            Width           =   5895
            Begin VB.HScrollBar scrlKillPlayer 
               Height          =   255
               Left            =   600
               Max             =   100
               Min             =   1
               TabIndex        =   153
               Top             =   1200
               Value           =   1
               Width           =   5055
            End
            Begin VB.Label Label 
               Caption         =   $"frmEditor_Quest.frx":0402
               ForeColor       =   &H000000FF&
               Height          =   615
               Index           =   39
               Left            =   120
               TabIndex        =   156
               Top             =   240
               Width           =   5655
            End
            Begin VB.Label Label 
               Caption         =   "Kill Player Count:"
               Height          =   255
               Index           =   41
               Left            =   600
               TabIndex        =   155
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label lblKillPlayerCount 
               Caption         =   "x1"
               Height          =   255
               Left            =   240
               TabIndex        =   154
               Top             =   1200
               Width           =   615
            End
         End
         Begin VB.Frame fraAquireDeliverItem 
            Caption         =   "Task Type: Aquire and Deliver Item(s)"
            Height          =   3855
            Left            =   120
            TabIndex        =   130
            Top             =   1320
            Width           =   5895
            Begin VB.TextBox txtAquireDeliverEventName 
               Height          =   390
               Left            =   120
               TabIndex        =   151
               Text            =   "txtAquireDeliverEventName"
               Top             =   3120
               Width           =   5535
            End
            Begin VB.ComboBox cmbAquireDeliverItem 
               Height          =   300
               Index           =   4
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   148
               Top             =   1920
               Width           =   1935
            End
            Begin VB.HScrollBar scrlAquireDeliverItemVal 
               Height          =   255
               Index           =   4
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   147
               Top             =   1920
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbAquireDeliverItem 
               Height          =   300
               Index           =   3
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   145
               Top             =   1560
               Width           =   1935
            End
            Begin VB.HScrollBar scrlAquireDeliverItemVal 
               Height          =   255
               Index           =   3
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   144
               Top             =   1560
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbAquireDeliverItem 
               Height          =   300
               Index           =   2
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   142
               Top             =   1200
               Width           =   1935
            End
            Begin VB.HScrollBar scrlAquireDeliverItemVal 
               Height          =   255
               Index           =   2
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   141
               Top             =   1200
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbAquireDeliverItem 
               Height          =   300
               Index           =   1
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   132
               Top             =   840
               Width           =   1935
            End
            Begin VB.HScrollBar scrlAquireDeliverItemVal 
               Height          =   255
               Index           =   1
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   131
               Top             =   840
               Value           =   1
               Width           =   1815
            End
            Begin VB.Label Label 
               BackStyle       =   0  'Transparent
               Caption         =   $"frmEditor_Quest.frx":04AE
               ForeColor       =   &H000000FF&
               Height          =   855
               Index           =   37
               Left            =   120
               TabIndex        =   150
               Top             =   2280
               Width           =   5655
            End
            Begin VB.Label lblAquireDeliverItemVal 
               Caption         =   "x1"
               Height          =   255
               Index           =   4
               Left            =   3240
               TabIndex        =   149
               Top             =   1920
               Width           =   615
            End
            Begin VB.Label lblAquireDeliverItemVal 
               Caption         =   "x1"
               Height          =   255
               Index           =   3
               Left            =   3240
               TabIndex        =   146
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label lblAquireDeliverItemVal 
               Caption         =   "x1"
               Height          =   255
               Index           =   2
               Left            =   3240
               TabIndex        =   143
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Label 
               BackStyle       =   0  'Transparent
               Caption         =   "This task is very much a like aquire items however you must end the task by delivering the items to an event."
               ForeColor       =   &H000000FF&
               Height          =   495
               Index           =   36
               Left            =   120
               TabIndex        =   140
               Top             =   240
               Width           =   5655
            End
            Begin VB.Label lblAquireDeliverItemVal 
               Caption         =   "x1"
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   139
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "Item"
               Height          =   255
               Index           =   35
               Left            =   1080
               TabIndex        =   138
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label 
               Caption         =   "Amount"
               Height          =   255
               Index           =   34
               Left            =   4440
               TabIndex        =   137
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "1."
               Height          =   255
               Index           =   33
               Left            =   120
               TabIndex        =   136
               Top             =   840
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "2."
               Height          =   255
               Index           =   32
               Left            =   120
               TabIndex        =   135
               Top             =   1200
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "3."
               Height          =   255
               Index           =   31
               Left            =   120
               TabIndex        =   134
               Top             =   1560
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "4."
               Height          =   255
               Index           =   27
               Left            =   120
               TabIndex        =   133
               Top             =   1920
               Width           =   375
            End
         End
         Begin VB.Frame fraAquireItems 
            Caption         =   "Task Type: Aquire Item(s)"
            Height          =   3855
            Left            =   120
            TabIndex        =   110
            Top             =   1320
            Width           =   5895
            Begin VB.HScrollBar scrlAquireItemVal 
               Height          =   255
               Index           =   4
               Left            =   3960
               Max             =   100
               Min             =   1
               TabIndex        =   127
               Top             =   2040
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbAquireItem 
               Height          =   300
               Index           =   4
               Left            =   600
               Style           =   2  'Dropdown List
               TabIndex        =   126
               Top             =   2040
               Width           =   1935
            End
            Begin VB.HScrollBar scrlAquireItemVal 
               Height          =   255
               Index           =   3
               Left            =   3960
               Max             =   100
               Min             =   1
               TabIndex        =   123
               Top             =   1680
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbAquireItem 
               Height          =   300
               Index           =   3
               Left            =   600
               Style           =   2  'Dropdown List
               TabIndex        =   122
               Top             =   1680
               Width           =   1935
            End
            Begin VB.HScrollBar scrlAquireItemVal 
               Height          =   255
               Index           =   2
               Left            =   3960
               Max             =   100
               Min             =   1
               TabIndex        =   119
               Top             =   1320
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbAquireItem 
               Height          =   300
               Index           =   2
               Left            =   600
               Style           =   2  'Dropdown List
               TabIndex        =   118
               Top             =   1320
               Width           =   1935
            End
            Begin VB.HScrollBar scrlAquireItemVal 
               Height          =   255
               Index           =   1
               Left            =   3960
               Max             =   100
               Min             =   1
               TabIndex        =   113
               Top             =   960
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbAquireItem 
               Height          =   300
               Index           =   1
               Left            =   600
               Style           =   2  'Dropdown List
               TabIndex        =   112
               Top             =   960
               Width           =   1935
            End
            Begin VB.Label Label 
               Caption         =   "4."
               Height          =   255
               Index           =   30
               Left            =   240
               TabIndex        =   129
               Top             =   2040
               Width           =   375
            End
            Begin VB.Label lblAquireItemVal 
               Caption         =   "x1"
               Height          =   255
               Index           =   4
               Left            =   3360
               TabIndex        =   128
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "3."
               Height          =   255
               Index           =   29
               Left            =   240
               TabIndex        =   125
               Top             =   1680
               Width           =   375
            End
            Begin VB.Label lblAquireItemVal 
               Caption         =   "x1"
               Height          =   255
               Index           =   3
               Left            =   3360
               TabIndex        =   124
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "2."
               Height          =   255
               Index           =   28
               Left            =   240
               TabIndex        =   121
               Top             =   1320
               Width           =   375
            End
            Begin VB.Label lblAquireItemVal 
               Caption         =   "x1"
               Height          =   255
               Index           =   2
               Left            =   3360
               TabIndex        =   120
               Top             =   1320
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "1."
               Height          =   255
               Index           =   26
               Left            =   240
               TabIndex        =   117
               Top             =   960
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "Amount"
               Height          =   255
               Index           =   24
               Left            =   4560
               TabIndex        =   116
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "Item"
               Height          =   255
               Index           =   23
               Left            =   1200
               TabIndex        =   115
               Top             =   720
               Width           =   495
            End
            Begin VB.Label lblAquireItemVal 
               Caption         =   "x1"
               Height          =   255
               Index           =   1
               Left            =   3360
               TabIndex        =   114
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   $"frmEditor_Quest.frx":05D6
               ForeColor       =   &H000000FF&
               Height          =   615
               Index           =   25
               Left            =   120
               TabIndex        =   111
               Top             =   240
               Width           =   5655
            End
         End
         Begin VB.Frame fraTalkToEvent 
            Caption         =   "Task Type: Talk To Event"
            Height          =   3855
            Left            =   120
            TabIndex        =   105
            Top             =   1320
            Width           =   5895
            Begin VB.TextBox txtEventTask 
               Height          =   375
               Left            =   120
               TabIndex        =   108
               Text            =   "txtEventTask"
               Top             =   1440
               Width           =   5655
            End
            Begin VB.Label Label 
               Caption         =   $"frmEditor_Quest.frx":0678
               ForeColor       =   &H000000FF&
               Height          =   615
               Index           =   22
               Left            =   120
               TabIndex        =   109
               Top             =   2040
               Width           =   5655
            End
            Begin VB.Label Label 
               Caption         =   "First, enter what you want the quest log to say the current task is. For example, ""Talk to Tom in the castle dungeon."""
               ForeColor       =   &H000000FF&
               Height          =   375
               Index           =   21
               Left            =   120
               TabIndex        =   107
               Top             =   960
               Width           =   5655
            End
            Begin VB.Label Label 
               Caption         =   $"frmEditor_Quest.frx":0714
               ForeColor       =   &H000000FF&
               Height          =   615
               Index           =   20
               Left            =   120
               TabIndex        =   106
               Top             =   240
               Width           =   5655
            End
         End
         Begin VB.Frame fraGatherResources 
            Caption         =   "Task Type: Gather Resources"
            Height          =   3855
            Left            =   120
            TabIndex        =   164
            Top             =   1320
            Width           =   5895
            Begin VB.ComboBox cmbGatherResource 
               Height          =   300
               Index           =   4
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   181
               Top             =   1560
               Width           =   1935
            End
            Begin VB.HScrollBar scrlGatherResourceAmount 
               Height          =   255
               Index           =   4
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   180
               Top             =   1560
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbGatherResource 
               Height          =   300
               Index           =   3
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   178
               Top             =   1200
               Width           =   1935
            End
            Begin VB.HScrollBar scrlGatherResourceAmount 
               Height          =   255
               Index           =   3
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   177
               Top             =   1200
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbGatherResource 
               Height          =   300
               Index           =   2
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   175
               Top             =   840
               Width           =   1935
            End
            Begin VB.HScrollBar scrlGatherResourceAmount 
               Height          =   255
               Index           =   2
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   174
               Top             =   840
               Value           =   1
               Width           =   1815
            End
            Begin VB.ComboBox cmbGatherResource 
               Height          =   300
               Index           =   1
               Left            =   480
               Style           =   2  'Dropdown List
               TabIndex        =   172
               Top             =   480
               Width           =   1935
            End
            Begin VB.HScrollBar scrlGatherResourceAmount 
               Height          =   255
               Index           =   1
               Left            =   3840
               Max             =   100
               Min             =   1
               TabIndex        =   171
               Top             =   480
               Value           =   1
               Width           =   1815
            End
            Begin VB.Label lblGatherResourceAmount 
               Caption         =   "x1"
               Height          =   255
               Index           =   4
               Left            =   3240
               TabIndex        =   182
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label lblGatherResourceAmount 
               Caption         =   "x1"
               Height          =   255
               Index           =   3
               Left            =   3240
               TabIndex        =   179
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label lblGatherResourceAmount 
               Caption         =   "x1"
               Height          =   255
               Index           =   2
               Left            =   3240
               TabIndex        =   176
               Top             =   840
               Width           =   615
            End
            Begin VB.Label lblGatherResourceAmount 
               Caption         =   "x1"
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   173
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "Resource"
               Height          =   255
               Index           =   48
               Left            =   960
               TabIndex        =   170
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label 
               Caption         =   "Amount"
               Height          =   255
               Index           =   47
               Left            =   4440
               TabIndex        =   169
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label 
               Caption         =   "1."
               Height          =   255
               Index           =   46
               Left            =   120
               TabIndex        =   168
               Top             =   480
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "2."
               Height          =   255
               Index           =   45
               Left            =   120
               TabIndex        =   167
               Top             =   840
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "3."
               Height          =   255
               Index           =   44
               Left            =   120
               TabIndex        =   166
               Top             =   1200
               Width           =   375
            End
            Begin VB.Label Label 
               Caption         =   "4."
               Height          =   255
               Index           =   43
               Left            =   120
               TabIndex        =   165
               Top             =   1560
               Width           =   375
            End
         End
         Begin VB.Label Label 
            Caption         =   "Quest Log Desc of Task:"
            Height          =   495
            Index           =   49
            Left            =   120
            TabIndex        =   195
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label 
            Caption         =   "Task Type:"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Tasks"
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   6135
         Begin VB.CommandButton cmdNextTask 
            Caption         =   "Next Task"
            Height          =   375
            Left            =   4440
            TabIndex        =   82
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton cmdPrevTask 
            Caption         =   "Previous Task"
            Height          =   375
            Left            =   120
            TabIndex        =   81
            Top             =   600
            Width           =   1575
         End
         Begin VB.HScrollBar scrlCurTask 
            Height          =   255
            Left            =   1440
            Max             =   10
            Min             =   1
            TabIndex        =   80
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.Label lblCurTask 
            Caption         =   "Task: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Frame fraStep2 
      Caption         =   "Step 2 - Setup Quest Start Options"
      Height          =   7095
      Left            =   2640
      TabIndex        =   25
      Top             =   120
      Width           =   6375
      Begin VB.Frame Frame 
         Caption         =   "Teleport Player"
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   69
         Top             =   5400
         Width           =   6135
         Begin VB.CheckBox chkTeleportOnStart 
            Caption         =   "Teleport Player on quest accept?"
            Height          =   255
            Left            =   120
            TabIndex        =   194
            Top             =   220
            Width           =   2895
         End
         Begin VB.HScrollBar scrlTeleY 
            Height          =   255
            Left            =   4080
            Max             =   255
            TabIndex        =   75
            Top             =   480
            Value           =   1
            Width           =   1815
         End
         Begin VB.HScrollBar scrlTeleX 
            Height          =   255
            Left            =   4080
            Max             =   255
            TabIndex        =   73
            Top             =   240
            Value           =   1
            Width           =   1815
         End
         Begin VB.HScrollBar scrlTeleMap 
            Height          =   255
            Left            =   1200
            Min             =   1
            TabIndex        =   71
            Top             =   480
            Value           =   1
            Width           =   1815
         End
         Begin VB.Label lblTeleY 
            Caption         =   "Y: 1"
            Height          =   255
            Left            =   3120
            TabIndex        =   74
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblTeleX 
            Caption         =   "X: 1"
            Height          =   255
            Left            =   3120
            TabIndex        =   72
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblTeleportMap 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Take Player Items:"
         Height          =   2415
         Left            =   120
         TabIndex        =   49
         Top             =   2880
         Width           =   6135
         Begin VB.ComboBox cmbTakeItem 
            Height          =   300
            Index           =   0
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   480
            Width           =   1935
         End
         Begin VB.HScrollBar scrlTakeItemVal 
            Height          =   255
            Index           =   0
            Left            =   3840
            Max             =   32000
            Min             =   1
            TabIndex        =   56
            Top             =   480
            Value           =   1
            Width           =   1815
         End
         Begin VB.ComboBox cmbTakeItem 
            Height          =   300
            Index           =   1
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   840
            Width           =   1935
         End
         Begin VB.HScrollBar scrlTakeItemVal 
            Height          =   255
            Index           =   1
            Left            =   3840
            Max             =   32000
            Min             =   1
            TabIndex        =   54
            Top             =   840
            Value           =   1
            Width           =   1815
         End
         Begin VB.ComboBox cmbTakeItem 
            Height          =   300
            Index           =   2
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1200
            Width           =   1935
         End
         Begin VB.HScrollBar scrlTakeItemVal 
            Height          =   255
            Index           =   2
            Left            =   3840
            Max             =   32000
            Min             =   1
            TabIndex        =   52
            Top             =   1200
            Value           =   1
            Width           =   1815
         End
         Begin VB.ComboBox cmbTakeItem 
            Height          =   300
            Index           =   3
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   1560
            Width           =   1935
         End
         Begin VB.HScrollBar scrlTakeItemVal 
            Height          =   255
            Index           =   3
            Left            =   3840
            Max             =   32000
            Min             =   1
            TabIndex        =   50
            Top             =   1560
            Value           =   1
            Width           =   1815
         End
         Begin VB.Label lblTakeItemVal 
            Caption         =   "x1"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   68
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label 
            Caption         =   "Item"
            Height          =   255
            Index           =   12
            Left            =   1080
            TabIndex        =   67
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label 
            Caption         =   "Value"
            Height          =   255
            Index           =   11
            Left            =   4440
            TabIndex        =   66
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label 
            Caption         =   "1."
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   65
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblTakeItemVal 
            Caption         =   "x1"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   64
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label 
            Caption         =   "2."
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   63
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lblTakeItemVal 
            Caption         =   "x1"
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   62
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label 
            Caption         =   "3."
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   61
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblTakeItemVal 
            Caption         =   "x1"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   60
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label 
            Caption         =   "4."
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "If the player does not have these items then it will say they do not meet the requirements needed to start the quest."
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   1920
            Width           =   5895
         End
      End
      Begin VB.Frame fraGiveItem 
         Caption         =   "Give Player Items:"
         Height          =   2415
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   6135
         Begin VB.HScrollBar scrlGiveItemVal 
            Height          =   255
            Index           =   3
            Left            =   3840
            Max             =   32000
            Min             =   1
            TabIndex        =   46
            Top             =   1560
            Value           =   1
            Width           =   1815
         End
         Begin VB.ComboBox cmbGiveItem 
            Height          =   300
            Index           =   3
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1560
            Width           =   1935
         End
         Begin VB.HScrollBar scrlGiveItemVal 
            Height          =   255
            Index           =   2
            Left            =   3840
            Max             =   32000
            Min             =   1
            TabIndex        =   42
            Top             =   1200
            Value           =   1
            Width           =   1815
         End
         Begin VB.ComboBox cmbGiveItem 
            Height          =   300
            Index           =   2
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1200
            Width           =   1935
         End
         Begin VB.HScrollBar scrlGiveItemVal 
            Height          =   255
            Index           =   1
            Left            =   3840
            Max             =   32000
            Min             =   1
            TabIndex        =   38
            Top             =   840
            Value           =   1
            Width           =   1815
         End
         Begin VB.ComboBox cmbGiveItem 
            Height          =   300
            Index           =   1
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   840
            Width           =   1935
         End
         Begin VB.HScrollBar scrlGiveItemVal 
            Height          =   255
            Index           =   0
            Left            =   3840
            Max             =   32000
            Min             =   1
            TabIndex        =   32
            Top             =   480
            Value           =   1
            Width           =   1815
         End
         Begin VB.ComboBox cmbGiveItem 
            Height          =   300
            Index           =   0
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   $"frmEditor_Quest.frx":07B1
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   48
            Top             =   1920
            Width           =   5895
         End
         Begin VB.Label Label 
            Caption         =   "4."
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   47
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label lblGiveItemVal 
            Caption         =   "x1"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   45
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label 
            Caption         =   "3."
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   43
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblGiveItemVal 
            Caption         =   "x1"
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   41
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label 
            Caption         =   "2."
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lblGiveItemVal 
            Caption         =   "x1"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   37
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label 
            Caption         =   "1."
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label 
            Caption         =   "Value"
            Height          =   255
            Index           =   2
            Left            =   4440
            TabIndex        =   34
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label 
            Caption         =   "Item"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblGiveItemVal 
            Caption         =   "x1"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   31
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Label Label7 
         Caption         =   "All of the below are simply extra options for starting a quest."
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private TempTask As Long

Private Sub chkAbandonable_Click()


   On Error GoTo errorhandler
    If chkAbandonable.Value = 1 Then
        quest(EditorIndex).Abandonable = 1
    Else
        quest(EditorIndex).Abandonable = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkAbandonable_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkEndQuestOnCompletion_Click()


   On Error GoTo errorhandler
    If chkEndQuestOnCompletion.Value = 1 Then
        quest(EditorIndex).Task(QuestEditorTask).EndOnCompletion = 1
    Else
        quest(EditorIndex).Task(QuestEditorTask).EndOnCompletion = 0
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkEndQuestOnCompletion_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkItemReq_Click()


   On Error GoTo errorhandler
    If chkItemReq.Value = 1 Then
        quest(EditorIndex).ItemReq = 1
    Else
        quest(EditorIndex).ItemReq = 0
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkItemReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkLevelReq_Click()


   On Error GoTo errorhandler
    If chkLevelReq.Value = 1 Then
        quest(EditorIndex).LevelReq = 1
    Else
        quest(EditorIndex).LevelReq = 0
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkLevelReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkNumQuestReq_Click()


   On Error GoTo errorhandler
    If chkNumQuestReq.Value = 1 Then
        quest(EditorIndex).NumQuestCompleteReq = 1
    Else
        quest(EditorIndex).NumQuestCompleteReq = 0
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkNumQuestReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkQuestReq_Click()


   On Error GoTo errorhandler
    If chkQuestReq.Value = 1 Then
        quest(EditorIndex).QuestCompleteReq = 1
    Else
        quest(EditorIndex).QuestCompleteReq = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkQuestReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkRepeat_Click()


   On Error GoTo errorhandler
    If chkRepeat.Value = 1 Then
        quest(EditorIndex).Repeatable = 1
    Else
        quest(EditorIndex).Repeatable = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkRepeat_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkRestoreHealth_Click()


   On Error GoTo errorhandler
    If chkRestoreHealth.Value = 1 Then
        quest(EditorIndex).RestoreHealth = 1
    Else
        quest(EditorIndex).RestoreHealth = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkRestoreHealth_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkRestoreMana_Click()


   On Error GoTo errorhandler
    If chkRestoreMana.Value = 1 Then
        quest(EditorIndex).RestoreMana = 1
    Else
        quest(EditorIndex).RestoreMana = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkRestoreMana_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkShowAfterQuest_Click()


   On Error GoTo errorhandler
    If chkShowAfterQuest.Value = 1 Then
        quest(EditorIndex).QuestLogAfter = 1
    Else
        quest(EditorIndex).QuestLogAfter = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkShowAfterQuest_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkShowBeforeQuest_Click()


   On Error GoTo errorhandler
    If chkShowBeforeQuest.Value = 1 Then
        quest(EditorIndex).QuestLogBefore = 1
    Else
        quest(EditorIndex).QuestLogBefore = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkShowBeforeQuest_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkSwitchReq_Click()
   On Error GoTo errorhandler
   
    If chkSwitchReq.Value = 1 Then
        quest(EditorIndex).SwitchReq = 1
    Else
        quest(EditorIndex).SwitchReq = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkSwitchReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkTeleAfter_Click()


   On Error GoTo errorhandler
    If chkTeleAfter.Value = 1 Then
        quest(EditorIndex).TeleportAfter = 1
    Else
        quest(EditorIndex).TeleportAfter = 0
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkTeleAfter_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkTeleportOnStart_Click()


   On Error GoTo errorhandler
    If chkTeleportOnStart.Value = 1 Then
        quest(EditorIndex).TeleportBefore = 1
    Else
        quest(EditorIndex).TeleportBefore = 0
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkTeleportOnStart_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkVariableReq_Click()
   On Error GoTo errorhandler
   
    If chkVariableReq.Value = 1 Then
        quest(EditorIndex).VariableReq = 1
    Else
        quest(EditorIndex).VariableReq = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkVariableReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub ClassReq_Click()


   On Error GoTo errorhandler
   
    If classReq.Value = 1 Then
        quest(EditorIndex).classReq = 1
    Else
        quest(EditorIndex).classReq = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClassReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbAquireDeliverItem_Click(Index As Integer)

   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).Task(QuestEditorTask).data(Index) = cmbAquireDeliverItem(Index).ListIndex


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbAquireDeliverItem_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbAquireItem_Click(Index As Integer)


   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).Task(QuestEditorTask).data(Index) = cmbAquireItem(Index).ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbAquireItem_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbClassReq_Click()


   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).RequiredClass = cmbClassReq.ListIndex + 1
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbGatherResource_Click(Index As Integer)
    

   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).Task(QuestEditorTask).data(Index) = cmbGatherResource(Index).ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbGatherResource_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbGiveItem_Click(Index As Integer)
   On Error GoTo errorhandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).GiveItemBefore(Index).Item = cmbGiveItem(Index).ListIndex
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbGiveItem_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbItemReq_Click()


   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).RequiredItem = cmbItemReq.ListIndex + 1
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbItemReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbKillNPC_Click(Index As Integer)
    

   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).Task(QuestEditorTask).data(Index) = cmbKillNPC(Index).ListIndex

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbKillNPC_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbPlayerSwitchReq_Click()
   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).RequiredSwitchNum = cmbPlayerSwitchReq.ListIndex
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbPlayerSwitchReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbPlayerVarReq_Click()


   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).RequiredVariableNum = cmbPlayerVarReq.ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbPlayerVarReqClick", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbQuestReq_Click()


   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).RequiredQuest = cmbQuestReq.ListIndex + 1
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbQuestReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbRewardItem_Click(Index As Integer)


   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).RewardItem(Index).Item = cmbRewardItem(Index).ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbRewardItem_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSetPlayerSwitch_Click()


   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).SetPlayerSwitch = cmbSetPlayerSwitch.ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSetPlayerSwitch_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSetPlayerVar_Click()


   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).SetPlayerVar = cmbSetPlayerVar.ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSetPlayerVar_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSetSwitchCompare_Click()
    

   On Error GoTo errorhandler
    quest(EditorIndex).SetPlayerSwitchValue = cmbSetSwitchCompare.ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSetSwitchCompare_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSetVariableCompare_Click()


   On Error GoTo errorhandler
    quest(EditorIndex).SetPlayerVarMod = cmbSetVariableCompare.ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSetVariableCompare_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSpellReward_Click(Index As Integer)
   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).RewardSpell(Index) = cmbSpellReward(Index).ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSpellReward_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSwitchReqCompare_Click()
   On Error GoTo errorhandler
    quest(EditorIndex).RequiredSwitchSet = cmbSwitchReqCompare.ListIndex
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSwitchReqCompare_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbTakeItem_Click(Index As Integer)
    

   On Error GoTo errorhandler
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    quest(EditorIndex).TakeItemBefore(Index).Item = cmbTakeItem(Index).ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbTakeItem_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbTaskType_Click()

   On Error GoTo errorhandler
    If cmbTaskType.ListIndex = 0 Then
        ClearQuestTask EditorIndex, QuestEditorTask
    End If
    quest(EditorIndex).Task(QuestEditorTask).type = cmbTaskType.ListIndex
    QuestEditorInitTask


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbTaskType_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbVariableReqCompare_Click()


   On Error GoTo errorhandler
    quest(EditorIndex).RequiredVariableCompare = cmbVariableReqCompare.ListIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbVariableReqCompare_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdHelp2_Click()


   On Error GoTo errorhandler
    MsgBox "The task description will show up in the task log with the current quest objectives. For example, lets say you have a task to go collect apples and oranges for farmer joe. The player is on this quest and they open up the quest editor.. the will see the items they need and how many they have but there should be a little more... The description will be shown above the objectives so for my task I would have put 'Farmer Joe has asked me to grab some apples and oranges in the field, I need to get 10 of each.' and the quest system would do the rest but give the player a nice explanation of the task at hand with the objectives as well."
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdHelp2_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdNextStep_Click()
Dim i As Long
   On Error GoTo errorhandler
   fraStep1.Visible = False
   fraStep2.Visible = False
   fraStep3.Visible = False
   fraStep4.Visible = False
   cmdNextStep.Enabled = True
   cmdPrevStep.Enabled = True
    Select Case QuestEditorPage
        Case 1
            fraStep2.Visible = True
            cmdPrevStep.Enabled = True
            QuestEditorPage = 2
            QuestEditorInitPage
        Case 2
            fraStep3.Visible = True
            QuestEditorPage = 3
            QuestEditorTask = 1
            QuestEditorInitPage
            QuestEditorInitTask
        Case 3
            fraStep4.Visible = True
            QuestEditorPage = 4
            QuestEditorInitPage
            cmdNextStep.Enabled = False
   
    End Select
   
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdNextStep_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdNextTask_Click()


   On Error GoTo errorhandler
    If QuestEditorTask < 10 Then
        QuestEditorTask = QuestEditorTask + 1
        scrlCurTask.Value = QuestEditorTask
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdNextTask_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdPrevStep_Click()
Dim i As Long

   On Error GoTo errorhandler

   fraStep1.Visible = False
   fraStep2.Visible = False
   fraStep3.Visible = False
   fraStep4.Visible = False
   cmdNextStep.Enabled = True
   cmdPrevStep.Enabled = True
    Select Case QuestEditorPage
        Case 2
            QuestEditorPage = 1
            fraStep1.Visible = True
            cmdPrevStep.Enabled = False
            QuestEditorInitPage
        Case 3
            QuestEditorPage = 2
            fraStep2.Visible = True
            QuestEditorInitPage
        Case 4
            fraStep3.Visible = True
            QuestEditorPage = 3
            QuestEditorTask = 1
            QuestEditorInitPage
            QuestEditorInitTask
            
   
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdPrevStep_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdPrevTask_Click()


   On Error GoTo errorhandler
    If QuestEditorTask > 1 Then
        QuestEditorTask = QuestEditorTask - 1
        scrlCurTask.Value = QuestEditorTask
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdPrevTask_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_Load()
Dim i As Long, X As Long
   On Error GoTo errorhandler
   
   EditorIndex = 0
   
    scrlCurTask.max = MAX_TASKS
    cmbClassReq.Clear
    If Max_Classes > 0 Then
        For i = 1 To Max_Classes
            cmbClassReq.AddItem i & ". " & Trim$(Class(i).Name)
        Next
        cmbClassReq.ListIndex = 0
    End If
    
    scrlLevelReq.max = MAX_LEVELS
    
    cmbQuestReq.Clear
    If MAX_QUESTS > 0 Then
        For i = 1 To MAX_QUESTS
            cmbQuestReq.AddItem i & ". " & Trim$(quest(i).Name)
        Next
        cmbQuestReq.ListIndex = 0
    End If
    
    cmbItemReq.Clear
    If MAX_ITEMS > 0 Then
        For i = 1 To MAX_ITEMS
            cmbItemReq.AddItem i & ". " & Trim$(Item(i).Name)
        Next
        cmbItemReq.ListIndex = 0
    End If
    
    cmbPlayerVarReq.Clear
    If MAX_VARIABLES > 0 Then
        For i = 1 To MAX_VARIABLES
            cmbPlayerVarReq.AddItem i & ". " & Trim$(Variables(i))
        Next
        cmbPlayerVarReq.ListIndex = 0
    End If

    cmbPlayerSwitchReq.Clear
    If MAX_SWITCHES > 0 Then
        For i = 1 To MAX_SWITCHES
            cmbPlayerSwitchReq.AddItem i & ". " & Trim$(Switches(i))
        Next
        cmbPlayerSwitchReq.ListIndex = 0
    End If
    
    For i = 0 To 3
        cmbGiveItem(i).Clear
        cmbTakeItem(i).Clear
        cmbGiveItem(i).AddItem "None"
        cmbTakeItem(i).AddItem "None"
        If MAX_ITEMS > 0 Then
            For X = 1 To MAX_ITEMS
                cmbGiveItem(i).AddItem X & ". " & Trim$(Item(X).Name)
                cmbTakeItem(i).AddItem X & ". " & Trim$(Item(X).Name)
            Next
            cmbGiveItem(i).ListIndex = 0
            cmbTakeItem(i).ListIndex = 0
        End If
    Next
    
    QuestEditorTask = 1
    
    For i = 1 To 4
        cmbKillNPC(i).Clear
        cmbKillNPC(i).AddItem "None"
        If MAX_NPCS > 0 Then
            For X = 1 To MAX_NPCS
                cmbKillNPC(i).AddItem X & ". " & Trim$(Npc(X).Name)
            Next
            cmbKillNPC(i).ListIndex = 0
        End If
    Next
    
    For i = 1 To 4
        cmbAquireItem(i).Clear
        cmbAquireItem(i).AddItem "None"
        If MAX_ITEMS > 0 Then
            For X = 1 To MAX_ITEMS
                cmbAquireItem(i).AddItem X & ". " & Trim$(Item(X).Name)
            Next
            cmbAquireItem(i).ListIndex = 0
        End If
    Next
    
    For i = 1 To 4
        cmbAquireDeliverItem(i).Clear
        cmbAquireDeliverItem(i).AddItem "None"
        If MAX_ITEMS > 0 Then
            For X = 1 To MAX_ITEMS
                cmbAquireDeliverItem(i).AddItem X & ". " & Trim$(Item(X).Name)
            Next
            cmbAquireDeliverItem(i).ListIndex = 0
        End If
    Next
    
    scrlGotoMap.max = MAX_MAPS
    
    For i = 1 To 4
        cmbGatherResource(i).Clear
        cmbGatherResource(i).AddItem "None"
        If MAX_RESOURCES > 0 Then
            For X = 1 To MAX_RESOURCES
                cmbGatherResource(i).AddItem X & ". " & Trim$(Resource(X).Name)
            Next
            cmbGatherResource(i).ListIndex = 0
        End If
    Next
    
    For i = 1 To 3
        cmbRewardItem(i).Clear
        cmbRewardItem(i).AddItem "None"
        If MAX_ITEMS > 0 Then
            For X = 1 To MAX_ITEMS
                cmbRewardItem(i).AddItem X & ". " & Trim$(Item(X).Name)
            Next
        End If
        cmbRewardItem(i).ListIndex = 0
    Next
    
    For i = 1 To 2
        cmbSpellReward(i).Clear
        cmbSpellReward(i).AddItem "None"
        If MAX_SPELLS > 0 Then
            For X = 1 To MAX_SPELLS
                cmbSpellReward(i).AddItem X & ". " & Trim$(spell(X).Name)
            Next
        End If
        cmbSpellReward(i).ListIndex = 0
    Next
    
    cmbSetPlayerVar.Clear
    cmbSetPlayerVar.AddItem "None"
    If MAX_VARIABLES > 0 Then
        For i = 1 To MAX_VARIABLES
            cmbSetPlayerVar.AddItem i & ". " & Trim$(Variables(i))
        Next
    End If
    cmbSetPlayerVar.ListIndex = 0
    
    cmbSetPlayerSwitch.Clear
    cmbSetPlayerSwitch.AddItem "None"
    If MAX_SWITCHES > 0 Then
        For i = 1 To MAX_SWITCHES
            cmbSetPlayerSwitch.AddItem i & ". " & Trim$(Switches(i))
        Next
    End If
    cmbSetPlayerSwitch.ListIndex = 0
    
    If EditorIndex = 0 Then EditorIndex = 1

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSave_Click()

   On Error GoTo errorhandler

    If LenB(Trim$(txtQuestName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        QuestEditorOk
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCancel_Click()

   On Error GoTo errorhandler

    QuestEditorCancel


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

   On Error GoTo errorhandler

    ClearQuest EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    QuestEditorInit


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lstIndex_Click()

   On Error GoTo errorhandler

    QuestEditorInit


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlAfterMap_Change()


   On Error GoTo errorhandler
    quest(EditorIndex).AfterMap = scrlAfterMap.Value
    lblAfterMap.Caption = "Map: " & scrlAfterMap.Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAfterMap_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlAfterX_Change()


   On Error GoTo errorhandler
    quest(EditorIndex).AfterX = scrlAfterX.Value
    lblAfterX.Caption = "X: " & scrlAfterX.Value
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAfterX_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlAfterY_Change()


   On Error GoTo errorhandler
    quest(EditorIndex).AfterY = scrlAfterY.Value
    lblAfterY.Caption = "Y: " & scrlAfterY.Value
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAfterY_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlAquireDeliverItemVal_Change(Index As Integer)


   On Error GoTo errorhandler
    quest(EditorIndex).Task(QuestEditorTask).data(4 + Index) = scrlAquireDeliverItemVal(Index).Value
    lblAquireDeliverItemVal(Index).Caption = "x" & scrlAquireDeliverItemVal(Index).Value
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAquireDeliverItemVal_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlAquireItemVal_Change(Index As Integer)


   On Error GoTo errorhandler
    quest(EditorIndex).Task(QuestEditorTask).data(4 + Index) = scrlAquireItemVal(Index).Value
    lblAquireItemVal(Index).Caption = "x" & scrlAquireItemVal(Index).Value
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAquireItemVal_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlGatherResourceAmount_Change(Index As Integer)


   On Error GoTo errorhandler
    quest(EditorIndex).Task(QuestEditorTask).data(4 + Index) = scrlGatherResourceAmount(Index).Value
    lblGatherResourceAmount(Index).Caption = "x" & scrlGatherResourceAmount(Index).Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlGatherResourceAmount_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlGiveExp_Change()
   On Error GoTo errorhandler
   
    lblGiveExp.Caption = "Give Exp: " & scrlGiveExp.Value
    quest(EditorIndex).GiveExp = scrlGiveExp.Value

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlGiveExp_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlGiveItemVal_Change(Index As Integer)
   On Error GoTo errorhandler
    
    quest(EditorIndex).GiveItemBefore(Index).Value = scrlGiveItemVal(Index).Value
    lblGiveItemVal(Index).Caption = "x" & scrlGiveItemVal(Index).Value

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlGiveItemVal_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlGotoMap_Change()
   On Error GoTo errorhandler

    lblGotoMap.Caption = "x" & scrlGotoMap.Value
    quest(EditorIndex).Task(QuestEditorTask).data(1) = scrlGotoMap.Value

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlGotoMap_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlItemReqVal_Change()
   On Error GoTo errorhandler

    lblItemReqVal.Caption = "x" & scrlItemReqVal.Value
    quest(EditorIndex).RequiredItemVal = scrlItemReqVal.Value

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlItemReqVal_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlKillNPCCount_Change(Index As Integer)


   On Error GoTo errorhandler
    quest(EditorIndex).Task(QuestEditorTask).data(4 + Index) = scrlKillNPCCount(Index).Value
    lblKillNPCValue(Index).Caption = "x" & scrlKillNPCCount(Index).Value
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlKillNPCCount_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlKillPlayer_Change()

   On Error GoTo errorhandler

    quest(EditorIndex).Task(QuestEditorTask).data(1) = scrlKillPlayer.Value
    lblKillPlayerCount.Caption = "x" & scrlKillPlayer.Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlKillPlayer_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlLevelReq_Change()

   On Error GoTo errorhandler

    lblLevelReq.Caption = "Level: " & scrlLevelReq.Value
    quest(EditorIndex).RequiredLevel = scrlLevelReq.Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlQuestCompleteCount_Change()

   On Error GoTo errorhandler

    lblQuestComplete.Caption = scrlQuestCompleteCount.Value
    quest(EditorIndex).RequiredQuestCount = scrlQuestCompleteCount.Value

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlQuestCompleteCount_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlRewardItemVal_Change(Index As Integer)


   On Error GoTo errorhandler
    quest(EditorIndex).RewardItem(Index).Value = scrlRewardItemVal(Index).Value
    lblRewardItemVal(Index).Caption = "x" & scrlRewardItemVal(Index).Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlRewardItemVal_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlTakeItemVal_Change(Index As Integer)


   On Error GoTo errorhandler
    quest(EditorIndex).TakeItemBefore(Index).Value = scrlTakeItemVal(Index).Value
    lblTakeItemVal(Index).Caption = "x" & scrlTakeItemVal(Index).Value
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlTakeItemVal_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlCurTask_Change()


   On Error GoTo errorhandler
    QuestEditorTask = scrlCurTask.Value
    QuestEditorInitTask


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlCurTask_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlTeleMap_Change()

   On Error GoTo errorhandler

    lblTeleportMap.Caption = "Map: " & scrlTeleMap.Value
    quest(EditorIndex).BeforeMap = scrlTeleMap.Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlTeleMap_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlTeleX_Change()

   On Error GoTo errorhandler

    lblTeleX.Caption = "X: " & scrlTeleX.Value
    quest(EditorIndex).BeforeX = scrlTeleX.Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlTeleX_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlTeleY_Change()

   On Error GoTo errorhandler

    lblTeleY.Caption = "Y: " & scrlTeleY.Value
    quest(EditorIndex).BeforeY = scrlTeleY.Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlTeleY_Change", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtAfterQuest_Validate(Cancel As Boolean)


   On Error GoTo errorhandler
    quest(EditorIndex).QuestLogAfterDesc = txtAfterQuest.Text
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtAfterQuest_Validate", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtAquireDeliverEventName_Validate(Cancel As Boolean)

   On Error GoTo errorhandler

    quest(EditorIndex).Task(QuestEditorTask).Text(1) = Trim$(txtAquireDeliverEventName.Text)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtAquireDeliverEventName_Validate", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtBeforeQuest_Validate(Cancel As Boolean)

   On Error GoTo errorhandler

    quest(EditorIndex).QuestLogBeforeDesc = txtBeforeQuest.Text


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtBeforeQuest_Validate", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtEventTask_Validate(Cancel As Boolean)


   On Error GoTo errorhandler
   
    quest(EditorIndex).Task(QuestEditorTask).Text(1) = Trim$(txtEventTask.Text)
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtEventTask_Validate", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtGoToMap_Validate(Cancel As Boolean)


   On Error GoTo errorhandler
    quest(EditorIndex).Task(QuestEditorTask).Text(1) = Trim$(txtGoToMap.Text)
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtGoToMap_Validate", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtQuestName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

   On Error GoTo errorhandler

    tmpIndex = lstIndex.ListIndex
    quest(EditorIndex).Name = Trim$(txtQuestName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtQuestName_Validate", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtQuestOffer_Validate(Cancel As Boolean)


   On Error GoTo errorhandler
    quest(EditorIndex).QuestDesc = Trim$(txtQuestOffer.Text)
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtQuestOffer_Validate", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtSetPlayerVarValue_Validate(Cancel As Boolean)


   On Error GoTo errorhandler
    If IsNumeric(txtSetPlayerVarValue.Text) Then
        quest(EditorIndex).SetPlayerVarValue = Val(txtSetPlayerVarValue.Text)
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtSetPlayerVarValue_Validate", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtTaskDesc_Validate(Cancel As Boolean)


   On Error GoTo errorhandler
    quest(EditorIndex).Task(QuestEditorTask).TaskDesc = Trim$(txtTaskDesc.Text)
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtTaskDesc_Validate", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtVariableReq_Validate(Cancel As Boolean)
   On Error GoTo errorhandler

    quest(EditorIndex).RequiredVariableCompareTo = Val(Trim$(txtVariableReq.Text))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtVariableReq_Validate", "frmEditor_Quest", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
