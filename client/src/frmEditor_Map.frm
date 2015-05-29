VERSION 5.00
Begin VB.Form frmEditor_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14655
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
   Icon            =   "frmEditor_Map.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   977
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picAttributes 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   7440
      ScaleHeight     =   7215
      ScaleWidth      =   7095
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Frame fraRandomDungeon 
         Caption         =   "Random Dungeon"
         Height          =   1935
         Left            =   1800
         TabIndex        =   102
         Top             =   2280
         Width           =   3375
         Begin VB.HScrollBar scrlRandomDungeon 
            Height          =   255
            Left            =   120
            Max             =   50
            Min             =   1
            TabIndex        =   105
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.HScrollBar scrlFloorNum 
            Height          =   255
            Left            =   120
            Max             =   100
            Min             =   1
            TabIndex        =   104
            Top             =   1080
            Value           =   1
            Width           =   3135
         End
         Begin VB.CommandButton cmdRandomDungeon 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   103
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblRandomDungeon 
            Caption         =   "Dungeon No: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label lblFloor 
            Caption         =   "Floor Num (1 as default): 1"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   840
            Width           =   3015
         End
      End
      Begin VB.Frame fraBuyHouse 
         Caption         =   "Buy House"
         Height          =   1935
         Left            =   1800
         TabIndex        =   94
         Top             =   2280
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlBuyHouse 
            Height          =   255
            Left            =   120
            Max             =   100
            Min             =   1
            TabIndex        =   96
            Top             =   720
            Value           =   1
            Width           =   3135
         End
         Begin VB.CommandButton cmdHouseTileOk 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   95
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblHouseName 
            Caption         =   "House:"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame fraSoundEffect 
         Caption         =   "Sound Effect"
         Height          =   1455
         Left            =   1800
         TabIndex        =   87
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdSoundEffectOk 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   89
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox cmbSoundEffect 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3332
            Left            =   240
            List            =   "frmEditor_Map.frx":3342
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraSlide 
         Caption         =   "Slide"
         Height          =   1455
         Left            =   1800
         TabIndex        =   83
         Top             =   2640
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbSlide 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":335D
            Left            =   240
            List            =   "frmEditor_Map.frx":336D
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdSlide 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   84
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraTrap 
         Caption         =   "Trap"
         Height          =   1575
         Left            =   1800
         TabIndex        =   79
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlTrap 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   81
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton cmdTrap 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   80
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblTrap 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraHeal 
         Caption         =   "Heal"
         Height          =   1815
         Left            =   1800
         TabIndex        =   74
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbHeal 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3388
            Left            =   240
            List            =   "frmEditor_Map.frx":3392
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdHeal 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   76
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlHeal 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   75
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblHeal 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame fraNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   2655
         Left            =   1800
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstNpc 
            Height          =   780
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlNpcDir 
            Height          =   255
            Left            =   240
            Max             =   3
            TabIndex        =   35
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton cmdNpcSpawn 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   34
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblNpcDir 
            Caption         =   "Direction: Up"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame fraResource 
         Caption         =   "Object"
         Height          =   1695
         Left            =   1800
         TabIndex        =   27
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdResourceOk 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   30
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   29
            Top             =   600
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblResource 
            Caption         =   "Object:"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraMapWarp 
         Caption         =   "Map Warp"
         Height          =   2775
         Left            =   1800
         TabIndex        =   56
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapWarp 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   63
            Top             =   2160
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapWarpY 
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1680
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarpX 
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarp 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   58
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label lblMapWarpY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblMapWarpX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblMapWarp 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraShop 
         Caption         =   "Shop"
         Height          =   1335
         Left            =   1920
         TabIndex        =   64
         Top             =   2640
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdShop 
            Caption         =   "Accept"
            Height          =   375
            Left            =   960
            TabIndex        =   66
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbShop 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraKeyOpen 
         Caption         =   "Key Open"
         Height          =   2055
         Left            =   1800
         TabIndex        =   50
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdKeyOpen 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   55
            Top             =   1440
            Width           =   1215
         End
         Begin VB.HScrollBar scrlKeyY 
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlKeyX 
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblKeyY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblKeyX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame fraMapKey 
         Caption         =   "Map Key"
         Height          =   1815
         Left            =   1800
         TabIndex        =   44
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.PictureBox picMapKey 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   49
            Top             =   600
            Width           =   480
         End
         Begin VB.CommandButton cmdMapKey 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   48
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkMapKey 
            Caption         =   "Take key away upon use."
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMapKey 
            Height          =   255
            Left            =   120
            Max             =   5
            Min             =   1
            TabIndex        =   46
            Top             =   600
            Value           =   1
            Width           =   2535
         End
         Begin VB.Label lblMapKey 
            Caption         =   "Item: None"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraMapItem 
         Caption         =   "Map Item"
         Height          =   1815
         Left            =   1800
         TabIndex        =   38
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapItem 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1200
            TabIndex        =   43
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapItemValue 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   42
            Top             =   840
            Value           =   1
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMapItem 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   41
            Top             =   480
            Value           =   1
            Width           =   2535
         End
         Begin VB.PictureBox picMapItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   40
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lblMapItem 
            Caption         =   "Item: None x0"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   3135
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type"
      Height          =   1335
      Left            =   5760
      TabIndex        =   22
      Top             =   5520
      Width           =   1455
      Begin VB.OptionButton optEvent 
         Alignment       =   1  'Right Justify
         Caption         =   "Events"
         Height          =   255
         Left            =   360
         TabIndex        =   86
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "Block"
         Height          =   255
         Left            =   480
         TabIndex        =   69
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optAttribs 
         Alignment       =   1  'Right Justify
         Caption         =   "Attributes"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optLayers 
         Alignment       =   1  'Right Justify
         Caption         =   "Layers"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.HScrollBar scrlPictureX 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5400
      Width           =   5295
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   120
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   14
      Top             =   120
      Width           =   5280
   End
   Begin VB.VScrollBar scrlPictureY 
      Height          =   5295
      Left            =   5400
      Max             =   255
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame fraTileSet 
      Caption         =   "Tileset: 0"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   5535
      Begin VB.HScrollBar scrlTileSet 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   1
         Top             =   360
         Value           =   1
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Attributes"
      Height          =   5415
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optRandomDungeon 
         Caption         =   "Dungeon"
         Height          =   270
         Left            =   120
         TabIndex        =   101
         Top             =   4320
         Width           =   1215
      End
      Begin VB.OptionButton optInstance 
         Caption         =   "Instance"
         Height          =   270
         Left            =   120
         TabIndex        =   100
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton optHouse 
         Caption         =   "Buy House"
         Height          =   270
         Left            =   120
         TabIndex        =   93
         Top             =   3840
         Width           =   1215
      End
      Begin VB.OptionButton optSound 
         Caption         =   "Sound"
         Height          =   270
         Left            =   120
         TabIndex        =   90
         Top             =   3600
         Width           =   1215
      End
      Begin VB.OptionButton optSlide 
         Caption         =   "Slide"
         Height          =   270
         Left            =   120
         TabIndex        =   73
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optTrap 
         Caption         =   "Trap"
         Height          =   270
         Left            =   120
         TabIndex        =   72
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Heal"
         Height          =   270
         Left            =   120
         TabIndex        =   71
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Bank"
         Height          =   270
         Left            =   120
         TabIndex        =   70
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optShop 
         Caption         =   "Shop"
         Height          =   270
         Left            =   120
         TabIndex        =   67
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   270
         Left            =   120
         TabIndex        =   32
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optDoor 
         Caption         =   "Door"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optResource 
         Caption         =   "Resource"
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optKeyOpen 
         Caption         =   "Key Open"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Blocked"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Warp"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         Height          =   390
         Left            =   120
         TabIndex        =   6
         Top             =   4920
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Npc Avoid"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optKey 
         Caption         =   "Key"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      Height          =   5415
      Left            =   5760
      TabIndex        =   15
      Top             =   120
      Width           =   1455
      Begin VB.HScrollBar scrlLayerNum 
         Height          =   255
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   98
         Top             =   1320
         Value           =   1
         Width           =   1215
      End
      Begin VB.HScrollBar scrlAutotile 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   91
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   390
         Left            =   120
         TabIndex        =   17
         Top             =   4800
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Ground"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblLayerNum 
         Alignment       =   2  'Center
         Caption         =   "Mask: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblAutotile 
         Alignment       =   2  'Center
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   3720
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Drag mouse to select multiple tiles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   5760
      Width           =   5535
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHeal_Click()

   On Error GoTo errorhandler

    MapEditorHealType = cmbHeal.ListIndex + 1
    MapEditorHealAmount = scrlHeal.Value
    picAttributes.Visible = False
    fraHeal.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdHouseTileOk_Click()

   On Error GoTo errorhandler

    HouseTileIndex = scrlBuyHouse.Value
    picAttributes.Visible = False
    fraBuyHouse.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdHouseTileOk_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdKeyOpen_Click()

   On Error GoTo errorhandler

    KeyOpenEditorX = scrlKeyX.Value
    KeyOpenEditorY = scrlKeyY.Value
    picAttributes.Visible = False
    fraKeyOpen.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdKeyOpen_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdMapItem_Click()

   On Error GoTo errorhandler

    ItemEditorNum = scrlMapItem.Value
    ItemEditorValue = scrlMapItemValue.Value
    picAttributes.Visible = False
    fraMapItem.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdMapItem_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdMapKey_Click()

   On Error GoTo errorhandler

    KeyEditorNum = scrlMapKey.Value
    KeyEditorTake = chkMapKey.Value
    picAttributes.Visible = False
    fraMapKey.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdMapKey_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdMapWarp_Click()

   On Error GoTo errorhandler

    EditorWarpMap = scrlMapWarp.Value
    EditorWarpX = scrlMapWarpX.Value
    EditorWarpY = scrlMapWarpY.Value
    picAttributes.Visible = False
    fraMapWarp.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdMapWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdNpcSpawn_Click()

   On Error GoTo errorhandler

    SpawnNpcNum = lstNpc.ListIndex + 1
    SpawnNpcDir = scrlNpcDir.Value
    picAttributes.Visible = False
    fraNpcSpawn.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdNpcSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdRandomDungeon_Click()
    

   On Error GoTo errorhandler

    MapEditorRandomDungeon = scrlRandomDungeon.Value
    MapEditorFloorNum = scrlFloorNum.Value
    picAttributes.Visible = False
    fraResource.Visible = False


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdRandomDungeon_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdResourceOk_Click()

   On Error GoTo errorhandler

    ResourceEditorNum = scrlResource.Value
    picAttributes.Visible = False
    fraResource.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdResourceOk_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdShop_Click()

   On Error GoTo errorhandler

    EditorShop = cmbShop.ListIndex
    picAttributes.Visible = False
    fraShop.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdShop_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdSlide_Click()

   On Error GoTo errorhandler

    MapEditorSlideDir = cmbSlide.ListIndex
    picAttributes.Visible = False
    fraSlide.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdSoundEffectOk_Click()

   On Error GoTo errorhandler

    MapEditorSound = soundCache(cmbSoundEffect.ListIndex + 1)
    picAttributes.Visible = False
    fraSoundEffect.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSoundEffectOk_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdTrap_Click()

   On Error GoTo errorhandler

    MapEditorHealAmount = scrlTrap.Value
    picAttributes.Visible = False
    fraTrap.Visible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Form_Load()
Dim i As Long
    ' move the entire attributes box on screen

   On Error GoTo errorhandler

    picAttributes.Left = 8
    picAttributes.Top = 8
    scrlBuyHouse.max = MAX_HOUSES
    scrlBuyHouse.Value = 1
    GraphicSelX = 0
    GraphicSelY = 0
    GraphicSelX2 = 0
    GraphicSelY2 = 0
    PopulateLists
    cmbSoundEffect.Clear
    scrlLayerNum.Value = 1
    scrlLayerNum.Enabled = False
    optLayer(1).Value = 1
    lblLayerNum.Caption = "Ground"
    If UBound(soundCache) > 0 Then
        For i = 1 To UBound(soundCache)
            cmbSoundEffect.AddItem (soundCache(i))
        Next
        cmbSoundEffect.ListIndex = 0
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optDoor_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    scrlMapWarp.max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.max = MAX_BYTE
    scrlMapWarpY.max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optDoor_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optHeal_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraHeal.Visible = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optHouse_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraBuyHouse.Visible = True
    scrlBuyHouse.max = MAX_HOUSES
    scrlBuyHouse.Value = 1
    scrlBuyHouse_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optHouse_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optInstance_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optInstance_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub optLayer_Click(Index As Integer)
    Select Case Index
        Case 1
            scrlLayerNum.Value = 1
            scrlLayerNum.Enabled = False
            lblLayerNum.Caption = "Ground"
        Case 2
            scrlLayerNum.Value = 1
            scrlLayerNum.Enabled = True
            lblLayerNum.Caption = "Mask: 1"
        Case 4
            scrlLayerNum.Value = 1
            scrlLayerNum.Enabled = True
            lblLayerNum.Caption = "Fringe: 1"
    End Select
End Sub

Private Sub optLayers_Click()


   On Error GoTo errorhandler

    If optLayers.Value Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optLayers_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optAttribs_Click()

   On Error GoTo errorhandler

    If optAttribs.Value Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optAttribs_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optNpcSpawn_Click()
Dim n As Long


   On Error GoTo errorhandler

    lstNpc.Clear
    For n = 1 To MAX_MAP_NPCS
        If Map.Npc(n) > 0 Then
            lstNpc.AddItem n & ": " & Npc(Map.Npc(n)).Name
        Else
            lstNpc.AddItem n & ": No Npc"
        End If
    Next n
    scrlNpcDir.Value = 0
    lstNpc.ListIndex = 0
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraNpcSpawn.Visible = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optNpcSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optRandomDungeon_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraRandomDungeon.Visible = True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optRandomDungeon_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub optResource_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraResource.Visible = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optResource_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optShop_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraShop.Visible = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optShop_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optSlide_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraSlide.Visible = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optSound_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraSoundEffect.Visible = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optSound_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optTrap_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraTrap.Visible = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdSend_Click()

   On Error GoTo errorhandler

    If frmEditor_Events.Visible Then
        If MsgBox("The event editor is open. Continuing to send the map will discard the changes to the event you are editing. Continue?", vbYesNo) = vbYes Then
            Call MapEditorSend
        End If
    Else
        Call MapEditorSend
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSend_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdCancel_Click()

   On Error GoTo errorhandler

    frmEditor_Events.Visible = False
    Call MapEditorCancel




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdProperties_Click()

   On Error GoTo errorhandler

    Load frmEditor_MapProperties
    MapEditorProperties
    frmEditor_MapProperties.Show vbModal




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdProperties_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optWarp_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    scrlMapWarp.max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.max = MAX_BYTE
    scrlMapWarpY.max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optItem_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapItem.Visible = True

    scrlMapItem.max = MAX_ITEMS
    scrlMapItem.Value = 1
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optItem_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optKey_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapKey.Visible = True
    scrlMapKey.max = MAX_ITEMS
    scrlMapKey.Value = 1
    chkMapKey.Value = 1
    lblMapKey.Caption = "Item: " & Trim$(Item(scrlMapKey.Value).Name)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optKey_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optKeyOpen_Click()

   On Error GoTo errorhandler

    ClearAttributeDialogue
    fraKeyOpen.Visible = True
    picAttributes.Visible = True
    scrlKeyX.max = Map.MaxX
    scrlKeyY.max = Map.MaxY
    scrlKeyX.Value = 0
    scrlKeyY.Value = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optKeyOpen_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdFill_Click()

   On Error GoTo errorhandler

    MapEditorFillLayer




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdFill_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdClear_Click()

   On Error GoTo errorhandler

    Call MapEditorClearLayer




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdClear_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdClear2_Click()

   On Error GoTo errorhandler

    Call MapEditorClearAttribs




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdClear2_Click", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo errorhandler

    X = X + (frmEditor_Map.scrlPictureX.Value * 32)
    Y = Y + (frmEditor_Map.scrlPictureY.Value * 32)
    Call MapEditorChooseTile(Button, X, Y)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "picBack_MouseDown", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo errorhandler

    X = X + (frmEditor_Map.scrlPictureX.Value * 32)
    Y = Y + (frmEditor_Map.scrlPictureY.Value * 32)
    Call MapEditorDrag(Button, X, Y)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "picBack_MouseMove", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlAutotile_Change()

   On Error GoTo errorhandler

    Select Case scrlAutotile.Value
        Case 0 ' normal
            lblAutotile.Caption = "Normal"
        Case 1 ' autotile
            lblAutotile.Caption = "Autotile (VX)"
        Case 2 ' fake autotile
            lblAutotile.Caption = "Fake (VX)"
        Case 3 ' animated
            lblAutotile.Caption = "Animated (VX)"
        Case 4 ' cliff
            lblAutotile.Caption = "Cliff (VX)"
        Case 5 ' waterfall
            lblAutotile.Caption = "Waterfall (VX)"
        Case 6 ' autotile
            lblAutotile.Caption = "Autotile (XP)"
        Case 7 ' fake autotile
            lblAutotile.Caption = "Fake (XP)"
        Case 8 ' animated
            lblAutotile.Caption = "Animated (XP)"
        Case 9 ' cliff
            lblAutotile.Caption = "Cliff (XP)"
        Case 10 ' waterfall
            lblAutotile.Caption = "Waterfall (XP)"
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAutotile_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlBuyHouse_Change()

   On Error GoTo errorhandler
    
    lblHouseName.Caption = scrlBuyHouse.Value & ". " & HouseConfig(scrlBuyHouse.Value).ConfigName


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlBuyHouse_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlFloorNum_Change()
    

   On Error GoTo errorhandler

    lblFloor.Caption = "Floor Num (1 as default): " & scrlFloorNum.Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlFloorNum_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlHeal_Change()

   On Error GoTo errorhandler

    lblHeal.Caption = "Amount: " & scrlHeal.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlHeal_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlKeyX_Change()

   On Error GoTo errorhandler

    lblKeyX.Caption = "X: " & scrlKeyX.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlKeyX_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlKeyX_Scroll()

   On Error GoTo errorhandler

    scrlKeyX_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlKeyX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlKeyY_Change()

   On Error GoTo errorhandler

    lblKeyY.Caption = "Y: " & scrlKeyY.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlKeyY_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlKeyY_Scroll()

   On Error GoTo errorhandler

    scrlKeyY_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlKeyY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlLayerNum_Change()


   On Error GoTo errorhandler
    If optLayer(2).Value = True Then
        lblLayerNum.Caption = "Mask: " & scrlLayerNum.Value
    ElseIf optLayer(4).Value = True Then
        lblLayerNum.Caption = "Fringe: " & scrlLayerNum.Value
    Else
    
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlLayerNum_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlRandomDungeon_Change()
    

   On Error GoTo errorhandler

    lblRandomDungeon.Caption = "Dungeon No: " & scrlRandomDungeon.Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlRandomDungeon_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlTrap_Change()

   On Error GoTo errorhandler

    lblTrap.Caption = "Amount: " & scrlTrap.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlTrap_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapItem_Change()

   On Error GoTo errorhandler

        If Item(scrlMapItem.Value).Stackable = 1 Then
        scrlMapItemValue.Enabled = True
    Else
        scrlMapItemValue.Value = 1
        scrlMapItemValue.Enabled = False
    End If

    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapItem_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapItem_Scroll()

   On Error GoTo errorhandler

    scrlMapItem_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapItem_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapItemValue_Change()

   On Error GoTo errorhandler

    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapItemValue_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapItemValue_Scroll()

   On Error GoTo errorhandler

    scrlMapItemValue_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapItemValue_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapKey_Change()

   On Error GoTo errorhandler

    lblMapKey.Caption = "Item: " & Trim$(Item(scrlMapKey.Value).Name)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapKey_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapKey_Scroll()

   On Error GoTo errorhandler

    scrlMapKey_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapKey_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapWarp_Change()

   On Error GoTo errorhandler

    lblMapWarp.Caption = "Map: " & scrlMapWarp.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapWarp_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapWarp_Scroll()

   On Error GoTo errorhandler

    scrlMapWarp_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapWarp_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapWarpX_Change()

   On Error GoTo errorhandler

    lblMapWarpX.Caption = "X: " & scrlMapWarpX.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapWarpX_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapWarpX_Scroll()

   On Error GoTo errorhandler

    scrlMapWarpX_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapWarpX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapWarpY_Change()

   On Error GoTo errorhandler

    lblMapWarpY.Caption = "Y: " & scrlMapWarpY.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapWarpY_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMapWarpY_Scroll()

   On Error GoTo errorhandler

    scrlMapWarpY_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMapWarpY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlNpcDir_Change()

   On Error GoTo errorhandler

    Select Case scrlNpcDir.Value
        Case DIR_DOWN
            lblNpcDir = "Direction: Down"
        Case DIR_UP
            lblNpcDir = "Direction: Up"
        Case DIR_LEFT
            lblNpcDir = "Direction: Left"
        Case DIR_RIGHT
            lblNpcDir = "Direction: Right"
    End Select




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlNpcDir_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlNpcDir_Scroll()

   On Error GoTo errorhandler

    scrlNpcDir_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlNpcDir_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlResource_Change()

   On Error GoTo errorhandler

    lblResource.Caption = "Resource: " & Resource(scrlResource.Value).Name




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlResource_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlResource_Scroll()

   On Error GoTo errorhandler

    scrlResource_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlResource_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlPictureX_Change()

   On Error GoTo errorhandler

    Call MapEditorTileScroll




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlPictureX_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlPictureY_Change()

   On Error GoTo errorhandler

    Call MapEditorTileScroll




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlPictureY_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlPictureX_Scroll()

   On Error GoTo errorhandler

    scrlPictureY_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlPictureX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlPictureY_Scroll()

   On Error GoTo errorhandler

    scrlPictureY_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlPictureY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlTileSet_Change()

   On Error GoTo errorhandler

    fraTileSet.Caption = "Tileset: " & scrlTileSet.Value
    If gTexture(Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Texture).IsLoaded = False Then
        LoadTexture1 Tex_Tileset(frmEditor_Map.scrlTileSet.Value)
    End If
    frmEditor_Map.scrlPictureY.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    MapEditorTileScroll
    EditorTileX = 0
    EditorTileY = 0
    EditorTileWidth = 1
    EditorTileHeight = 1




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlTileSet_Change", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlTileSet_Scroll()

   On Error GoTo errorhandler

    scrlTileSet_Change




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlTileSet_Scroll", "frmEditor_Map", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
