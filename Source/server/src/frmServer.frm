VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmServer.frx":1708A
   ScaleHeight     =   6600
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraControlPanel 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Server Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      TabIndex        =   94
      Top             =   1920
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CheckBox chkDisableRestart 
         BackColor       =   &H00000000&
         Caption         =   "Disable Remote Restart?"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   172
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CheckBox chkStaffOnly 
         BackColor       =   &H00000000&
         Caption         =   "Staff Only?"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   171
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00000000&
         Caption         =   "Server"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   154
         Top             =   2760
         Width           =   5895
         Begin VB.TextBox txtUpdateUrl 
            Height          =   285
            Left            =   1800
            TabIndex        =   159
            Text            =   "txtUpdateUrl"
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox txtDataFolder 
            Height          =   285
            Left            =   3600
            MaxLength       =   20
            TabIndex        =   156
            Text            =   "txtDataFolder"
            Top             =   160
            Width           =   2175
         End
         Begin VB.CommandButton cmdSaveDataFolder 
            Caption         =   "Save"
            Height          =   200
            Left            =   120
            TabIndex        =   155
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblUpdateHelp 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Click here for info on setting up and using the in-game updater!"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   170
            Top             =   960
            Width           =   5655
         End
         Begin VB.Label Label14 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "No URL will result in the client using default data."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   160
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label Label12 
            BackColor       =   &H00000000&
            Caption         =   "Update.ini URL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblDataFolder 
            BackColor       =   &H00000000&
            Caption         =   "Game Data Folder (Should be unique.)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   157
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Game Credits"
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   6240
         TabIndex        =   151
         Top             =   1800
         Width           =   2655
         Begin VB.TextBox txtCredits 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   153
            Text            =   "frmServer.frx":D868C
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton cmdSaveCredits 
            Caption         =   "Save Credits"
            Height          =   195
            Left            =   120
            TabIndex        =   152
            Top             =   1080
            Width           =   2415
         End
      End
      Begin VB.Frame fraNews 
         BackColor       =   &H00000000&
         Caption         =   "Game Menu News"
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   6240
         TabIndex        =   148
         Top             =   480
         Width           =   2655
         Begin VB.CommandButton cmdSaveNews 
            Caption         =   "Save News"
            Height          =   195
            Left            =   120
            TabIndex        =   150
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox txtNews 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   149
            Text            =   "frmServer.frx":D8694
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         Caption         =   "Map Report"
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   3840
         TabIndex        =   118
         Top             =   480
         Width           =   2295
         Begin VB.ListBox lstMaps 
            Height          =   1815
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame fraDatabase 
         BackColor       =   &H00000000&
         Caption         =   "Reload"
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   240
         TabIndex        =   109
         Top             =   480
         Width           =   1455
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Classes"
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   1920
            Width           =   1215
         End
      End
      Begin VB.Frame fraServer 
         BackColor       =   &H00000000&
         Caption         =   "Server"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   1920
         TabIndex        =   106
         Top             =   480
         Width           =   1815
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Shut Down"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkServerLog 
            BackColor       =   &H00000000&
            Caption         =   "Server Log"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "Mass-Name Maps"
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   1920
         TabIndex        =   97
         Top             =   1320
         Width           =   1815
         Begin VB.TextBox txtRMap 
            Height          =   285
            Left            =   720
            TabIndex        =   101
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtRMaps2 
            Height          =   285
            Left            =   720
            TabIndex        =   100
            Text            =   "0"
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton cmdReserveMaps 
            Caption         =   "Mass-Name"
            Height          =   195
            Left            =   120
            TabIndex        =   99
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtMapName 
            Height          =   285
            Left            =   720
            TabIndex        =   98
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "to"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   105
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "First #"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Last #"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdCloseControlPanel 
         Caption         =   "Close Control Panel"
         Height          =   255
         Left            =   6240
         TabIndex        =   95
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label lblNotifications 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Notifications...."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   174
         Top             =   4080
         Width           =   5655
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Control Panel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   96
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame fraConsole 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Server Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CommandButton cmdCloseConsole 
         Caption         =   "Close Console"
         Height          =   495
         Left            =   7320
         TabIndex        =   3
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtText 
         Height          =   2175
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   840
         Width           =   8655
      End
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   3120
         Width           =   8655
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Console"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblCPS 
         BackColor       =   &H00000000&
         Caption         =   "CPS: 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblCpsLock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "[Unlock]"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Frame fraHousing 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Server Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      TabIndex        =   76
      Top             =   1920
      Visible         =   0   'False
      Width           =   9015
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "House Setup"
         ForeColor       =   &H00FFFFFF&
         Height          =   2895
         Left            =   240
         TabIndex        =   79
         Top             =   600
         Width           =   8535
         Begin VB.ListBox lstHouses 
            Height          =   2400
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtHouseName 
            Height          =   255
            Left            =   4200
            TabIndex        =   86
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtBaseMap 
            Height          =   285
            Left            =   4200
            TabIndex        =   85
            Top             =   645
            Width           =   2655
         End
         Begin VB.TextBox txtHouseFurniture 
            Height          =   285
            Left            =   4200
            TabIndex        =   84
            Top             =   2085
            Width           =   2655
         End
         Begin VB.CommandButton cmdSaveHouse 
            Caption         =   "Save Changes"
            Height          =   375
            Left            =   6960
            TabIndex        =   83
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtHousePrice 
            Height          =   285
            Left            =   4200
            TabIndex        =   82
            Top             =   1725
            Width           =   2655
         End
         Begin VB.TextBox txtXEntrance 
            Height          =   285
            Left            =   4200
            TabIndex        =   81
            Top             =   1005
            Width           =   2655
         End
         Begin VB.TextBox txtYEntrance 
            Height          =   285
            Left            =   4200
            TabIndex        =   80
            Top             =   1365
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Name of House:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   93
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblHouseMap 
            BackStyle       =   0  'Transparent
            Caption         =   "Base Map:"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2520
            TabIndex        =   92
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Max Pieces of Furniture (0 for no max):"
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2520
            TabIndex        =   91
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblHousePrice 
            BackStyle       =   0  'Transparent
            Caption         =   "Price:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2520
            TabIndex        =   90
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrance X:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2520
            TabIndex        =   89
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrance Y:"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2520
            TabIndex        =   88
            Top             =   1320
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdCloseHousing 
         Caption         =   "Close Housing Setup"
         Height          =   495
         Left            =   7320
         TabIndex        =   77
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Player Housing Setup"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   78
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Frame fraAccount 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Server Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      TabIndex        =   120
      Top             =   1920
      Visible         =   0   'False
      Width           =   9015
      Begin VB.Frame fraLogin 
         BackColor       =   &H00000000&
         Caption         =   "Login"
         ForeColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   2040
         TabIndex        =   122
         Top             =   720
         Width           =   5175
         Begin VB.CommandButton cmdCancelLogin 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3720
            TabIndex        =   128
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CommandButton cmdLogin 
            Caption         =   "Login"
            Height          =   375
            Left            =   1440
            TabIndex        =   127
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtLPass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1440
            PasswordChar    =   "*"
            TabIndex        =   126
            Text            =   "txtLEmail"
            Top             =   720
            Width           =   3615
         End
         Begin VB.TextBox txtLEmail 
            Height          =   285
            Left            =   1440
            TabIndex        =   125
            Text            =   "txtLEmail"
            Top             =   330
            Width           =   3615
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmServer.frx":D869C
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   120
            TabIndex        =   129
            Top             =   1080
            Width           =   4935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   124
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   123
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.PictureBox picEditAccount 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   8895
         TabIndex        =   130
         Top             =   600
         Visible         =   0   'False
         Width           =   8895
         Begin VB.Frame Frame8 
            BackColor       =   &H00000000&
            Caption         =   "Account Info"
            ForeColor       =   &H00FFFFFF&
            Height          =   1815
            Left            =   0
            TabIndex        =   142
            Top             =   0
            Width           =   3375
            Begin VB.Label lblEmail 
               BackStyle       =   0  'Transparent
               Caption         =   "Email: xyz@eclipseorigins.com"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   146
               Top             =   240
               Width           =   3255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Password: N/A"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   145
               Top             =   480
               Width           =   3135
            End
            Begin VB.Label lblAccountFName 
               BackStyle       =   0  'Transparent
               Caption         =   "First Name:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   144
               Top             =   720
               Width           =   3255
            End
            Begin VB.Label lblAccountLName 
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   143
               Top             =   960
               Width           =   3255
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00000000&
            Caption         =   "Current Gold License(s)"
            ForeColor       =   &H00FFFFFF&
            Height          =   3015
            Left            =   3600
            TabIndex        =   136
            Top             =   0
            Width           =   5055
            Begin VB.ListBox lstLicenses 
               Height          =   1230
               Left            =   120
               TabIndex        =   139
               Top             =   480
               Width           =   4815
            End
            Begin VB.CommandButton cmdCopyKey 
               Caption         =   "Copy License Key to Clipboard"
               Height          =   495
               Left            =   120
               TabIndex        =   138
               Top             =   1800
               Width           =   2055
            End
            Begin VB.CommandButton cmdResetKey 
               Caption         =   "Reset License Key for selected License."
               Height          =   495
               Left            =   2880
               TabIndex        =   137
               Top             =   1800
               Width           =   2055
            End
            Begin VB.Label lblLicenseCount 
               BackStyle       =   0  'Transparent
               Caption         =   "License Count: 0"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   141
               Top             =   240
               Width           =   3135
            End
            Begin VB.Label lblLicenseInfo 
               BackStyle       =   0  'Transparent
               Caption         =   $"frmServer.frx":D8740
               ForeColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   120
               TabIndex        =   140
               Top             =   2400
               Width           =   4935
            End
         End
         Begin VB.Frame fraActivate 
            BackColor       =   &H00000000&
            Caption         =   "Redeem Gold License Code"
            ForeColor       =   &H00FFFFFF&
            Height          =   1215
            Left            =   0
            TabIndex        =   132
            Top             =   1800
            Width           =   3375
            Begin VB.TextBox txtActivateCode 
               Height          =   285
               Left            =   720
               TabIndex        =   134
               Text            =   "txtActivateCode"
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdActivate 
               Caption         =   "Activate"
               Height          =   375
               Left            =   960
               TabIndex        =   133
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Code:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   135
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdCloseAccount 
            Caption         =   "Close Account"
            Height          =   495
            Left            =   7200
            TabIndex        =   131
            Top             =   3240
            Width           =   1455
         End
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   121
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Server Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      TabIndex        =   161
      Top             =   1920
      Width           =   9015
      Begin VB.Label lblServerMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   0
         TabIndex        =   169
         Top             =   840
         Width           =   9015
      End
      Begin VB.Label lblServerMenuOpt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Console"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   168
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblServerMenuOpt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Player List"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   167
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblServerMenuOpt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "House Setup"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   166
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblServerMenuOpt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Control Panel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   165
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblServerMenuOpt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   6120
         TabIndex        =   164
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblServerMenuOpt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   7080
         TabIndex        =   163
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblServerMenuOpt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   7680
         TabIndex        =   162
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame fraPlayers 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Server Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CommandButton cmdClosePlayers 
         Caption         =   "Close Player List"
         Height          =   375
         Left            =   6960
         TabIndex        =   8
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Frame fraEditPlayer 
         BackColor       =   &H00000000&
         Caption         =   "Edit Player"
         ForeColor       =   &H00FFFFFF&
         Height          =   4335
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   9000
         Begin VB.TextBox txtLogin 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   50
            Text            =   "Login Name"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtPassword 
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   49
            Text            =   "Password"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtCharName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   48
            Text            =   "Character Name"
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cmbSex 
            Height          =   315
            ItemData        =   "frmServer.frx":D87D1
            Left            =   1200
            List            =   "frmServer.frx":D87DB
            TabIndex        =   47
            Text            =   "cmbSex"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cmbClass 
            Height          =   315
            ItemData        =   "frmServer.frx":D87ED
            Left            =   1200
            List            =   "frmServer.frx":D87EF
            TabIndex        =   46
            Text            =   "cmbClass"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtLevel 
            Height          =   285
            Left            =   1200
            TabIndex        =   45
            Text            =   "0"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtExp 
            Height          =   285
            Left            =   1200
            TabIndex        =   44
            Text            =   "0"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.ComboBox cmbPK 
            Height          =   315
            ItemData        =   "frmServer.frx":D87F1
            Left            =   1200
            List            =   "frmServer.frx":D87FB
            TabIndex        =   43
            Text            =   "cmbPK"
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txtHP 
            Height          =   285
            Left            =   1200
            TabIndex        =   42
            Text            =   "0"
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox txtMP 
            Height          =   285
            Left            =   1200
            TabIndex        =   41
            Text            =   "0"
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox txtStrength 
            Height          =   285
            Left            =   3840
            TabIndex        =   40
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtEndurance 
            Height          =   285
            Left            =   3840
            TabIndex        =   39
            Text            =   "0"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtIntelligence 
            Height          =   285
            Left            =   3840
            TabIndex        =   38
            Text            =   "0"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtAgility 
            Height          =   285
            Left            =   3840
            TabIndex        =   37
            Text            =   "0"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtWillPower 
            Height          =   285
            Left            =   3840
            TabIndex        =   36
            Text            =   "0"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtPoints 
            Height          =   285
            Left            =   3840
            TabIndex        =   35
            Text            =   "0"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.ComboBox cmbWeapon 
            Height          =   315
            ItemData        =   "frmServer.frx":D8808
            Left            =   3840
            List            =   "frmServer.frx":D880A
            TabIndex        =   34
            Text            =   "cmbWeapon"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.ComboBox cmbArmor 
            Height          =   315
            ItemData        =   "frmServer.frx":D880C
            Left            =   3840
            List            =   "frmServer.frx":D880E
            TabIndex        =   33
            Text            =   "cmbArmor"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.ComboBox cmbHelmet 
            Height          =   315
            ItemData        =   "frmServer.frx":D8810
            Left            =   3840
            List            =   "frmServer.frx":D8812
            TabIndex        =   32
            Text            =   "cmbHelmet"
            Top             =   3120
            Width           =   1335
         End
         Begin VB.ComboBox cmbShield 
            Height          =   315
            ItemData        =   "frmServer.frx":D8814
            Left            =   3840
            List            =   "frmServer.frx":D8816
            TabIndex        =   31
            Text            =   "cmbShield"
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00000000&
            Caption         =   "Inventory Editing"
            ForeColor       =   &H00FFFFFF&
            Height          =   2055
            Left            =   5400
            TabIndex        =   24
            Top             =   1800
            Width           =   1695
            Begin VB.ComboBox cmbInvSlot 
               Height          =   315
               ItemData        =   "frmServer.frx":D8818
               Left            =   120
               List            =   "frmServer.frx":D881A
               TabIndex        =   27
               Text            =   "cmbInvSlot"
               Top             =   480
               Width           =   1335
            End
            Begin VB.ComboBox cmbItems 
               Height          =   315
               ItemData        =   "frmServer.frx":D881C
               Left            =   120
               List            =   "frmServer.frx":D881E
               TabIndex        =   26
               Text            =   "cmbItems"
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox txtItemQuantity 
               Height          =   285
               Left            =   120
               TabIndex        =   25
               Text            =   "0"
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label lblRandom 
               BackStyle       =   0  'Transparent
               Caption         =   "Slot:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   22
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lblRandom 
               BackStyle       =   0  'Transparent
               Caption         =   "Item:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   23
               Left            =   120
               TabIndex        =   29
               Top             =   840
               Width           =   975
            End
            Begin VB.Label lblRandom 
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   24
               Left            =   120
               TabIndex        =   28
               Top             =   1440
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00000000&
            Caption         =   "Spell Editing"
            ForeColor       =   &H00FFFFFF&
            Height          =   2055
            Left            =   7320
            TabIndex        =   19
            Top             =   1800
            Width           =   1455
            Begin VB.ComboBox cmbSpells 
               Height          =   315
               ItemData        =   "frmServer.frx":D8820
               Left            =   120
               List            =   "frmServer.frx":D8822
               TabIndex        =   21
               Text            =   "cmbSpells"
               Top             =   1080
               Width           =   1335
            End
            Begin VB.ComboBox cmbSpellSlot 
               Height          =   315
               ItemData        =   "frmServer.frx":D8824
               Left            =   120
               List            =   "frmServer.frx":D8826
               TabIndex        =   20
               Text            =   "cmbSpellSlot"
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lblRandom 
               BackStyle       =   0  'Transparent
               Caption         =   "Spell:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   26
               Left            =   120
               TabIndex        =   23
               Top             =   840
               Width           =   975
            End
            Begin VB.Label lblRandom 
               BackStyle       =   0  'Transparent
               Caption         =   "Slot:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   27
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.TextBox txtMap 
            Height          =   285
            Left            =   6480
            TabIndex        =   18
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtX 
            Height          =   285
            Left            =   6480
            TabIndex        =   17
            Text            =   "0"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtY 
            Height          =   285
            Left            =   6480
            TabIndex        =   16
            Text            =   "0"
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdEditPlayerOk 
            Caption         =   "Save and Close"
            Height          =   255
            Left            =   5400
            TabIndex        =   15
            Top             =   3960
            Width           =   1815
         End
         Begin VB.CommandButton cmdCancelEditPlayer 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   7440
            TabIndex        =   14
            Top             =   3960
            Width           =   1335
         End
         Begin VB.ComboBox cmbDir 
            Height          =   315
            ItemData        =   "frmServer.frx":D8828
            Left            =   6480
            List            =   "frmServer.frx":D8838
            TabIndex        =   13
            Text            =   "cmbDir"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cmbAccess 
            Height          =   315
            ItemData        =   "frmServer.frx":D8853
            Left            =   1200
            List            =   "frmServer.frx":D8866
            TabIndex        =   12
            Text            =   "cmbAccess"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   74
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Char Name:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   73
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   72
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Class:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   71
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Level:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   70
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Exp:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   69
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Access:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   68
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "PKer:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   67
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "HP:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   66
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "MP:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   65
            Top             =   3840
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Strength:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   12
            Left            =   2760
            TabIndex        =   64
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Endurance:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   13
            Left            =   2760
            TabIndex        =   63
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Intelligence:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   14
            Left            =   2760
            TabIndex        =   62
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Agility:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   15
            Left            =   2760
            TabIndex        =   61
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Willpower:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   16
            Left            =   2760
            TabIndex        =   60
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Stat Points:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   17
            Left            =   2760
            TabIndex        =   59
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Weapon:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   18
            Left            =   2760
            TabIndex        =   58
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Armor:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   19
            Left            =   2760
            TabIndex        =   57
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Helmet:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   20
            Left            =   2760
            TabIndex        =   56
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Shield:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   21
            Left            =   2760
            TabIndex        =   55
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Map:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   25
            Left            =   5400
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "X:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   28
            Left            =   5400
            TabIndex        =   53
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Y:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   29
            Left            =   5400
            TabIndex        =   52
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblRandom 
            BackStyle       =   0  'Transparent
            Caption         =   "Dir:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   30
            Left            =   5400
            TabIndex        =   51
            Top             =   1320
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   3135
         Left            =   240
         TabIndex        =   173
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5530
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "You may right click on a player for more options."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Players"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label lblServerType 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   147
      Top             =   0
      Width           =   2895
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuEditPlayer 
         Caption         =   "Edit Player"
      End
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDisableRestart_Click()


   On Error GoTo errorhandler
    If chkDisableRestart.Value = 1 Then
        Options.DisableRemoteRestart = 1
        SaveOptions
    Else
        Options.DisableRemoteRestart = 0
        SaveOptions
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkDisableRestart_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkStaffOnly_Click()
Dim i As Long

   On Error GoTo errorhandler
    If chkStaffOnly.Value = 1 Then
        Options.StaffOnly = 1
        SaveOptions
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If GetPlayerAccess(i) = 0 Then
                    AlertMsg i, "Sorry, the server was switched to staff-only mode. Please check back later!"
                End If
            End If
        Next
    Else
        Options.StaffOnly = 0
        SaveOptions
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkStaffOnly_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbInvSlot_Click()

   On Error GoTo errorhandler

    cmbItems.ListIndex = EditInv(cmbInvSlot.ListIndex + 1).Num
    txtItemQuantity.Text = Val(EditInv(cmbInvSlot.ListIndex + 1).Value)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbInvSlot_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbItems_Click()

   On Error GoTo errorhandler

    EditInv(cmbInvSlot.ListIndex + 1).Num = cmbItems.ListIndex


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbItems_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSpells_Change()

   On Error GoTo errorhandler

    EditSpell(cmbSpellSlot.ListIndex + 1) = cmbSpells.ListIndex


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSpells_Change", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbSpellSlot_Click()

   On Error GoTo errorhandler

    cmbSpells.ListIndex = EditSpell(cmbSpellSlot.ListIndex + 1)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSpellSlot_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCancelEditPlayer_Click()

   On Error GoTo errorhandler

    EditingPlayer = 0
    fraEditPlayer.Visible = False
    lblNotifications.Caption = "Player Editing Canceled!"
    lblNotifications.ForeColor = &HFF&


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancelEditPlayer_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdAccess_Click()
    Dim help As String

   On Error GoTo errorhandler

    help = "Access is defined by 5 numbers..."
    help = help & vbNewLine & "0: Normal Player"
    help = help & vbNewLine & "1: Moderator - Can kick/warp and simple admin functions."
    help = help & vbNewLine & "2: Mapper - Mod Powers + Mapping Abilities"
    help = help & vbNewLine & "3: Developer - Mapper Powers and ability to edit all game content."
    help = help & vbNewLine & "4: Creator - All Powers, for owner(s) of the game)"
    
    MsgBox help


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAccess_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Private Sub cmdCancelLogin_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancelLogin_Click", "frumServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseAccount_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseAccount_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseConsole_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseConsole_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseControlPanel_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseControlPanel_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCloseHousing_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCloseHousing_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdClosePlayers_Click()


   On Error GoTo errorhandler
    ClearServerWindows True
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdClosePlayers_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCopyKey_Click()
Dim a() As String, i As Long

   On Error GoTo errorhandler
    i = frmServer.lstLicenses.ListIndex
    If i > -1 Then
        a = Split(frmServer.lstLicenses.List(i), "-")
        Clipboard.Clear
        Clipboard.SetText Trim$(a(1))
        lblNotifications.Caption = "License code copied to clipboard!"
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCopyKey_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdEditPlayerOk_Click()
Dim i As Long

On Error GoTo errorhandler
    If IsPlaying(EditingPlayer) Then
        'Check Everything First.
        If Val(txtLevel.Text) > MAX_LEVELS Then
            lblNotifications.Caption = "Player Saving Failed: Level is greater than " & MAX_LEVELS
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        If Val(txtStrength.Text) > 255 Then
            lblNotifications.Caption = "Player Saving Failed: Strength is greater than 255"
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        If Val(txtEndurance.Text) > 255 Then
            lblNotifications.Caption = "Player Saving Failed: Endurance is greater than 255"
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        If Val(txtIntelligence.Text) > 255 Then
            lblNotifications.Caption = "Player Saving Failed: Intelligence is greater than 255"
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        If Val(txtAgility.Text) > 255 Then
            lblNotifications.Caption = "Player Saving Failed: Agility is greater than 255"
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        If Val(txtWillPower.Text) > 255 Then
            lblNotifications.Caption = "Player Saving Failed: Willpower is greater than 255"
            lblNotifications.ForeColor = &HFF&
            Exit Sub
        End If
        Player(EditingPlayer).Password = Trim$(txtPassword.Text)
        With Player(EditingPlayer).characters(TempPlayer(EditingPlayer).CurChar)
            .Name = Trim$(txtCharName.Text)
            .Sex = cmbSex.ListIndex
            .Class = cmbClass.ListIndex
            '.Sprite = Val(txtSprite.Text)
            .Level = Val(txtLevel.Text)
            .Exp = Val(txtExp.Text)
            .access = cmbAccess.ListIndex
            .PK = cmbPK.ListIndex
            .Vital(Vitals.HP) = Val(txtHP.Text)
            .Vital(Vitals.MP) = Val(txtMP.Text)
            .stat(Stats.Strength) = Val(txtStrength.Text)
            .stat(Stats.Endurance) = Val(txtEndurance.Text)
            .stat(Stats.Intelligence) = Val(txtIntelligence.Text)
            .stat(Stats.Agility) = Val(txtAgility.Text)
            .stat(Stats.Willpower) = Val(txtWillPower.Text)
            .Points = Val(txtPoints.Text)
            .Equipment(Equipment.Weapon) = cmbWeapon.ListIndex
            .Equipment(Equipment.armor) = cmbArmor.ListIndex
            .Equipment(Equipment.Helmet) = cmbHelmet.ListIndex
            .Equipment(Equipment.Shield) = cmbShield.ListIndex
            .Map = Val(txtMap.Text)
            .x = Val(txtX.Text)
            .y = Val(txtY.Text)
            .Dir = cmbDir.ListIndex
            
            For i = 1 To MAX_INV
                Player(EditingPlayer).characters(TempPlayer(EditingPlayer).CurChar).Inv(i) = EditInv(i)
            Next
            For i = 1 To MAX_PLAYER_SPELLS
                Player(EditingPlayer).characters(TempPlayer(EditingPlayer).CurChar).Spell(i) = EditSpell(i)
            Next
            SavePlayer EditingPlayer
            ' send vitals, exp + stats
            For i = 1 To Vitals.Vital_Count - 1
                Call SendVital(EditingPlayer, i)
            Next
            SendEXP EditingPlayer
            Call SendStats(EditingPlayer)
            Call SendInventory(EditingPlayer)
            SendDataToMap GetPlayerMap(EditingPlayer), PlayerData(EditingPlayer)
        End With
        PlayerWarp EditingPlayer, GetPlayerMap(EditingPlayer), GetPlayerX(EditingPlayer), GetPlayerY(EditingPlayer), False
    Else
        EditingPlayer = 0
        fraEditPlayer.Visible = False
        lblNotifications.Caption = "Player Saving Failed: Player not Found Online"
        lblNotifications.ForeColor = &HFF&
    End If
    EditingPlayer = 0
    fraEditPlayer.Visible = False
    lblNotifications.Caption = "Player Saved Sucessfully!"
    lblNotifications.ForeColor = &HC000&
    Exit Sub
    
errorhandler:
    lblNotifications.Caption = "Player Saving Failed: Unknown Error!"
    lblNotifications.ForeColor = &HFF&
    HandleError "cmdEditPlayerOk_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdReserveMaps_Click()
Dim map1 As Long, map2 As Long, i As Long

   On Error GoTo errorhandler

    If IsNumeric(txtRMap.Text) Then
        If IsNumeric(txtRMaps2.Text) Then
            map1 = Val(txtRMap.Text)
            map2 = Val(txtRMaps2.Text)
            If map1 > map2 Or map1 < 1 Or map1 > MAX_MAPS Or map2 < 1 Or map2 > MAX_MAPS Then
                lblNotifications.Caption = "An error occured. One of the map values are invalid. The first value must be the smaller one."
                lblNotifications.ForeColor = &HFF&
                Exit Sub
            Else
                For i = map1 To map2
                    Map(i).Name = txtMapName.Text
                    Map(i).Revision = Map(i).Revision + 1
                    MapCache_Create i
                    SaveMap i
                Next
                lblNotifications.Caption = "Maps reserved."
                lblNotifications.ForeColor = &HC000&
                UpdateMapReport
                Exit Sub
            End If
        End If
    End If
    lblNotifications.Caption = "Non-numeric value entered for map number... maps not reserved."
    lblNotifications.ForeColor = &HFF&


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReserveMaps_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveCredits_Click()


   On Error GoTo errorhandler
   
    Credits = txtCredits.Text
    Dim iFileNumber As Integer
    iFileNumber = FreeFile
    Open App.path & "\data\credits.txt" For Output As #iFileNumber
    Print #iFileNumber, Credits
    Close #iFileNumber
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveCredits_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveDataFolder_Click()


   On Error GoTo errorhandler
   
   Options.UpdateURL = txtUpdateUrl.Text
   SaveOptions
    
    If Len(Trim$(txtDataFolder.Text)) > 0 And Trim$(LCase(txtDataFolder.Text)) <> "default" Then
        If IsValidFileName(Trim$(txtDataFolder.Text)) Then
            Options.DataFolder = txtDataFolder.Text
            SaveOptions
            lblNotifications.Caption = "Saved new data folder and update.ini URL!"
        Else
            lblNotifications.Caption = "Data folder not valid! (Saved URL)"
        End If
    Else
        lblNotifications.Caption = "Using 'Default' data folder."
        Options.DataFolder = "default"
        txtDataFolder.Text = "default"
        SaveOptions
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveDataFolder_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Private Function IsValidFileName(strName As String) As Boolean
    IsValidFileName = True
    
    If InStrB(1, strName, "\", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, "/", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, ":", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, "?", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, "*", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, "|", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, Chr(34), vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, "<", vbBinaryCompare) Then IsValidFileName = False
    If InStrB(1, strName, ">", vbBinaryCompare) Then IsValidFileName = False
End Function

Private Sub cmdSaveHouse_Click()

   On Error GoTo errorhandler

    If Val(txtBaseMap.Text) <= 0 Or Val(txtBaseMap.Text) > MAX_MAPS Then
        lblNotifications.Caption = "Base Map value invalid. Must be a number between 1 and " & MAX_MAPS
        lblNotifications.ForeColor = &HFF&
        Exit Sub
    End If
    If Val(txtHouseFurniture.Text) < 0 Or Val(txtHouseFurniture.Text) > 1000 Then
        lblNotifications.Caption = "Value of max furnitures invalid. Must be a number between 0 (infinite) and 1000"
        lblNotifications.ForeColor = &HFF&
        Exit Sub
    End If
    
    If Val(txtXEntrance.Text) < 0 Or Val(txtXEntrance.Text) > Map(txtBaseMap.Text).MaxX Then
        lblNotifications.Caption = "Value of x coordinate is invalid. Must be a number between 0  and map max x value."
        lblNotifications.ForeColor = &HFF&
        Exit Sub
    End If
    
    If Val(txtYEntrance.Text) < 0 Or Val(txtYEntrance.Text) > Map(txtBaseMap.Text).MaxY Then
        lblNotifications.Caption = "Value of y coordinate is invalid. Must be a number between 0  and map max y value."
        lblNotifications.ForeColor = &HFF&
        Exit Sub
    End If
    
    
    If frmServer.lstHouses.ListIndex > -1 And frmServer.lstHouses.ListIndex < MAX_HOUSES Then
        HouseConfig(frmServer.lstHouses.ListIndex + 1).BaseMap = Val(txtBaseMap.Text)
        HouseConfig(frmServer.lstHouses.ListIndex + 1).ConfigName = txtHouseName.Text
        HouseConfig(frmServer.lstHouses.ListIndex + 1).MaxFurniture = Val(txtHouseFurniture.Text)
        HouseConfig(frmServer.lstHouses.ListIndex + 1).price = Val(txtHousePrice.Text)
        HouseConfig(frmServer.lstHouses.ListIndex + 1).x = Val(txtXEntrance.Text)
        HouseConfig(frmServer.lstHouses.ListIndex + 1).y = Val(txtYEntrance.Text)
        SaveHouse frmServer.lstHouses.ListIndex + 1
        lblNotifications.Caption = "House Saved."
        lblNotifications.ForeColor = &HC000&
    Else
        lblNotifications.Caption = "Error: No house configuration selected in lst config. House not saved."
        lblNotifications.ForeColor = &HFF&
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveHouse_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSaveNews_Click()


   On Error GoTo errorhandler

    News = txtNews.Text
    Dim iFileNumber As Integer
    iFileNumber = FreeFile
    Open App.path & "\data\news.txt" For Output As #iFileNumber
    Print #iFileNumber, News
    Close #iFileNumber


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSaveNews_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Command_Click()
    
End Sub

Private Sub fraMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long

   On Error GoTo errorhandler

    For i = 0 To 6
        lblServerMenuOpt(i).Font.Underline = False
    Next
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "fraMenu_MouseMove", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblCPSLock_Click()

   On Error GoTo errorhandler

    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblCPSLock_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblServerMenuOpt_Click(Index As Integer)
    

   On Error GoTo errorhandler

    Select Case Index
        Case 0
            ClearServerWindows False
            fraConsole.Visible = True
        Case 1
            ClearServerWindows False
            fraPlayers.Visible = True
        Case 2
            ClearServerWindows False
            fraHousing.Visible = True
        Case 3
            ClearServerWindows False
            fraControlPanel.Visible = True
            txtNews.Text = News
            txtCredits.Text = Credits
            txtUpdateUrl.Text = Options.UpdateURL
            txtDataFolder.Text = Options.DataFolder
        Case 4
            ClearServerWindows False
            fraAccount.Visible = True
            fraLogin.Visible = True
        Case 5
            Call ShellExecute(0, vbNullString, "http://eclipseorigins.com/index.php?topic=48.0", vbNullString, vbNullString, vbNormalFocus)
        Case 6
            DestroyServer
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblServerMenuOpt_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblServerMenuOpt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long

   On Error GoTo errorhandler

    For i = 0 To 6
        If Index <> i Then
            lblServerMenuOpt(i).Font.Underline = False
        Else
            lblServerMenuOpt(i).Font.Underline = True
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblServerMenuOpt_MouseMove", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblUpdateHelp_Click()


   On Error GoTo errorhandler
    
    Call ShellExecute(0, vbNullString, "http://eclipseorigins.com/smf1/index.php?topic=51", vbNullString, vbNullString, vbNormalFocus)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblUpdateHelp_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lstHouses_Click()

   On Error GoTo errorhandler

    If lstHouses.ListIndex > -1 And lstHouses.ListIndex < MAX_HOUSES Then
        txtBaseMap.Text = HouseConfig(lstHouses.ListIndex + 1).BaseMap
        txtHouseName.Text = HouseConfig(lstHouses.ListIndex + 1).ConfigName
        txtHouseFurniture.Text = HouseConfig(lstHouses.ListIndex + 1).MaxFurniture
        txtHousePrice.Text = HouseConfig(lstHouses.ListIndex + 1).price
        txtXEntrance.Text = HouseConfig(lstHouses.ListIndex + 1).x
        txtYEntrance.Text = HouseConfig(lstHouses.ListIndex + 1).y
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstHouses_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub mnuEditPlayer_Click()
    Dim Name As String
    Dim i As Long

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        If Len(Trim$(Name)) <= 0 Then Exit Sub
        i = FindPlayer(Trim$(Name))
        EditingPlayer = i
        For i = 1 To MAX_INV
            EditInv(i) = Player(EditingPlayer).characters(TempPlayer(EditingPlayer).CurChar).Inv(i)
        Next
        For i = 1 To MAX_PLAYER_SPELLS
            EditSpell(i) = Player(EditingPlayer).characters(TempPlayer(EditingPlayer).CurChar).Spell(i)
        Next
        LoadEditPlayer EditingPlayer
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuEditPlayer_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)

   On Error GoTo errorhandler

    Call AcceptConnection(Index, requestID)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Socket_ConnectionRequest", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)

   On Error GoTo errorhandler

    Call AcceptConnection(Index, SocketId)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Socket_Accept", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)


   On Error GoTo errorhandler

    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Socket_Close(Index As Integer)

   On Error GoTo errorhandler

    Call CloseSocket(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Socket_Close", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ********************
Private Sub chkServerLog_Click()

    ' if its not 0, then its true

   On Error GoTo errorhandler

    If Not chkServerLog.Value Then
        ServerLog = True
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkServerLog_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdExit_Click()

   On Error GoTo errorhandler

    Call DestroyServer


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdExit_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadClasses_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadClasses
    Call TextAdd("All classes reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendClasses i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadClasses_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadItems_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadItems
    Call TextAdd("All items reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendItems i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadItems_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadMaps_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadMaps
    Call TextAdd("All maps reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadMaps_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadNPCs_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendNpcs i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadNPCs_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadShops_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadShops
    Call TextAdd("All shops reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendShops i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadShops_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadSpells_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadSpells_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadResources_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadResources
    Call TextAdd("All Resources reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendResources i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadResources_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub cmdReloadAnimations_Click()
Dim i As Long

   On Error GoTo errorhandler

    Call LoadAnimations
    Call TextAdd("All Animations reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendAnimations i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdReloadAnimations_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdShutDown_Click()

   On Error GoTo errorhandler

    If isShuttingDown Then
        isShuttingDown = False
        shutDownType = 0
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
        shutDownType = 0
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdShutDown_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_Load()

   On Error GoTo errorhandler

    lblNotifications.Caption = ""
    lblNotifications.ForeColor = &HFF&


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearServerWindows(Optional showMenu As Boolean = True)
    fraConsole.Visible = False
    fraPlayers.Visible = False
    fraHousing.Visible = False
    fraControlPanel.Visible = False
    fraAccount.Visible = False
    picEditAccount.Visible = False
    fraLogin.Visible = True
    txtLEmail.Text = ""
    txtLPass.Text = ""
    
    If showMenu Then
        fraMenu.Visible = True
    Else
        fraMenu.Visible = False
    End If
End Sub

Private Sub Form_Resize()


   On Error GoTo errorhandler

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Resize", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error GoTo errorhandler

    Cancel = True
    Call DestroyServer


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.

   On Error GoTo errorhandler

    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.Index - 1
    lvwInfo.Sorted = True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lvwInfo_ColumnClick", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub


Private Sub txtItemQuantity_Change()

   On Error GoTo errorhandler

    EditInv(cmbInvSlot.ListIndex + 1).Value = Val(txtItemQuantity.Text)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtItemQuantity_Change", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtText_GotFocus()

   On Error GoTo errorhandler

    txtChat.SetFocus


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtText_GotFocus", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)


   On Error GoTo errorhandler

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtChat_KeyPress", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub UsersOnline_Start()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (i)

        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If

        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UsersOnline_Start", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)


   On Error GoTo errorhandler

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lvwInfo_MouseDown", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server owner!")
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuKickPlayer_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        CloseSocket (FindPlayer(Name))
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuDisconnectPlayer_Click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        If Ban(FindPlayer(Name), "", True, "Banned by server console.") = False Then
            frmServer.lblNotifications.Caption = Trim$(Name) & " is already banned!"
        Else
            frmServer.lblNotifications.Caption = Trim$(Name) & " and his IP has been banned from this server!"
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuBanPlayer_click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted administrator access.", BrightCyan)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuAdminPlayer_click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String

   On Error GoTo errorhandler

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have had your administrator access revoked.", BrightRed)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "mnuRemoveAdmin_click", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long, i As Long

   On Error GoTo errorhandler

    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select
    
    For i = 0 To 6
        lblServerMenuOpt(i).Font.Underline = False
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub LoadEditPlayer(Index As Long)
Dim i As Long

   On Error GoTo errorhandler

    If IsPlaying(Index) Then
        fraEditPlayer.ZOrder 0
        fraEditPlayer.Visible = True
        
        'Load all of the players info :D
        With Player(Index)
            txtLogin.Text = Trim$(.login)
            txtPassword.Text = Trim$(.Password)
        End With
        With Player(Index).characters(TempPlayer(Index).CurChar)
            txtCharName.Text = Trim$(.Name)
            cmbSex.ListIndex = .Sex
            cmbClass.ListIndex = .Class + 1
            'txtSprite.Text = Val(.Sprite)
            txtLevel.Text = Val(.Level)
            txtExp.Text = Val(.Exp)
            cmbAccess.ListIndex = .access
            cmbPK.ListIndex = .PK
            txtHP.Text = Val(.Vital(Vitals.HP))
            txtMP.Text = Val(.Vital(Vitals.MP))
            txtStrength.Text = Val(.stat(Stats.Strength))
            txtEndurance.Text = Val(.stat(Stats.Endurance))
            txtIntelligence.Text = Val(.stat(Stats.Intelligence))
            txtAgility.Text = Val(.stat(Stats.Agility))
            txtWillPower.Text = Val(.stat(Stats.Willpower))
            txtPoints.Text = Val(.Points)
            cmbWeapon.ListIndex = .Equipment(Equipment.Weapon)
            cmbArmor.ListIndex = .Equipment(Equipment.armor)
            cmbHelmet.ListIndex = .Equipment(Equipment.Helmet)
            cmbShield.ListIndex = .Equipment(Equipment.Shield)
            txtMap.Text = Val(.Map)
            txtX.Text = Val(.x)
            txtY.Text = Val(.y)
            cmbDir.ListIndex = .Dir
            
            cmbInvSlot.Clear
            For i = 1 To MAX_INV
                cmbInvSlot.AddItem i
            Next
            cmbInvSlot.ListIndex = 0
            
            cmbSpellSlot.Clear
            For i = 1 To MAX_PLAYER_SPELLS
                cmbSpellSlot.AddItem i
            Next
            cmbSpellSlot.ListIndex = 0
        
        End With
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadEditPlayer", "frmServer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
