VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picCharSelect 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   555
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   40
      Top             =   180
      Visible         =   0   'False
      Width           =   6630
      Begin VB.ListBox lstCharacters 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   600
         TabIndex        =   42
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label lblDelChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4800
         TabIndex        =   44
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblUseChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblNewChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   41
         Top             =   2640
         Width           =   1215
      End
   End
   Begin VB.PictureBox picNewCharacter 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   555
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   16
      Top             =   180
      Visible         =   0   'False
      Width           =   6630
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   5400
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   25
         Top             =   1800
         Width           =   480
      End
      Begin VB.ComboBox cmbClass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1680
         Width           =   2175
      End
      Begin VB.OptionButton optMale 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   2160
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFemale 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtCName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   21
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lblNextShoe 
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   6000
         TabIndex        =   39
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label lblLastShoe 
         BackStyle       =   0  'Transparent
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   5160
         TabIndex        =   38
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label lblNextLeg 
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   6000
         TabIndex        =   37
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label lblNextShirt 
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   6000
         TabIndex        =   36
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label lblNextEye 
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   6000
         TabIndex        =   35
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label lblNextHair 
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   6000
         TabIndex        =   34
         Top             =   1770
         Width           =   135
      End
      Begin VB.Label lblNextBody 
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   5670
         TabIndex        =   33
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label lblLastBody 
         BackStyle       =   0  'Transparent
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   5520
         TabIndex        =   32
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label lblLastLeg 
         BackStyle       =   0  'Transparent
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   5160
         TabIndex        =   31
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label lblLastShirt 
         BackStyle       =   0  'Transparent
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   5160
         TabIndex        =   30
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label lblLastEye 
         BackStyle       =   0  'Transparent
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   5160
         TabIndex        =   29
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label lblLastHair 
         BackStyle       =   0  'Transparent
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   5160
         TabIndex        =   28
         Top             =   1770
         Width           =   135
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   24
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   23
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   22
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblCAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.PictureBox picRegister 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   555
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   7
      Top             =   180
      Visible         =   0   'False
      Width           =   6630
      Begin VB.TextBox txtRPass2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   20
         PasswordChar    =   "•"
         TabIndex        =   13
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtRPass 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   20
         PasswordChar    =   "•"
         TabIndex        =   10
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtRUser 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2520
         MaxLength       =   12
         TabIndex        =   8
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Retype:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label txtRAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.PictureBox picLogin 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   555
      ScaleHeight     =   242
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   6630
      Begin VB.CheckBox chkPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Save Password?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtLPass 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   20
         PasswordChar    =   "•"
         TabIndex        =   3
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtLUser 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2520
         MaxLength       =   12
         TabIndex        =   1
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label lblLAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.PictureBox picCredits 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   555
      ScaleHeight     =   3645
      ScaleWidth      =   6630
      TabIndex        =   15
      Top             =   180
      Visible         =   0   'False
      Width           =   6630
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   555
      ScaleHeight     =   3645
      ScaleWidth      =   6630
      TabIndex        =   26
      Top             =   180
      Width           =   6630
      Begin VB.Label lblNews 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This is an example of the news. Not very exciting, I know, but it's better than nothing, amirite? "
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   1680
         TabIndex        =   27
         Top             =   1200
         Width           =   3135
      End
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   4
      Left            =   5460
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   3
      Left            =   3960
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   2
      Left            =   2460
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   1
      Left            =   960
      Top             =   4305
      Width           =   1335
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClass_Click()
    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
End Sub

Private Sub Form_Load()
    Dim tmpTxt As String, tmpArray() As String, i As Long


    ' general menu stuff
    Me.Caption = Options.Game_Name
    ' load news
    Open App.Path & "\data files\news.txt" For Input As #1
        Line Input #1, tmpTxt
    Close #1
    ' split breaks
    tmpArray() = Split(tmpTxt, "<br />")
    lblNews.Caption = vbNullString
    For i = 0 To UBound(tmpArray)
        lblNews.Caption = lblNews.Caption & tmpArray(i) & vbNewLine
    Next

    ' Load the username + pass
    txtLUser.Text = Trim$(Options.Username)
    If Options.SavePass = 1 Then
        txtLPass.Text = Trim$(Options.Password)
        chkPass.value = Options.SavePass
    End If



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    resetButtons_Menu



End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not EnteringGame Then DestroyGame



End Sub

Private Sub imgButton_Click(Index As Integer)

    Select Case Index
        Case 1
            If Not picLogin.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.Visible = False
                picLogin.Visible = True
                picRegister.Visible = False
                picNewCharacter.Visible = False
                picCharSelect.Visible = False
                picMain.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        Case 2
            If Not picRegister.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.Visible = False
                picLogin.Visible = False
                picRegister.Visible = True
                picNewCharacter.Visible = False
                picCharSelect.Visible = False
                picMain.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        Case 3
            If Not picCredits.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.Visible = True
                picLogin.Visible = False
                picRegister.Visible = False
                picNewCharacter.Visible = False
                picCharSelect.Visible = False
                picMain.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        Case 4
            Call DestroyGame
    End Select



End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    ' reset other buttons
    resetButtons_Menu Index
    ' change the button we're hovering on
    changeButtonState_Menu Index, 2 ' clicked



End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    ' reset other buttons
    resetButtons_Menu Index
    ' change the button we're hovering on
    If Not MenuButton(Index).state = 2 Then ' make sure we're not clicking
        changeButtonState_Menu Index, 1 ' hover
    End If
    ' play sound
    If Not LastButtonSound_Menu = Index Then
        PlaySound Sound_ButtonHover, -1, -1
        LastButtonSound_Menu = Index
    End If



End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

        ' reset all buttons
    resetButtons_Menu -1



End Sub

Private Sub Label6_Click()

End Sub

Private Sub lblDelChar_Click()
    If frmMenu.lstCharacters.ListIndex > -1 Then
        If MsgBox("Are you sure you want to delete the character?", vbYesNo, "Character Deletion") = vbYes Then
            Dim buffer As clsBuffer
            Set buffer = New clsBuffer
            buffer.WriteLong CUseChar
            buffer.WriteLong frmMenu.lstCharacters.ListIndex + 1
            buffer.WriteLong 1
            SendData buffer.ToArray
            Set buffer = Nothing
        End If
    End If
End Sub

Private Sub lblLAccept_Click()

    If isLoginLegal(txtLUser.Text, txtLPass.Text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If




End Sub


Private Sub lblSprite_Click()
Dim spritecount As Long

    If optMale.value Then
        spritecount = UBound(Class(cmbClass.ListIndex + 1).MaleSprite)
    Else
        spritecount = UBound(Class(cmbClass.ListIndex + 1).FemaleSprite)
    End If

    If newCharSprite >= spritecount Then
        newCharSprite = 0
    Else
        newCharSprite = newCharSprite + 1
    End If



End Sub

Private Sub lblLastBody_Click()
    NewCharBody = NewCharBody - 1
    If NewCharSex = 0 Then
        If NewCharBody = 0 Then NewCharBody = UBound(Tex_MaleBodies)
    Else
        If NewCharBody = 0 Then NewCharBody = UBound(Tex_FemaleBodies)
    End If
End Sub

Private Sub lblLastEye_Click()
    NewCharEyes = NewCharEyes - 1
    If NewCharSex = 0 Then
        If NewCharEyes = 0 Then NewCharEyes = UBound(Tex_MaleEyes)
    Else
        If NewCharEyes = 0 Then NewCharEyes = UBound(Tex_FemaleEyes)
    End If
End Sub

Private Sub lblLastEyebrow_Click()
    NewCharEyeBrows = NewCharEyeBrows - 1
    If NewCharSex = 0 Then
        If NewCharEyeBrows = 0 Then NewCharEyeBrows = UBound(Tex_MaleEyeBrows)
    Else
        If NewCharEyeBrows = 0 Then NewCharEyeBrows = UBound(Tex_FemaleEyeBrows)
    End If
End Sub

Private Sub lblLastHair_Click()
    NewCharHair = NewCharHair - 1
    If NewCharSex = 0 Then
        If NewCharHair = 0 Then NewCharHair = UBound(Tex_MaleHair)
    Else
        If NewCharHair = 0 Then NewCharHair = UBound(Tex_FemaleHair)
    End If
End Sub

Private Sub lblLastLeg_Click()
    NewCharLegs = NewCharLegs - 1
    If NewCharSex = 0 Then
        If NewCharLegs = 0 Then NewCharLegs = UBound(Tex_MaleLegs)
    Else
        If NewCharLegs = 0 Then NewCharLegs = UBound(Tex_FemaleLegs)
    End If
End Sub

Private Sub lblLastShirt_Click()
    NewCharShirt = NewCharShirt - 1
    If NewCharSex = 0 Then
        If NewCharShirt = 0 Then NewCharShirt = UBound(Tex_MaleShirts)
    Else
        If NewCharShirt = 0 Then NewCharShirt = UBound(Tex_FemaleShirts)
    End If
End Sub

Private Sub lblLastShoe_Click()
    NewCharShoes = NewCharShoes - 1
    If NewCharSex = 0 Then
        If NewCharShoes = 0 Then NewCharShoes = UBound(Tex_MaleShoes)
    Else
        If NewCharShoes = 0 Then NewCharShoes = UBound(Tex_FemaleShoes)
    End If
End Sub

Private Sub lblNewChar_Click()
    If frmMenu.lstCharacters.ListIndex > -1 Then
        Dim buffer As clsBuffer
        Set buffer = New clsBuffer
        buffer.WriteLong CUseChar
        buffer.WriteLong frmMenu.lstCharacters.ListIndex + 1
        buffer.WriteLong 0
        SendData buffer.ToArray
        picCharSelect.Visible = False
        Set buffer = Nothing
    End If
End Sub

Private Sub lblNextBody_Click()
    NewCharBody = NewCharBody + 1
    If NewCharSex = 0 Then
        If NewCharBody = UBound(Tex_MaleBodies) + 1 Then NewCharBody = 1
    Else
        If NewCharBody = UBound(Tex_FemaleBodies) + 1 Then NewCharBody = 1
    End If
End Sub

Private Sub lblNextEye_Click()
    NewCharEyes = NewCharEyes + 1
    If NewCharSex = 0 Then
        If NewCharEyes = UBound(Tex_MaleEyes) + 1 Then NewCharEyes = 1
    Else
        If NewCharEyes = UBound(Tex_FemaleEyes) + 1 Then NewCharEyes = 1
    End If
End Sub

Private Sub lblNextEyebrow_Click()
    NewCharEyeBrows = NewCharEyeBrows + 1
    If NewCharSex = 0 Then
        If NewCharEyeBrows = UBound(Tex_MaleEyeBrows) + 1 Then NewCharEyeBrows = 1
    Else
        If NewCharEyeBrows = UBound(Tex_FemaleEyeBrows) + 1 Then NewCharEyeBrows = 1
    End If
End Sub

Private Sub lblNextHair_Click()
    NewCharHair = NewCharHair + 1
    If NewCharSex = 0 Then
        If NewCharHair = UBound(Tex_MaleHair) + 1 Then NewCharHair = 1
    Else
        If NewCharHair = UBound(Tex_FemaleHair) + 1 Then NewCharHair = 1
    End If
End Sub

Private Sub lblNextLeg_Click()
    NewCharLegs = NewCharLegs + 1
    If NewCharSex = 0 Then
        If NewCharLegs = UBound(Tex_MaleLegs) + 1 Then NewCharLegs = 1
    Else
        If NewCharLegs = UBound(Tex_FemaleLegs) + 1 Then NewCharLegs = 1
    End If
End Sub

Private Sub lblNextShirt_Click()
    NewCharShirt = NewCharShirt + 1
    If NewCharSex = 0 Then
        If NewCharShirt = UBound(Tex_MaleShirts) + 1 Then NewCharShirt = 1
    Else
        If NewCharShirt = UBound(Tex_FemaleShirts) + 1 Then NewCharShirt = 1
    End If
End Sub

Private Sub lblNextShoe_Click()
    NewCharShoes = NewCharShoes + 1
    If NewCharSex = 0 Then
        If NewCharShoes = UBound(Tex_MaleShoes) + 1 Then NewCharShoes = 1
    Else
        If NewCharShoes = UBound(Tex_FemaleShoes) + 1 Then NewCharShoes = 1
    End If
End Sub

Private Sub lblUseChar_Click()
    If frmMenu.lstCharacters.ListIndex > -1 Then
        Dim buffer As clsBuffer
        Set buffer = New clsBuffer
        buffer.WriteLong CUseChar
        buffer.WriteLong frmMenu.lstCharacters.ListIndex + 1
        buffer.WriteLong 0
        SendData buffer.ToArray
        picCharSelect.Visible = False
        Set buffer = Nothing
    End If
End Sub

Private Sub optFemale_Click()


    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    NewCharSex = SEX_FEMALE
    NewCharBody = 1
    NewCharEyes = 1
    NewCharHair = 1
    NewCharLegs = 1
    NewCharShirt = 1
    NewCharShoes = 1



End Sub

Private Sub optMale_Click()


    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    NewCharSex = SEX_MALE
    NewCharBody = 1
    NewCharEyes = 1
    NewCharHair = 1
    NewCharLegs = 1
    NewCharShirt = 1
    NewCharShoes = 1



End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    resetButtons_Menu



End Sub

Private Sub picCredits_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    resetButtons_Menu



End Sub

Private Sub picLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    resetButtons_Menu



End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    resetButtons_Menu



End Sub

Private Sub picRegister_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    resetButtons_Menu



End Sub

' Register
Private Sub txtRAccept_Click()
    Dim name As String
    Dim Password As String
    Dim PasswordAgain As String

    name = Trim$(txtRUser.Text)
    Password = Trim$(txtRPass.Text)
    PasswordAgain = Trim$(txtRPass2.Text)

    If isLoginLegal(name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("Passwords don't match.")
            Exit Sub
        End If

        If Not isStringLegal(name) Then
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If



End Sub

' New Char
Private Sub lblCAccept_Click()

    Call MenuState(MENU_STATE_ADDCHAR)



End Sub
