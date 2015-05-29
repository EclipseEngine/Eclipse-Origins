VERSION 5.00
Begin VB.Form frmEditor_Projectile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Projectile Editor"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Projectile Properties"
      Height          =   4455
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   3255
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         Left            =   120
         Max             =   1000
         TabIndex        =   15
         Top             =   3240
         Width           =   3015
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   13
         Top             =   2640
         Width           =   3015
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   11
         Top             =   2040
         Width           =   3015
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   9
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3015
      End
      Begin VB.PictureBox picProjectile 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1200
         ScaleHeight     =   32
         ScaleMode       =   0  'User
         ScaleWidth      =   512
         TabIndex        =   5
         Top             =   840
         Width           =   1920
      End
      Begin VB.Label lblDamage 
         Caption         =   "Additional Damage: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label lblSpeed 
         Caption         =   "Speed: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblRange 
         Caption         =   "Range: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label lblPic 
         Caption         =   "Pic: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Projectile List"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   4545
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Projectile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
   On Error GoTo errorhandler
   
    ProjectileEditorCancel
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Projectile", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdSave_Click()
   On Error GoTo errorhandler
   
    ProjectileEditorOk
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Projectile", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_Load()
   On Error GoTo errorhandler
   
    scrlPic.max = NumProjectiles
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Projectile", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lstIndex_Click()
   On Error GoTo errorhandler
   
    ProjectileEditorInit
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Projectile", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlDamage_Change()
   On Error GoTo errorhandler
   
    If EditorIndex < 1 Or EditorIndex > MAX_PROJECTILES Then Exit Sub
    Projectiles(EditorIndex).Damage = scrlDamage.Value
    lblDamage.Caption = "Additional Damage: " & scrlDamage.Value
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_Projectile", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlPic_Change()
   On Error GoTo errorhandler
   
    If EditorIndex < 1 Or EditorIndex > MAX_PROJECTILES Then Exit Sub
    Projectiles(EditorIndex).Sprite = scrlPic.Value
    lblPic.Caption = "Pic: " & scrlPic.Value
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Projectile", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlRange_Change()
   On Error GoTo errorhandler
   
    If EditorIndex < 1 Or EditorIndex > MAX_PROJECTILES Then Exit Sub
    Projectiles(EditorIndex).Range = scrlRange.Value
    lblRange.Caption = "Range: " & scrlRange.Value
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_Projectile", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlSpeed_Change()
   On Error GoTo errorhandler
   
    If EditorIndex < 1 Or EditorIndex > MAX_PROJECTILES Then Exit Sub
    Projectiles(EditorIndex).speed = scrlSpeed.Value
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Projectile", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

   On Error GoTo errorhandler
   
    If EditorIndex < 1 Or EditorIndex > MAX_PROJECTILES Then Exit Sub
    Projectiles(EditorIndex).Name = Trim$(txtName.Text)
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Projectiles(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Projectile", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
