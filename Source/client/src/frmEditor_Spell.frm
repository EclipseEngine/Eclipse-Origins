VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10335
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
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spell Properties"
      Height          =   7335
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   6960
         Width           =   2535
      End
      Begin VB.TextBox txtDesc 
         Height          =   735
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   6480
         Width           =   3255
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data"
         Height          =   6615
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   3255
         Begin VB.HScrollBar scrlAOE 
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   3720
            Width           =   3015
         End
         Begin VB.HScrollBar scrlProjectile 
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   4440
            Width           =   3015
         End
         Begin VB.CheckBox chkProjectile 
            Caption         =   "Projectile?"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   3960
            Width           =   3015
         End
         Begin VB.HScrollBar scrlStun 
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   6240
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnim 
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   5640
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnimCast 
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   5040
            Width           =   2895
         End
         Begin VB.CheckBox chkAOE 
            Caption         =   "Area of Effect spell?"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   3240
            Width           =   3015
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlInterval 
            Height          =   255
            Left            =   1680
            Max             =   60
            TabIndex        =   38
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDuration 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   36
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlVital 
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   1680
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblProjectile 
            Caption         =   "Projectile: 0 None"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   4200
            Width           =   3015
         End
         Begin VB.Label lblStun 
            Caption         =   "Stun Duration: None"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   6000
            Width           =   2895
         End
         Begin VB.Label lblAnim 
            Caption         =   "Animation: None"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   5400
            Width           =   2895
         End
         Begin VB.Label lblAnimCast 
            Caption         =   "Cast Anim: None"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   4800
            Width           =   2895
         End
         Begin VB.Label lblAOE 
            Caption         =   "AoE: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   3480
            Width           =   3015
         End
         Begin VB.Label lblRange 
            Caption         =   "Range: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2640
            Width           =   3015
         End
         Begin VB.Label lblInterval 
            Caption         =   "Interval: 0s"
            Height          =   255
            Left            =   1680
            TabIndex        =   37
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblDuration 
            Caption         =   "Duration: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblVital 
            Caption         =   "Vital: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label lblDir 
            Caption         =   "Dir: Down"
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Basic Information"
         Height          =   5895
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3255
         Begin VB.PictureBox picSprite 
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
            Height          =   480
            Left            =   2640
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   50
            Top             =   5160
            Width           =   480
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   5400
            Width           =   2415
         End
         Begin VB.HScrollBar scrlCool 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   32
            Top             =   4680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   30
            Top             =   4080
            Width           =   3015
         End
         Begin VB.ComboBox cmbClass 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   3480
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   26
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   24
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMP 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   3015
         End
         Begin VB.ComboBox cmbType 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":0000
            Left            =   120
            List            =   "frmEditor_Spell.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblIcon 
            Caption         =   "Icon: None"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   5160
            Width           =   3015
         End
         Begin VB.Label lblCool 
            Caption         =   "Cooldown Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   4440
            Width           =   2535
         End
         Begin VB.Label lblCast 
            Caption         =   "Casting Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Class Required:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label lblAccess 
            Caption         =   "Access Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblMP 
            Caption         =   "MP Cost: None"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3480
         TabIndex        =   55
         Top             =   6960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   6240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Spell List"
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6900
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAOE_Click()


   On Error GoTo errorhandler

    If chkAOE.Value = 0 Then
        spell(EditorIndex).IsAoE = False
    Else
        spell(EditorIndex).IsAoE = True
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkAOE_Click", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub chkProjectile_Click()


   On Error GoTo errorhandler

    If chkProjectile.Value = 0 Then
        spell(EditorIndex).IsProjectile = False
    Else
        spell(EditorIndex).IsProjectile = True
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkProjectile_Click", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmbClass_Click()


   On Error GoTo errorhandler

    spell(EditorIndex).classReq = cmbClass.ListIndex




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbClass_Click", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmbType_Click()


   On Error GoTo errorhandler

    spell(EditorIndex).type = cmbType.ListIndex

    If spell(EditorIndex).type = SPELL_TYPE_PET Then
        scrlVital.Value = spell(EditorIndex).Pet
        scrlVital_Change
    Else
        scrlVital.Value = spell(EditorIndex).Vital
        scrlVital_Change
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long



   On Error GoTo errorhandler

    ClearSpell EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    SpellEditorInit




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdSave_Click()


   On Error GoTo errorhandler

    SpellEditorOk




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub lstIndex_Click()


   On Error GoTo errorhandler

    SpellEditorInit




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdCancel_Click()


   On Error GoTo errorhandler

    SpellEditorCancel




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlAccess_Change()


   On Error GoTo errorhandler

    If scrlAccess.Value > 0 Then
        lblAccess.Caption = "Access Required: " & scrlAccess.Value
    Else
        lblAccess.Caption = "Access Required: None"
    End If
    spell(EditorIndex).AccessReq = scrlAccess.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAccess_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlAnim_Change()


   On Error GoTo errorhandler

    If scrlAnim.Value > 0 Then
        lblAnim.Caption = "Animation: " & Trim$(Animation(scrlAnim.Value).Name)
    Else
        lblAnim.Caption = "Animation: None"
    End If
    spell(EditorIndex).SpellAnim = scrlAnim.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlAnimCast_Change()


   On Error GoTo errorhandler

    If scrlAnimCast.Value > 0 Then
        lblAnimCast.Caption = "Cast Anim: " & Trim$(Animation(scrlAnimCast.Value).Name)
    Else
        lblAnimCast.Caption = "Cast Anim: None"
    End If
    spell(EditorIndex).CastAnim = scrlAnimCast.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAnimCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlAOE_Change()


   On Error GoTo errorhandler

    If scrlAOE.Value > 0 Then
        lblAOE.Caption = "AoE: " & scrlAOE.Value & " tiles."
    Else
        lblAOE.Caption = "AoE: Self-cast"
    End If
    spell(EditorIndex).AoE = scrlAOE.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAOE_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlCast_Change()


   On Error GoTo errorhandler

    lblCast.Caption = "Casting Time: " & scrlCast.Value & "s"
    spell(EditorIndex).CastTime = scrlCast.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlCool_Change()


   On Error GoTo errorhandler

    lblCool.Caption = "Cooldown Time: " & scrlCool.Value & "s"
    spell(EditorIndex).CDTime = scrlCool.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlCool_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlDir_Change()
Dim sDir As String


   On Error GoTo errorhandler

    Select Case scrlDir.Value
        Case DIR_UP
            sDir = "Up"
        Case DIR_DOWN
            sDir = "Down"
        Case DIR_RIGHT
            sDir = "Right"
        Case DIR_LEFT
            sDir = "Left"
    End Select
    lblDir.Caption = "Dir: " & sDir
    spell(EditorIndex).dir = scrlDir.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlDir_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlDuration_Change()


   On Error GoTo errorhandler

    lblDuration.Caption = "Duration: " & scrlDuration.Value & "s"
    spell(EditorIndex).Duration = scrlDuration.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlDuration_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlIcon_Change()


   On Error GoTo errorhandler

    If scrlIcon.Value > 0 Then
        lblIcon.Caption = "Icon: " & scrlIcon.Value
    Else
        lblIcon.Caption = "Icon: None"
    End If
    spell(EditorIndex).Icon = scrlIcon.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlIcon_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlInterval_Change()


   On Error GoTo errorhandler

    lblInterval.Caption = "Interval: " & scrlInterval.Value & "s"
    spell(EditorIndex).Interval = scrlInterval.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlInterval_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlLevel_Change()


   On Error GoTo errorhandler

    If scrlLevel.Value > 0 Then
        lblLevel.Caption = "Level Required: " & scrlLevel.Value
    Else
        lblLevel.Caption = "Level Required: None"
    End If
    spell(EditorIndex).LevelReq = scrlLevel.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlLevel_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub ScrlMap_Change()


   On Error GoTo errorhandler

    lblMap.Caption = "Map: " & ScrlMap.Value
    spell(EditorIndex).Map = ScrlMap.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMap_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlMP_Change()


   On Error GoTo errorhandler

    If scrlMP.Value > 0 Then
        lblMP.Caption = "MP Cost: " & scrlMP.Value
    Else
        lblMP.Caption = "MP Cost: None"
    End If
    spell(EditorIndex).MPCost = scrlMP.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMP_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlProjectile_Change()
   On Error GoTo errorhandler

    If scrlProjectile.Value > 0 Then
        lblProjectile.Caption = "Projectile: " & scrlProjectile.Value & " " & Trim$(Projectiles(scrlProjectile.Value).Name)
    Else
        lblRange.Caption = "Projectile: 0 None"
    End If
    spell(EditorIndex).Projectile = scrlProjectile.Value

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlProjectile_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlRange_Change()


   On Error GoTo errorhandler

    If scrlRange.Value > 0 Then
        lblRange.Caption = "Range: " & scrlRange.Value & " tiles."
    Else
        lblRange.Caption = "Range: Self-cast"
    End If
    spell(EditorIndex).Range = scrlRange.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlStun_Change()


   On Error GoTo errorhandler

    If scrlStun.Value > 0 Then
        lblStun.Caption = "Stun Duration: " & scrlStun.Value & "s"
    Else
        lblStun.Caption = "Stun Duration: None"
    End If
    spell(EditorIndex).StunDuration = scrlStun.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlStun_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlVital_Change()


   On Error GoTo errorhandler

    If spell(EditorIndex).type = SPELL_TYPE_PET Then
        If scrlVital.Value = 0 Then
            lblVital.Caption = "Pet: None"
        Else
            lblVital.Caption = "Pet: " & spell(EditorIndex).Pet & ". " & Trim$(Pet(scrlVital.Value).Name)
        End If
        spell(EditorIndex).Pet = scrlVital.Value
    Else
        lblVital.Caption = "Vital: " & scrlVital.Value
        spell(EditorIndex).Vital = scrlVital.Value
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlVital_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Scrlx_Change()


   On Error GoTo errorhandler

    lblx.Caption = "X: " & Scrlx.Value
    spell(EditorIndex).X = Scrlx.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlX_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Scrly_Change()


   On Error GoTo errorhandler

    lbly.Caption = "Y: " & Scrly.Value
    spell(EditorIndex).Y = Scrly.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlY_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtDesc_Change()


   On Error GoTo errorhandler

    spell(EditorIndex).Desc = txtDesc.Text




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long



   On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    spell(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmbSound_Click()

   On Error GoTo errorhandler

    If cmbSound.ListIndex >= 0 Then
        spell(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        spell(EditorIndex).sound = "None."
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSound_Click", "frmEditor_Spell", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
