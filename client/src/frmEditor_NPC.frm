VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
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
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   32
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   31
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   30
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "NPC Properties"
      Height          =   7575
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.CheckBox chkDropItems 
         Caption         =   "This NPC picks up dropped items found on the map."
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   6720
         Width           =   4695
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   50
         Text            =   "0"
         Top             =   7080
         Width           =   1815
      End
      Begin VB.Frame fraDrop 
         Caption         =   "Drop"
         Height          =   1815
         Left            =   120
         TabIndex        =   41
         Top             =   4800
         Width           =   4815
         Begin VB.TextBox txtChance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   45
            Text            =   "0"
            Top             =   720
            Width           =   1815
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   1200
            Max             =   100
            TabIndex        =   44
            Top             =   1200
            Width           =   3495
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            Max             =   32000
            TabIndex        =   43
            Top             =   1440
            Width           =   3495
         End
         Begin VB.HScrollBar scrlDrop 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   42
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Chance: (X out of 100)"
            Height          =   180
            Left            =   120
            TabIndex        =   49
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   1725
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "Num: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   48
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   46
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   4680
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   3240
         TabIndex        =   36
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   960
         TabIndex        =   35
         Top             =   2880
         Width           =   1575
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   3240
         Width           =   2175
      End
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
         Left            =   4440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   22
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   21
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   1320
         List            =   "frmEditor_NPC.frx":3345
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1680
         Width           =   3615
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   18
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   600
         Width           =   3975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stats"
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   3600
         Width           =   4815
         Begin VB.CheckBox chkProjectile 
            Caption         =   "Projectile"
            Height          =   255
            Left            =   3600
            TabIndex        =   53
            Top             =   720
            Width           =   1035
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   1680
            Max             =   255
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   3240
            Max             =   255
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   120
            Max             =   255
            TabIndex        =   8
            Top             =   720
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   1680
            Max             =   255
            TabIndex        =   7
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Str: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "End: 0"
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   15
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Int: 0"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   14
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Agi: 0"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   465
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Will: 0"
            Height          =   180
            Index           =   5
            Left            =   1680
            TabIndex        =   12
            Top             =   960
            Width           =   480
         End
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn Rate (in seconds)"
         Height          =   300
         Left            =   120
         TabIndex        =   51
         Top             =   7080
         UseMnemonic     =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Damage:"
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   2640
         TabIndex        =   37
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   2640
         TabIndex        =   24
         Top             =   2520
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7080
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
      Top             =   7800
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DropIndex As Byte

Private Sub chkDropItems_Click()
   On Error GoTo errorhandler

    If chkDropItems.Value = 1 Then
        Npc(EditorIndex).ItemBehaviour = 1
    Else
        Npc(EditorIndex).ItemBehaviour = 0
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkDropItems_Click", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub chkProjectile_Click()

   On Error GoTo errorhandler
   
     Npc(EditorIndex).Projectile = chkProjectile.Value
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "chkProjectile_Click", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmbBehaviour_Click()

   On Error GoTo errorhandler

    Npc(EditorIndex).Behaviour = cmbBehaviour.ListIndex




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

   On Error GoTo errorhandler

    ClearNPC EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    NpcEditorInit




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Form_Load()

   On Error GoTo errorhandler

    DropIndex = scrlDrop.Value
    scrlDrop.max = MAX_NPC_DROPS
    scrlDrop.min = 1
    fraDrop.Caption = "Drop - " & DropIndex
    scrlSprite.max = NumCharacters
    scrlAnimation.max = MAX_ANIMATIONS
    scrlNum.max = MAX_ITEMS




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdSave_Click()

   On Error GoTo errorhandler

    Call NpcEditorOk




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdCancel_Click()

   On Error GoTo errorhandler

    Call NpcEditorCancel




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub lstIndex_Click()

   On Error GoTo errorhandler

    NpcEditorInit




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlAnimation_Change()
Dim sString As String

   On Error GoTo errorhandler

    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnimation.Caption = "Anim: " & sString
    Npc(EditorIndex).Animation = scrlAnimation.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlDrop_Change()

   On Error GoTo errorhandler

    DropIndex = scrlDrop.Value
    fraDrop.Caption = "Drop - " & DropIndex
    txtChance.Text = Npc(EditorIndex).DropChances(DropIndex)
    scrlNum.Value = Npc(EditorIndex).DropItems(DropIndex)
    scrlValue.Value = Npc(EditorIndex).DropItemValues(DropIndex)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlDrop_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlSprite_Change()

   On Error GoTo errorhandler

    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Npc(EditorIndex).Sprite = scrlSprite.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlRange_Change()

   On Error GoTo errorhandler

    lblRange.Caption = "Range: " & scrlRange.Value
    Npc(EditorIndex).Range = scrlRange.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlNum_Change()

   On Error GoTo errorhandler

    lblNum.Caption = "Num: " & scrlNum.Value

    If scrlNum.Value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.Value).Name)
    End If
    Npc(EditorIndex).DropItems(DropIndex) = scrlNum.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlStat_Change(Index As Integer)
Dim prefix As String

   On Error GoTo errorhandler

    Select Case Index
        Case 1
            prefix = "Str: "
        Case 2
            prefix = "End: "
        Case 3
            prefix = "Int: "
        Case 4
            prefix = "Agi: "
        Case 5
            prefix = "Will: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).Value
    Npc(EditorIndex).stat(Index) = scrlStat(Index).Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlValue_Change()

   On Error GoTo errorhandler

    lblValue.Caption = "Value: " & scrlValue.Value
    Npc(EditorIndex).DropItemValues(DropIndex) = scrlValue.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtAttackSay_Change()

   On Error GoTo errorhandler

    Npc(EditorIndex).AttackSay = txtAttackSay.Text




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtChance_Change()
    On Error GoTo chanceErr
    If Not IsNumeric(txtChance.Text) Then
        Exit Sub
    End If
    If txtChance.Text < 0 Or txtChance.Text > 100 Then
        Exit Sub
    End If
    Npc(EditorIndex).DropChances(DropIndex) = txtChance.Text
    Exit Sub
chanceErr:
    MsgBox "Invalid entry for chance! " & Err.Description
    txtChance.Text = "0"
    Npc(EditorIndex).DropChances(DropIndex) = 0
End Sub

Private Sub txtDamage_Change()

   On Error GoTo errorhandler

    If Not Len(txtDamage.Text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.Text) Then Npc(EditorIndex).Damage = Val(txtDamage.Text)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtEXP_Change()

   On Error GoTo errorhandler

    If Not Len(txtExp.Text) > 0 Then Exit Sub
    If IsNumeric(txtExp.Text) Then Npc(EditorIndex).Exp = Val(txtExp.Text)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtHP_Change()

   On Error GoTo errorhandler

    If Not Len(txtHP.Text) > 0 Then Exit Sub
    If IsNumeric(txtHP.Text) Then Npc(EditorIndex).HP = Val(txtHP.Text)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtLevel_Change()

   On Error GoTo errorhandler

    If Not Len(txtLevel.Text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.Text) Then Npc(EditorIndex).Level = Val(txtLevel.Text)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtLevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long


   On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Npc(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtSpawnSecs_Change()

   On Error GoTo errorhandler

    If Not Len(txtSpawnSecs.Text) > 0 Then Exit Sub
    Npc(EditorIndex).SpawnSecs = Val(txtSpawnSecs.Text)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmbSound_Click()

   On Error GoTo errorhandler

    If cmbSound.ListIndex >= 0 Then
        Npc(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Npc(EditorIndex).sound = "None."
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
