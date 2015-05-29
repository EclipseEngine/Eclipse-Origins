VERSION 5.00
Begin VB.Form frmEditor_Pet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pet Editor"
   ClientHeight    =   6750
   ClientLeft      =   945
   ClientTop       =   480
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
   Icon            =   "frmEditor_Pet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Frame u 
      Caption         =   "Pet Properties"
      Height          =   6015
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.Frame Frame4 
         Caption         =   "Leveling"
         Height          =   1215
         Left            =   120
         TabIndex        =   23
         Top             =   3240
         Width           =   4815
         Begin VB.OptionButton optDoNotLevel 
            Caption         =   "Does not level."
            Height          =   255
            Left            =   2160
            TabIndex        =   40
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optLevel 
            Caption         =   "Level by Exp."
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1695
         End
         Begin VB.PictureBox picPetlevel 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   4575
            TabIndex        =   41
            Top             =   480
            Width           =   4575
            Begin VB.HScrollBar scrlMaxLevel 
               Height          =   255
               Left            =   3000
               Max             =   100
               Min             =   1
               TabIndex        =   46
               Top             =   120
               Value           =   1
               Width           =   1215
            End
            Begin VB.HScrollBar scrlPetPnts 
               Height          =   255
               Left            =   1440
               Max             =   100
               TabIndex        =   44
               Top             =   120
               Value           =   5
               Width           =   1215
            End
            Begin VB.HScrollBar scrlPetExp 
               Height          =   255
               Left            =   0
               Max             =   500
               Min             =   -500
               TabIndex        =   42
               Top             =   120
               Value           =   100
               Width           =   1335
            End
            Begin VB.Label lblmaxlevel 
               AutoSize        =   -1  'True
               Caption         =   "Max Level: 1"
               Height          =   180
               Left            =   3000
               TabIndex        =   47
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblPetPnts 
               AutoSize        =   -1  'True
               Caption         =   "Points Per Level: 5"
               Height          =   180
               Left            =   1440
               TabIndex        =   45
               Top             =   360
               Width           =   1440
            End
            Begin VB.Label lblPetExp 
               AutoSize        =   -1  'True
               Caption         =   "Exp Gain: 100%"
               Height          =   180
               Left            =   0
               TabIndex        =   43
               Top             =   360
               Width           =   1260
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Spells"
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   4560
         Width           =   4815
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Index           =   3
            Left            =   120
            Max             =   255
            TabIndex        =   21
            Top             =   840
            Width           =   1455
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Index           =   2
            Left            =   2400
            Max             =   255
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Index           =   4
            Left            =   2400
            Max             =   255
            TabIndex        =   16
            Top             =   840
            Width           =   1455
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblSpell 
            AutoSize        =   -1  'True
            Caption         =   "Spell 3: None"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label lblSpell 
            AutoSize        =   -1  'True
            Caption         =   "Spell 2: None"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   20
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label lblSpell 
            AutoSize        =   -1  'True
            Caption         =   "Spell 4: None"
            Height          =   180
            Index           =   4
            Left            =   2400
            TabIndex        =   19
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label lblSpell 
            AutoSize        =   -1  'True
            Caption         =   "Spell 1: None"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   1020
         End
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
         TabIndex        =   7
         Top             =   600
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   6
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   4
         Top             =   960
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Caption         =   "Starting Stats"
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   4815
         Begin VB.OptionButton optAdoptStats 
            Caption         =   "Adopt Owner's Stats"
            Height          =   255
            Left            =   2160
            TabIndex        =   25
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton optCustomStats 
            Caption         =   "Custom Stats"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1695
         End
         Begin VB.PictureBox picCustomStats 
            BorderStyle     =   0  'None
            Height          =   1215
            Left            =   120
            ScaleHeight     =   1215
            ScaleWidth      =   4575
            TabIndex        =   26
            Top             =   480
            Width           =   4575
            Begin VB.HScrollBar scrlStat 
               Height          =   255
               Index           =   5
               Left            =   1560
               Max             =   255
               TabIndex        =   32
               Top             =   720
               Width           =   1455
            End
            Begin VB.HScrollBar scrlStat 
               Height          =   255
               Index           =   4
               Left            =   0
               Max             =   255
               TabIndex        =   31
               Top             =   720
               Width           =   1455
            End
            Begin VB.HScrollBar scrlStat 
               Height          =   255
               Index           =   3
               Left            =   3120
               Max             =   255
               TabIndex        =   30
               Top             =   120
               Width           =   1455
            End
            Begin VB.HScrollBar scrlStat 
               Height          =   255
               Index           =   2
               Left            =   1560
               Max             =   255
               TabIndex        =   29
               Top             =   120
               Width           =   1455
            End
            Begin VB.HScrollBar scrlStat 
               Height          =   255
               Index           =   1
               Left            =   0
               Max             =   255
               TabIndex        =   28
               Top             =   120
               Width           =   1455
            End
            Begin VB.HScrollBar scrlStat 
               Height          =   255
               Index           =   6
               Left            =   3120
               Max             =   255
               TabIndex        =   27
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label lblStat 
               AutoSize        =   -1  'True
               Caption         =   "Will: 0"
               Height          =   180
               Index           =   5
               Left            =   1560
               TabIndex        =   38
               Top             =   960
               Width           =   480
            End
            Begin VB.Label lblStat 
               AutoSize        =   -1  'True
               Caption         =   "Agi: 0"
               Height          =   180
               Index           =   4
               Left            =   0
               TabIndex        =   37
               Top             =   960
               Width           =   465
            End
            Begin VB.Label lblStat 
               AutoSize        =   -1  'True
               Caption         =   "Int: 0"
               Height          =   180
               Index           =   3
               Left            =   3120
               TabIndex        =   36
               Top             =   360
               Width           =   435
            End
            Begin VB.Label lblStat 
               AutoSize        =   -1  'True
               Caption         =   "End: 0"
               Height          =   180
               Index           =   2
               Left            =   1560
               TabIndex        =   35
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblStat 
               AutoSize        =   -1  'True
               Caption         =   "Str: 0"
               Height          =   180
               Index           =   1
               Left            =   0
               TabIndex        =   34
               Top             =   360
               Width           =   435
            End
            Begin VB.Label lblStat 
               AutoSize        =   -1  'True
               Caption         =   "Level: 0"
               Height          =   180
               Index           =   6
               Left            =   3120
               TabIndex        =   33
               Top             =   960
               Width           =   615
            End
         End
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pet List"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6180
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Pet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    

   On Error GoTo errorhandler

    ClearPet EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pet(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    PetEditorInit


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Form_Load()
    

   On Error GoTo errorhandler

    scrlSprite.max = NumCharacters
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdSave_Click()
    

   On Error GoTo errorhandler

    Call PetEditorOk
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdCancel_Click()
    

   On Error GoTo errorhandler

    Call PetEditorCancel


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lstIndex_Click()
    

   On Error GoTo errorhandler

    PetEditorInit
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub optAdoptStats_Click()


   On Error GoTo errorhandler
    If optAdoptStats.Value = True Then
        picCustomStats.Visible = False
        Pet(EditorIndex).StatType = 0
    Else
        picCustomStats.Visible = True
        Pet(EditorIndex).StatType = 1
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optAdoptStats_Click", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub optCustomStats_Click()


   On Error GoTo errorhandler

    If optAdoptStats.Value = True Then
        picCustomStats.Visible = False
        Pet(EditorIndex).StatType = 0
    Else
        picCustomStats.Visible = True
        Pet(EditorIndex).StatType = 1
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optCustomStats_Click", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub optDoNotLevel_Click()


   On Error GoTo errorhandler
    If optDoNotLevel.Value = True Then
        picPetlevel.Visible = False
        Pet(EditorIndex).LevelingType = 1
    Else
        picPetlevel.Visible = True
        Pet(EditorIndex).LevelingType = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optDoNotLevel_Click", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub optLevel_Click()


   On Error GoTo errorhandler
    If optDoNotLevel.Value = True Then
        picPetlevel.Visible = False
        Pet(EditorIndex).LevelingType = 1
    Else
        picPetlevel.Visible = True
        Pet(EditorIndex).LevelingType = 0
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "optLevel_Click", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlMaxLevel_Change()


   On Error GoTo errorhandler
    lblmaxlevel.Caption = "Max Level: " & scrlMaxLevel.Value
    Pet(EditorIndex).MaxLevel = scrlMaxLevel.Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlMaxLevel_Change", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlPetExp_Change()


   On Error GoTo errorhandler
    lblPetExp.Caption = "Exp Gain: " & scrlPetExp.Value & "%"
    Pet(EditorIndex).ExpGain = scrlPetExp.Value
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlPetExp_Change", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlPetPnts_Change()


   On Error GoTo errorhandler
    lblPetPnts.Caption = "Points Per Level: " & scrlPetPnts.Value
    Pet(EditorIndex).LevelPnts = scrlPetPnts.Value
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlPetPnts_Change", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub scrlSpell_Change(Index As Integer)
Dim prefix As String
    
    

   On Error GoTo errorhandler

    prefix = "Spell " & Index & ": "
    
    If scrlSpell(Index).Value = 0 Then
        lblSpell(Index).Caption = prefix & "None"
    Else
        lblSpell(Index).Caption = prefix & Trim$(spell(scrlSpell(Index).Value).Name)
    End If
    
    
    
    Pet(EditorIndex).spell(Index) = scrlSpell(Index).Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Private Sub scrlSprite_Change()
    

   On Error GoTo errorhandler

    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Pet(EditorIndex).Sprite = scrlSprite.Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Private Sub scrlRange_Change()
    

   On Error GoTo errorhandler

    lblRange.Caption = "Range: " & scrlRange.Value
    Pet(EditorIndex).Range = scrlRange.Value
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_Pet", Err.Number, Err.Description, Erl
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
        Case 6
            prefix = "Level: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).Value
    
    If Index = 0 Then
    ElseIf Index = 6 Then
        Pet(EditorIndex).Level = scrlStat(Index).Value
    Else
        Pet(EditorIndex).stat(Index) = scrlStat(Index).Value
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    

   On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Pet(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pet(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Pet", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
