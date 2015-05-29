VERSION 5.00
Begin VB.Form frmEditor_House 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "House Editor"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
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
   Icon            =   "frmEditor_House.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "House Properties"
      Height          =   4455
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtYEntrance 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   1365
         Width           =   2655
      End
      Begin VB.TextBox txtXEntrance 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   1005
         Width           =   2655
      End
      Begin VB.TextBox txtHousePrice 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   1725
         Width           =   2655
      End
      Begin VB.TextBox txtHouseFurniture 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   2085
         Width           =   2655
      End
      Begin VB.TextBox txtBaseMap 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   645
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrance Y:"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrance X:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblHousePrice 
         BackStyle       =   0  'Transparent
         Caption         =   "Price:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Pieces of Furniture (0 for no max):"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblHouseMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Base Map:"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "House List"
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   4560
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
End
Attribute VB_Name = "frmEditor_House"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()

   On Error GoTo errorhandler

    If LenB(Trim$(txtName.Text)) = 0 Then
        Call MsgBox("Name required.")
        Exit Sub
    End If
    
    If LenB(Trim$(txtBaseMap.Text)) = 0 Then
        Call MsgBox("Base map required.")
        Exit Sub
    End If
    
    HouseEditorOk





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_House", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdCancel_Click()

   On Error GoTo errorhandler

    
    HouseEditorCancel



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_House", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub lstIndex_Click()

   On Error GoTo errorhandler

    HouseEditorInit




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_House", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtBaseMap_Change()


   On Error GoTo errorhandler
    If EditorIndex = 0 Then Exit Sub
    If IsNumeric(txtBaseMap.Text) Then
        If Val(txtBaseMap.Text) < 1 Or Val(txtBaseMap.Text) > MAX_MAPS Then Exit Sub
        House(EditorIndex).BaseMap = Val(txtBaseMap.Text)
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtBaseMap_Change", "frmEditor_House", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtHouseFurniture_Change()
    If EditorIndex = 0 Then Exit Sub
    If IsNumeric(txtHouseFurniture.Text) Then
        If Val(txtHouseFurniture.Text) < 0 Then Exit Sub
        House(EditorIndex).MaxFurniture = Val(txtHouseFurniture.Text)
    End If
End Sub

Private Sub txtHousePrice_Change()


   On Error GoTo errorhandler
    If EditorIndex = 0 Then Exit Sub
    If IsNumeric(txtHousePrice.Text) Then
        House(EditorIndex).Price = Val(txtHousePrice.Text)
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtHousePrice_Change", "frmEditor_House", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long


   On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    House(EditorIndex).ConfigName = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & House(EditorIndex).ConfigName, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_House", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub txtXEntrance_Change()


   On Error GoTo errorhandler
    If EditorIndex = 0 Then Exit Sub
    If IsNumeric(txtXEntrance.Text) Then
        If Val(txtXEntrance.Text) < 0 Or Val(txtXEntrance.Text) > 255 Then Exit Sub
        House(EditorIndex).X = Val(txtXEntrance.Text)
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtXEntrance_Change", "frmEditor_House", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtYEntrance_Change()


   On Error GoTo errorhandler
    If EditorIndex = 0 Then Exit Sub
    If IsNumeric(txtYEntrance.Text) Then
        If Val(txtYEntrance.Text) < 0 Or Val(txtYEntrance.Text) > 255 Then Exit Sub
        House(EditorIndex).Y = Val(txtYEntrance.Text)
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtYEntrance_Change", "frmEditor_House", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
