VERSION 5.00
Begin VB.Form frmEditor_Zone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zone Editor"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8520
      TabIndex        =   17
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Zone Editor"
      Height          =   6975
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame5 
         Caption         =   "Zone Properties"
         Height          =   2415
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   6495
         Begin VB.Frame Frame6 
            Caption         =   "Weather Chances"
            Height          =   2055
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   3735
            Begin VB.HScrollBar scrlWeatherIntensity 
               Height          =   255
               Left            =   2040
               Max             =   100
               TabIndex        =   30
               Top             =   1680
               Width           =   1575
            End
            Begin VB.HScrollBar scrlWeather 
               Height          =   255
               Index           =   5
               Left            =   2040
               Max             =   100
               TabIndex        =   28
               Top             =   1080
               Width           =   1575
            End
            Begin VB.HScrollBar scrlWeather 
               Height          =   255
               Index           =   4
               Left            =   2040
               Max             =   100
               TabIndex        =   26
               Top             =   480
               Width           =   1575
            End
            Begin VB.HScrollBar scrlWeather 
               Height          =   255
               Index           =   3
               Left            =   120
               Max             =   100
               TabIndex        =   24
               Top             =   1680
               Width           =   1575
            End
            Begin VB.HScrollBar scrlWeather 
               Height          =   255
               Index           =   2
               Left            =   120
               Max             =   100
               TabIndex        =   22
               Top             =   1080
               Width           =   1575
            End
            Begin VB.HScrollBar scrlWeather 
               Height          =   255
               Index           =   1
               Left            =   120
               Max             =   100
               TabIndex        =   20
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label lblWeatherIntensity 
               Caption         =   "Intensity: 0/100"
               Height          =   255
               Left            =   2040
               TabIndex        =   29
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label lblWeather 
               Caption         =   "Storm: 0/100"
               Height          =   255
               Index           =   5
               Left            =   2040
               TabIndex        =   27
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label lblWeather 
               Caption         =   "Sand Storm: 0/100"
               Height          =   255
               Index           =   4
               Left            =   2040
               TabIndex        =   25
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label lblWeather 
               Caption         =   "Hail: 0/100"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   23
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label lblWeather 
               Caption         =   "Snow: 0/100"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   21
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label lblWeather 
               Caption         =   "Rain: 0/100"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   1575
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Zone NPCS"
         Height          =   3495
         Left            =   3480
         TabIndex        =   11
         Top             =   840
         Width           =   3135
         Begin VB.CommandButton cmdRemoveNpc 
            Caption         =   "Remove From List"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   3120
            Width           =   2895
         End
         Begin VB.ComboBox cmbNpc 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2760
            Width           =   2895
         End
         Begin VB.ListBox lstNpcs 
            Height          =   2400
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Zone Maps"
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3135
         Begin VB.CommandButton cmdRemoveMap 
            Caption         =   "Remove From List"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   3120
            Width           =   2895
         End
         Begin VB.CommandButton cmdAddMap 
            Caption         =   "Add"
            Height          =   255
            Left            =   2400
            TabIndex        =   9
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtAddMap 
            Height          =   255
            Left            =   720
            TabIndex        =   8
            Text            =   "0"
            Top             =   2760
            Width           =   1575
         End
         Begin VB.ListBox lstMaps 
            Height          =   2400
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "Map #:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   2760
            Width           =   735
         End
      End
      Begin VB.TextBox txtZoneName 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   330
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Zone Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Zones"
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7080
         ItemData        =   "frmEditor_Zone.frx":0000
         Left            =   120
         List            =   "frmEditor_Zone.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Zone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbNpc_Click()
Dim i As Long, X As Long

   On Error GoTo errorhandler

    i = EditorIndex
    If frmEditor_Zone.lstNpcs.ListIndex > -1 Then
        MapZones(EditorIndex).NPCs(frmEditor_Zone.lstNpcs.ListIndex + 1) = cmbNpc.ListIndex
        frmEditor_Zone.lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS * 2
            If MapZones(i).NPCs(X) > 0 Then
                frmEditor_Zone.lstNpcs.AddItem CStr(MapZones(i).NPCs(X)) & ". " & Trim$(Npc(MapZones(i).NPCs(X)).Name)
            Else
                frmEditor_Zone.lstNpcs.AddItem "No NPC"
            End If
        Next
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmbNpc_Click", "frmEditor_Zone", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdAddMap_Click()
Dim i As Long, X As Long, newmap As Long

   On Error GoTo errorhandler

    i = EditorIndex
    If IsNumeric(txtAddMap.Text) Then
        If Val(txtAddMap.Text) > 0 And Val(txtAddMap.Text) <= MAX_MAPS Then
            newmap = Val(txtAddMap.Text)
        Else
            MsgBox "Map number invalid!"
            Exit Sub
        End If
    Else
        MsgBox "Map number must be numeric!"
        Exit Sub
    End If
    If MapZones(i).MapCount > 0 Then
        For X = 1 To MapZones(i).MapCount
            If MapZones(i).Maps(X) = newmap Then
                MsgBox "Map already exists in zone!"
                Exit Sub
            End If
        Next
        MapZones(i).MapCount = MapZones(i).MapCount + 1
        ReDim Preserve MapZones(i).Maps(MapZones(i).MapCount)
        MapZones(i).Maps(MapZones(i).MapCount) = newmap
    Else
        MapZones(i).MapCount = 1
        ReDim Preserve MapZones(i).Maps(MapZones(i).MapCount)
        MapZones(i).Maps(1) = newmap
    End If
    frmEditor_Zone.lstMaps.Clear
    If MapZones(i).MapCount > 0 Then
        For X = 1 To MapZones(i).MapCount
            frmEditor_Zone.lstMaps.AddItem "Map #" & MapZones(i).Maps(X)
        Next
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdAddMap_Click", "frmEditor_Zone", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdCancel_Click()

   On Error GoTo errorhandler

    Call ZoneEditorCancel




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Zone", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdRemoveMap_Click()
Dim i As Long, X As Long, z As Long

   On Error GoTo errorhandler

    i = EditorIndex
    If frmEditor_Zone.lstMaps.ListIndex > -1 Then
        z = frmEditor_Zone.lstMaps.ListIndex + 1
        For X = z + 1 To MapZones(i).MapCount
            MapZones(i).Maps(X - 1) = MapZones(i).Maps(X)
        Next
        MapZones(i).MapCount = MapZones(i).MapCount - 1
        ReDim Preserve MapZones(i).Maps(MapZones(i).MapCount)
    End If
    frmEditor_Zone.lstMaps.Clear
    If MapZones(i).MapCount > 0 Then
        For X = 1 To MapZones(i).MapCount
            frmEditor_Zone.lstMaps.AddItem "Map #" & MapZones(i).Maps(X)
        Next
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdRemoveMap_Click", "frmEditor_Zone", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdRemoveNpc_Click()
Dim i As Long, X As Long

   On Error GoTo errorhandler

    i = EditorIndex
    If frmEditor_Zone.lstNpcs.ListIndex > -1 Then
        MapZones(EditorIndex).NPCs(frmEditor_Zone.lstNpcs.ListIndex + 1) = 0
        frmEditor_Zone.lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS * 2
            If MapZones(i).NPCs(X) > 0 Then
                frmEditor_Zone.lstNpcs.AddItem CStr(MapZones(i).NPCs(X)) & ". " & Trim$(Npc(MapZones(i).NPCs(X)).Name)
            Else
                frmEditor_Zone.lstNpcs.AddItem "No NPC"
            End If
        Next
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdRemoveNpc_Click", "frmEditor_Zone", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub cmdSave_Click()

   On Error GoTo errorhandler

    Call ZoneEditorOk




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Zone", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub lstIndex_Click()

   On Error GoTo errorhandler

    ZoneEditorInit




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Zone", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlWeatherIntensity_Change()

   On Error GoTo errorhandler

    lblWeatherIntensity.Caption = "Intensity: " & scrlWeatherIntensity.Value & "/100"
    MapZones(EditorIndex).WeatherIntensity = scrlWeatherIntensity.Value




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlWeatherIntensity_Change", "frmEditor_Zone", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub scrlWeather_Change(Index As Integer)
Dim X As String

   On Error GoTo errorhandler

    Select Case Index
        Case 1
            X = "Rain"
        Case 2
            X = "Snow"
        Case 3
            X = "Hail"
        Case 4
            X = "Sand Storm"
        Case 5
            X = "Storm"
    End Select
    lblWeather(Index).Caption = X & ": " & scrlWeather(Index).Value & "/100"
    MapZones(EditorIndex).Weather(Index) = scrlWeather(Index).Value


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "scrlWeather_Change", "frmEditor_Zone", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub txtZoneName_Validate(Cancel As Boolean)
Dim tmpIndex As Long


   On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ZONES Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    MapZones(EditorIndex).Name = Trim$(txtZoneName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & MapZones(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "txtZoneName_Validate", "frmEditor_Zone", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
