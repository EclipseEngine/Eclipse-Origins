VERSION 5.00
Begin VB.Form frmServers 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eclipse Origins Game Engine"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmServers.frx":0000
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDefault 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   6480
      Width           =   195
   End
   Begin VB.PictureBox picDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   6480
      Picture         =   "frmServers.frx":E1042
      ScaleHeight     =   450
      ScaleWidth      =   1350
      TabIndex        =   3
      Top             =   6480
      Width           =   1350
   End
   Begin VB.PictureBox picUse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   5040
      Picture         =   "frmServers.frx":E3064
      ScaleHeight     =   450
      ScaleWidth      =   1350
      TabIndex        =   2
      Top             =   6480
      Width           =   1350
   End
   Begin VB.PictureBox picAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   3600
      Picture         =   "frmServers.frx":E5086
      ScaleHeight     =   450
      ScaleWidth      =   1350
      TabIndex        =   1
      Top             =   6480
      Width           =   1350
   End
   Begin VB.ListBox lstServers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3930
      Left            =   1800
      TabIndex        =   0
      Top             =   2460
      Width           =   6000
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Server List"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Make Default"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   6480
      Width           =   1455
   End
End
Attribute VB_Name = "frmServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)


   On Error GoTo errorhandler
    If KeyAscii = vbKeyEscape Then
        End
    End If

    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmServers", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_Load()


   On Error GoTo errorhandler
    
    frmMain.Socket.Close
    ' load options
    Call SetStatus("Loading Options...")
    LoadOptions
    PopulateServerList
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmServers", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub PopulateServerList()
Dim i As Long
   On Error GoTo errorhandler
    lstServers.Clear
    
    If ServerCount > 0 Then
        For i = 1 To ServerCount
            If Trim$(Servers(i).Game_Name) = "" Then
                lstServers.AddItem "Unnamed Game - " & " IP: " & Trim$(Servers(i).ip) & " Port: " & Servers(i).port
            Else
                lstServers.AddItem Trim$(Servers(i).Game_Name) & " - " & " IP: " & Trim$(Servers(i).ip) & " Port: " & Servers(i).port
            End If
        Next
    Else
        lstServers.AddItem "No Servers!"
    End If
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PopulateServerList", "frmServers", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


   On Error GoTo errorhandler
    End
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_QueryUnload", "frmServers", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Label_Click()


   On Error GoTo errorhandler

    If chkDefault.Value = 0 Then
        chkDefault.Value = 1
    Else
        chkDefault.Value = 0
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Label_Click", "frmServers", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub picAdd_Click()
Dim ip As String, port As Long

   On Error GoTo errorhandler
    ip = InputBox("What is the address of the Eclipse Origins game server that you are trying to connect to?", "Add Eclipse Origins game server.")
    If Len(Trim$(ip)) > 0 Then
        port = Val(InputBox("What is the port of the Eclipse Origins game server that you are trying to connect to?", "Add Eclipse Origins game server."))
        If port > 0 Then
            ServerCount = ServerCount + 1
            ReDim Preserve Servers(ServerCount)
            Servers(ServerCount).Game_Name = ""
            Servers(ServerCount).ip = ip
            Servers(ServerCount).port = port
            Servers(ServerCount).SavePass = 0
            Servers(ServerCount).Username = ""
            Servers(ServerCount).Password = ""
            SaveServers
            PopulateServerList
        End If
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "picAdd_Click", "frmServers", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub picDelete_Click()
Dim i As Long, X As Long

   On Error GoTo errorhandler

    If ServerCount > 0 Then
        If lstServers.ListIndex > -1 Then
            If MsgBox("Are you sure you want to remove this server?", vbYesNo) = vbYes Then
                i = lstServers.ListIndex + 1
                For X = i + 1 To ServerCount
                    Servers(X - 1).Game_Name = Servers(X).Game_Name
                    Servers(X - 1).ip = Servers(X).ip
                    Servers(X - 1).port = Servers(X).port
                    Servers(X - 1).Username = Servers(X).Username
                    Servers(X - 1).Password = Servers(X).Password
                    Servers(X - 1).SavePass = Servers(X).SavePass
                Next
                ServerCount = ServerCount - 1
                ReDim Preserve Servers(ServerCount)
                PopulateServerList
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "picDelete_Click", "frmServers", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub picUse_Click()


   On Error GoTo errorhandler

    If ServerCount > 0 Then
        If lstServers.ListIndex > -1 Then
            ServerIndex = lstServers.ListIndex + 1
            If chkDefault.Value = 1 Then
                Options.DefaultServer = ServerIndex
                SaveOptions
            End If
            frmServers.Visible = False
            QuitConnecting = False
            Main
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "picUse_Click", "frmServers", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
