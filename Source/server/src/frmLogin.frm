VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eclipse Origins Server Login"
   ClientHeight    =   6600
   ClientLeft      =   6375
   ClientTop       =   4110
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrNotifications 
      Interval        =   1000
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer tmrConnect 
      Interval        =   1000
      Left            =   360
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sockAuth 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraConnecting 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Connecting..."
      Height          =   4095
      Left            =   6240
      TabIndex        =   12
      Top             =   -240
      Visible         =   0   'False
      Width           =   9015
      Begin MSComctlLib.ProgressBar pbarLoading 
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   3600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblProg 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Loading...."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   3240
         Width           =   8535
      End
      Begin VB.Label lblConnecting 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Connecting!!!"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   1920
         TabIndex        =   13
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.Frame fraLogin 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Authorization Login"
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   4920
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   6015
      Begin VB.OptionButton optUseSilver 
         BackColor       =   &H00000000&
         Caption         =   "Silver License"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   1080
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optUseGold 
         BackColor       =   &H00000000&
         Caption         =   "Gold License"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Back"
         Height          =   375
         Left            =   4200
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox chkAuto 
         BackColor       =   &H00000000&
         Caption         =   "Auto-Login From Now On"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CheckBox chkRemember 
         BackColor       =   &H00000000&
         Caption         =   "Remember Login"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtLoginPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtLoginEmail 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "License:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label 
         BackColor       =   &H00000000&
         Caption         =   "Login:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Password:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Email Address:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraActivate 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Authorization Login"
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   1680
      TabIndex        =   36
      Top             =   2040
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdActivate 
         Caption         =   "Activate"
         Height          =   375
         Left            =   2760
         TabIndex        =   39
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelActivation 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4200
         TabIndex        =   38
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtActivate 
         Height          =   285
         Left            =   2040
         TabIndex        =   37
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label lblNewCode 
         BackColor       =   &H00000000&
         Caption         =   $"frmLogin.frx":C1602
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   43
         Top             =   1800
         Width           =   5295
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Activate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Activation Code:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   $"frmLogin.frx":C16B0
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Width           =   5295
      End
   End
   Begin VB.Frame fraRegister 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Authorization Login"
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   1680
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtRPass2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   29
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   2040
         TabIndex        =   24
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   2040
         TabIndex        =   23
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtREmail 
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtRPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   28
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancelRegistration 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4200
         TabIndex        =   33
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register"
         Height          =   375
         Left            =   2760
         TabIndex        =   31
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label lblRegisterDisclaim 
         BackColor       =   &H00000000&
         Caption         =   $"frmLogin.frx":C17A2
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         TabIndex        =   32
         Top             =   2280
         Width           =   5295
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Re-Type Password:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Last Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "First Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Email Address:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Password:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Register:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   9015
      Begin VB.Label lblNews 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Loading News..."
         ForeColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   1080
         TabIndex        =   35
         Top             =   600
         Width           =   6855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "News:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   720
         TabIndex        =   34
         Top             =   0
         Width           =   7575
      End
      Begin VB.Label lblExitServer 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Server"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   6000
         TabIndex        =   18
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label lblExistingUser 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Existing User"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   2640
         TabIndex        =   16
         Top             =   3600
         Width           =   3615
      End
      Begin VB.Label lblNewUser 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "New User"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   3600
         Width           =   2655
      End
   End
   Begin VB.Label lblNotifications 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6360
      Width           =   9015
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActivate_Click()
Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CActivate
    Buffer.WriteString Trim$(txtActivate.Text)
    SendDataToAuth Buffer.ToArray
    Set Buffer = Nothing
    
    fraActivate.Visible = False
    fraConnecting.Visible = False
    fraLogin.Visible = False
    fraRegister.Visible = False
    fraMenu.Visible = True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdActivate_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCancelActivation_Click()
    

   On Error GoTo errorhandler
    fraActivate.Visible = False
    txtActivate.Text = ""
    fraMenu.Visible = True
    fraLogin.Visible = False
    fraConnecting.Visible = False


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancelActivation_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdCancelRegistration_Click()


   On Error GoTo errorhandler
    fraRegister.Visible = False
    fraMenu.Visible = True
    txtFirstName.Text = ""
    txtLastName.Text = ""
    txtREmail.Text = ""
    txtRPass.Text = ""
    txtRPass2.Text = ""
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdCancelRegistration_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdExit_Click()

   On Error GoTo errorhandler

    fraMenu.Visible = True
    fraLogin.Visible = False
    txtLoginPass.Text = ""


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdExit_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdLogin_Click()
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    If IsAuthConnected Then
        lblConnecting.Caption = "Connecting!!!"
        Set Buffer = New clsBuffer
        Buffer.WriteLong CAuthLogin
        Buffer.WriteString Trim$(frmLogin.txtLoginEmail.Text)
        Buffer.WriteString Trim$(frmLogin.txtLoginPass.Text)
        Buffer.WriteString Trim$(Options.Game_Name)
        AuthSeed = rand(100000, 10000000)
        Buffer.WriteLong AuthSeed
        If optUseSilver.Value = True Then
            Buffer.WriteLong 0
        Else
            Buffer.WriteLong 1
        End If
        SendDataToAuth Buffer.ToArray
        Set Buffer = Nothing
        fraLogin.Visible = False
        fraMenu.Visible = True
        lblNotifications.Caption = "Attempting to login....."
    Else
        lblNotifications.Caption = "Not connected to authorization server. Re-trying now. Is your internet online?"
        lblNotifications.ForeColor = &HFF&
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "cmdLogin_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendDataToAuth(ByRef data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte


   On Error GoTo errorhandler

    If IsAuthConnected Then
        Set Buffer = New clsBuffer
        TempData = data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        sockAuth.SendData Buffer.ToArray
        Set Buffer = Nothing
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDataToAuth", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub cmdRegister_Click()
Dim Buffer As clsBuffer
   On Error GoTo cmdRegister_Click_Error

    If Len(Trim$(txtREmail.Text)) = 0 Or InStr(1, txtREmail.Text, "@") = 0 Then
        MsgBox "You must enter a valid email address!", vbOKOnly, "Eclipse Origins Account Registration"
        Exit Sub
    End If
    
    If Len(Trim$(txtFirstName.Text)) = 0 Then
        MsgBox "Please give a real First/Last name!", vbOKOnly, "Eclipse Origins Account Registration"
        Exit Sub
    End If
    
    If Len(Trim$(txtLastName.Text)) = 0 Then
        MsgBox "Please give a real First/Last name!", vbOKOnly, "Eclipse Origins Account Registration"
        Exit Sub
    End If
    
    If Len(Trim$(txtRPass.Text)) < 6 Then
        MsgBox "Your password must be at least 6 characters long!", vbOKOnly, "Eclipse Origins Account Registration"
        Exit Sub
    End If
    
    If Len(Trim$(txtRPass2.Text)) < 6 Then
        MsgBox "Your password must be at least 6 characters long!", vbOKOnly, "Eclipse Origins Account Registration"
        Exit Sub
    End If
    
    If Trim$(txtRPass2.Text) <> Trim$(txtRPass.Text) Then
        MsgBox "Your passwords must match!", vbOKOnly, "Eclipse Origins Account Registration"
        Exit Sub
    End If
    
    If MsgBox("By clicking yes, you verify that all of the entered information is accurate. An activation email will be sent and any accounts with fake names will be deleted!", vbYesNo, "Eclipse Origins Account Registration") = vbYes Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CNewAcc
        Buffer.WriteString txtREmail.Text
        Buffer.WriteString txtFirstName.Text
        Buffer.WriteString txtLastName.Text
        Buffer.WriteString txtRPass.Text
        SendDataToAuth Buffer.ToArray
        Set Buffer = Nothing
        fraRegister.Visible = False
        fraMenu.Visible = True
        txtREmail.Text = ""
        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtRPass.Text = ""
        txtRPass2.Text = ""
    End If

   On Error GoTo 0
   Exit Sub

cmdRegister_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRegister_Click of Form frmLogin"
End Sub

Private Sub Form_Load()
Dim i As Long, x As Long
   On Error GoTo errorhandler
   ReDim TempPlayer(0)
    i = DateTime.Second(Now) * DateTime.Minute(Now) * DateTime.Hour(Now)
    Randomize i
    InitMessages
    Set TempPlayer(0).Buffer = New clsBuffer
    sockAuth.RemoteHost = AuthIP
    sockAuth.RemotePort = AuthPort
    ConnectToAuthServer (1)
        ' load options, set if they dont exist
    If Not FileExist(App.path & "\data\options.ini", True) Then
        Options.Game_Name = "Eclipse Origins"
        Options.Port = 7001
        Options.MOTD = "Welcome to Eclipse Origins."
        Options.Website = "http://www.eclipseorigins.com"
        Options.SilentStartup = 0
        Options.Key = GenerateOptionsKey
        PutVar App.path & "\data\options.ini", "OPTIONS", "MapCount", "300"
        SaveOptions
    Else
        LoadOptions
    End If
    
    If GetSetting("Eclipse Origins", "Server" & Options.Key, "Auto") = "1" Then
        frmLogin.txtLoginEmail.Text = GetSetting("Eclipse Origins", "Server" & Options.Key, "Email")
        frmLogin.txtLoginPass.Text = GetSetting("Eclipse Origins", "Server" & Options.Key, "Pass")
        frmLogin.chkAuto.Value = 1
        frmLogin.chkRemember.Value = 1
        If GetSetting("Eclipse Origins", "Server" & Options.Key, "Gold") = "0" Then
            frmLogin.optUseSilver.Value = True
        Else
            frmLogin.optUseGold.Value = True
        End If
        AutoLogin = True
        ConnectToAuthServer (1)
    Else
        If GetSetting("Eclipse Origins", "Server" & Options.Key, "Remember") = "1" Then
            frmLogin.txtLoginEmail.Text = GetSetting("Eclipse Origins", "Server" & Options.Key, "Email")
            frmLogin.txtLoginPass.Text = GetSetting("Eclipse Origins", "Server" & Options.Key, "Pass")
            frmLogin.chkRemember.Value = 1
            If GetSetting("Eclipse Origins", "Server" & Options.Key, "Gold") = "0" Then
                frmLogin.optUseSilver.Value = True
            Else
                frmLogin.optUseGold.Value = True
            End If
        Else
            frmLogin.txtLoginEmail.Text = ""
            frmLogin.txtLoginPass.Text = ""
        End If
    End If
    
    If AutoLogin = True And Options.SilentStartup = 1 Then
        Me.Hide
    Else
        Me.Show
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   On Error GoTo errorhandler

    lblNewUser.Font.Bold = False
    lblExistingUser.Font.Bold = False
    lblExitServer.Font.Bold = False


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   On Error GoTo errorhandler

    End


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_QueryUnload", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub fraMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   On Error GoTo errorhandler

    lblNewUser.Font.Bold = False
    lblExistingUser.Font.Bold = False
    lblExitServer.Font.Bold = False

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "fraMenu_MouseMove", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblNewCode_Click()
Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CNewCode
    SendDataToAuth Buffer.ToArray
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblNewCode_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblExistingUser_Click()

   On Error GoTo errorhandler

    fraMenu.Visible = False
    fraLogin.Visible = True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblExistingUser_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblExistingUser_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   On Error GoTo errorhandler

    lblExistingUser.Font.Bold = True
    lblNewUser.Font.Bold = False
    lblExitServer.Font.Bold = False

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblExistingUser_MouseMove", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblExitServer_Click()
    

   On Error GoTo errorhandler
    End
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblExitServer_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblExitServer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    

   On Error GoTo errorhandler
   
   lblExitServer.Font.Bold = True
   lblNewUser.Font.Bold = False
   lblExistingUser.Font.Bold = False
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblExitServer_MouseMove", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblNewUser_Click()


   On Error GoTo errorhandler
    fraMenu.Visible = False
    fraRegister.Visible = True
    txtFirstName.SetFocus


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblNewUser_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblNewUser_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   On Error GoTo errorhandler

    lblNewUser.Font.Bold = True
    lblExistingUser.Font.Bold = False
    lblExitServer.Font.Bold = False

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblNewUser_MouseMove", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub sockAuth_DataArrival(ByVal bytesTotal As Long)

   On Error GoTo errorhandler

    If IsAuthConnected Then
        Call IncomingData(0, bytesTotal)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "sockAuth_DataArrival", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub tmrConnect_Timer()
Static b As Boolean, c As Long
Static initialfalse As Boolean


   On Error GoTo errorhandler
    If NoAuth = True Then Exit Sub
    If initialfalse = False Then b = True: initialfalse = True
    If AuthLost > 0 Then b = True
    If IsAuthConnected = False Or b = True Then
        If IsAuthConnected = True Then
            If b = True Then
                If ServerOnline Then
                    frmServer.lblNotifications.Caption = "Connection to Authorization Server restored!"
                    frmServer.lblNotifications.ForeColor = &HC000&
                    AuthLost = 0
                    cmdLogin_Click
                    If ReAuth = 1 Then ReAuth = 0
                    b = False
                Else
                    If AutoLogin Then
                        autolost = 0
                        cmdLogin_Click
                    End If
                    b = False
                End If
            End If
        Else
            ConnectToAuthServer 1
            If ServerOnline = False Then
                If IsAuthConnected = False Then c = c + 1
                If c = 5 Then
                    'Success
                    frmLogin.fraMenu.Visible = False
                    frmLogin.fraLogin.Visible = False
                    frmLogin.fraConnecting.Visible = True
                    frmLogin.lblConnecting.Caption = "Could not connect to auth server. Auth server maybe down or your may not have an internet connection. Running in silver mode."
                    NoAuth = True
                    Main
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "tmrConnect_Timer", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function ConnectToAuthServer(ByVal i As Long) As Boolean
Dim Wait As Long
Static authLostTmr As Long
    
    ' Check to see if we are already connected, if so just exit

   On Error GoTo errorhandler

    If IsAuthConnected Then
        ConnectToAuthServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    sockAuth.Close
    sockAuth.Connect
    
    If authLostTmr < GetTickCount Then
        If AuthLost > 0 Then
            frmServer.lblNotifications.Caption = "Reconnecting to Authorization Server! Attempt #" & AuthLost
            lblNotifications.ForeColor = &HFF&
            AuthLost = AuthLost + 1
            authLostTmr = GetTickCount + 480000
        Else
            frmServer.lblNotifications.Caption = "Connecting to Authorization Server!"
            lblNotifications.ForeColor = &HFF&
        End If
    End If
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsAuthConnected) And (GetTickCount <= Wait + 50)
        DoEvents
    Loop
    Randomize
    Rnd
    Rnd
    
    ConnectToAuthServer = IsAuthConnected


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ConnectToAuthServer", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function IsAuthConnected() As Boolean
    

   On Error GoTo errorhandler

    If sockAuth.State = sckConnected Then
        IsAuthConnected = True
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsAuthConnected", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
    
End Function

Private Sub tmrNotifications_Timer()
    Static LastNotification As String
    Static TimeShown As Long
    

   On Error GoTo errorhandler

    If lblNotifications.Caption <> "" Then
        If lblNotifications.Caption = LastNotification Then
            If TimeShown >= 6 Then
                LastNotification = ""
                lblNotifications.Caption = ""
                TimeShown = 0
            Else
                TimeShown = TimeShown + 1
            End If
        Else
            LastNotification = lblNotifications.Caption
            TimeShown = 0
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "tmrNotifications_Timer", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

