VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9570
   ClientLeft      =   -345
   ClientTop       =   2310
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   640
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************
' ** Events **
' ************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo errorhandler
   
   If InGame Then
        If KeyCode = 116 Then
            HideChat = Not HideChat
        End If
        
        If KeyCode = 117 Then
            HideMenu = Not HideMenu
        End If
        
        If KeyCode = 118 Then
            HideBars = Not HideBars
            BarWidth_GuiHP = 0
            BarWidth_GuiSP = 0
            BarWidth_GuiEXP = 0
        End If
        
        If KeyCode = 119 Then
            HideHotbar = Not HideHotbar
        End If
    End If
   
    If KeyCode = 123 Then
         DebugMode = Not DebugMode
         If DebugMode Then
             Options.Debug = 1
             SaveOptions
         Else
             Options.Debug = 0
             SaveOptions
         End If
         UpdateDebugCaption
    End If

    If KeyCode = 17 Then
        If HideChat = False Or HideHotbar = False Or HideMenu = False Or HideBars = False Then
            hideGUI = True
        Else
            hideGUI = False
        End If
        If hideGUI = True Then
            HideChat = True
            HideMenu = True
            HideHotbar = True
            HideBars = True
            CurrentGameMenu = 0
            BarWidth_GuiHP = 0
            BarWidth_GuiSP = 0
            BarWidth_GuiEXP = 0
        Else
            HideChat = False
            HideMenu = False
            HideHotbar = False
            HideBars = False
            BarWidth_GuiHP = 0
            BarWidth_GuiSP = 0
            BarWidth_GuiEXP = 0
        End If
    ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyDelete Or KeyCode = vbKeyTab Then
        HandleMenuKeypress KeyCode
    ElseIf KeyCode = vbKeyTab Then
        If TabDown1 = False Then
            HandleMenuKeypress vbKeyTab
            TabDown1 = True
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_KeyDown", "frmMain", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_Resize()
    MAX_MAPX = (Me.Width / Screen.TwipsPerPixelX) / 32
    MAX_MAPY = (Me.Height / Screen.TwipsPerPixelY) / 32
    MAX_MAPY = MAX_MAPY - 2
    MAX_MAPX = MAX_MAPX - 1
    HalfX = ((MAX_MAPX + 1) / 2) * PIC_X
    HalfY = ((MAX_MAPY + 1) / 2) * PIC_Y
    ScreenX = (MAX_MAPX + 1) * PIC_X
    ScreenY = (MAX_MAPY + 1) * PIC_Y
    StartXValue = ((MAX_MAPX + 1) / 2)
    StartYValue = ((MAX_MAPY + 1) / 2)
    EndXValue = (MAX_MAPX + 1)
    EndYValue = (MAX_MAPY + 1)
    Half_PIC_X = (PIC_X / 2)
    Half_PIC_Y = (PIC_Y / 2)
    BBWidth = (MAX_MAPX + 1) * 32
    BBHeight = (MAX_MAPY + 1) * 32
    GameScreenWidth = BBWidth
    GameScreenHeight = BBHeight
    GameWindowWidth = BBWidth
    GameWindowHeight = BBHeight
    Render_Graphics
End Sub

Private Sub Form_Unload(Cancel As Integer)


   On Error GoTo errorhandler
    
    If InGame = True Then
        Cancel = True
        logoutGame
    Else
        DestroyGame
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub form_dblClick()
    HandleGame_DblClick
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' hide the descriptions

   On Error GoTo errorhandler

    ItemDescVisible = False
    SpellDescVisible = False

    If InGame Then
        'X = X * (GameScreenWidth / frmMain.ScaleWidth)
        'Y = Y * (GameScreenHeight / frmMain.ScaleHeight)
    Else
        X = X * (MenuWidth / frmMain.ScaleWidth)
        Y = Y * (MenuHeight / frmMain.ScaleHeight)
    End If
    If X >= GameScreenBounds.Left And X <= GameScreenBounds.Right Then
        If Y >= GameScreenBounds.Top And Y <= GameScreenBounds.Bottom Then
            CurX = TileView.Left + (((X - GameScreenBounds.Left) + Camera.Left) \ PIC_X)
            CurY = TileView.Top + (((Y - GameScreenBounds.Top) + Camera.Top) \ PIC_Y)
        End If
    End If
    GlobalX = X
    GlobalY = Y

    If InMapEditor Then
        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, X, Y)
        End If
    End If
    HandleGame_MouseMove



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo errorhandler

    MouseBtn = Button
    If InGame = False Then
        HandleMenu_MouseDown
    Else
        If HandleGame_MouseDown = False Then
            If InMapEditor Then
                Call MapEditorMouseDown(Button, X, Y, False)
            Else
                ' left click
                If Button = vbLeftButton Then
                    If ShiftDown Then
                        If Player(MyIndex).Pet.Alive Then
                            If isInBounds Then
                                Call PetMove(CurX, CurY)
                            End If
                        Else
                            If Options.ClicktoWalk = 1 Then
                                If CheckClickArrow(X, Y) = False Then
                                    WalkToX = CurX
                                    WalkToY = CurY
                                End If
                            End If
                            Call PlayerSearch(CurX, CurY)
                        End If
                    Else
                        If Options.ClicktoWalk = 1 Then
                            If CheckClickArrow(X, Y) = False Then
                                WalkToX = CurX
                                WalkToY = CurY
                            End If
                        End If
                        Call PlayerSearch(CurX, CurY)
                    End If
                ' right click
                ElseIf Button = vbRightButton Then
                    If ShiftDown Then
                        ' admin warp if we're pressing shift and right clicking
                        If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
                    Else
                        CheckMapGetItem
                    End If
                End If
            End If
        End If
        If frmEditor_Events.Visible Then frmEditor_Events.SetFocus
    End If



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_MouseDown", "frmMain", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function CheckClickArrow(X As Single, Y As Single) As Boolean

   On Error GoTo errorhandler

    If Map.Up > 0 Then
        If CurY = 0 Then
            WalkToY = -1
            WalkToX = CurX
            CheckClickArrow = True
        End If
    End If
    If Map.Right > 0 Then
        If CurX >= Map.MaxX Then
            WalkToX = Map.MaxX + 1
            WalkToY = CurY
            CheckClickArrow = True
        End If
    End If
    If Map.Down > 0 Then
        If CurY >= Map.MaxY Then
            WalkToY = Map.MaxY + 1
            WalkToX = CurX
            CheckClickArrow = True
        End If
    End If
    If Map.Left > 0 Then
        If CurX = 0 Then
            WalkToX = -1
            WalkToY = CurY
            CheckClickArrow = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CheckClickArrow", "frmMain", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo errorhandler

    If InGame = False Then
        HandleMenu_MouseUp
    Else
        HandleGame_MouseUp
        ResetGUIButtons
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_MouseUp", "frmMain", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)



   On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)


   On Error GoTo errorhandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim buffer As clsBuffer
   On Error GoTo errorhandler

    If InGame = False Then Exit Sub
        
        Select Case KeyCode
            Case vbKeyInsert
            If Player(MyIndex).Access > 0 Then
                If frmAdmin.Visible = False Then
                    Set buffer = New clsBuffer
                    buffer.WriteLong CAdmin
                    SendData buffer.ToArray
                    Set buffer = Nothing
                Else
                    frmAdmin.Visible = False
                End If
            End If
        End Select
        If chatOn = False And CurrencyMenu = 0 And dialogueIndex = 0 And EventChat = False Then
            If KeyCode >= 49 And KeyCode <= 58 Then
                SendHotbarUse 1 + (KeyCode - 49)
            End If
            If KeyCode = 189 Then
                SendHotbarUse 11
            End If
            If KeyCode = 190 Then
                SendHotbarUse 12
            End If
        End If
        ' handles delete events
        If KeyCode = vbKeyDelete Then
            If InMapEditor Then DeleteEvent CurX, CurY
        End If
        ' handles copy + pasting events
        If KeyCode = vbKeyC Then
            If InMapEditor Then
                CopyEvent_Map CurX, CurY
            End If
        End If
        If KeyCode = vbKeyV Then
            If InMapEditor Then
                PasteEvent_Map CurX, CurY
            End If
        End If
        If KeyCode = vbKeyTab Then
            TabDown1 = False
        End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
