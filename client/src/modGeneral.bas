Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1


Public Sub Main()
Dim reconnectCount As Long, curVersion As Long, myVersion As Long, i As Long
    ' set loading screen

   On Error GoTo errorhandler
   
   MapCacheX = -1
   MapCacheY = -1
   Dim Width As Long, Height As Long
    DoEvents
    'Width = GetSystemMetrics(SM_CXSCREEN)
    'Height = GetSystemMetrics(SM_CYSCREEN)
    MAX_MAPX = Width / PIC_X
    MAX_MAPY = Height / PIC_Y
    'frmMain.Width = Width
    'frmMain.Height = Height
    HalfX = ((MAX_MAPX + 1) / 2) * PIC_X
    HalfY = ((MAX_MAPY + 1) / 2) * PIC_Y
    ScreenX = (MAX_MAPX + 1) * PIC_X
    ScreenY = (MAX_MAPY + 1) * PIC_Y
    StartXValue = ((MAX_MAPX + 1) / 2)
    StartYValue = (MAX_MAPY / 2)
    EndXValue = (MAX_MAPX + 1) + 1
    EndYValue = MAX_MAPY + 1
    Half_PIC_X = PIC_X / 2
    Half_PIC_Y = PIC_Y / 2
    
   If Trim$(GetVar(App.path & "\data files\config.ini", "Options", "ServerListImage")) <> "" Then
        'frmServers.Picture = LoadPicture(App.path & "\data files\" & Trim$(GetVar(App.path & "\data files\config.ini", "Options", "ServerListImage")))
   End If
   
   If Trim$(GetVar(App.path & "\data files\config.ini", "Options", "LoadingFormImage")) <> "" Then
        frmLoad.picLoad.Picture = LoadPicture(App.path & "\data files\" & Trim$(GetVar(App.path & "\data files\config.ini", "Options", "LoadingFormImage")))
   End If
   
   
   ChkDir App.path & "\", "data files"
   ChkDir App.path & "\data files\" & ServerDir & "\", "logs"
   
    ' load options
    If ServerIndex = 0 Then
        Call SetStatus("Loading Options...")
        LoadOptions
        
        If Options.DefaultServer > 0 Then
            If ServerCount > 0 And ServerCount >= Options.DefaultServer Then
                ServerIndex = Options.DefaultServer
            End If
        End If
    End If
   
    
    If ServerIndex = 0 Then
        If Options.HideServerList = 1 Then
            End
        End If
        frmServers.Show
        Exit Sub
    End If
    
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    Call InitMessages
    
    frmLoad.Visible = True
    
    
    
    Call SetStatus("Attempting to connect to game server...")
    Do Until GotServerInfo = True
        If QuitConnecting Then Call DestroyTCP: Exit Sub
        If ConnectToServer(2) = True Then
            'Do not do anything.... data will arrive soon!
            If GotServerInfo Then Exit Do
        Else
            reconnectCount = reconnectCount + 1
            If reconnectCount >= 6 Then
                MsgBox "Could not connect to game server. The server maybe offline. Make sure the address and port is correct. Also be sure that you have internet access."
                frmLoad.Visible = False
                If Options.HideServerList = 1 Then
                    End
                End If
                frmServers.Visible = True
                frmMain.Socket.Close
                Options.DefaultServer = 0
                SaveOptions
                Exit Sub
            Else
                Call SetStatus("Failed to connect to server... Retrying, attempt " & reconnectCount & " / 5.")
            End If
        End If
        DoEvents
    Loop
    
    'Check for updates
    If Trim$(ServerDir) = "" Then
        ServerDir = "default"
        ServerDir = LCase(ServerDir)
    Else
        ServerDir = LCase(ServerDir)
    End If
    

    ' load main menu
    'Call SetStatus("Loading Menu...")
    'Load frmMenu

    ' Check if the directory is there, if its not make it
    ChkDir App.path & "\", "data files"
    ChkDir App.path & "\data files\", LCase(ServerDir)
    ChkDir App.path & "\data files\" & ServerDir & "\", "graphics"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "animations"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "characters"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "pictures"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "character creation"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\", "faces"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\faces\", "clothes"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\faces\", "ears"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\faces\", "etc"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\faces\", "eyebrows"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\faces\", "eyes"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\faces\", "hair"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\faces\", "heads"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\faces\", "mouths"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\faces\", "noses"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\", "characters"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\characters\", "bodies"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\characters\", "hair"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\characters\", "pants"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\characters\pants\", "male"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\characters\pants\", "female"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\characters\", "shirts"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\characters\", "shoes"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\characters\shoes\", "male"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\character creation\characters\shoes\", "female"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "items"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "paperdolls"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "resources"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "spellicons"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "tilesets"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "faces"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "gui"
    ChkDir App.path & "\data files\" & ServerDir & "\graphics\", "projectiles"
    ChkDir App.path & "\data files\" & ServerDir & "\", "maps"
    ChkDir App.path & "\data files\" & ServerDir & "\", "music"
    ChkDir App.path & "\data files\" & ServerDir & "\", "sound"
    
    SOUND_PATH = "\Data Files\" & ServerDir & "\sound\"
    MUSIC_PATH = "\Data Files\" & ServerDir & "\music\"
    MAP_PATH = "\Data Files\" & ServerDir & "\maps\"
    GFX_PATH = "\Data Files\" & ServerDir & "\graphics\"
    FONT_PATH = "\Data Files\" & ServerDir & "\graphics\fonts\"
    
    'Check for updates
    If Trim$(ServerDir) = "" Then
        ServerDir = "default"
        ServerDir = LCase(ServerDir)
    Else
        'Check for updates now!
        Call SetStatus("Checking for content updates...")
        If DownloadFile(Trim$(UpdateURL), App.path & "\data files\" & ServerDir & "\temp.ini") = True Then
            curVersion = Val(GetVar(App.path & "\data files\" & ServerDir & "\temp.ini", "Updates", "CurrentVersion"))
            myVersion = Val(GetVar(App.path & "\data files\" & ServerDir & "\updates.ini", "Updates", "CurrentVersion"))
            If myVersion > curVersion Then
                myVersion = 0
            End If
            If curVersion > 0 Then
                If myVersion < curVersion Then
                    For i = myVersion + 1 To curVersion
                        Call SetStatus("Downloading update " & i & "/" & curVersion & ".")
                        If DownloadFile(GetVar(App.path & "\data files\" & ServerDir & "\temp.ini", "Update" & i, "UpdateURL"), App.path & "\data files\" & ServerDir & "\tempupdate.rar", "Downloading update " & i & "/" & curVersion, True) = True Then
                            If RARExecute(OP_EXTRACT, App.path & "\data files\" & ServerDir & "\tempupdate.rar", "", "Data Files\" & ServerDir & "\") = True Then
                                Call SetStatus("Extracting update " & i & "/" & curVersion & ".")
                                Kill App.path & "\data files\" & ServerDir & "\tempupdate.rar"
                                PutVar App.path & "\data files\" & ServerDir & "\updates.ini", "Updates", "CurrentVersion", str(i)
                            Else
                                MsgBox "There was an error trying to update! The update maybe offline or the server maybe configured incorrectly!"
                                frmLoad.Hide
                                If Options.HideServerList = 1 Then
                                    End
                                End If
                                frmServers.Show
                                QuitConnecting = True
                                Exit Sub
                            End If
                        Else
                            MsgBox "There was an error trying to update! The update maybe offline or the server maybe configured incorrectly!"
                            frmLoad.Hide
                            If Options.HideServerList = 1 Then
                                End
                            End If
                            frmServers.Show
                            QuitConnecting = True
                            Exit Sub
                        End If
                    Next
                End If
            Else
                MsgBox "There was an error trying to update! The update maybe offline or the server maybe configured incorrectly!"
                frmLoad.Hide
                If Options.HideServerList = 1 Then
                    End
                End If
                frmServers.Show
                QuitConnecting = True
                Exit Sub
            End If
        Else
            If Trim$(UpdateURL) = "" Then
            
            Else
                MsgBox "There was an error trying to update! The update maybe offline or the server maybe configured incorrectly!"
                frmLoad.Hide
                If Options.HideServerList = 1 Then
                    End
                End If
                frmServers.Show
                QuitConnecting = True
                Exit Sub
            End If
        End If
    End If
    
    LoadGUI True
    
    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34) ' "
    ' Update the form with the game's name before it's loaded
    UpdateDebugCaption
    ' load gui
    Call SetStatus("Loading graphics...")
    EngineInitFontSettings
    InitDX8
    ' randomize rnd's seed
    Randomize
    Call SetStatus("Initializing DirectX...")
    ' load music/sound engine
    InitFmod
    ' check if we have main-menu music
    'If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMusic Trim$(Options.MenuMusic)
    ' Reset values
    Ping = -1
    ' cache the buttons then reset & render them
        ' load gui
    Call SetStatus("Loading interface...")
        LoadGUI
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.Weapon
    ' hide the load form
    frmLoad.Visible = False
    frmMain.Visible = True
    InitGUI
    MenuLoop


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MenuState(ByVal state As Long)

   On Error GoTo errorhandler

    frmLoad.Visible = True

    Select Case state
        Case MENU_STATE_ADDCHAR
            'frmMain.Visible = False
            MenuStage = 0

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending character addition data...")
                            Call SendAddChar(TxtUsername, NewCharSex, newCharClass)
            End If
                Case MENU_STATE_NEWACCOUNT
            'frmMain.Visible = False
            MenuStage = 0

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(TxtUsername, txtPassword)
            End If

        Case MENU_STATE_LOGIN
            'frmMain.Visible = False
            MenuStage = 0

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(TxtUsername, txtPassword)
                Exit Sub
            End If
    End Select

    If frmLoad.Visible Then
        If Not IsConnected Then
            'frmMain.Visible = True
            MenuStage = 0
            frmLoad.Visible = False
            Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or contact an administrator. Thanks.", vbOKOnly, Trim$(Servers(ServerIndex).Game_Name))
        End If
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub logoutGame()
Dim buffer As clsBuffer, i As Long


   On Error GoTo errorhandler

    isLogging = True
    InGame = False
    InEvent = False
    HoldPlayer = False
    frmAdmin.Visible = False
    MailBoxMenu = 1
    MailToFrom = ""
    MailContent = ""
    MailItem = 0
    MailItemValue = 0
    Set buffer = New clsBuffer
    buffer.WriteLong CQuit
    SendData buffer.ToArray()
    Set buffer = Nothing
    Call DestroyTCP
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    ' destroy temp values
    DragInvSlotNum = 0
    InvX = 0
    InvY = 0
    EqX = 0
    EqY = 0
    SpellX = 0
    SpellY = 0
    LastItemDesc = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0
    PetSpellBuffer = 0
    PetSpellBufferTimer = 0
    InTrade = 0
    InBank = False
    InShop = 0
    EventChat = False
    CurrencyMenu = 0
    dialogueIndex = 0
    For i = 1 To ChatTextBufferSize
        ChatTextBuffer(i).Text = ""
        ChatTextBuffer(i).color = 0
    Next
    totalChatLines = 0
    ChatScroll = ChatLines
    UpdateChatArray
    For i = 1 To 10
        With Pictures(i)
            .pic = 0
        End With
    Next
    If Options.sound = 1 Then StopAllSounds: StopMusic
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "logoutGame", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub GameInit()

   On Error GoTo errorhandler

    EnteringGame = True
    EnteringGame = False
    ' hide gui
    InBank = False
    InShop = False
    InTrade = False
    ' Set font
    'Call SetFont(FONT_NAME, FONT_SIZE)
    frmMain.font = "Arial Bold"
    frmMain.FontSize = 10
    ' show the main form
    frmLoad.Visible = False
    ' get ping
    GetPing
    DrawPing
    ' set values for amdin panel
    frmAdmin.scrlAItem.max = MAX_ITEMS
    frmAdmin.scrlAItem.Value = 1
    'stop the song playing
    StopMusic




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DestroyGame()
    ' break out of GameLoop

   On Error GoTo errorhandler

    InGame = False
    Call DestroyTCP
    ServerIndex = 0
    'destroy objects in reverse order
    DestroyDX8
    'Call UnloadAllForms
    frmMain.Hide
    frmLoad.Hide
    frmAdmin.Hide
    If Options.HideServerList = 1 Then
        frmMain.Socket.Close
        End
    End If
    frmServers.Visible = True
    StopMusic
    frmMain.Socket.Close




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DestroyGame", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub UnloadAllForms()
Dim frm As Form



   On Error GoTo errorhandler

    For Each frm In VB.Forms
        Unload frm
    Next





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SetStatus(ByVal Caption As String)

   On Error GoTo errorhandler

    frmLoad.lblStatus.Caption = Caption
    DoEvents




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)


   On Error GoTo errorhandler

    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If

    Txt.SelStart = Len(Txt.Text) - 1




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long

   On Error GoTo errorhandler

    Rand = Int((High - Low + 1) * Rnd) + Low



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim GlobalX As Long
Dim GlobalY As Long


   On Error GoTo errorhandler

    GlobalX = PB.Left
    GlobalY = PB.Top

    If Button = 1 Then
        PB.Left = GlobalX + X - SOffsetX
        PB.Top = GlobalY + Y - SOffsetY
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MovePicture", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean

   On Error GoTo errorhandler

    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If




   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim i As Long

    ' Prevent high ascii chars

   On Error GoTo errorhandler

    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Trim$(Servers(ServerIndex).Game_Name))
            Exit Function
        End If

    Next

    isStringLegal = True



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub PopulateLists()
Dim strLoad As String, i As Long


   On Error GoTo errorhandler

    ReDim soundCache(0)
    ReDim musicCache(0)
    ' Cache music list
    strLoad = dir(App.path & MUSIC_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        If UBound(musicCache) = 0 Then ReDim musicCache(1)
        ReDim Preserve musicCache(i) As String
        musicCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop
    ' Cache sound list
    strLoad = dir(App.path & SOUND_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        If UBound(soundCache) = 0 Then ReDim soundCache(1)
        ReDim Preserve soundCache(i) As String
        soundCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
