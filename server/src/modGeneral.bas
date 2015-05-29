Attribute VB_Name = "modGeneral"
Option Explicit
' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long


Public Sub Main()

   On Error GoTo errorhandler

    DebugMode = True
    Call InitServer


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub InitServer()
    Dim i As Long
    Dim F As Long
    Dim time1 As Long
    Dim time2 As Long
   
   On Error GoTo errorhandler

    Call InitMessages
    time1 = GetTickCount

    ' Initialize the random-number generator
    Randomize ', seed
    
    SetLoadingProgress "Checking Directories", 1, 1
    ' Check if the directory is there, if its not make it
    ChkDir App.path & "\", "data"
    ChkDir App.path & "\Data\", "accounts"
    ChkDir App.path & "\Data\", "animations"
    ChkDir App.path & "\Data\", "items"
    ChkDir App.path & "\Data\", "logs"
    ChkDir App.path & "\Data\", "maps"
    ChkDir App.path & "\Data\", "npcs"
    ChkDir App.path & "\Data\", "resources"
    ChkDir App.path & "\Data\", "shops"
    ChkDir App.path & "\Data\", "quests"
    ChkDir App.path & "\Data\", "spells"
    ChkDir App.path & "\Data\", "pets"
    ChkDir App.path & "\Data\", "projectiles"

    ' set quote character
    vbQuote = ChrW$(34) ' "
 
    SetLoadingProgress "Loading Options", 2, 1
    
    ' load options, set if they dont exist
    If FileExist(App.path & "\options.ini", True) = False And FileExist(App.path & "\data\options.ini", True) = False Then
        Options.Game_Name = "Eclipse Origins Silver Edition"
        Options.Port = 7001
        Options.MOTD = "Welcome to Eclipse Origins."
        Options.Website = "http://www.eclipseorigins.com/smf"
        PutVar App.path & "\options.ini", "OPTIONS", "MapCount", "300"
        SaveOptions
    Else
        
    End If
    
    LoadOptions
    MAX_MAPS = Val(GetVar(App.path & "\options.ini", "OPTIONS", "MapCount"))
    
    If MAX_MAPS <= 0 Then
        MAX_MAPS = 300
        PutVar App.path & "\options.ini", "OPTIONS", "MapCount", "300"
    End If
    
        MAX_PLAYERS = 1500
        MAX_ITEMS = 1000
        MAX_NPCS = 1000
        MAX_ANIMATIONS = 500
        MAX_SHOPS = 500
        MAX_SPELLS = 1000
        MAX_RESOURCES = 500
        MAX_ZONES = 255
        MAX_PLAYER_CHARS = 10
        MAX_HOUSES = 100
        MAX_QUESTS = 250
        MAX_PETS = 1000
    
    ReDim Player(0 To MAX_PLAYERS)
    ReDim Preserve TempPlayer(0 To MAX_PLAYERS)
    ReDim Bank(1 To MAX_PLAYERS)
    ReDim Item(1 To MAX_ITEMS)
    ReDim Npc(1 To MAX_ITEMS)
    ReDim Shop(1 To MAX_SHOPS)
    ReDim Spell(1 To MAX_SPELLS)
    ReDim Resource(1 To MAX_RESOURCES)
    ReDim Animation(1 To MAX_ANIMATIONS)
    ReDim Switches(1 To MAX_SWITCHES)
    ReDim Variables(1 To MAX_VARIABLES)
    ReDim HouseConfig(1 To MAX_HOUSES)
    ReDim MapZones(1 To MAX_ZONES)
    ReDim ZoneNpc(1 To MAX_ZONES)
    ReDim Pet(1 To MAX_PETS)
    
    For i = 1 To MAX_PLAYERS
        ReDim Player(i).characters(1 To MAX_PLAYER_CHARS)
    Next
    
    ReDim MonkeyPlayer.characters(1 To MAX_PLAYER_CHARS)
    
    Call frmServer.UsersOnline_Start
    
    If FileExist(App.path & "\data\spawn.ini", True) Then
        START_MAP = Val(GetVar(App.path & "\data\spawn.ini", "Spawn", "Map"))
        START_X = Val(GetVar(App.path & "\data\spawn.ini", "Spawn", "X"))
        START_Y = Val(GetVar(App.path & "\data\spawn.ini", "Spawn", "Y"))
        If START_MAP > 0 And START_MAP < MAX_MAPS And START_X >= 0 And START_Y >= 0 Then
        Else
            Call PutVar(App.path & "\data\spawn.ini", "Spawn", "Map", "1")
            Call PutVar(App.path & "\data\spawn.ini", "Spawn", "X", "1")
            Call PutVar(App.path & "\data\spawn.ini", "Spawn", "Y", "1")
        End If
    Else
        Call PutVar(App.path & "\data\spawn.ini", "Spawn", "Map", "1")
        Call PutVar(App.path & "\data\spawn.ini", "Spawn", "X", "1")
        Call PutVar(App.path & "\data\spawn.ini", "Spawn", "Y", "1")
    End If
    

    
    ReDim Map(MIN_MAPS To MAX_MAPS)
    ReDim TempEventMap(MIN_MAPS To MAX_MAPS)
    ReDim MapCache(MIN_MAPS To MAX_MAPS)
    ReDim temptile(MIN_MAPS To MAX_MAPS)
    ReDim PlayersOnMap(MIN_MAPS To MAX_MAPS)
    ReDim ResourceCache(MIN_MAPS To MAX_MAPS)
    ReDim MapItem(MIN_MAPS To MAX_MAPS, MAX_MAP_ITEMS)
    ReDim MapNpc(MIN_MAPS To MAX_MAPS)
    ReDim ZoneNpc(MAX_ZONES)
    ReDim MapBlocks(MIN_MAPS To MAX_MAPS)
    ReDim MapProjectiles(MIN_MAPS To MAX_MAPS, 1 To MAX_PROJECTILES)
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")
    SetLoadingProgress "Initializing player array.", 3, 1
    
    'Let's show the server and what's being loaded, eh?
    frmServer.Show
    frmServer.fraConsole.Visible = True
    
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call ClearGameData

    Call LoadGameData
    
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Spawning global events...")
    Call SpawnAllMapGlobalEvents
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache
    Call SetStatus("Loading System Tray...")
    SetLoadingProgress "Loading System Tray.", 34, 1
    DoEvents
    Call LoadSystemTray
    SetLoadingProgress "Loading Accounts...", 35, 1
    LoadAccounts
    SetLoadingProgress "Loading Bans...", 36, 1
    LoadBans
    SetLoadingProgress "Finished loading.", 37, 1
    DoEvents

    ' Check if  master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If
   
   
    frmServer.Show
  
    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    time2 = GetTickCount
    Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & "ms.")
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop


   On Error GoTo 0
   Exit Sub
errorhandler:
    If InStr(1, Err.Description, "Address in use") > 0 Then
        MsgBox "Port " & Options.Port & " is already in use! Please make sure you do not already have an Eclipse Origins server or other application running that is using port " & Options.Port & "."
        End
    End If
    HandleError "InitServer", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub DestroyServer()
    Dim i As Long

   On Error GoTo errorhandler

    ServerOnline = False
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")
    
   
    Unload frmServer
    
    End


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DestroyServer", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SetStatus(ByVal Status As String)

   On Error GoTo errorhandler
    
   
    Call TextAdd(Status)
    DoEvents


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub ClearGameData()

   On Error GoTo errorhandler

    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTiles
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing Zones...")
    Call ClearZones
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map projectiles...")
    Call ClearMapProjectiles
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing zone npcs...")
    Call ClearZoneNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing Resources...")
    Call ClearResources
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing quests...")
    Call ClearQuests
    Call SetStatus("Clearing pets...")
    Call Clearpets
    Call SetStatus("Clearing projectiles...")
    Call ClearProjectiles

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearGameData", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub LoadGameData()

   On Error GoTo errorhandler

    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading zones...")
    Call LoadZones
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading Resources...")
    Call LoadResources
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    Call SetStatus("Loading switches...")
    Call LoadSwitches
    Call SetStatus("Loading variables...")
    Call LoadVariables
    Call SetStatus("Loading House Configurations...")
    Call LoadHouses
    Call SetStatus("Loading quests...")
    Call LoadQuests
    Call SetStatus("Loading pets...")
    Call Loadpets
    Call SetStatus("Loading projectiles...")
    Call LoadProjectiles

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadGameData", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub TextAdd(Msg As String)

   On Error GoTo errorhandler

    Msg = "[" & TimeValue(Now) & "] " & Msg
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If
    
    If Trim$(frmServer.txtText.Text) = "" Then
        frmServer.txtText.Text = Msg
    Else
        frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Msg
    End If
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean


   On Error GoTo errorhandler

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "isNameLegal", "modGeneral", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SetLoadingProgress(Caption As String, Stage As Long, StageProg As Double)
    Dim maxstages As Long, prog As Double, a() As String, b As String, c As String
   On Error GoTo SetLoadingProgress_Error

    maxstages = 38
    prog = ((Stage / maxstages) + (StageProg * (1 / maxstages))) * 100
    

   On Error GoTo 0
   Exit Sub

SetLoadingProgress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetLoadingProgress of Module modGeneral"
End Sub


