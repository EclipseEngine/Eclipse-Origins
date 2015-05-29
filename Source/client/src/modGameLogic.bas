Attribute VB_Name = "modGameLogic"
Option Explicit
Public TabDown1 As Boolean

Public Sub GameLoop()
Dim FrameTime As Long
Dim tick As Long
Dim TickFPS As Long
Dim fps As Long
Dim i As Long, X As Long
Dim WalkTimer As Long
Dim tmr25 As Long
Dim tmr100 As Long
Dim tmr10000 As Long
Dim tmr500, Fadetmr As Long
Dim chattmr As Long
Dim fogtmr As Long, bartmr As Long, Tab1 As Long, Tab2 As Long


    ' *** Start GameLoop ***

   On Error GoTo errorhandler
    ElapsedMTime = GetTickCount
    Do While InGame
        tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = tick                               ' Set the time second loop time to the first.

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < tick Then


                    ' check ping
            Call GetPing
            Call DrawPing
            tmr10000 = tick + 10000
        End If

        If tmr25 < tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hwnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
                    ' check if we need to end the CD icon
            If NumSpellIcons > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i) > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (spell(PlayerSpells(i)).CDTime * 1000) < tick Then
                                SpellCD(i) = 0
                            End If
                        End If
                    End If
                Next
            End If
            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (spell(PlayerSpells(SpellBuffer)).CastTime * 1000) < tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If
            ' check if we need to unlock the pets's spell casting restriction
            If PetSpellBuffer > 0 Then
                If PetSpellBufferTimer + (spell(Pet(Player(MyIndex).Pet.Num).spell(PetSpellBuffer)).CastTime * 1000) < tick Then
                    PetSpellBuffer = 0
                    PetSpellBufferTimer = 0
                End If
            End If
            
            If HoldPlayer = False Then
                If CanMoveNow Then
                    Call CheckMovement ' Check if player is trying to move
                    Call CheckAttack   ' Check to see if player is trying to attack
                End If
            End If

            ' Change map animation every 250 milliseconds
            If MapAnimTimer < tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = tick + 250
            End If
                    ' Update inv animation
            If numitems > 0 Then
                If tmr100 < tick Then
                    DrawAnimatedInvItems
                    tmr100 = tick + 250
                End If
            End If
                    For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next
                    tmr25 = tick + 25
        End If
            If tick > EventChatTimer Then
            If EventText = "" Then
                If EventChat = True Then
                    EventChat = False
                End If
            End If
        End If
            If chattmr < tick Then
            If ChatUpBtnState = 2 Then
                ScrollChatBox 0
            End If
            If ChatDownBtnState = 2 Then
                ScrollChatBox 1
            End If
            chattmr = tick + 50
        End If
    
        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < tick Then
            ElapsedMTime = tick - ElapsedMTime
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                    If Player(i).Pet.Alive = True Then
                        ProcessPetMovement (i)
                    End If
                End If
            Next i

            ' Process npc movements (actually move them)
            For i = 1 To Npc_HighIndex
                If Map.Npc(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If
            Next i
                    ' Process zone npc movements (actually move them)
            For i = 1 To MAX_ZONES
                For X = 1 To MAX_MAP_NPCS * 2
                    If ZoneNPC(i).Npc(X).Num > 0 Then
                        If ZoneNPC(i).Npc(X).Map = GetPlayerMap(MyIndex) Then
                            Call ProcessZoneNpcMovement(i, X)
                        End If
                    End If
                Next
            Next i
            
            If Map.CurrentEvents > 0 Then
                For i = 1 To Map.CurrentEvents
                    Call ProcessEventMovement(i)
                Next i
            End If
            ElapsedMTime = tick
            WalkTimer = tick + 30 ' edit this value to change WalkTimer
        End If
            ' fog scrolling
        If fogtmr < tick Then
            If CurrentFogSpeed > 0 Then
                ' move
                fogOffsetX = fogOffsetX - 1
                fogOffsetY = fogOffsetY - 1
                ' reset
                If fogOffsetX < -256 Then fogOffsetX = 0
                If fogOffsetY < -256 Then fogOffsetY = 0
                fogtmr = tick + 255 - CurrentFogSpeed
            End If
        End If
            ' elastic bars
        If bartmr < tick Then
            If hideGUI = False Then
                SetBarWidth BarWidth_GuiHP_Max, BarWidth_GuiHP
                SetBarWidth BarWidth_GuiSP_Max, BarWidth_GuiSP
                SetBarWidth BarWidth_GuiEXP_Max, BarWidth_GuiEXP
            End If
                    ' reset timer
            bartmr = tick + 10
        End If
            If tmr500 < tick Then
            ' animate waterfalls
            Select Case waterfallFrame
                Case 0
                    waterfallFrame = 1
                Case 1
                    waterfallFrame = 2
                Case 2
                    waterfallFrame = 0
            End Select
                    ' animate autotiles
            Select Case autoTileFrame
                Case 0
                    autoTileFrame = 1
                Case 1
                    autoTileFrame = 2
                Case 2
                    autoTileFrame = 0
            End Select
                    ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = " "
            Else
                chatShowLine = "|"
            End If
                    tmr500 = tick + 500
        End If
            ProcessWeather
            If Fadetmr < tick Then
            If FadeType <> 2 Then
                If FadeType = 1 Then
                    If FadeAmount = 255 Then
                                        Else
                        FadeAmount = FadeAmount + 5
                    End If
                ElseIf FadeType = 0 Then
                    If FadeAmount = 0 Then
                                    Else
                        FadeAmount = FadeAmount - 5
                    End If
                End If
            End If
            Fadetmr = tick + 30
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics
        Call UpdateSounds
        DoEvents
        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < tick + 30
                DoEvents
                Sleep 1
            Loop
        End If
        ' Calculate fps
        If TickFPS < tick Then
            GameFPS = fps
            TickFPS = tick + 1000
            fps = 0
        Else
            fps = fps + 1
        End If

    Loop

    If isLogging Then
        isLogging = False
        'frmMain.Visible = False
        GettingMap = True
        StopMusic
    Else
        ' Shutdown the game
        frmLoad.Visible = True
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

    MenuLoop



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MenuLoop()
Dim FrameTime As Long
Dim tick As Long
Dim TickFPS As Long
Dim fps As Long
Dim i As Long, X As Long
Dim WalkTimer As Long
Dim tmr25 As Long
Dim tmr100 As Long
Dim tmr10000 As Long
Dim tmr500, Fadetmr As Long
Dim fogtmr As Long
Dim Tab1 As Long, Tab2 As Long

    'Menu Init Stuff

   On Error GoTo errorhandler

  
    'Clear Menu Variables
    MenuStage = 0
    TxtUsername = ""
    txtPassword = ""
    TxtPassword2 = ""
    SelectedChar = 1
    UpdateDebugCaption
    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMusic Trim$(Options.MenuMusic)
    'End Menu Init

    ' *** Start GameLoop ***
    Do While Not InGame And ServerIndex > 0
        tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = tick
            If tmr500 < tick Then
                    ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = " "
            Else
                chatShowLine = "|"
            End If
            tmr500 = tick + 500
        End If
        If Fadetmr < tick Then
            If FadeType <> 2 Then
                If FadeType = 1 Then
                    If FadeAmount = 255 Then
                                        Else
                        FadeAmount = FadeAmount + 5
                    End If
                ElseIf FadeType = 0 Then
                    If FadeAmount = 0 Then
                                    Else
                        FadeAmount = FadeAmount - 5
                    End If
                End If
            End If
            Fadetmr = tick + 40
        End If
            'Slideshow Setup
        If InIntro = 1 Then
            If IntroTick = 0 Then
                If FadeType = 0 And FadeAmount = 0 Then
                    IntroTick = GetTickCount + IntroTimer
                    If IntroFade = 0 Then
                        IntroStep = IntroStep + 1
                        If IntroStep > UBound(IntroImages) Then
                            InIntro = 0
                            FadeType = 0
                        Else
                            FadeType = 0
                        End If
                    End If
                End If
                If IntroFade = 1 Then
                    If FadeType = 1 And FadeAmount = 255 Then
                        IntroStep = IntroStep + 1
                        If IntroStep >= UBound(IntroImages) Then
                            InIntro = 0
                            FadeType = 0
                        Else
                            FadeType = 0
                        End If
                    End If
                End If
            Else
                If IntroFade = 1 Then
                    If IntroTick < GetTickCount Then
                        FadeType = 1
                        IntroTick = 0
                    End If
                Else
                    If IntroTick < GetTickCount Then
                        IntroTick = 0
                    End If
                End If
            End If
        End If


        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Menu
        DoEvents

        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < tick + 30
                DoEvents
                Sleep 1
            Loop
        End If
            ' Calculate fps
        If TickFPS < tick Then
            GameFPS = fps
            TickFPS = tick + 1000
            fps = 0
        Else
            fps = fps + 1
        End If

    Loop
    ' Shutdown the game
    frmLoad.Visible = True
    Call SetStatus("Destroying game data...")
    Call DestroyGame





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MenuLoop", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long
Static eTime As Long

    ' Check if player is walking, and if so process moving them over
   On Error GoTo errorhandler

    Select Case Player(Index).Moving
        Case MOVING_WALKING: MovementSpeed = ((ElapsedMTime / 1000) * (RUN_SPEED * SIZE_X))
        Case MOVING_RUNNING: MovementSpeed = ((ElapsedMTime / 1000) * (WALK_SPEED * SIZE_X))
        Case Else: Exit Sub
    End Select
    If Player(Index).Step = 0 Then Player(Index).Step = 1
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).YOffset = Player(Index).YOffset - MovementSpeed
            If Player(Index).YOffset < 0 Then Player(Index).YOffset = 0
        Case DIR_DOWN
            Player(Index).YOffset = Player(Index).YOffset + MovementSpeed
            If Player(Index).YOffset > 0 Then Player(Index).YOffset = 0
        Case DIR_LEFT
            Player(Index).XOffset = Player(Index).XOffset - MovementSpeed
            If Player(Index).XOffset < 0 Then Player(Index).XOffset = 0
        Case DIR_RIGHT
            Player(Index).XOffset = Player(Index).XOffset + MovementSpeed
            If Player(Index).XOffset > 0 Then Player(Index).XOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (Player(Index).XOffset >= 0) And (Player(Index).YOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        Else
            If (Player(Index).XOffset <= 0) And (Player(Index).YOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProcessMovement", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)


    ' Check if NPC is walking, and if so process moving them over

   On Error GoTo errorhandler

    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
            Select Case MapNpc(MapNpcNum).dir
            Case DIR_UP
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - ((ElapsedMTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).YOffset < 0 Then MapNpc(MapNpcNum).YOffset = 0
                        Case DIR_DOWN
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + ((ElapsedMTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).YOffset > 0 Then MapNpc(MapNpcNum).YOffset = 0
                        Case DIR_LEFT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - ((ElapsedMTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).XOffset < 0 Then MapNpc(MapNpcNum).XOffset = 0
                        Case DIR_RIGHT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + ((ElapsedMTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).XOffset > 0 Then MapNpc(MapNpcNum).XOffset = 0
                    End Select
        ' Check if completed walking over to the next tile
        If MapNpc(MapNpcNum).Moving > 0 Then
            If MapNpc(MapNpcNum).dir = DIR_RIGHT Or MapNpc(MapNpcNum).dir = DIR_DOWN Then
                If (MapNpc(MapNpcNum).XOffset >= 0) And (MapNpc(MapNpcNum).YOffset >= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 1 Then
                        MapNpc(MapNpcNum).Step = 3
                    Else
                        MapNpc(MapNpcNum).Step = 1
                    End If
                End If
            Else
                If (MapNpc(MapNpcNum).XOffset <= 0) And (MapNpc(MapNpcNum).YOffset <= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 1 Then
                        MapNpc(MapNpcNum).Step = 3
                    Else
                        MapNpc(MapNpcNum).Step = 1
                    End If
                End If
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ProcessZoneNpcMovement(ByVal zonenum As Long, zoneNpcNum As Long)


    ' Check if NPC is walking, and if so process moving them over

   On Error GoTo errorhandler

    If ZoneNPC(zonenum).Npc(zoneNpcNum).Moving = MOVING_WALKING Then
            Select Case ZoneNPC(zonenum).Npc(zoneNpcNum).dir
            Case DIR_UP
                ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset = ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset - ((ElapsedMTime / 1000) * (WALK_SPEED * SIZE_X))
                If ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset < 0 Then ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset = 0
                        Case DIR_DOWN
                ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset = ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset + ((ElapsedMTime / 1000) * (WALK_SPEED * SIZE_X))
                If ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset > 0 Then ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset = 0
                        Case DIR_LEFT
                ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset = ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset - ((ElapsedMTime / 1000) * (WALK_SPEED * SIZE_X))
                If ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset < 0 Then ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset = 0
                        Case DIR_RIGHT
                ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset = ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset + ((ElapsedMTime / 1000) * (WALK_SPEED * SIZE_X))
                If ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset > 0 Then ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset = 0
                    End Select
        ' Check if completed walking over to the next tile
        If ZoneNPC(zonenum).Npc(zoneNpcNum).Moving > 0 Then
            If ZoneNPC(zonenum).Npc(zoneNpcNum).dir = DIR_RIGHT Or ZoneNPC(zonenum).Npc(zoneNpcNum).dir = DIR_DOWN Then
                If (ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset >= 0) And (ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset >= 0) Then
                    ZoneNPC(zonenum).Npc(zoneNpcNum).Moving = 0
                    If ZoneNPC(zonenum).Npc(zoneNpcNum).Step = 1 Then
                        ZoneNPC(zonenum).Npc(zoneNpcNum).Step = 3
                    Else
                        ZoneNPC(zonenum).Npc(zoneNpcNum).Step = 1
                    End If
                End If
            Else
                If (ZoneNPC(zonenum).Npc(zoneNpcNum).XOffset <= 0) And (ZoneNPC(zonenum).Npc(zoneNpcNum).YOffset <= 0) Then
                    ZoneNPC(zonenum).Npc(zoneNpcNum).Moving = 0
                    If ZoneNPC(zonenum).Npc(zoneNpcNum).Step = 1 Then
                        ZoneNPC(zonenum).Npc(zoneNpcNum).Step = 3
                    Else
                        ZoneNPC(zonenum).Npc(zoneNpcNum).Step = 1
                    End If
                End If
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProcessZoneNpcMovement", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub CheckMapGetItem()
Dim buffer As New clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer

    If GetTickCount > Player(MyIndex).MapGetTimer + 500 Then
        If Trim$(MyText) = vbNullString Then
            Player(MyIndex).MapGetTimer = GetTickCount
            buffer.WriteLong CMapGetItem
            SendData buffer.ToArray()
        End If
    End If

    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckAttack()
Dim buffer As clsBuffer
Dim attackspeed As Long, X As Long, Y As Long, i As Long


   On Error GoTo errorhandler

    If SpaceDown Then
        If InEvent = True Then Exit Sub
        If InMailbox = True Then Exit Sub
        If frmEditor_Map.Visible Then Exit Sub
        If chatOn = True Then Exit Sub
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = GetTickCount
                End With

                Set buffer = New clsBuffer
                buffer.WriteLong CAttack
                SendData buffer.ToArray()
                Set buffer = Nothing
            End If
        End If
            Select Case Player(MyIndex).dir
            Case DIR_UP
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) - 1
            Case DIR_DOWN
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) + 1
            Case DIR_LEFT
                X = GetPlayerX(MyIndex) - 1
                Y = GetPlayerY(MyIndex)
            Case DIR_RIGHT
                X = GetPlayerX(MyIndex) + 1
                Y = GetPlayerY(MyIndex)
        End Select
            If GetTickCount > Player(MyIndex).EventTimer Then
            For i = 1 To Map.CurrentEvents
                If Map.MapEvents(i).Visible = 1 Then
                    If Map.MapEvents(i).X = X And Map.MapEvents(i).Y = Y Then
                        Set buffer = New clsBuffer
                        buffer.WriteLong CEvent
                        buffer.WriteLong i
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        Player(MyIndex).EventTimer = GetTickCount + 200
                    End If
                End If
            Next
        End If

    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function IsOneBlockAway(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Boolean

   On Error GoTo errorhandler

    If X1 = X2 Then
        If Y1 = Y2 - 1 Or Y1 = Y2 + 1 Then
            IsOneBlockAway = True
        Else
            IsOneBlockAway = False
        End If
    ElseIf Y1 = Y2 Then
        If X1 = X2 - 1 Or X1 = X2 + 1 Then
            IsOneBlockAway = True
        Else
            IsOneBlockAway = False
        End If
    Else
        IsOneBlockAway = False
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsOneBlockAway", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function IsTryingToMove() As Boolean
Dim dir As Long

   On Error GoTo errorhandler

    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
        WalkToX = -1
        WalkToY = -1
        Exit Function
    End If
    If Options.ClicktoWalk = 1 Then
        If WalkToX <> -1 Or WalkToY <> -1 Then
            If WalkToX <> GetPlayerX(MyIndex) Or WalkToY <> GetPlayerY(MyIndex) Then
                If IsOneBlockAway(GetPlayerX(MyIndex), GetPlayerY(MyIndex), WalkToX, WalkToY) = True Then
                    If GetPlayerX(MyIndex) > WalkToX Then
                        DirLeft = True
                    ElseIf GetPlayerX(MyIndex) < WalkToX Then
                        DirRight = True
                    ElseIf GetPlayerY(MyIndex) > WalkToY Then
                        DirUp = True
                    ElseIf GetPlayerY(MyIndex) < WalkToY Then
                        DirDown = True
                    End If
                Else
                    dir = GetPlayerNextStep()
                    If dir < 5 Then
                        Select Case dir
                            Case DIR_UP
                                DirUp = True
                            Case DIR_DOWN
                                DirDown = True
                            Case DIR_RIGHT
                                DirRight = True
                            Case DIR_LEFT
                                DirLeft = True
                        End Select
                    End If
                End If
            End If
        End If
    End If
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If




   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function CanMove() As Boolean
Dim d As Long

   On Error GoTo errorhandler

    CanMove = True
    
    If frmEditor_Events.Visible Then CanMove = False: Exit Function

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    If HomeUp = False And HomeDown = True Then
        CanMove = False
        Exit Function
    End If
    If InEvent Then
        CanMove = False
        Exit Function
    End If
    If InTrade Then
        CanMove = False
        Exit Function
    End If
    If Options.GfxMode = 0 Then
        If CacheMap = False Then
            CanMove = False
            Exit Function
        End If
    End If
            

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If
    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    ' not in bank
    If InBank Then
        If IsTryingToMove Then
            CloseBank
        End If
    End If
    If InShop > 0 Then
        If IsTryingToMove Then
            CloseShop
        End If
    End If
    If InMailbox Then
        CanMove = False
        Exit Function
    End If
    If CurrencyMenu > 0 Then CanMove = False: Exit Function
    If dialogueIndex > 0 Then CanMove = False: Exit Function
    If EventChat Then CanMove = False:   Exit Function

    d = GetPlayerDir(MyIndex)

    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Up > 0 And Map.Up <= MAX_MAPS Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Down > 0 And Map.Down <= MAX_MAPS Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Left > 0 And Map.Left <= MAX_MAPS Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Right > 0 And Map.Right <= MAX_MAPS Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If




   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function CheckDirection(ByVal direction As Byte) As Boolean
Dim X As Long
Dim Y As Long
Dim i As Long, z As Long, buffer As clsBuffer


   On Error GoTo errorhandler

    CheckDirection = False

    Select Case direction
        Case DIR_UP
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
    End Select

    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map.Tile(X, Y).type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the key door is open or not
    If Map.Tile(X, Y).type = TILE_TYPE_KEY Then

        ' This actually checks if its open or not
        If TempTile(X, Y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If
    If Map.Moral <> MAP_MORAL_SAFE Then
        ' Check to see if a player is already on that tile
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If Player(i).InHouse = Player(MyIndex).InHouse Then
                    If GetPlayerX(i) = X Then
                        If GetPlayerY(i) = Y Then
                            CheckDirection = True
                            Exit Function
                        ElseIf Player(i).Pet.X = X And Player(i).Pet.Alive = True Then
                            If Player(i).Pet.Y = Y Then
                                CheckDirection = True
                                Exit Function
                            End If
                        End If
                    ElseIf Player(i).Pet.X = X And Player(i).Pet.Alive = True Then
                        If Player(i).Pet.Y = Y Then
                            CheckDirection = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next i
    End If
    If FurnitureHouse > 0 Then
        If FurnitureHouse = Player(MyIndex).InHouse Then
            If FurnitureCount > 0 Then
                For i = 1 To FurnitureCount
                    If Item(Furniture(i).ItemNum).Data3 = 0 Then
                        If X >= Furniture(i).X And X <= Furniture(i).X + Item(Furniture(i).ItemNum).FurnitureWidth - 1 Then
                            If Y <= Furniture(i).Y And Y >= Furniture(i).Y - Item(Furniture(i).ItemNum).FurnitureHeight Then
                                z = Item(Furniture(i).ItemNum).FurnitureBlocks(X - Furniture(i).X, ((Furniture(i).Y - Y) * -1) + Item(Furniture(i).ItemNum).FurnitureHeight)
                                If z = 1 Then CheckDirection = True: Exit Function
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End If
    ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
    For i = 1 To MAX_ZONES
        For z = 1 To MAX_MAP_NPCS * 2
            If ZoneNPC(i).Npc(z).Num > 0 Then
                If ZoneNPC(i).Npc(z).X = X Then
                    If ZoneNPC(i).Npc(z).Y = Y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next
    Next
    
    For i = 1 To Map.CurrentEvents
        If Map.MapEvents(i).Visible = 1 Then
            If Map.MapEvents(i).X = X Then
                If Map.MapEvents(i).Y = Y Then
                    'We are walking on top of OR tried to touch an event. Time to Handle the commands
                    Set buffer = New clsBuffer
                    buffer.WriteLong CEventTouch
                    buffer.WriteLong i
                    SendData buffer.ToArray
                    Set buffer = Nothing
                    If Map.MapEvents(i).WalkThrough = 0 Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next




   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CheckDirection", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub CheckMovement()

   On Error GoTo errorhandler
    If GettingMap = True Then Exit Sub
    If IsTryingToMove Then
        If CanMove Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = MOVING_RUNNING
            Else
                Player(MyIndex).Moving = MOVING_WALKING
            End If
                
            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select
                    If WalkToX = GetPlayerX(MyIndex) And WalkToY = GetPlayerY(MyIndex) Then
                WalkToX = -1
                WalkToY = -1
            End If

            If Player(MyIndex).XOffset = 0 Then
                If Player(MyIndex).YOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
        End If
    Else
        If CanMove Then
            If HomeDown Then
                Select Case GetPlayerDir(MyIndex)
                    Case DIR_UP
                        Player(MyIndex).dir = DIR_RIGHT
                    Case DIR_RIGHT
                        Player(MyIndex).dir = DIR_DOWN
                    Case DIR_DOWN
                        Player(MyIndex).dir = DIR_LEFT
                    Case DIR_LEFT
                        Player(MyIndex).dir = DIR_UP
                End Select
                SendPlayerMove
                HomeUp = False
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function isInBounds()

   On Error GoTo errorhandler

    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If




   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub UpdateDrawMapName()

   On Error GoTo errorhandler

    DrawMapNameX = ((MAX_MAPX + 1) * PIC_X / 2) - (getWidth(Font_Default, Trim$(Map.Name)) / 2)
    DrawMapNameY = 1

    Select Case Map.Moral
        Case MAP_MORAL_NONE
            DrawMapNameColor = BrightRed
        Case MAP_MORAL_SAFE
            DrawMapNameColor = White
        Case Else
            DrawMapNameColor = White
    End Select





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateDrawMapName", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub UseItem()
    ' Check for subscript out of range

   On Error GoTo errorhandler

    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ForgetSpell(ByVal SpellSlot As Long)
Dim buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If SpellSlot < 1 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    ' dont let them forget a spell which is in CD
    If SpellCD(SpellSlot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    ' dont let them forget a spell which is buffered
    If SpellBuffer = SpellSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    If PlayerSpells(SpellSlot) > 0 Then
        Set buffer = New clsBuffer
        buffer.WriteLong CForgetSpell
        buffer.WriteLong SpellSlot
        SendData buffer.ToArray()
        Set buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CastSpell(ByVal SpellSlot As Long)
Dim buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If SpellSlot < 1 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    If SpellCD(SpellSlot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    If PlayerSpells(SpellSlot) = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < spell(PlayerSpells(SpellSlot)).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(spell(PlayerSpells(SpellSlot)).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(SpellSlot) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CCast
                buffer.WriteLong SpellSlot
                SendData buffer.ToArray()
                Set buffer = Nothing
                SpellBuffer = SpellSlot
                SpellBufferTimer = GetTickCount
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CastSpell", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearTempTile()
Dim X As Long
Dim Y As Long


   On Error GoTo errorhandler

    ReDim TempTile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            TempTile(X, Y).DoorOpen = NO
        Next
    Next





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearTempTile", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DevMsg(ByVal Text As String, ByVal color As Byte)

   On Error GoTo errorhandler

    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(Text, color)
        End If
    End If

    Debug.Print Text




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DevMsg", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Function PixelsToTwips(pixels As Integer)

   On Error GoTo errorhandler

PixelsToTwips = pixels * Screen.TwipsPerPixelX


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "PixelsToTwips", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String

   On Error GoTo errorhandler

    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub DrawPing()
Dim PingToDraw As String


   On Error GoTo errorhandler

    PingToDraw = Ping

    Select Case Ping
        Case -1
            PingToDraw = "Syncing"
        Case 0 To 5
            PingToDraw = "Local"
    End Select

    'frmMain.lblPing.Caption = PingToDraw




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawPing", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub UpdateSpellWindow(ByVal Spellnum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long
    ' don't show desc when dragging

   On Error GoTo errorhandler

    If DragSpell > 0 Then Exit Sub
    ' get out
    If Spellnum = 0 Then Exit Sub
    If LastSpellDesc = Spellnum Then Exit Sub
    For i = 1 To 8
        SpellDescLbls(i) = ""
    Next
    ' item type
    i = 1
    Select Case spell(Spellnum).type
        Case SPELL_TYPE_DAMAGEHP
            SpellDescLbls(i) = "Damage HP"
        Case SPELL_TYPE_DAMAGEMP
            SpellDescLbls(i) = "Damage SP"
        Case SPELL_TYPE_HEALHP
            SpellDescLbls(i) = "Heal HP"
        Case SPELL_TYPE_HEALMP
            SpellDescLbls(i) = "Heal SP"
        Case SPELL_TYPE_WARP
            SpellDescLbls(i) = "Warp"
        Case SPELL_TYPE_PET
            SpellDescLbls(i) = "Summon"
    End Select
    ' more info
    Select Case spell(Spellnum).type
        Case SPELL_TYPE_DAMAGEHP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
            ' damage
            i = i + 1
                    SpellDescLbls(i) = "Vital: " & spell(Spellnum).Vital
                    ' mp cost
            i = i + 1
                    SpellDescLbls(i) = "Cost: " & spell(Spellnum).MPCost & " SP"
                    ' cast time
            i = i + 1
                    SpellDescLbls(i) = "Cast Time: " & spell(Spellnum).CastTime & "s"
                    ' cd time
            i = i + 1
                    SpellDescLbls(i) = "Cooldown: " & spell(Spellnum).CDTime & "s"
                    ' aoe
            If spell(Spellnum).AoE > 0 Then
                i = i + 1
                            SpellDescLbls(i) = "AoE: " & spell(Spellnum).AoE
            End If
                    ' stun
            If spell(Spellnum).StunDuration > 0 Then
                i = i + 1
                            SpellDescLbls(i) = "Stun: " & spell(Spellnum).StunDuration & "s"
            End If
                    ' dot
            If spell(Spellnum).Duration > 0 And spell(Spellnum).Interval > 0 Then
                i = i + 1
                            SpellDescLbls(i) = "DoT: " & (spell(Spellnum).Duration / spell(Spellnum).Interval) & " tick"
            End If
    End Select
    SpellDescVisible = True
    LastSpellDesc = Spellnum




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateSpellWindow", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub UpdateDescWindow(ByVal ItemNum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long
Dim FirstLetter As String * 1
Dim Name As String
Dim colour As Long
     ' work out name colour

   On Error GoTo errorhandler

    Select Case Item(ItemNum).Rarity
        Case 0 ' white
            colour = FontColor
        Case 1 ' green
            colour = Yellow
        Case 2 ' blue
            colour = Green
        Case 3 ' maroon
            colour = Blue
        Case 4 ' purple
            colour = Red
        Case 5 ' orange
            colour = Purple
    End Select
    ItemNameColor = colour
    ' class req
    If Item(ItemNum).classReq > 0 Then
        ItemClassReq = Trim$(Class(Item(ItemNum).classReq).Name)
        ' do we match it?
        If GetPlayerClass(MyIndex) = Item(ItemNum).classReq Then
            colour = Green
        Else
            colour = BrightRed
        End If
    Else
        ItemClassReq = "No class req."
        colour = Green
    End If
    ItemClassColor = colour
    ' level
    If Item(ItemNum).LevelReq > 0 Then
        ItemLevelReq = "Level " & Item(ItemNum).LevelReq
        ' do we match it?
        If GetPlayerLevel(MyIndex) >= Item(ItemNum).LevelReq Then
            colour = Green
        Else
            colour = BrightRed
        End If
    Else
        ItemLevelReq = "No level req."
        colour = Green
    End If
    ItemLevelColor = colour
    ' first we cache all information strings then loop through and render them
    For i = 1 To 8
        ItemDescLbls(i) = ""
    Next

    ' item type
    i = 1
    Select Case Item(ItemNum).type
        Case ITEM_TYPE_NONE
            ItemDescLbls(i) = "No type"
        Case ITEM_TYPE_WEAPON
            ItemDescLbls(i) = "Weapon"
        Case ITEM_TYPE_ARMOR
            ItemDescLbls(i) = "Armour"
        Case ITEM_TYPE_HELMET
            ItemDescLbls(i) = "Helmet"
        Case ITEM_TYPE_SHIELD
            ItemDescLbls(i) = "Shield"
        Case ITEM_TYPE_CONSUME
            ItemDescLbls(i) = "Consume"
        Case ITEM_TYPE_KEY
            ItemDescLbls(i) = "Key"
        Case ITEM_TYPE_CURRENCY
            ItemDescLbls(i) = "Currency"
        Case ITEM_TYPE_SPELL
            ItemDescLbls(i) = "Spell"
        Case ITEM_TYPE_FURNITURE
            ItemDescLbls(i) = "Furniture"
    End Select
    ' more info
    Select Case Item(ItemNum).type
        Case ITEM_TYPE_NONE, ITEM_TYPE_KEY, ITEM_TYPE_CURRENCY
            ' binding
            If Item(ItemNum).BindType = 1 Then
                i = i + 1
                ItemDescLbls(i) = "Bind on Pickup"
            ElseIf Item(ItemNum).BindType = 2 Then
                i = i + 1
                ItemDescLbls(i) = "Bind on Equip"
            End If
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SHIELD
            ' damage/defence
            If Item(ItemNum).type = ITEM_TYPE_WEAPON Then
                i = i + 1
                If Item(ItemNum).speed > 0 Then
                    ItemDescLbls(i) = "DPS: " & Int((Item(ItemNum).data2 * (1000 / Item(ItemNum).speed)))
                End If
            Else
                If Item(ItemNum).data2 > 0 Then
                    i = i + 1
                    ItemDescLbls(i) = "Defence: " & Item(ItemNum).data2
                End If
            End If
            ' binding
            If Item(ItemNum).BindType = 1 Then
                i = i + 1
                ItemDescLbls(i) = "Bind on Pickup"
            ElseIf Item(ItemNum).BindType = 2 Then
                i = i + 1
                ItemDescLbls(i) = "Bind on Equip"
            End If
            ' stat bonuses
            If Item(ItemNum).Add_Stat(Stats.Strength) > 0 Then
                i = i + 1
                ItemDescLbls(i) = "+" & Item(ItemNum).Add_Stat(Stats.Strength) & " Str"
            End If
            If Item(ItemNum).Add_Stat(Stats.Endurance) > 0 Then
                i = i + 1
                ItemDescLbls(i) = "+" & Item(ItemNum).Add_Stat(Stats.Endurance) & " End"
            End If
            If Item(ItemNum).Add_Stat(Stats.Intelligence) > 0 Then
                i = i + 1
                ItemDescLbls(i) = "+" & Item(ItemNum).Add_Stat(Stats.Intelligence) & " Int"
            End If
            If Item(ItemNum).Add_Stat(Stats.Agility) > 0 Then
                i = i + 1
                ItemDescLbls(i) = "+" & Item(ItemNum).Add_Stat(Stats.Agility) & " Agi"
            End If
            If Item(ItemNum).Add_Stat(Stats.Willpower) > 0 Then
                i = i + 1
                ItemDescLbls(i) = "+" & Item(ItemNum).Add_Stat(Stats.Willpower) & " Will"
            End If
        Case ITEM_TYPE_CONSUME
            If Item(ItemNum).CastSpell > 0 Then
                i = i + 1
                ItemDescLbls(i) = "Casts Spell"
            End If
            If Item(ItemNum).AddHP > 0 Then
                i = i + 1
                ItemDescLbls(i) = "+" & Item(ItemNum).AddHP & " HP"
            End If
            If Item(ItemNum).AddMP > 0 Then
                i = i + 1
                ItemDescLbls(i) = "+" & Item(ItemNum).AddMP & " SP"
            End If
            If Item(ItemNum).AddEXP > 0 Then
                i = i + 1
                ItemDescLbls(i) = "+" & Item(ItemNum).AddEXP & " EXP"
            End If
            ' price
            i = i + 1
            ItemDescLbls(i) = "Value: " & Item(ItemNum).Price & "g"
        Case ITEM_TYPE_SPELL
            ' price
            i = i + 1
            ItemDescLbls(i) = "Value: " & Item(ItemNum).Price & "g"
    End Select

    ItemDescVisible = True





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateDescWindow", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CacheResources()
Dim X As Long, Y As Long, Resource_Count As Long


   On Error GoTo errorhandler

    Resource_Count = 0

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If Map.Tile(X, Y).type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If
        Next
    Next

    Resource_Index = Resource_Count




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CreateActionMsg(ByVal Message As String, ByVal color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
Dim i As Long


   On Error GoTo errorhandler

    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .Message = Message
        .color = color
        .type = MsgType
        .Created = GetTickCount
        .Scroll = 1
        .X = X
        .Y = Y
    End With

    If ActionMsg(ActionMsgIndex).type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    End If
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim i As Long


   On Error GoTo errorhandler

    ActionMsg(Index).Message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).type = 0
    ActionMsg(Index).color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).X = 0
    ActionMsg(Index).Y = 0
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
Dim looptime As Long
Dim Layer As Long
Dim FrameCount As Long
Dim lockindex As Long

    ' if doesn't exist then exit sub

   On Error GoTo errorhandler

    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
                    ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(Layer) = 0 Then AnimInstance(Index).frameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
                    ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).Timer(Layer) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimInstance(Index).frameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).frameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).frameIndex(Layer) = AnimInstance(Index).frameIndex(Layer) + 1
                End If
                AnimInstance(Index).Timer(Layer) = GetTickCount
            End If
        End If
    Next
    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckAnimInstance", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub OpenShop(ByVal ShopNum As Long)

   On Error GoTo errorhandler

    InShop = ShopNum
    ShopAction = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long

   On Error GoTo errorhandler

    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    GetBankItemNum = Bank.Item(bankslot).Num



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal ItemNum As Long)

   On Error GoTo errorhandler

    Bank.Item(bankslot).Num = ItemNum




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long

   On Error GoTo errorhandler

    GetBankItemValue = Bank.Item(bankslot).Value



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)

   On Error GoTo errorhandler

    Bank.Item(bankslot).Value = ItemValue




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef dir As Byte, ByVal block As Boolean)

   On Error GoTo errorhandler

    If block Then
        blockvar = blockvar Or (2 ^ dir)
    Else
        blockvar = blockvar And Not (2 ^ dir)
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean

   On Error GoTo errorhandler

    If Not blockvar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Function IsHotbarSlot(ByVal X As Single, ByVal Y As Single) As Long
Dim Top As Long, Left As Long
Dim i As Long


   On Error GoTo errorhandler

    IsHotbarSlot = 0
    If X >= HotbarIcon1Bounds.Left And X <= HotbarIcon1Bounds.Left + HotbarIcon1Bounds.Right Then
        If Y >= HotbarIcon1Bounds.Top And Y <= HotbarIcon1Bounds.Top + HotbarIcon1Bounds.Bottom Then
            IsHotbarSlot = 1
            Exit Function
        End If
    End If
    If X >= HotbarIcon2Bounds.Left And X <= HotbarIcon2Bounds.Left + HotbarIcon2Bounds.Right Then
        If Y >= HotbarIcon2Bounds.Top And Y <= HotbarIcon2Bounds.Top + HotbarIcon2Bounds.Bottom Then
            IsHotbarSlot = 2
            Exit Function
        End If
    End If
    If X >= HotbarIcon3Bounds.Left And X <= HotbarIcon3Bounds.Left + HotbarIcon3Bounds.Right Then
        If Y >= HotbarIcon3Bounds.Top And Y <= HotbarIcon3Bounds.Top + HotbarIcon3Bounds.Bottom Then
            IsHotbarSlot = 3
            Exit Function
        End If
    End If
    If X >= HotbarIcon4Bounds.Left And X <= HotbarIcon4Bounds.Left + HotbarIcon4Bounds.Right Then
        If Y >= HotbarIcon4Bounds.Top And Y <= HotbarIcon4Bounds.Top + HotbarIcon4Bounds.Bottom Then
            IsHotbarSlot = 4
            Exit Function
        End If
    End If
    If X >= HotbarIcon5Bounds.Left And X <= HotbarIcon5Bounds.Left + HotbarIcon5Bounds.Right Then
        If Y >= HotbarIcon5Bounds.Top And Y <= HotbarIcon5Bounds.Top + HotbarIcon5Bounds.Bottom Then
            IsHotbarSlot = 5
            Exit Function
        End If
    End If
    If X >= HotbarIcon6Bounds.Left And X <= HotbarIcon6Bounds.Left + HotbarIcon6Bounds.Right Then
        If Y >= HotbarIcon6Bounds.Top And Y <= HotbarIcon6Bounds.Top + HotbarIcon6Bounds.Bottom Then
            IsHotbarSlot = 6
            Exit Function
        End If
    End If
    If X >= HotbarIcon7Bounds.Left And X <= HotbarIcon7Bounds.Left + HotbarIcon7Bounds.Right Then
        If Y >= HotbarIcon7Bounds.Top And Y <= HotbarIcon7Bounds.Top + HotbarIcon7Bounds.Bottom Then
            IsHotbarSlot = 7
            Exit Function
        End If
    End If
    If X >= HotbarIcon8Bounds.Left And X <= HotbarIcon8Bounds.Left + HotbarIcon8Bounds.Right Then
        If Y >= HotbarIcon8Bounds.Top And Y <= HotbarIcon8Bounds.Top + HotbarIcon8Bounds.Bottom Then
            IsHotbarSlot = 8
            Exit Function
        End If
    End If
    If X >= HotbarIcon9Bounds.Left And X <= HotbarIcon9Bounds.Left + HotbarIcon9Bounds.Right Then
        If Y >= HotbarIcon9Bounds.Top And Y <= HotbarIcon9Bounds.Top + HotbarIcon9Bounds.Bottom Then
            IsHotbarSlot = 9
            Exit Function
        End If
    End If
    If X >= HotbarIcon10Bounds.Left And X <= HotbarIcon10Bounds.Left + HotbarIcon10Bounds.Right Then
        If Y >= HotbarIcon10Bounds.Top And Y <= HotbarIcon10Bounds.Top + HotbarIcon10Bounds.Bottom Then
            IsHotbarSlot = 10
            Exit Function
        End If
    End If
    If X >= HotbarIcon11Bounds.Left And X <= HotbarIcon11Bounds.Left + HotbarIcon11Bounds.Right Then
        If Y >= HotbarIcon11Bounds.Top And Y <= HotbarIcon11Bounds.Top + HotbarIcon11Bounds.Bottom Then
            IsHotbarSlot = 11
            Exit Function
        End If
    End If
    If X >= HotbarIcon12Bounds.Left And X <= HotbarIcon12Bounds.Left + HotbarIcon12Bounds.Right Then
        If Y >= HotbarIcon12Bounds.Top And Y <= HotbarIcon12Bounds.Top + HotbarIcon12Bounds.Bottom Then
            IsHotbarSlot = 12
            Exit Function
        End If
    End If



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim soundName As String


   On Error GoTo errorhandler

    If entityNum <= 0 Then Exit Sub
    ' find the sound
    Select Case entityType
        ' animations
        Case SoundEntity.seAnimation
            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).sound)
                ' items
        Case SoundEntity.seItem
            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).sound)
        ' npcs
        Case SoundEntity.seNpc
            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(Npc(entityNum).sound)
        ' resources
        Case SoundEntity.seResource
            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).sound)
        ' spells
        Case SoundEntity.seSpell
            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(spell(entityNum).sound)
        ' other
        Case Else
            Exit Sub
    End Select
    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    PlaySound soundName, X, Y




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal Data1 As Long = 0, Optional ByVal data2 As String = "")
    ' exit out if we've already got a dialogue open

   On Error GoTo errorhandler

    If dialogueIndex > 0 Then Exit Sub
    ' set global dialogue index
    dialogueIndex = diIndex
    ' set the global dialogue data
    dialogueData1 = Data1
    dialogueData2 = data2

    ' set the captions
    dialogueTitle = diTitle
    dialogueText = diText
    If isYesNo Then
        dialogueFunction = 1
    Else
        dialogueFunction = 0
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "dialogue", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub dialogueHandler(ByVal Index As Long)
Dim buffer As clsBuffer
    ' find out which button

   On Error GoTo errorhandler

    If Index = 1 Then ' okay button
        ' dialogue index
    ElseIf Index = 2 Then ' yes button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendAcceptTradeRequest
            Case DIALOGUE_TYPE_FORGET
                ForgetSpell dialogueData1
            Case DIALOGUE_TYPE_PARTY
                SendAcceptParty
            Case DIALOGUE_TYPE_BUYHOUSE
                Set buffer = New clsBuffer
                buffer.WriteLong CBuyHouse
                buffer.WriteLong 1
                SendData buffer.ToArray
                Set buffer = Nothing
            Case DIALOGUE_TYPE_VISIT
                Set buffer = New clsBuffer
                buffer.WriteLong CAcceptVisit
                buffer.WriteLong 1
                SendData buffer.ToArray
                Set buffer = Nothing
            Case DIALOGUE_TYPE_REMOVEFRIEND
                Set buffer = New clsBuffer
                buffer.WriteLong CEditFriend
                buffer.WriteString dialogueData2
                buffer.WriteLong 1
                SendData buffer.ToArray
                Set buffer = Nothing
        End Select
    ElseIf Index = 3 Then ' no button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendDeclineTradeRequest
            Case DIALOGUE_TYPE_PARTY
                SendDeclineParty
            Case DIALOGUE_TYPE_BUYHOUSE
                Set buffer = New clsBuffer
                buffer.WriteLong CBuyHouse
                buffer.WriteLong 0
                SendData buffer.ToArray
                Set buffer = Nothing
            Case DIALOGUE_TYPE_VISIT
                Set buffer = New clsBuffer
                buffer.WriteLong CAcceptVisit
                buffer.WriteLong 0
                SendData buffer.ToArray
                Set buffer = Nothing
        End Select
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "dialogueHandler", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ProcessEventMovement(ByVal id As Long)


    ' Check if NPC is walking, and if so process moving them over

   On Error GoTo errorhandler

    If Map.MapEvents(id).Moving = 1 Then
            Select Case Map.MapEvents(id).dir
            Case DIR_UP
                Map.MapEvents(id).YOffset = Map.MapEvents(id).YOffset - ((ElapsedMTime / 1000) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).YOffset < 0 Then Map.MapEvents(id).YOffset = 0
                        Case DIR_DOWN
                Map.MapEvents(id).YOffset = Map.MapEvents(id).YOffset + ((ElapsedMTime / 1000) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).YOffset > 0 Then Map.MapEvents(id).YOffset = 0
                        Case DIR_LEFT
                Map.MapEvents(id).XOffset = Map.MapEvents(id).XOffset - ((ElapsedMTime / 1000) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).XOffset < 0 Then Map.MapEvents(id).XOffset = 0
                        Case DIR_RIGHT
                Map.MapEvents(id).XOffset = Map.MapEvents(id).XOffset + ((ElapsedMTime / 1000) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).XOffset > 0 Then Map.MapEvents(id).XOffset = 0
                    End Select
        ' Check if completed walking over to the next tile
        If Map.MapEvents(id).Moving > 0 Then
            If Map.MapEvents(id).dir = DIR_RIGHT Or Map.MapEvents(id).dir = DIR_DOWN Then
                If (Map.MapEvents(id).XOffset >= 0) And (Map.MapEvents(id).YOffset >= 0) Then
                    Map.MapEvents(id).Moving = 0
                    If Map.MapEvents(id).Step = 1 Then
                        Map.MapEvents(id).Step = 3
                    Else
                        Map.MapEvents(id).Step = 1
                    End If
                End If
            Else
                If (Map.MapEvents(id).XOffset <= 0) And (Map.MapEvents(id).YOffset <= 0) Then
                    Map.MapEvents(id).Moving = 0
                    If Map.MapEvents(id).Step = 1 Then
                        Map.MapEvents(id).Step = 3
                    Else
                        Map.MapEvents(id).Step = 1
                    End If
                End If
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProcessEventMovement", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function GetColorString(color As Long)

   On Error GoTo errorhandler

    Select Case color
        Case 0
            GetColorString = "Black"
        Case 1
            GetColorString = "Blue"
        Case 2
            GetColorString = "Green"
        Case 3
            GetColorString = "Cyan"
        Case 4
            GetColorString = "Red"
        Case 5
            GetColorString = "Magenta"
        Case 6
            GetColorString = "Brown"
        Case 7
            GetColorString = "Grey"
        Case 8
            GetColorString = "Dark Grey"
        Case 9
            GetColorString = "Bright Blue"
        Case 10
            GetColorString = "Bright Green"
        Case 11
            GetColorString = "Bright Cyan"
        Case 12
            GetColorString = "Bright Red"
        Case 13
            GetColorString = "Pink"
        Case 14
            GetColorString = "Yellow"
        Case 15
            GetColorString = "White"

    End Select


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetColorString", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub ClearEventChat()
    Dim i As Long

   On Error GoTo errorhandler

    If AnotherChat = 1 Then
        For i = 1 To 4
            EventChoiceVisible(i) = False
        Next
        EventText = ""
        EventChatType = 1
        EventChatTimer = GetTickCount + 100
    ElseIf AnotherChat = 2 Then
        For i = 1 To 4
            EventChoiceVisible(i) = False
        Next
        EventText = ""
        EventChatType = 1
        EventChatTimer = GetTickCount + 100
    Else
        EventChat = False
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearEventChat", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ProcessWeather()
Dim i As Long

   On Error GoTo errorhandler

    If CurrentWeather > 0 Then
        i = Rand(1, 101 - CurrentWeatherIntensity)
        If i = 1 Then
            'Add a new particle
            For i = 1 To MAX_WEATHER_PARTICLES
                If WeatherParticle(i).InUse = False Then
                    If Rand(1, 2) = 1 Then
                        WeatherParticle(i).InUse = True
                        WeatherParticle(i).type = CurrentWeather
                        WeatherParticle(i).Velocity = Rand(8, 14)
                        WeatherParticle(i).X = (TileView.Left * 32) - 32
                        WeatherParticle(i).Y = (TileView.Top * 32) + Rand(-32, frmMain.ScaleHeight)
                    Else
                        WeatherParticle(i).InUse = True
                        WeatherParticle(i).type = CurrentWeather
                        WeatherParticle(i).Velocity = Rand(10, 15)
                        WeatherParticle(i).X = (TileView.Left * 32) + Rand(-32, frmMain.ScaleWidth)
                        WeatherParticle(i).Y = (TileView.Top * 32) - 32
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    If CurrentWeather = WEATHER_TYPE_STORM Then
        i = Rand(1, 400 - CurrentWeatherIntensity)
        If i = 1 Then
            'Draw Thunder
            DrawThunder = Rand(15, 22)
            PlaySound Sound_Thunder, -1, -1
        End If
    End If
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).X > TileView.Right * 32 Or WeatherParticle(i).Y > TileView.Bottom * 32 Then
                WeatherParticle(i).InUse = False
            Else
                WeatherParticle(i).X = WeatherParticle(i).X + WeatherParticle(i).Velocity
                WeatherParticle(i).Y = WeatherParticle(i).Y + WeatherParticle(i).Velocity
            End If
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProcessWeather", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub AddChatBubble(ByVal target As Long, ByVal targetType As Byte, ByVal Msg As String, ByVal colour As Long)
Dim i As Long, Index As Long

    ' set the global index

   On Error GoTo errorhandler

    chatBubbleIndex = chatBubbleIndex + 1
    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    ' default to new bubble
    Index = chatBubbleIndex
    ' loop through and see if that player/npc already has a chat bubble
    For i = 1 To MAX_BYTE
        If chatBubble(i).targetType = targetType Then
            If chatBubble(i).target = target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                Index = i
                Exit For
            End If
        End If
    Next
    ' set the bubble up
    With chatBubble(Index)
        .target = target
        .targetType = targetType
        .Msg = Msg
        .colour = colour
        .Timer = GetTickCount
        .active = True
    End With


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddChatBubble", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerNextStep() As Long
    Dim tim As Long, sx As Long, sy As Long, Pos() As Long, reachable As Boolean, X As Long, Y As Long, j As Long, LastSum As Long, Sum As Long, fx As Long, fy As Long, i As Long
    Dim path() As Vector, LastX As Long, LastY As Long, did As Boolean, pos1() As Long, test1 As Boolean
    Dim dir As Long, ExitFunc As Boolean

    On Error GoTo errorhandler

    dir = 5
    If GettingMap Then Exit Function
    ReDim pos1(0 To Map.MaxX, 0 To Map.MaxY)
    ReDim Pos(0 To Map.MaxX, 0 To Map.MaxY)
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If Map.Tile(X, Y).type = TILE_TYPE_BLOCKED Then
                pos1(X, Y) = 9
            End If
        Next
    Next
    For X = 1 To MAX_PLAYERS
        If IsPlaying(X) Then
            If X <> MyIndex Then
                If Player(X).Map = Player(MyIndex).Map Then
                    pos1(Player(X).X, Player(X).Y) = 9
                    If Player(X).Pet.Alive Then
                        pos1(Player(X).Pet.X, Player(X).Pet.Y) = 9
                    End If
                End If
            Else
                If Player(X).Pet.Alive Then
                    pos1(Player(X).Pet.X, Player(X).Pet.Y) = 9
                End If
            End If
        End If
    Next
    For X = 1 To Map.CurrentEvents
        If Map.MapEvents(X).Visible Then
            If Map.MapEvents(X).WalkThrough = 0 Then
                pos1(Map.MapEvents(X).X, Map.MapEvents(X).Y) = 9
            End If
        End If
    Next
    If Resource_Index > 0 Then
        For i = 1 To Resource_Index
            pos1(MapResource(i).X, MapResource(i).Y) = 9
        Next
    End If
    For X = 1 To MAX_MAP_NPCS
        If MapNpc(X).Vital(Vitals.HP) > 0 Then
            If MapNpc(X).Num > 0 Then
                pos1(MapNpc(X).X, MapNpc(X).Y) = 9
            End If
        End If
    Next
    For X = 1 To MAX_ZONES
        For Y = 1 To MAX_MAP_NPCS * 2
            If ZoneNPC(X).Npc(Y).Map = GetPlayerMap(MyIndex) Then
                If ZoneNPC(X).Npc(Y).Vital(Vitals.HP) > 0 Then
                    If ZoneNPC(X).Npc(Y).Num > 0 Then
                        pos1(ZoneNPC(X).Npc(Y).X, ZoneNPC(X).Npc(Y).Y) = 9
                    End If
                End If
            End If
        Next
    Next
    Do While ExitFunc = False
        Pos = pos1
        sx = GetPlayerX(MyIndex)
        sy = GetPlayerY(MyIndex)
        fx = WalkToX
        fy = WalkToY
        If fx < 0 Then fx = 0: test1 = True
        If fy < 0 Then fy = 0: test1 = True
        If fx > Map.MaxX Then fx = Map.MaxX: test1 = True
        If fy > Map.MaxY Then fy = Map.MaxY: test1 = True
        If test1 Then
            If IsOneBlockAway(fx, fy, sx, sy) Then
                If sx > fx Then
                    dir = DIR_LEFT
                ElseIf sx < fx Then
                    dir = DIR_RIGHT
                ElseIf sy > fy Then
                    dir = DIR_UP
                ElseIf sy < fy Then
                    dir = DIR_DOWN
                End If
                If CheckDirection(dir) = False Then
                    Exit Do
                End If
            End If
        End If
        tim = 0
        Pos(sx, sy) = 100 + tim
        Pos(fx, fy) = 2
        'reset reachable
        reachable = False
            'Do while reachable is false... if its set true in progress, we jump out
        'If the path is decided unreachable in process, we will use exit sub. Not proper,
        'but faster ;-)
        Do While reachable = False
            'we loop through all squares
            For j = 0 To Map.MaxY
                For i = 0 To Map.MaxX
                    'If j = 8 And i = 5 Then MsgBox "hi!"
                    'If they are to be extended, the pointer TIM is on them
                    If Pos(i, j) = 100 + tim Then
                    'The part is to be extended, so do it
                        'We have to make sure that there is a pos(i+1,j) BEFORE we actually use it,
                        'because then we get error... If the square is on side, we dont test for this one!
                        If i < Map.MaxX Then
                            'If there isnt a wall, or any other... thing
                            If Pos(i + 1, j) = 0 Then
                                'Expand it, and make its pos equal to tim+1, so the next time we make this loop,
                                'It will exapand that square too! This is crucial part of the program
                                Pos(i + 1, j) = 100 + tim + 1
                            ElseIf Pos(i + 1, j) = 2 Then
                                'If the position is no 0 but its 2 (FINISH) then Reachable = true!!! We found end
                                reachable = True
                            End If
                        End If
                        'This is the same as the last one, as i said a lot of copy paste work and editing that
                        'This is simply another side that we have to test for... so instead of i+1 we have i-1
                        'Its actually pretty same then... I wont comment it therefore, because its only repeating
                        'same thing with minor changes to check sides
                        If i > 0 Then
                            If Pos((i - 1), j) = 0 Then
                                Pos(i - 1, j) = 100 + tim + 1
                            ElseIf Pos(i - 1, j) = 2 Then
                                reachable = True
                            End If
                        End If
                                        If j < Map.MaxY Then
                            If Pos(i, j + 1) = 0 Then
                                Pos(i, j + 1) = 100 + tim + 1
                            ElseIf Pos(i, j + 1) = 2 Then
                                reachable = True
                            End If
                        End If
                                        If j > 0 Then
                            If Pos(i, j - 1) = 0 Then
                                Pos(i, j - 1) = 100 + tim + 1
                            ElseIf Pos(i, j - 1) = 2 Then
                                reachable = True
                            End If
                        End If
                    End If
                    DoEvents
                Next i
            Next j
                    'If the reachable is STILL false, then
            If reachable = False Then
                'reset sum
                Sum = 0
                For j = 0 To Map.MaxY
                    For i = 0 To Map.MaxX
                    'we add up ALL the squares
                    Sum = Sum + Pos(i, j)
                    Next i
                Next j
                            'Now if the sum is euqal to the last sum, its not reachable, if it isnt, then we store
                'sum to lastsum
                If Sum = LastSum Then
                    dir = 5
                    Exit Function
                Else
                    LastSum = Sum
                End If
            End If
                    'we increase the pointer to point to the next squares to be expanded
            tim = tim + 1
        Loop
            'We work backwards to find the way...
        LastX = fx
        LastY = fy
            ReDim path(tim + 1)
            'The following code may be a little bit confusing but ill try my best to explain it.
        'We are working backwards to find ONE of the shortest ways back to Start.
        'So we repeat the loop until the LastX and LastY arent in start. Look in the code to see
        'how LastX and LasY change
        Do While LastX <> sx Or LastY <> sy
            'We decrease tim by one, and then we are finding any adjacent square to the final one, that
            'has that value. So lets say the tim would be 5, because it takes 5 steps to get to the target.
            'Now everytime we decrease that, so we make it 4, and we look for any adjacent square that has
            'that value. When we find it, we just color it yellow as for the solution
            tim = tim - 1
            'reset did to false
            did = False
                    'If we arent on edge
            If LastX < Map.MaxX Then
                'check the square on the right of the solution. Is it a tim-1 one? or just a blank one
                If Pos(LastX + 1, LastY) = 100 + tim Then
                    'if it, then make it yellow, and change did to true
                    LastX = LastX + 1
                    did = True
                End If
            End If
                    'This will then only work if the previous part didnt execute, and did is still false. THen
            'we want to check another square, the on left. Is it a tim-1 one ?
            If did = False Then
                If LastX > 0 Then
                    If Pos(LastX - 1, LastY) = 100 + tim Then
                        LastX = LastX - 1
                        did = True
                    End If
                End If
            End If
                    'We check the one below it
            If did = False Then
                If LastY < Map.MaxY Then
                    If Pos(LastX, LastY + 1) = 100 + tim Then
                        LastY = LastY + 1
                        did = True
                    End If
                End If
            End If
                    'And above it. One of these have to be it, since we have found the solution, we know that already
            'there is a way back.
            If did = False Then
                If LastY > 0 Then
                    If Pos(LastX, LastY - 1) = 100 + tim Then
                        LastY = LastY - 1
                    End If
                End If
            End If
                    path(tim).X = LastX
            path(tim).Y = LastY
                    'Now we loop back and decrease tim, and look for the next square with lower value
            DoEvents
        Loop
            'Ok we got a path. Now, lets look at the first step and see what direction we should take.
        If path(1).X > LastX Then
            dir = DIR_RIGHT
        ElseIf path(1).Y > LastY Then
            dir = DIR_DOWN
        ElseIf path(1).Y < LastY Then
            dir = DIR_UP
        ElseIf path(1).X < LastX Then
            dir = DIR_LEFT
        End If
        'check directional blocking
        If isDirBlocked(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, dir + 1) Then
            pos1(path(1).X, path(1).Y) = 9
        Else
            ExitFunc = True
        End If
    Loop
    GetPlayerNextStep = dir


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerNextStep", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub UpdateDebugCaption()
    If Options.FullScreen = 1 Then Exit Sub
    If DebugMode Then
        If InGame Then
            If Player(MyIndex).InHouse > 0 Then
                frmMain.Caption = Trim$(Servers(ServerIndex).Game_Name) & " -  " & Trim$(Player(Player(MyIndex).InHouse).Name) & "'s House" & " - Debug Mode: " & ErrorCount & " errors occured."
            Else
                frmMain.Caption = Trim$(Servers(ServerIndex).Game_Name) & " -  " & Trim$(Map.Name) & " - Debug Mode: " & ErrorCount & " errors occured."
            End If
        Else
            frmMain.Caption = Trim$(Servers(ServerIndex).Game_Name) & " - Debug Mode: " & ErrorCount & " errors occured."
        End If
    Else
        If InGame Then
            If Player(MyIndex).InHouse > 0 Then
                frmMain.Caption = Trim$(Servers(ServerIndex).Game_Name) & " -  " & Trim$(Player(Player(MyIndex).InHouse).Name) & "'s House"
            Else
                frmMain.Caption = Trim$(Servers(ServerIndex).Game_Name) & " -  " & Trim$(Map.Name)
            End If
        Else
            frmMain.Caption = Trim$(Servers(ServerIndex).Game_Name)
        End If
    End If
End Sub

Public Function uptimeToDHMS(ByVal inSeconds As Long) As String
        uptimeToDHMS = ""
        Dim seconds As Integer
        seconds = inSeconds Mod 60
        inSeconds = (inSeconds - seconds) / 60
        Dim minutes As Integer
        minutes = inSeconds Mod 60
        inSeconds = (inSeconds - minutes) / 60
        Dim hours As Integer
        hours = inSeconds Mod 24
        inSeconds = (inSeconds - hours) / 24
        Dim days As Integer
        days = inSeconds
        
        If days < 0 Then days = days * -1
        If hours < 0 Then hours = hours * -1
        If minutes < 0 Then minutes = minutes * -1
        If seconds < 0 Then seconds = seconds * -1
        
        uptimeToDHMS = ""
        If days < 10 Then
            uptimeToDHMS = uptimeToDHMS & "0" & days
        Else
            uptimeToDHMS = uptimeToDHMS & days
        End If
        uptimeToDHMS = uptimeToDHMS & ":"
        If hours < 10 Then
            uptimeToDHMS = uptimeToDHMS & "0" & hours
        Else
            uptimeToDHMS = uptimeToDHMS & hours
        End If
        uptimeToDHMS = uptimeToDHMS & ":"
        If minutes < 10 Then
            uptimeToDHMS = uptimeToDHMS & "0" & minutes
        Else
            uptimeToDHMS = uptimeToDHMS & minutes
        End If
        uptimeToDHMS = uptimeToDHMS & ":"
        If seconds < 10 Then
            uptimeToDHMS = uptimeToDHMS & "0" & seconds
        Else
            uptimeToDHMS = uptimeToDHMS & seconds
        End If
End Function

