Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim MapNum As Long, LastUpdateMapLogic As Long
    Dim i As Long, x As Long, Buffer As clsBuffer
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers() As Long, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals() As Long
    
    ReDim LastUpdatePlayerVitals(1 To MAX_PLAYERS)
    ReDim LastUpdateSavePlayers(1 To MAX_PLAYERS)


   On Error GoTo errorhandler

    ServerOnline = True
    StartTime = GetTickCount

    Do While ServerOnline
        Tick = GetTickCount
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
         ' Player loop
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Tick > tmr25 Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If GetTickCount > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).characters(TempPlayer(i).CurChar).Spell(TempPlayer(i).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.Target, TempPlayer(i).spellBuffer.tType, TempPlayer(i).spellBuffer.TargetZone
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.Target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < GetTickCount Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For x = 1 To MAX_DOTS
                        HandleDoT_Player i, x
                        HandleHoT_Player i, x
                    Next
                    If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive = True Then
                        If TempPlayer(i).PetspellBuffer.Spell > 0 Then
                            If GetTickCount > TempPlayer(i).PetspellBuffer.Timer + (Spell(Player(i).characters(TempPlayer(i).CurChar).Pet.Spell(TempPlayer(i).PetspellBuffer.Spell)).CastTime * 1000) Then
                                PetCastSpell i, TempPlayer(i).PetspellBuffer.Spell, TempPlayer(i).PetspellBuffer.Target, TempPlayer(i).PetspellBuffer.tType, True, TempPlayer(i).PetspellBuffer.TargetZone
                                TempPlayer(i).PetspellBuffer.Spell = 0
                                TempPlayer(i).PetspellBuffer.Timer = 0
                                TempPlayer(i).PetspellBuffer.Target = 0
                                TempPlayer(i).PetspellBuffer.tType = 0
                            End If
                        End If
                        
                        ' check if need to turn off stunned
                        If TempPlayer(i).PetStunDuration > 0 Then
                            If GetTickCount > TempPlayer(i).PetStunTimer + (TempPlayer(i).PetStunDuration * 1000) Then
                                TempPlayer(i).PetStunDuration = 0
                                TempPlayer(i).PetStunTimer = 0
                            End If
                        End If
                        
                        ' check regen timer
                        If TempPlayer(i).PetstopRegen Then
                            If TempPlayer(i).PetstopRegenTimer + 5000 < GetTickCount Then
                                TempPlayer(i).PetstopRegen = False
                                TempPlayer(i).PetstopRegenTimer = 0
                            End If
                        End If
                        
                        ' HoT and DoT logic
                        For x = 1 To MAX_DOTS
                            HandleDoT_Pet i, x
                            HandleHoT_Pet i, x
                        Next
                    End If
                End If
            
                ' Checks to update player vitals every 5 seconds - Can be tweaked
                If Tick > LastUpdatePlayerVitals(i) Then
                    UpdatePlayerVitals i
                    LastUpdatePlayerVitals(i) = GetTickCount + 5000
                End If
                
                ' Checks to save players every 5 minutes - Can be tweaked
                If Tick > LastUpdateSavePlayers(i) Then
                    UpdateSavePlayers i
                    LastUpdateSavePlayers(i) = GetTickCount + 300000
                End If
            End If
        Next
        
        If GetTickCount > tmr25 Then
            UpdateEventLogic
            tmr25 = GetTickCount + 25
        End If
        
      

        ' Check for disconnections every half second
        If Tick > tmr500 Then
        
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If Player(i).characters(TempPlayer(i).CurChar).InHouse > 0 Then
                        If IsPlaying(Player(i).characters(TempPlayer(i).CurChar).InHouse) Then
                            If Player(Player(i).characters(TempPlayer(i).CurChar).InHouse).characters(TempPlayer(i).CurChar).InHouse <> Player(i).characters(TempPlayer(i).CurChar).InHouse Then
                                Player(i).characters(TempPlayer(i).CurChar).InHouse = 0
                                PlayerWarp i, Player(i).characters(TempPlayer(i).CurChar).LastMap, Player(i).characters(TempPlayer(i).CurChar).LastX, Player(i).characters(TempPlayer(i).CurChar).LastY
                                PlayerMsg i, "Your visitation has ended. Possibly due to a disconnection. You are being warped back to your previous location.", White
                            End If
                        End If
                    End If
                End If
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            
            UpdateZoneLogic

          
        End If

        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            tmr1000 = GetTickCount + 1000
        End If
        
        For MapNum = MIN_MAPS To MAX_MAPS
            ' Checks to spawn map items every 5 minutes - Can be tweaked
            If Tick > LastUpdateMapSpawnItems Then
                UpdateMapSpawnItems MapNum
                LastUpdateMapSpawnItems = GetTickCount + 300000
            End If
        
            ' update map logic
            If Tick > LastUpdateMapLogic Then
                UpdateMapLogic MapNum
            End If
        Next
        
        If Tick > LastUpdateMapSpawnItems Then
            LastUpdateMapSpawnItems = GetTickCount + 300000
        End If
        
        If Tick > LastUpdateMapLogic Then
            LastUpdateMapLogic = GetTickCount + 500
        End If
        
        

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
        ' Set server CPS on label
        frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
    Loop
    
    End
    DestroyServer


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ServerLoop", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub UpdateMapSpawnItems(y As Long)
    Dim x As Long
    ' Make sure no one is on the map when it respawns

   On Error GoTo errorhandler

    If Not PlayersOnMap(y) Then

        ' Clear out unnecessary junk
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next

        ' Spawn the items
        Call SpawnMapItems(y)
        Call SendMapItemsToAll(y)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateMapSpawnItems", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub UpdateMapLogic(MapNum As Long)
    Dim i As Long, x As Long, n As Long, x1 As Long, y1 As Long, y As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, npcnum As Long
    Dim Target As Long, TargetType As Byte, TargetZone As Byte, didwalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim targetx As Long, targety As Long, target_verify As Boolean, d As Long, z As Long
    Dim shortestZone As Long
    Dim shortestNum As Long
    Dim shortestD As Long, invslot As Long, chase As Boolean

        ' items appearing to everyone

   On Error GoTo errorhandler
   
        ' Clear old map projectiles
        For i = 1 To MAX_PROJECTILES
            If MapProjectiles(MapNum, i).Timer < GetTickCount Then
                ClearMapProjectile MapNum, i
            End If
        Next

        For i = 1 To MAX_MAP_ITEMS
            If MapItem(MapNum, i).Num > 0 Then
                If MapItem(MapNum, i).playerName <> vbNullString Then
                    ' make item public?
                    If MapItem(MapNum, i).playerTimer < GetTickCount Then
                        ' make it public
                        MapItem(MapNum, i).playerName = vbNullString
                        MapItem(MapNum, i).playerTimer = 0
                        ' send updates to everyone
                        SendMapItemsToAll MapNum
                    End If
                End If
                ' despawn item?
                If MapItem(MapNum, i).canDespawn Then
                    If MapItem(MapNum, i).despawnTimer < GetTickCount Then
                        ' despawn it
                        ClearMapItem i, MapNum
                        ' send updates to everyone
                        SendMapItemsToAll MapNum
                    End If
                End If
                If MapItem(MapNum, i).Num > 0 Then 'Item Still Exists
                    chase = True
                    For x = 1 To MAX_ZONES
                        For y = 1 To MAX_MAP_NPCS * 2
                            If ZoneNpc(x).Npc(y).Num > 0 Then
                                If ZoneNpc(x).Npc(y).Vital(HP) > 0 Then
                                    If ZoneNpc(x).Npc(y).Map = MapNum Then
                                        If ZoneNpc(x).Npc(y).TargetType = TARGET_TYPE_ITEM Then
                                            If ZoneNpc(x).Npc(y).Target = i Then
                                                'NPC Already on it
                                                chase = False
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                    For x = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(x).Num > 0 Then
                            If MapNpc(MapNum).Npc(x).Vital(HP) > 0 Then
                                If MapNpc(MapNum).Npc(x).TargetType = TARGET_TYPE_ITEM Then
                                    If MapNpc(MapNum).Npc(x).Target = i Then
                                        chase = False
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If chase Then
                        shortestD = 10000
                        'Get NPCs Going
                        If 1 = 0 Then 'NPCS don't grab items.
                            'Do Nothing
                        Else
                            For x = 1 To MAX_ZONES
                                For y = 1 To MAX_MAP_NPCS * 2
                                    If ZoneNpc(x).Npc(y).Num > 0 Then
                                        If ZoneNpc(x).Npc(y).Vital(Vitals.HP) > 0 And Npc(ZoneNpc(x).Npc(y).Num).ItemBehaviour = 1 Then
                                            If ZoneNpc(x).Npc(y).Map = MapNum Then
                                                If ZoneNpc(x).Npc(y).TargetType = 0 Then
                                                    If FindOpenNpcInvSlot(0, y, x) > 0 Then
                                                        d = dist(MapItem(MapNum, i).x, MapItem(MapNum, i).y, ZoneNpc(x).Npc(y).x, ZoneNpc(x).Npc(y).y)
                                                        If d < shortestD Then
                                                            shortestZone = x
                                                            shortestNum = y
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            Next
                            For x = 1 To MAX_MAP_NPCS
                                If MapNpc(MapNum).Npc(x).Num > 0 Then
                                    If MapNpc(MapNum).Npc(x).Vital(Vitals.HP) > 0 And Npc(MapNpc(MapNum).Npc(x).Num).ItemBehaviour = 1 Then
                                        If MapNpc(MapNum).Npc(x).TargetType = 0 Then
                                            If FindOpenNpcInvSlot(MapNum, x, 0) > 0 Then
                                                d = dist(MapItem(MapNum, i).x, MapItem(MapNum, i).y, MapNpc(MapNum).Npc(x).x, MapNpc(MapNum).Npc(x).y)
                                                If d < shortestD Then
                                                    shortestZone = 0
                                                    shortestNum = x
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            If shortestNum > 0 Then
                                If shortestZone > 0 Then
                                    ZoneNpc(shortestZone).Npc(shortestNum).TargetType = TARGET_TYPE_ITEM
                                    ZoneNpc(shortestZone).Npc(shortestNum).Target = i
                                Else
                                    MapNpc(MapNum).Npc(shortestNum).TargetType = TARGET_TYPE_ITEM
                                    MapNpc(MapNum).Npc(shortestNum).Target = i
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        
        '  Close the doors
        If TickCount > temptile(MapNum).DoorTimer + 5000 Then
            For x1 = 0 To Map(MapNum).MaxX
                For y1 = 0 To Map(MapNum).MaxY
                    If Map(MapNum).Tile(x1, y1).type = TILE_TYPE_KEY And temptile(MapNum).DoorOpen(x1, y1) = YES Then
                        temptile(MapNum).DoorOpen(x1, y1) = NO
                        SendMapKeyToMap MapNum, x1, y1, 0
                    End If
                Next
            Next
        End If
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(MapNum).Npc(i).Num > 0 Then
                For x = 1 To MAX_DOTS
                    HandleDoT_Npc MapNum, i, x
                    HandleHoT_Npc MapNum, i, x
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(MapNum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(MapNum).Resource_Count
                Resource_index = Map(MapNum).Tile(ResourceCache(MapNum).ResourceData(i).x, ResourceCache(MapNum).ResourceData(i).y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(MapNum).ResourceData(i).ResourceState = 1 Or ResourceCache(MapNum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(MapNum).ResourceData(i).ResourceTimer = GetTickCount
                            ResourceCache(MapNum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(MapNum).ResourceData(i).cur_health = Resource(Resource_index).Health
                            SendResourceCacheToMap MapNum, i
                        End If
                    End If
                End If
            Next
        End If


            TickCount = GetTickCount
            
            For x = 1 To MAX_MAP_NPCS
                npcnum = MapNpc(MapNum).Npc(x).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(npcnum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(npcnum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(MapNum).Npc(x).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum And MapNpc(MapNum).Npc(x).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive Then
                                            n = Npc(npcnum).Range
                                            DistanceX = MapNpc(MapNum).Npc(x).x - Player(i).characters(TempPlayer(i).CurChar).Pet.x
                                            DistanceY = MapNpc(MapNum).Npc(x).y - Player(i).characters(TempPlayer(i).CurChar).Pet.y
        
                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                            If DistanceY < 0 Then DistanceY = DistanceY * -1
        
                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n And DistanceY <= n Then
                                                If Npc(npcnum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    If Len(Trim$(Npc(npcnum).AttackSay)) > 0 Then
                                                        Call PlayerMsg(i, Trim$(Npc(npcnum).Name) & " says: " & Trim$(Npc(npcnum).AttackSay), SayColor)
                                                    End If
                                                    MapNpc(MapNum).Npc(x).TargetType = TARGET_TYPE_PET
                                                    MapNpc(MapNum).Npc(x).Target = i
                                                End If
                                            End If
                                        Else
                                            n = Npc(npcnum).Range
                                            DistanceX = MapNpc(MapNum).Npc(x).x - GetPlayerX(i)
                                            DistanceY = MapNpc(MapNum).Npc(x).y - GetPlayerY(i)
        
                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                            If DistanceY < 0 Then DistanceY = DistanceY * -1
        
                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n And DistanceY <= n Then
                                                If Npc(npcnum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    If Len(Trim$(Npc(npcnum).AttackSay)) > 0 Then
                                                        Call PlayerMsg(i, Trim$(Npc(npcnum).Name) & " says: " & Trim$(Npc(npcnum).AttackSay), SayColor)
                                                    End If
                                                    MapNpc(MapNum).Npc(x).TargetType = 1 ' player
                                                    MapNpc(MapNum).Npc(x).Target = i
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then
                    If MapNpc(MapNum).Npc(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(MapNum).Npc(x).StunTimer + (MapNpc(MapNum).Npc(x).StunDuration * 1000) Then
                            MapNpc(MapNum).Npc(x).StunDuration = 0
                            MapNpc(MapNum).Npc(x).StunTimer = 0
                        End If
                    Else
                            
                        Target = MapNpc(MapNum).Npc(x).Target
                        TargetType = MapNpc(MapNum).Npc(x).TargetType
    
                        ' Check to see if its time for the npc to walk
                        If Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        
                            If TargetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If Target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                        didwalk = False
                                        target_verify = True
                                        targety = GetPlayerY(Target)
                                        targetx = GetPlayerX(Target)
                                    Else
                                        MapNpc(MapNum).Npc(x).TargetType = 0 ' clear
                                        MapNpc(MapNum).Npc(x).Target = 0
                                    End If
                                End If
                            
                            ElseIf TargetType = 2 Then 'npc
                                
                                If Target > 0 Then
                                    
                                    If MapNpc(MapNum).Npc(Target).Num > 0 Then
                                        didwalk = False
                                        target_verify = True
                                        targety = MapNpc(MapNum).Npc(Target).y
                                        targetx = MapNpc(MapNum).Npc(Target).x
                                    Else
                                        MapNpc(MapNum).Npc(x).TargetType = 0 ' clear
                                        MapNpc(MapNum).Npc(x).Target = 0
                                    End If
                                End If
                            ElseIf TargetType = TARGET_TYPE_ITEM Then
                                If Target > 0 Then
                                    If MapItem(MapNum, Target).Num > 0 Then
                                        didwalk = False
                                        target_verify = True
                                        targetx = MapItem(MapNum, Target).x
                                        targety = MapItem(MapNum, Target).y
                                    Else
                                        MapNpc(MapNum).Npc(x).TargetType = 0 ' clear
                                        MapNpc(MapNum).Npc(x).Target = 0
                                    End If
                                End If
                            ElseIf TargetType = TARGET_TYPE_PET Then
                                If Target > 0 Then
                                    
                                    If IsPlaying(Target) = True And GetPlayerMap(Target) = MapNum And Player(Target).characters(TempPlayer(Target).CurChar).Pet.Alive = True Then
                                        didwalk = False
                                        target_verify = True
                                        targety = Player(Target).characters(TempPlayer(Target).CurChar).Pet.y
                                        targetx = Player(Target).characters(TempPlayer(Target).CurChar).Pet.x
                                    Else
                                        MapNpc(MapNum).Npc(x).TargetType = 0 ' clear
                                        MapNpc(MapNum).Npc(x).Target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                'Gonna make the npcs smarter.. Implementing a pathfinding algorithm.. we shall see what happens.
                                If IsOneBlockAway(targetx, targety, CLng(MapNpc(MapNum).Npc(x).x), CLng(MapNpc(MapNum).Npc(x).y)) = False Then
                                    If PathfindingType = 1 Then
                                        i = Int(Rnd * 5)
            
                                        ' Lets move the npc
                                        Select Case i
                                            Case 0
            
                                                ' Up
                                                If MapNpc(MapNum).Npc(x).y > targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_UP, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If MapNpc(MapNum).Npc(x).y < targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_DOWN, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If MapNpc(MapNum).Npc(x).x > targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_LEFT, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If MapNpc(MapNum).Npc(x).x < targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_RIGHT, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 1
            
                                                ' Right
                                                If MapNpc(MapNum).Npc(x).x < targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_RIGHT, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If MapNpc(MapNum).Npc(x).x > targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_LEFT, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If MapNpc(MapNum).Npc(x).y < targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_DOWN, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If MapNpc(MapNum).Npc(x).y > targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_UP, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 2
            
                                                ' Down
                                                If MapNpc(MapNum).Npc(x).y < targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_DOWN, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If MapNpc(MapNum).Npc(x).y > targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_UP, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If MapNpc(MapNum).Npc(x).x < targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_RIGHT, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If MapNpc(MapNum).Npc(x).x > targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_LEFT, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 3
            
                                                ' Left
                                                If MapNpc(MapNum).Npc(x).x > targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_LEFT, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If MapNpc(MapNum).Npc(x).x < targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_RIGHT, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If MapNpc(MapNum).Npc(x).y > targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_UP, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If MapNpc(MapNum).Npc(x).y < targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_DOWN, False, 0) Then
                                                        Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING, False, 0)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                        End Select
            
                                        ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                        If Not didwalk Then
                                            If MapNpc(MapNum).Npc(x).x - 1 = targetx And MapNpc(MapNum).Npc(x).y = targety Then
                                                If MapNpc(MapNum).Npc(x).Dir <> DIR_LEFT Then
                                                    Call NpcDir(MapNum, x, DIR_LEFT, False, 0)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If MapNpc(MapNum).Npc(x).x + 1 = targetx And MapNpc(MapNum).Npc(x).y = targety Then
                                                If MapNpc(MapNum).Npc(x).Dir <> DIR_RIGHT Then
                                                    Call NpcDir(MapNum, x, DIR_RIGHT, False, 0)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If MapNpc(MapNum).Npc(x).x = targetx And MapNpc(MapNum).Npc(x).y - 1 = targety Then
                                                If MapNpc(MapNum).Npc(x).Dir <> DIR_UP Then
                                                    Call NpcDir(MapNum, x, DIR_UP, False, 0)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If MapNpc(MapNum).Npc(x).x = targetx And MapNpc(MapNum).Npc(x).y + 1 = targety Then
                                                If MapNpc(MapNum).Npc(x).Dir <> DIR_DOWN Then
                                                    Call NpcDir(MapNum, x, DIR_DOWN, False, 0)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            ' We could not move so Target must be behind something, walk randomly.
                                            If Not didwalk Then
                                                i = Int(Rnd * 2)
            
                                                If i = 1 Then
                                                    i = Int(Rnd * 4)
            
                                                    If CanNpcMove(MapNum, x, i, False, 0) Then
                                                        Call NpcMove(MapNum, x, i, MOVING_WALKING, False, 0)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        i = FindNpcPath(MapNum, x, targetx, targety, False, 0)
                                        If i < 4 Then 'Returned an answer. Move the NPC
                                            If CanNpcMove(MapNum, x, i, False, 0) Then
                                                NpcMove MapNum, x, i, MOVING_WALKING, False, 0
                                            End If
                                        Else 'No good path found. Move randomly
                                            i = Int(Rnd * 4)
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
                
                                                If CanNpcMove(MapNum, x, i, False, 0) Then
                                                    Call NpcMove(MapNum, x, i, MOVING_WALKING, False, 0)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    If MapNpc(MapNum).Npc(x).TargetType = TARGET_TYPE_ITEM Then
                                        If MapNpc(MapNum).Npc(x).Target > 0 Then
                                            If MapItem(MapNum, MapNpc(MapNum).Npc(x).Target).Num > 0 Then
                                                'NPC Needs to take item...
                                                If FindOpenNpcInvSlot(MapNum, x, 0) > 0 Then
                                                    invslot = FindOpenNpcInvSlot(MapNum, x, 0)
                                                    MapNpc(MapNum).Npc(x).Inventory(invslot).Num = MapItem(MapNum, MapNpc(MapNum).Npc(x).Target).Num
                                                    MapNpc(MapNum).Npc(x).Inventory(invslot).Value = MapItem(MapNum, MapNpc(MapNum).Npc(x).Target).Value
                                                    ' despawn it
                                                    ClearMapItem MapNpc(MapNum).Npc(x).Target, MapNum
                                                    ' send updates to everyone
                                                    SendMapItemsToAll MapNum
                                                    MapNpc(MapNum).Npc(x).Target = 0
                                                    MapNpc(MapNum).Npc(x).TargetType = 0
                                                End If
                                            Else
                                                MapNpc(MapNum).Npc(x).Target = 0
                                                MapNpc(MapNum).Npc(x).TargetType = 0
                                            End If
                                        Else
                                            MapNpc(MapNum).Npc(x).Target = 0
                                            MapNpc(MapNum).Npc(x).TargetType = 0
                                        End If
                                    End If
                                    Call NpcDir(MapNum, x, GetNpcDir(targetx, targety, CLng(MapNpc(MapNum).Npc(x).x), CLng(MapNpc(MapNum).Npc(x).y)), False, 0)
                                End If
                            Else
                                i = Int(Rnd * 4)
    
                                If i = 1 Then
                                    i = Int(Rnd * 4)
    
                                    If CanNpcMove(MapNum, x, i, False, 0) Then
                                        Call NpcMove(MapNum, x, i, MOVING_WALKING, False, 0)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then
                    Target = MapNpc(MapNum).Npc(x).Target
                    TargetType = MapNpc(MapNum).Npc(x).TargetType

                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                    
                        If TargetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                If Npc(MapNpc(MapNum).Npc(x).Num).Projectile = 0 Then
                                    TryNpcAttackPlayer x, Target
                                Else
                                    NPCFireProjectile MapNum, x, Npc(MapNpc(MapNum).Npc(x).Num).Projectile
                                End If
                            Else
                            ' Player left map or game, set target to 0
                                MapNpc(MapNum).Npc(x).Target = 0
                                MapNpc(MapNum).Npc(x).TargetType = 0 ' clear
                            End If
                        ElseIf TargetType = TARGET_TYPE_PET Then
                            If IsPlaying(Target) And GetPlayerMap(Target) = MapNum And Player(Target).characters(TempPlayer(Target).CurChar).Pet.Alive Then
                                If Npc(MapNpc(MapNum).Npc(x).Num).Projectile = 0 Then
                                    TryNpcAttackPet x, Target
                                Else
                                    NPCFireProjectile MapNum, x, Npc(MapNpc(MapNum).Npc(x).Num).Projectile
                                End If
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(MapNum).Npc(x).Target = 0
                                MapNpc(MapNum).Npc(x).TargetType = 0 ' clear
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(MapNum).Npc(x).stopRegen Then
                    If MapNpc(MapNum).Npc(x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(MapNum).Npc(x).Vital(Vitals.HP) > 0 Then
                            MapNpc(MapNum).Npc(x).Vital(Vitals.HP) = MapNpc(MapNum).Npc(x).Vital(Vitals.HP) + GetNpcVitalRegen(npcnum, Vitals.HP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(MapNum).Npc(x).Vital(Vitals.HP) > GetNpcMaxVital(npcnum, Vitals.HP) Then
                                MapNpc(MapNum).Npc(x).Vital(Vitals.HP) = GetNpcMaxVital(npcnum, Vitals.HP)
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(MapNum).Npc(x).Num = 0 And Map(MapNum).Npc(x) > 0 Then
                    If TickCount > MapNpc(MapNum).Npc(x).SpawnWait + (Npc(Map(MapNum).Npc(x)).SpawnSecs * 1000) Then
                        Call SpawnNpc(x, MapNum)
                    End If
                End If
                
                

            Next
            
            For x = 1 To Player_HighIndex
                If GetPlayerMap(x) = MapNum Then
                    If Player(x).characters(TempPlayer(x).CurChar).Pet.Alive = True Then
                            ' /////////////////////////////////////////
                            ' // This is used for ATTACKING ON SIGHT //
                            ' /////////////////////////////////////////
        
                            ' If the npc is a attack on sight, search for a player on the map
                            If Player(x).characters(TempPlayer(x).CurChar).Pet.AttackBehaviour <> PET_ATTACK_BEHAVIOUR_DONOTHING Then
                            
                                ' make sure it's not stunned
                                If Not TempPlayer(x).PetStunDuration > 0 Then
            
                                    For i = 1 To Player_HighIndex
                                        If TempPlayer(x).PetTargetType > 0 Then
                                            If TempPlayer(x).PetTargetType = 1 And TempPlayer(x).PetTarget = x Then
                                            
                                            Else
                                                Exit For
                                            End If
                                        End If
                                        If IsPlaying(i) And i <> x Then
                                            If GetPlayerMap(i) = MapNum And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                                If Player(x).characters(TempPlayer(x).CurChar).Pet.Alive Then
                                                    n = Pet(Player(x).characters(TempPlayer(x).CurChar).Pet.Num).Range
                                                    DistanceX = Player(x).characters(TempPlayer(x).CurChar).Pet.x - Player(i).characters(TempPlayer(i).CurChar).Pet.x
                                                    DistanceY = Player(x).characters(TempPlayer(x).CurChar).Pet.y - Player(i).characters(TempPlayer(i).CurChar).Pet.y
                    
                                                    ' Make sure we get a positive value
                                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                    
                                                    ' Are they in range?  if so GET'M!
                                                    If DistanceX <= n And DistanceY <= n Then
                                                        If Player(x).characters(TempPlayer(x).CurChar).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT Then
                                                            TempPlayer(x).PetTargetType = TARGET_TYPE_PET ' pet
                                                            TempPlayer(x).PetTarget = i
                                                            TempPlayer(x).PetTargetZone = 0
                                                        End If
                                                    End If
                                                Else
                                                    n = Pet(Player(x).characters(TempPlayer(x).CurChar).Pet.Num).Range
                                                    DistanceX = Player(x).characters(TempPlayer(x).CurChar).Pet.x - GetPlayerX(i)
                                                    DistanceY = Player(x).characters(TempPlayer(x).CurChar).Pet.y - GetPlayerY(i)
                    
                                                    ' Make sure we get a positive value
                                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                    
                                                    ' Are they in range?  if so GET'M!
                                                    If DistanceX <= n And DistanceY <= n Then
                                                        If Player(x).characters(TempPlayer(x).CurChar).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT Then
                                                                TempPlayer(x).PetTargetType = 1 ' player
                                                                TempPlayer(x).PetTarget = i
                                                                TempPlayer(x).PetTargetZone = 0
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                    
                                    If TempPlayer(x).PetTargetType = 0 Then
                                        For i = 1 To MAX_MAP_NPCS
                                            If TempPlayer(x).PetTargetType > 0 Then Exit For
                                            If Player(x).characters(TempPlayer(x).CurChar).Pet.Alive Then
                                                n = Pet(Player(x).characters(TempPlayer(x).CurChar).Pet.Num).Range
                                                DistanceX = Player(x).characters(TempPlayer(x).CurChar).Pet.x - MapNpc(GetPlayerMap(x)).Npc(i).x
                                                DistanceY = Player(x).characters(TempPlayer(x).CurChar).Pet.y - MapNpc(GetPlayerMap(x)).Npc(i).y
                    
                                                ' Make sure we get a positive value
                                                If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                If DistanceY < 0 Then DistanceY = DistanceY * -1
                    
                                                ' Are they in range?  if so GET'M!
                                                If DistanceX <= n And DistanceY <= n Then
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT Then
                                                        TempPlayer(x).PetTargetType = 2 ' npc
                                                        TempPlayer(x).PetTarget = i
                                                        TempPlayer(x).PetTargetZone = 0
                                                    End If
                                                End If
                                            End If
                                        Next
                                    End If
                                    
                                    If TempPlayer(x).PetTargetType = 0 Then
                                        For i = 1 To MAX_ZONES
                                            If TempPlayer(x).PetTargetType > 0 Then Exit For
                                            For z = 1 To MAX_MAP_NPCS * 2
                                                If TempPlayer(x).PetTargetType > 0 Then Exit For
                                                If ZoneNpc(i).Npc(z).Num > 0 Then
                                                    If ZoneNpc(i).Npc(z).Vital(HP) > 0 Then
                                                        If ZoneNpc(i).Npc(z).Map = GetPlayerMap(x) Then
                                                            If Player(x).characters(TempPlayer(x).CurChar).Pet.Alive Then
                                                                n = Pet(Player(x).characters(TempPlayer(x).CurChar).Pet.Num).Range
                                                                DistanceX = Player(x).characters(TempPlayer(x).CurChar).Pet.x - ZoneNpc(i).Npc(z).x
                                                                DistanceY = Player(x).characters(TempPlayer(x).CurChar).Pet.y - ZoneNpc(i).Npc(z).y
                                    
                                                                ' Make sure we get a positive value
                                                                If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                                If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                                                ' Are they in range?  if so GET'M!
                                                                If DistanceX <= n And DistanceY <= n Then
                                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT Then
                                                                        TempPlayer(x).PetTargetType = TARGET_TYPE_ZONENPC ' npc
                                                                        TempPlayer(x).PetTarget = z
                                                                        TempPlayer(x).PetTargetZone = i
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        Next
                                    End If
                                End If
                            End If
            
                        target_verify = False
                        ' /////////////////////////////////////////////
                        ' // This is used for Pet walking/targetting //
                        ' /////////////////////////////////////////////
                        ' Make sure theres a npc with the map
                            If TempPlayer(x).PetStunDuration > 0 Then
                                ' check if we can unstun them
                                If GetTickCount > TempPlayer(x).PetStunTimer + (TempPlayer(x).PetStunDuration * 1000) Then
                                    TempPlayer(x).PetStunDuration = 0
                                    TempPlayer(x).PetStunTimer = 0
                                End If
                            Else
                                Target = TempPlayer(x).PetTarget
                                TargetType = TempPlayer(x).PetTargetType
                                TargetZone = TempPlayer(x).PetTargetZone
                                ' Check to see if its time for the npc to walk
                                If Player(x).characters(TempPlayer(x).CurChar).Pet.AttackBehaviour <> PET_ATTACK_BEHAVIOUR_DONOTHING Then
                                
                                    If TargetType = 1 Then ' player
            
                                        ' Check to see if we are following a player or not
                                        If Target > 0 Then
                
                                            ' Check if the player is even playing, if so follow'm
                                            If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                                If Target <> x Then
                                                    didwalk = False
                                                    target_verify = True
                                                    targety = GetPlayerY(Target)
                                                    targetx = GetPlayerX(Target)
                                                End If
                                            Else
                                                TempPlayer(x).PetTargetType = 0 ' clear
                                                TempPlayer(x).PetTarget = 0
                                            End If
                                        End If
                                    
                                    ElseIf TargetType = 2 Then 'npc
                                        
                                        If Target > 0 Then
                                            
                                            If MapNpc(MapNum).Npc(Target).Num > 0 Then
                                                didwalk = False
                                                target_verify = True
                                                targety = MapNpc(MapNum).Npc(Target).y
                                                targetx = MapNpc(MapNum).Npc(Target).x
                                            Else
                                                TempPlayer(x).PetTargetType = 0 ' clear
                                                TempPlayer(x).PetTarget = 0
                                            End If
                                        End If
                                    ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                                        If Target > 0 Then
                                            
                                            If ZoneNpc(TargetZone).Npc(Target).Num > 0 And ZoneNpc(TargetZone).Npc(Target).Map = MapNum Then
                                                didwalk = False
                                                target_verify = True
                                                targety = ZoneNpc(TargetZone).Npc(Target).y
                                                targetx = ZoneNpc(TargetZone).Npc(Target).x
                                            Else
                                                TempPlayer(x).PetTargetType = 0 ' clear
                                                TempPlayer(x).PetTarget = 0
                                                TempPlayer(x).PetTargetZone = 0
                                            End If
                                        End If
                                    ElseIf TargetType = TARGET_TYPE_PET Then 'other pet
                                        If Target > 0 Then
                                            
                                            If IsPlaying(Target) And GetPlayerMap(Target) = MapNum And Player(Target).characters(TempPlayer(Target).CurChar).Pet.Alive Then
                                                didwalk = False
                                                target_verify = True
                                                targety = Player(Target).characters(TempPlayer(Target).CurChar).Pet.y
                                                targetx = Player(Target).characters(TempPlayer(Target).CurChar).Pet.x
                                            Else
                                                TempPlayer(x).PetTargetType = 0 ' clear
                                                TempPlayer(x).PetTarget = 0
                                            End If
                                        End If
                                    End If
                                End If
                                    
                                If target_verify Then
                                    didwalk = False
                                    If IsOneBlockAway(Player(x).characters(TempPlayer(x).CurChar).Pet.x, Player(x).characters(TempPlayer(x).CurChar).Pet.y, targetx, targety) Then
                                        If Player(x).characters(TempPlayer(x).CurChar).Pet.x < targetx Then
                                            PetDir x, DIR_RIGHT
                                            didwalk = True
                                        ElseIf Player(x).characters(TempPlayer(x).CurChar).Pet.x > targetx Then
                                            PetDir x, DIR_LEFT
                                            didwalk = True
                                        ElseIf Player(x).characters(TempPlayer(x).CurChar).Pet.y < targety Then
                                            PetDir x, DIR_UP
                                            didwalk = True
                                        ElseIf Player(x).characters(TempPlayer(x).CurChar).Pet.y > targety Then
                                            PetDir x, DIR_DOWN
                                            didwalk = True
                                        End If
                                    Else
                                        didwalk = PetTryWalk(x, targetx, targety)
                                    End If
                                ElseIf TempPlayer(x).PetBehavior = PET_BEHAVIOUR_GOTO And target_verify = False Then
                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x = TempPlayer(x).GoToX And Player(x).characters(TempPlayer(x).CurChar).Pet.y = TempPlayer(x).GoToY Then
                                        'Unblock these for the random turning
'                                        i = Int(Rnd * 4)
                                        'Call PetDir(x, i)
                                    Else
                                        didwalk = False
                                        targetx = TempPlayer(x).GoToX
                                        targety = TempPlayer(x).GoToY
                                        didwalk = PetTryWalk(x, targetx, targety)
                                            
                                        If didwalk = False Then
                                            i = Int(Rnd * 4)
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
                                                If CanPetMove(MapNum, x, i) Then
                                                    Call PetMove(MapNum, x, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
                                ElseIf TempPlayer(x).PetBehavior = PET_BEHAVIOUR_FOLLOW Then
                                    If IsPetByPlayer(x) Then
                                        'Unblock these to enable random turning
                                        'i = Int(Rnd * 4)
                                        'Call PetDir(x, i)
                                    Else
                                        didwalk = False
                                        targetx = GetPlayerX(x)
                                        targety = GetPlayerY(x)
                                        didwalk = PetTryWalk(x, targetx, targety)
                                        If didwalk = False Then
                                            i = Int(Rnd * 4)
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
                                                If CanPetMove(MapNum, x, i) Then
                                                    Call PetMove(MapNum, x, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            
                             ' /////////////////////////////////////////////
                            ' // This is used for pets to attack targets //
                            ' /////////////////////////////////////////////
                            ' Make sure theres a npc with the map
                                Target = TempPlayer(x).PetTarget
                                TargetType = TempPlayer(x).PetTargetType
                                TargetZone = TempPlayer(x).PetTargetZone
            
                                ' Check if the npc can attack the targeted player player
                                If Target > 0 Then
                                
                                    If TargetType = 1 Then ' player
                                        ' Is the target playing and on the same map?
                                        If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                            If x <> Target Then TryPetAttackPlayer x, Target
                                        Else
                                            ' Player left map or game, set target to 0
                                            TempPlayer(x).PetTarget = 0
                                            TempPlayer(x).PetTargetType = 0 ' clear
                                        End If
                                    ElseIf TargetType = 2 Then 'npc
                                        If MapNpc(GetPlayerMap(x)).Npc(TempPlayer(x).PetTarget).Num > 0 Then
                                           Call TryPetAttackNpc(x, TempPlayer(x).PetTarget)
                                        Else
                                            ' Player left map or game, set target to 0
                                            TempPlayer(x).PetTarget = 0
                                            TempPlayer(x).PetTargetType = 0 ' clear
                                        End If
                                    ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                                        If ZoneNpc(TargetZone).Npc(TempPlayer(x).PetTarget).Num > 0 And ZoneNpc(TargetZone).Npc(TempPlayer(x).PetTarget).Map = GetPlayerMap(x) Then
                                           Call TryPetAttackZoneNpc(x, TempPlayer(x).PetTargetZone, TempPlayer(x).PetTarget)
                                        Else
                                            ' Player left map or game, set target to 0
                                            TempPlayer(x).PetTarget = 0
                                            TempPlayer(x).PetTargetType = 0 ' clear
                                            TempPlayer(x).PetTargetZone = 0
                                        End If
                                    ElseIf TargetType = TARGET_TYPE_PET Then 'pet
                                        ' Is the target playing and on the same map? And is pet alive??
                                        If IsPlaying(Target) And GetPlayerMap(Target) = MapNum And Player(Target).characters(TempPlayer(Target).CurChar).Pet.Alive = True Then
                                            TryPetAttackPet x, Target
                                        Else
                                            ' Player left map or game, set target to 0
                                            TempPlayer(x).PetTarget = 0
                                            TempPlayer(x).PetTargetType = 0 ' clear
                                        End If
                                    End If
                                End If
                                                
                                ' ////////////////////////////////////////////
                                ' // This is used for regenerating PET's HP //
                                ' ////////////////////////////////////////////
                                ' Check to see if we want to regen some of the npc's hp
                                If Not TempPlayer(x).PetstopRegen Then
                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.Alive = True And TickCount > GivePetHPTimer + 10000 Then
                                        If Player(x).characters(TempPlayer(x).CurChar).Pet.Health > 0 Then
                                            Player(x).characters(TempPlayer(x).CurChar).Pet.Health = Player(x).characters(TempPlayer(x).CurChar).Pet.Health + GetPetVitalRegen(x, Vitals.HP)
                                            Player(x).characters(TempPlayer(x).CurChar).Pet.Mana = Player(x).characters(TempPlayer(x).CurChar).Pet.Mana + GetPetVitalRegen(x, Vitals.MP)
                                            ' Check if they have more then they should and if so just set it to max
                                            If Player(x).characters(TempPlayer(x).CurChar).Pet.Health > GetPetMaxVital(x, HP) Then
                                                Player(x).characters(TempPlayer(x).CurChar).Pet.Health = GetPetMaxVital(x, HP)
                                            End If
                                            If Player(x).characters(TempPlayer(x).CurChar).Pet.Mana > GetPetMaxVital(x, MP) Then
                                                Player(x).characters(TempPlayer(x).CurChar).Pet.Mana = GetPetMaxVital(x, MP)
                                            End If
                                            Call SendPetVital(x, HP)
                                            Call SendPetVital(x, MP)
                                        End If
                                    End If
                                End If
                        End If
                    End If
                Next

         DoEvents

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateMapLogic", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function FindOpenNpcInvSlot(MapNum As Long, npcnum As Long, ZoneNum As Long) As Long
Dim i As Long, x As Long

   On Error GoTo errorhandler

    If ZoneNum = 0 Then
        For i = 1 To 20
            If MapNpc(MapNum).Npc(npcnum).Inventory(i).Num = 0 Then
                FindOpenNpcInvSlot = i
            End If
        Next
    Else
        For i = 1 To 20
            If ZoneNpc(ZoneNum).Npc(npcnum).Inventory(i).Num = 0 Then
                FindOpenNpcInvSlot = i
            End If
        Next
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindOpenNpcInvSlot", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Private Sub UpdateZoneLogic()
    Dim i As Long, x As Long, n As Long, x1 As Long, y1 As Long, MapNum As Long, j As Long, xOffset As Long, yOffset As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, npcnum As Long
    Dim Target As Long, TargetType As Byte, didwalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim targetx As Long, targety As Long, target_verify As Boolean, ZoneNum As Long, d As Long, z As Long
    Dim shortestZone As Long
    Dim shortestNum As Long
    Dim shortestD As Long, invslot As Long


   On Error GoTo errorhandler

    For ZoneNum = 1 To MAX_ZONES
        If MapZones(ZoneNum).WeatherTimer < GetTickCount Then
            MapZones(ZoneNum).WeatherTimer = GetTickCount + 1800000 + Random(60000, 600000)
            If MapZones(ZoneNum).CurrentWeather > 0 Then
                MapZones(ZoneNum).CurrentWeather = 0
                If MapZones(ZoneNum).MapCount > 0 Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            For x = 1 To MapZones(ZoneNum).MapCount
                                If GetPlayerMap(i) = MapZones(ZoneNum).Maps(x) Then
                                    PlayerMsg i, "The weather clears up around you.", BrightBlue
                                End If
                            Next
                        End If
                    Next
                End If
            Else
                i = Random(1, 2)
                If i > 2 Then i = 2
                If i < 1 Then i = 1
                If i = 2 Then
                    For i = 1 To 5
                        If MapZones(ZoneNum).Weather(i) > 0 Then
                            x = Random(1, 100)
                            If x < 1 Then x = 1
                            If x > 100 Then x = 100
                            If x < MapZones(ZoneNum).Weather(i) Then
                                MapZones(ZoneNum).CurrentWeather = i
                                If MapZones(ZoneNum).MapCount > 0 Then
                                    For z = 1 To Player_HighIndex
                                        If IsPlaying(z) Then
                                            For x = 1 To MapZones(ZoneNum).MapCount
                                                If GetPlayerMap(z) = MapZones(ZoneNum).Maps(x) Then
                                                    Select Case i
                                                        Case 1
                                                            PlayerMsg z, "Rain starts falling from the sky!", BrightBlue
                                                        Case 2
                                                            PlayerMsg z, "Snow begins to fall to the ground.", BrightBlue
                                                        Case 3
                                                            PlayerMsg z, "Look out! Hail is falling from the sky!", BrightBlue
                                                        Case 4
                                                            PlayerMsg z, "A sand-storm starts to brew...", BrightBlue
                                                        Case 5
                                                            PlayerMsg z, "The sky darkens as a storm approaches!", BrightBlue
                                                    End Select
                                                    'Send the Zone Weather
                                                    SendSpecialEffect z, EFFECT_TYPE_WEATHER, CByte(MapZones(ZoneNum).CurrentWeather), CLng(MapZones(ZoneNum).WeatherIntensity)
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                                i = 5
                            End If
                        End If
                    Next
                End If
            End If
            MapZones(ZoneNum).WeatherTimer = GetTickCount + 1800000 + Random(60000, 600000)
        End If
    
    
        For x = 1 To MAX_MAP_NPCS * 2
            
            npcnum = MapZones(ZoneNum).NPCs(x)
            MapNum = ZoneNpc(ZoneNum).Npc(x).Map
            'If PlayersOnMap(mapNum) = YES Then
                TickCount = GetTickCount
            
            

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapZones(ZoneNum).NPCs(x) > 0 And MapZones(ZoneNum).NPCs(x) > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(npcnum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(npcnum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not ZoneNpc(ZoneNum).Npc(x).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum And ZoneNpc(ZoneNum).Npc(x).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive Then
                                            n = Npc(npcnum).Range
                                            DistanceX = ZoneNpc(ZoneNum).Npc(x).x - Player(i).characters(TempPlayer(i).CurChar).Pet.x
                                            DistanceY = ZoneNpc(ZoneNum).Npc(x).y - Player(i).characters(TempPlayer(i).CurChar).Pet.y
        
                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                            If DistanceY < 0 Then DistanceY = DistanceY * -1
        
                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n And DistanceY <= n Then
                                                If Npc(npcnum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    If Len(Trim$(Npc(npcnum).AttackSay)) > 0 Then
                                                        Call PlayerMsg(i, Trim$(Npc(npcnum).Name) & " says: " & Trim$(Npc(npcnum).AttackSay), SayColor)
                                                    End If
                                                    ZoneNpc(ZoneNum).Npc(x).TargetType = TARGET_TYPE_PET
                                                    ZoneNpc(ZoneNum).Npc(x).Target = i
                                                End If
                                            End If
                                        Else
                                            n = Npc(npcnum).Range
                                            DistanceX = ZoneNpc(ZoneNum).Npc(x).x - GetPlayerX(i)
                                            DistanceY = ZoneNpc(ZoneNum).Npc(x).y - GetPlayerY(i)
        
                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                            If DistanceY < 0 Then DistanceY = DistanceY * -1
        
                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n And DistanceY <= n Then
                                                If Npc(npcnum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    If Len(Trim$(Npc(npcnum).AttackSay)) > 0 Then
                                                        Call PlayerMsg(i, Trim$(Npc(npcnum).Name) & " says: " & Trim$(Npc(npcnum).AttackSay), SayColor)
                                                    End If
                                                    ZoneNpc(ZoneNum).Npc(x).TargetType = 1 ' player
                                                    ZoneNpc(ZoneNum).Npc(x).Target = i
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapZones(ZoneNum).NPCs(x) > 0 And ZoneNpc(ZoneNum).Npc(x).Num > 0 Then
                    If ZoneNpc(ZoneNum).Npc(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > ZoneNpc(ZoneNum).Npc(x).StunTimer + (ZoneNpc(ZoneNum).Npc(x).StunDuration * 1000) Then
                            ZoneNpc(ZoneNum).Npc(x).StunDuration = 0
                            ZoneNpc(ZoneNum).Npc(x).StunTimer = 0
                        End If
                    Else
                            
                        Target = ZoneNpc(ZoneNum).Npc(x).Target
                        TargetType = ZoneNpc(ZoneNum).Npc(x).TargetType
    
                        ' Check to see if its time for the npc to walk
                        If Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        
                            If TargetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If Target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                        didwalk = False
                                        target_verify = True
                                        targety = GetPlayerY(Target)
                                        targetx = GetPlayerX(Target)
                                    Else
                                        If IsPlaying(Target) Then
                                            If Map(ZoneNpc(ZoneNum).Npc(x).Map).Up = GetPlayerMap(Target) Or Map(ZoneNpc(ZoneNum).Npc(x).Map).Down = GetPlayerMap(Target) Or Map(ZoneNpc(ZoneNum).Npc(x).Map).Left = GetPlayerMap(Target) Or Map(ZoneNpc(ZoneNum).Npc(x).Map).Right = GetPlayerMap(Target) Then
                                                If MapZones(ZoneNum).MapCount > 0 Then
                                                    For j = 1 To MapZones(ZoneNum).MapCount
                                                        If MapZones(ZoneNum).Maps(j) = GetPlayerMap(Target) Then
                                                            If Map(ZoneNpc(ZoneNum).Npc(x).Map).Up = GetPlayerMap(Target) Then
                                                                targety = -1
                                                                targetx = GetPlayerX(Target)
                                                                yOffset = 1
                                                                xOffset = 0
                                                            End If
                                                            If Map(ZoneNpc(ZoneNum).Npc(x).Map).Down = GetPlayerMap(Target) Then
                                                                targety = Map(ZoneNpc(ZoneNum).Npc(x).Map).MaxY + 1
                                                                targetx = GetPlayerX(Target)
                                                                yOffset = -1
                                                                xOffset = 0
                                                            End If
                                                            If Map(ZoneNpc(ZoneNum).Npc(x).Map).Left = GetPlayerMap(Target) Then
                                                                targetx = -1
                                                                targety = GetPlayerY(Target)
                                                                xOffset = 1
                                                                yOffset = 0
                                                            End If
                                                            If Map(ZoneNpc(ZoneNum).Npc(x).Map).Right = GetPlayerMap(Target) Then
                                                                targetx = Map(ZoneNpc(ZoneNum).Npc(x).Map).MaxX + 1
                                                                targety = GetPlayerY(Target)
                                                                xOffset = -1
                                                                yOffset = 0
                                                            End If
                                                            didwalk = False
                                                            target_verify = True
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        Else
                                            ZoneNpc(ZoneNum).Npc(x).TargetType = 0 ' clear
                                            ZoneNpc(ZoneNum).Npc(x).Target = 0
                                        End If
                                    End If
                                End If
                            
                            ElseIf TargetType = 2 Then 'npc
                                
                                If Target > 0 Then
                                    
                                    If MapNpc(MapNum).Npc(Target).Num > 0 Then
                                        didwalk = False
                                        target_verify = True
                                        targety = MapNpc(MapNum).Npc(Target).y
                                        targetx = MapNpc(MapNum).Npc(Target).x
                                    Else
                                        ZoneNpc(ZoneNum).Npc(x).TargetType = 0 ' clear
                                        ZoneNpc(ZoneNum).Npc(x).Target = 0
                                    End If
                                End If
                            ElseIf TargetType = 6 Then
                                
                                If Target > 0 Then
                                    
                                    If MapItem(MapNum, Target).Num > 0 Then
                                        didwalk = False
                                        target_verify = True
                                        targety = MapItem(MapNum, Target).y
                                        targetx = MapItem(MapNum, Target).x
                                    Else
                                        ZoneNpc(ZoneNum).Npc(x).TargetType = 0 ' clear
                                        ZoneNpc(ZoneNum).Npc(x).Target = 0
                                    End If
                                End If
                            ElseIf TargetType = TARGET_TYPE_PET Then 'PET
                                If Target > 0 Then
                                    
                                    If IsPlaying(Target) = True And GetPlayerMap(Target) = MapNum And Player(Target).characters(TempPlayer(Target).CurChar).Pet.Alive = True Then
                                        didwalk = False
                                        target_verify = True
                                        targety = Player(Target).characters(TempPlayer(Target).CurChar).Pet.y
                                        targetx = Player(Target).characters(TempPlayer(Target).CurChar).Pet.x
                                    Else
                                        ZoneNpc(ZoneNum).Npc(x).TargetType = 0 ' clear
                                        ZoneNpc(ZoneNum).Npc(x).Target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                'Gonna make the npcs smarter.. Implementing a pathfinding algorithm.. we shall see what happens.
                                If IsOneBlockAway(targetx + xOffset, targety + yOffset, CLng(ZoneNpc(ZoneNum).Npc(x).x), CLng(ZoneNpc(ZoneNum).Npc(x).y)) = False Then
                                    If PathfindingType = 1 Then
                                        i = Int(Rnd * 5)
            
                                        ' Lets move the npc
                                        Select Case i
                                            Case 0
            
                                                ' Up
                                                If ZoneNpc(ZoneNum).Npc(x).y > targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_UP, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If ZoneNpc(ZoneNum).Npc(x).y < targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_DOWN, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If ZoneNpc(ZoneNum).Npc(x).x > targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_LEFT, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If ZoneNpc(ZoneNum).Npc(x).x < targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_RIGHT, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 1
            
                                                ' Right
                                                If ZoneNpc(ZoneNum).Npc(x).x < targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_RIGHT, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If ZoneNpc(ZoneNum).Npc(x).x > targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_LEFT, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If ZoneNpc(ZoneNum).Npc(x).y < targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_DOWN, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If ZoneNpc(ZoneNum).Npc(x).y > targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_UP, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 2
            
                                                ' Down
                                                If ZoneNpc(ZoneNum).Npc(x).y < targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_DOWN, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If ZoneNpc(ZoneNum).Npc(x).y > targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_UP, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If ZoneNpc(ZoneNum).Npc(x).x < targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_RIGHT, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If ZoneNpc(ZoneNum).Npc(x).x > targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_LEFT, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 3
            
                                                ' Left
                                                If ZoneNpc(ZoneNum).Npc(x).x > targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_LEFT, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If ZoneNpc(ZoneNum).Npc(x).x < targetx And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_RIGHT, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If ZoneNpc(ZoneNum).Npc(x).y > targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_UP, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If ZoneNpc(ZoneNum).Npc(x).y < targety And Not didwalk Then
                                                    If CanNpcMove(MapNum, x, DIR_DOWN, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING, True, ZoneNum)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                        End Select
            
                                        ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                        If Not didwalk Then
                                            If ZoneNpc(ZoneNum).Npc(x).x - 1 = targetx And ZoneNpc(ZoneNum).Npc(x).y = targety Then
                                                If ZoneNpc(ZoneNum).Npc(x).Dir <> DIR_LEFT Then
                                                    Call NpcDir(MapNum, x, DIR_LEFT, True, ZoneNum)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If ZoneNpc(ZoneNum).Npc(x).x + 1 = targetx And ZoneNpc(ZoneNum).Npc(x).y = targety Then
                                                If ZoneNpc(ZoneNum).Npc(x).Dir <> DIR_RIGHT Then
                                                    Call NpcDir(MapNum, x, DIR_RIGHT, True, ZoneNum)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If ZoneNpc(ZoneNum).Npc(x).x = targetx And ZoneNpc(ZoneNum).Npc(x).y - 1 = targety Then
                                                If ZoneNpc(ZoneNum).Npc(x).Dir <> DIR_UP Then
                                                    Call NpcDir(MapNum, x, DIR_UP, True, ZoneNum)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If ZoneNpc(ZoneNum).Npc(x).x = targetx And ZoneNpc(ZoneNum).Npc(x).y + 1 = targety Then
                                                If ZoneNpc(ZoneNum).Npc(x).Dir <> DIR_DOWN Then
                                                    Call NpcDir(MapNum, x, DIR_DOWN, True, ZoneNum)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            ' We could not move so Target must be behind something, walk randomly.
                                            If Not didwalk Then
                                                i = Int(Rnd * 2)
            
                                                If i = 1 Then
                                                    i = Int(Rnd * 4)
            
                                                    If CanNpcMove(MapNum, x, i, True, ZoneNum) Then
                                                        Call NpcMove(MapNum, x, i, MOVING_WALKING, True, ZoneNum)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        i = FindNpcPath(MapNum, x, targetx + xOffset, targety + yOffset, True, ZoneNum)
                                        If i < 4 Then 'Returned an answer. Move the NPC
                                            If CanNpcMove(MapNum, x, i, True, ZoneNum) Then
                                                NpcMove MapNum, x, i, MOVING_WALKING, True, ZoneNum
                                            End If
                                        Else 'No good path found. Move randomly
                                            i = Int(Rnd * 4)
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
                
                                                If CanNpcMove(MapNum, x, i, True, ZoneNum) Then
                                                    Call NpcMove(MapNum, x, i, MOVING_WALKING, True, ZoneNum)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    If ZoneNpc(ZoneNum).Npc(x).TargetType = TARGET_TYPE_ITEM Then
                                        If ZoneNpc(ZoneNum).Npc(x).Target > 0 Then
                                            If MapItem(MapNum, ZoneNpc(ZoneNum).Npc(x).Target).Num > 0 Then
                                                'NPC Needs to take item...
                                                If FindOpenNpcInvSlot(0, x, ZoneNum) > 0 Then
                                                    invslot = FindOpenNpcInvSlot(0, x, ZoneNum)
                                                    ZoneNpc(ZoneNum).Npc(x).Inventory(invslot).Num = MapItem(MapNum, ZoneNpc(ZoneNum).Npc(x).Target).Num
                                                    ZoneNpc(ZoneNum).Npc(x).Inventory(invslot).Value = MapItem(MapNum, ZoneNpc(ZoneNum).Npc(x).Target).Value
                                                    ' despawn it
                                                    ClearMapItem ZoneNpc(ZoneNum).Npc(x).Target, MapNum
                                                    ' send updates to everyone
                                                    SendMapItemsToAll MapNum
                                                    ZoneNpc(ZoneNum).Npc(x).Target = 0
                                                    ZoneNpc(ZoneNum).Npc(x).TargetType = 0
                                                End If
                                            Else
                                                ZoneNpc(ZoneNum).Npc(x).Target = 0
                                                ZoneNpc(ZoneNum).Npc(x).TargetType = 0
                                            End If
                                        Else
                                            ZoneNpc(ZoneNum).Npc(x).Target = 0
                                            ZoneNpc(ZoneNum).Npc(x).TargetType = 0
                                        End If
                                    Else
                                        If targetx < ZoneNpc(ZoneNum).Npc(x).x Then
                                            Call NpcDir(MapNum, x, DIR_LEFT, True, ZoneNum)
                                        ElseIf targetx > ZoneNpc(x).Npc(x).x Then
                                            Call NpcDir(MapNum, x, DIR_RIGHT, True, ZoneNum)
                                        ElseIf targety < ZoneNpc(ZoneNum).Npc(x).y Then
                                            Call NpcDir(MapNum, x, DIR_UP, True, ZoneNum)
                                        ElseIf targety > ZoneNpc(ZoneNum).Npc(x).y Then
                                            Call NpcDir(MapNum, x, DIR_DOWN, True, ZoneNum)
                                        End If
                                    End If
                                    If targetx < 0 Then
                                        Call CanNpcMove(MapNum, x, DIR_LEFT, True, ZoneNum)
                                    ElseIf targetx = Map(MapNum).MaxX + 1 Then
                                        Call CanNpcMove(MapNum, x, DIR_RIGHT, True, ZoneNum)
                                    End If
                                    If targety < 0 Then
                                        Call CanNpcMove(MapNum, x, DIR_UP, True, ZoneNum)
                                    ElseIf targety = Map(MapNum).MaxX + 1 Then
                                        Call CanNpcMove(MapNum, x, DIR_DOWN, True, ZoneNum)
                                    End If
                                End If
                            Else
                                i = Int(Rnd * 4)
    
                                If i = 1 Then
                                    i = Int(Rnd * 4)
    
                                    If CanNpcMove(MapNum, x, i, True, ZoneNum) Then
                                        Call NpcMove(MapNum, x, i, MOVING_WALKING, True, ZoneNum)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapZones(ZoneNum).NPCs(x) > 0 And ZoneNpc(ZoneNum).Npc(x).Num > 0 Then
                    Target = ZoneNpc(ZoneNum).Npc(x).Target
                    TargetType = ZoneNpc(ZoneNum).Npc(x).TargetType

                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                    
                        If TargetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                TryZoneNpcAttackPlayer ZoneNum, x, Target
                            Else
                                ' Player left map or game, set target to 0
                                'ZoneNpc(ZoneNum).NPC(x).target = 0
                                'ZoneNpc(ZoneNum).NPC(x).targetType = 0 ' clear
                            End If
                        ElseIf TargetType = TARGET_TYPE_PET Then
                            If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                TryZoneNpcAttackPet ZoneNum, x, Target
                            Else
                            
                            End If
                        Else
                            ' lol no npc combat :(
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not ZoneNpc(ZoneNum).Npc(x).stopRegen Then
                    If ZoneNpc(ZoneNum).Npc(x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If ZoneNpc(ZoneNum).Npc(x).Vital(Vitals.HP) > 0 Then
                            ZoneNpc(ZoneNum).Npc(x).Vital(Vitals.HP) = ZoneNpc(ZoneNum).Npc(x).Vital(Vitals.HP) + GetNpcVitalRegen(npcnum, Vitals.HP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If ZoneNpc(ZoneNum).Npc(x).Vital(Vitals.HP) > GetNpcMaxVital(npcnum, Vitals.HP) Then
                                ZoneNpc(ZoneNum).Npc(x).Vital(Vitals.HP) = GetNpcMaxVital(npcnum, Vitals.HP)
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                
            'End If
            
            If ZoneNpc(ZoneNum).Npc(x).Num = 0 And MapZones(ZoneNum).NPCs(x) > 0 Then
                If GetTickCount > ZoneNpc(ZoneNum).Npc(x).SpawnWait + (Npc(MapZones(ZoneNum).NPCs(x)).SpawnSecs * 1000) Then
                    Call SpawnZoneNpc(ZoneNum, x)
                End If
            End If

        Next

        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateZoneLogic", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Private Sub UpdatePlayerVitals(ByVal Index As Long)


   On Error GoTo errorhandler

    If Not TempPlayer(Index).stopRegen Then
        If GetPlayerVital(Index, Vitals.HP) <> GetPlayerMaxVital(Index, Vitals.HP) Then
            Call SetPlayerVital(Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + GetPlayerVitalRegen(Index, Vitals.HP))
            Call SendVital(Index, Vitals.HP)
            ' send vitals to party if in one
            If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
        End If
        
        If GetPlayerVital(Index, Vitals.MP) <> GetPlayerMaxVital(Index, Vitals.MP) Then
            Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + GetPlayerVitalRegen(Index, Vitals.MP))
            Call SendVital(Index, Vitals.MP)
            ' send vitals to party if in one
            If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdatePlayerVitals", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub UpdateSavePlayers(ByVal Index As Long)
Dim i As Long

' Prevent subscript out range

   On Error GoTo errorhandler

If Not IsPlaying(Index) Then Exit Sub

' Save player
Call TextAdd("Saving all online players...")
Call SavePlayer(Index)
Call SaveBank(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateSavePlayers", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleShutdown()


   On Error GoTo errorhandler

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        If shutDownType = 0 Then
            Call GlobalMsg("Server Shutdown.", BrightRed)
            Call DestroyServer
        Else
            Call GlobalMsg("Server Shutdown.", BrightRed)
            Shell App.path & "\Eclipse Origins Server Launcher.exe", vbNormalFocus
            Call DestroyServer
            End
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleShutdown", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function CanEventMoveTowardsPlayer(playerID As Long, MapNum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long
Dim tim As Long, sX As Long, sY As Long, pos() As Long, reachable As Boolean, j As Long, LastSum As Long, Sum As Long, FX As Long, FY As Long
Dim path() As Vector, LastX As Long, LastY As Long, did As Boolean
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 4 is not a valid direction so we assume fail unless otherwise told.

   On Error GoTo errorhandler

    CanEventMoveTowardsPlayer = 4
    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    WalkThrough = Map(MapNum).Events(TempPlayer(playerID).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(playerID).EventMap.EventPages(eventID).pageID).WalkThrough
    'Add option for pathfinding to random guessing option.
    
    If PathfindingType = 1 Then
        i = Int(Rnd * 5)
        didwalk = False
        
        ' Lets move the event
        Select Case i
            Case 0
        
                ' Up
                If y1 > y And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        didwalk = True
                    End If
                End If
        
                ' Down
                If y1 < y And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        didwalk = True
                    End If
                End If
        
                ' Left
                If x1 > x And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        didwalk = True
                    End If
                End If
        
                ' Right
                If x1 < x And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        didwalk = True
                    End If
                End If
        
            Case 1
            
                ' Right
                If x1 < x And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Left
                If x1 > x And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Down
                If y1 < y And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Up
                If y1 > y And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        didwalk = True
                    End If
                End If
        
            Case 2
            
                ' Down
                If y1 < y And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Up
                If y1 > y And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Right
                If x1 < x And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Left
                If x1 > x And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        didwalk = True
                    End If
                End If
        
            Case 3
            
                ' Left
                If x1 > x And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Right
                If x1 < x And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Up
                If y1 > y And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Down
                If y1 < y And Not didwalk Then
                    If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        didwalk = True
                    End If
                End If
        End Select
        CanEventMoveTowardsPlayer = Random(0, 3)
    ElseIf PathfindingType = 2 Then
        'Initialization phase
        tim = 0
        sX = x1
        sY = y1
        FX = x
        FY = y
        
        ReDim pos(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
        
        'CacheMapBlocks mapnum
        
        pos = MapBlocks(MapNum).Blocks
        
        For i = 1 To TempPlayer(playerID).EventMap.CurrentEvents
            If TempPlayer(playerID).EventMap.EventPages(i).Visible Then
                If TempPlayer(playerID).EventMap.EventPages(i).WalkThrough = 1 Then
                    pos(TempPlayer(playerID).EventMap.EventPages(i).x, TempPlayer(playerID).EventMap.EventPages(i).y) = 9
                End If
            End If
        Next
        
        pos(sX, sY) = 100 + tim
        pos(FX, FY) = 2
        
        'reset reachable
        reachable = False
        
        'Do while reachable is false... if its set true in progress, we jump out
        'If the path is decided unreachable in process, we will use exit sub. Not proper,
        'but faster ;-)
        Do While reachable = False
            'we loop through all squares
            For j = 0 To Map(MapNum).MaxY
                For i = 0 To Map(MapNum).MaxX
                    'If j = 10 And i = 0 Then MsgBox "hi!"
                    'If they are to be extended, the pointer TIM is on them
                    If pos(i, j) = 100 + tim Then
                    'The part is to be extended, so do it
                        'We have to make sure that there is a pos(i+1,j) BEFORE we actually use it,
                        'because then we get error... If the square is on side, we dont test for this one!
                        If i < Map(MapNum).MaxX Then
                            'If there isnt a wall, or any other... thing
                            If pos(i + 1, j) = 0 Then
                                'Expand it, and make its pos equal to tim+1, so the next time we make this loop,
                                'It will exapand that square too! This is crucial part of the program
                                pos(i + 1, j) = 100 + tim + 1
                            ElseIf pos(i + 1, j) = 2 Then
                                'If the position is no 0 but its 2 (FINISH) then Reachable = true!!! We found end
                                reachable = True
                            End If
                        End If
                    
                        'This is the same as the last one, as i said a lot of copy paste work and editing that
                        'This is simply another side that we have to test for... so instead of i+1 we have i-1
                        'Its actually pretty same then... I wont comment it therefore, because its only repeating
                        'same thing with minor changes to check sides
                        If i > 0 Then
                            If pos((i - 1), j) = 0 Then
                                pos(i - 1, j) = 100 + tim + 1
                            ElseIf pos(i - 1, j) = 2 Then
                                reachable = True
                            End If
                        End If
                    
                        If j < Map(MapNum).MaxY Then
                            If pos(i, j + 1) = 0 Then
                                pos(i, j + 1) = 100 + tim + 1
                            ElseIf pos(i, j + 1) = 2 Then
                                reachable = True
                            End If
                        End If
                    
                        If j > 0 Then
                            If pos(i, j - 1) = 0 Then
                                pos(i, j - 1) = 100 + tim + 1
                            ElseIf pos(i, j - 1) = 2 Then
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
                For j = 0 To Map(MapNum).MaxY
                    For i = 0 To Map(MapNum).MaxX
                    'we add up ALL the squares
                    Sum = Sum + pos(i, j)
                    Next i
                Next j
                
                'Now if the sum is euqal to the last sum, its not reachable, if it isnt, then we store
                'sum to lastsum
                If Sum = LastSum Then
                    CanEventMoveTowardsPlayer = 4
                    Exit Function
                Else
                    LastSum = Sum
                End If
            End If
            
            'we increase the pointer to point to the next squares to be expanded
            tim = tim + 1
        Loop
        
        'We work backwards to find the way...
        LastX = FX
        LastY = FY
        
        ReDim path(tim + 1)
        
        'The following code may be a little bit confusing but ill try my best to explain it.
        'We are working backwards to find ONE of the shortest ways back to Start.
        'So we repeat the loop until the LastX and LastY arent in start. Look in the code to see
        'how LastX and LasY change
        Do While LastX <> sX Or LastY <> sY
            'We decrease tim by one, and then we are finding any adjacent square to the final one, that
            'has that value. So lets say the tim would be 5, because it takes 5 steps to get to the target.
            'Now everytime we decrease that, so we make it 4, and we look for any adjacent square that has
            'that value. When we find it, we just color it yellow as for the solution
            tim = tim - 1
            'reset did to false
            did = False
            
            'If we arent on edge
            If LastX < Map(MapNum).MaxX Then
                'check the square on the right of the solution. Is it a tim-1 one? or just a blank one
                If pos(LastX + 1, LastY) = 100 + tim Then
                    'if it, then make it yellow, and change did to true
                    LastX = LastX + 1
                    did = True
                End If
            End If
            
            'This will then only work if the previous part didnt execute, and did is still false. THen
            'we want to check another square, the on left. Is it a tim-1 one ?
            If did = False Then
                If LastX > 0 Then
                    If pos(LastX - 1, LastY) = 100 + tim Then
                        LastX = LastX - 1
                        did = True
                    End If
                End If
            End If
            
            'We check the one below it
            If did = False Then
                If LastY < Map(MapNum).MaxY Then
                    If pos(LastX, LastY + 1) = 100 + tim Then
                        LastY = LastY + 1
                        did = True
                    End If
                End If
            End If
            
            'And above it. One of these have to be it, since we have found the solution, we know that already
            'there is a way back.
            If did = False Then
                If LastY > 0 Then
                    If pos(LastX, LastY - 1) = 100 + tim Then
                        LastY = LastY - 1
                    End If
                End If
            End If
            
            path(tim).x = LastX
            path(tim).y = LastY
            
            'Now we loop back and decrease tim, and look for the next square with lower value
            DoEvents
        Loop
        
        'Ok we got a path. Now, lets look at the first step and see what direction we should take.
        If path(1).x > LastX Then
            CanEventMoveTowardsPlayer = DIR_RIGHT
        ElseIf path(1).y > LastY Then
            CanEventMoveTowardsPlayer = DIR_DOWN
        ElseIf path(1).y < LastY Then
            CanEventMoveTowardsPlayer = DIR_UP
        ElseIf path(1).x < LastX Then
            CanEventMoveTowardsPlayer = DIR_LEFT
        End If
        
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanEventMoveTowardsPlayer", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function CanEventMoveAwayFromPlayer(playerID As Long, MapNum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.

   On Error GoTo errorhandler

    CanEventMoveAwayFromPlayer = 5
    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    WalkThrough = Map(MapNum).Events(TempPlayer(playerID).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(playerID).EventMap.EventPages(eventID).pageID).WalkThrough
    
    i = Int(Rnd * 5)
    didwalk = False
    
    ' Lets move the event
    Select Case i
        Case 0
    
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 1
        
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 2
        
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 3
        
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
    
        End Select
        
        CanEventMoveAwayFromPlayer = Random(0, 3)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanEventMoveAwayFromPlayer", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetDirToPlayer(playerID As Long, MapNum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long, distance As Long
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.

   On Error GoTo errorhandler

    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    
    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_RIGHT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_LEFT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If y - y1 > 0 Then
        If y - y1 > distance Then
            i = DIR_DOWN
            distance = y - y1
        End If
    ElseIf y - y1 < 0 Then
        If ((y - y1) * -1) > distance Then
            i = DIR_UP
            distance = ((y - y1) * -1)
        End If
    End If
    
    GetDirToPlayer = i


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetDirToPlayer", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
    
End Function

Function GetDirAwayFromPlayer(playerID As Long, MapNum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long, distance As Long
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.

   On Error GoTo errorhandler

    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    
    
    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_LEFT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_RIGHT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If y - y1 > 0 Then
        If y - y1 > distance Then
            i = DIR_UP
            distance = y - y1
        End If
    ElseIf y - y1 < 0 Then
        If ((y - y1) * -1) > distance Then
            i = DIR_DOWN
            distance = ((y - y1) * -1)
        End If
    End If
    
    GetDirAwayFromPlayer = i


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetDirAwayFromPlayer", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function FindNpcPath(MapNum As Long, mapnpcnum As Long, targetx As Long, targety As Long, IsZoneNpc As Boolean, ZoneNum As Long) As Long
Dim tim As Long, sX As Long, sY As Long, pos() As Long, reachable As Boolean, x As Long, y As Long, j As Long, LastSum As Long, Sum As Long, FX As Long, FY As Long, i As Long
Dim path() As Vector, LastX As Long, LastY As Long, did As Boolean

'Initialization phase

   On Error GoTo errorhandler

tim = 0
If IsZoneNpc Then
    sX = ZoneNpc(ZoneNum).Npc(mapnpcnum).x
    sY = ZoneNpc(ZoneNum).Npc(mapnpcnum).y
Else
    sX = MapNpc(MapNum).Npc(mapnpcnum).x
    sY = MapNpc(MapNum).Npc(mapnpcnum).y
End If
FX = targetx
FY = targety

If FX = -1 Then FX = 0
If FY = -1 Then FY = 0
If IsZoneNpc Then
    If FX > Map(ZoneNpc(ZoneNum).Npc(mapnpcnum).Map).MaxX Then FX = Map(ZoneNpc(ZoneNum).Npc(mapnpcnum).Map).MaxX - 1
    If FY > Map(ZoneNpc(ZoneNum).Npc(mapnpcnum).Map).MaxY Then FY = Map(ZoneNpc(ZoneNum).Npc(mapnpcnum).Map).MaxY - 1
End If

ReDim pos(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
pos = MapBlocks(MapNum).Blocks

pos(sX, sY) = 100 + tim
pos(FX, FY) = 2

'reset reachable
reachable = False

'Do while reachable is false... if its set true in progress, we jump out
'If the path is decided unreachable in process, we will use exit sub. Not proper,
'but faster ;-)
Do While reachable = False
    'we loop through all squares
    For j = 0 To Map(MapNum).MaxY
        For i = 0 To Map(MapNum).MaxX
            'If j = 10 And i = 0 Then MsgBox "hi!"
            'If they are to be extended, the pointer TIM is on them
            If pos(i, j) = 100 + tim Then
            'The part is to be extended, so do it
                'We have to make sure that there is a pos(i+1,j) BEFORE we actually use it,
                'because then we get error... If the square is on side, we dont test for this one!
                If i < Map(MapNum).MaxX Then
                    'If there isnt a wall, or any other... thing
                    If pos(i + 1, j) = 0 Then
                        'Expand it, and make its pos equal to tim+1, so the next time we make this loop,
                        'It will exapand that square too! This is crucial part of the program
                        pos(i + 1, j) = 100 + tim + 1
                    ElseIf pos(i + 1, j) = 2 Then
                        'If the position is no 0 but its 2 (FINISH) then Reachable = true!!! We found end
                        reachable = True
                    End If
                End If
            
                'This is the same as the last one, as i said a lot of copy paste work and editing that
                'This is simply another side that we have to test for... so instead of i+1 we have i-1
                'Its actually pretty same then... I wont comment it therefore, because its only repeating
                'same thing with minor changes to check sides
                If i > 0 Then
                    If pos((i - 1), j) = 0 Then
                        pos(i - 1, j) = 100 + tim + 1
                    ElseIf pos(i - 1, j) = 2 Then
                        reachable = True
                    End If
                End If
            
                If j < Map(MapNum).MaxY Then
                    If pos(i, j + 1) = 0 Then
                        pos(i, j + 1) = 100 + tim + 1
                    ElseIf pos(i, j + 1) = 2 Then
                        reachable = True
                    End If
                End If
            
                If j > 0 Then
                    If pos(i, j - 1) = 0 Then
                        pos(i, j - 1) = 100 + tim + 1
                    ElseIf pos(i, j - 1) = 2 Then
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
        For j = 0 To Map(MapNum).MaxY
            For i = 0 To Map(MapNum).MaxX
            'we add up ALL the squares
            Sum = Sum + pos(i, j)
            Next i
        Next j
        
        'Now if the sum is euqal to the last sum, its not reachable, if it isnt, then we store
        'sum to lastsum
        If Sum = LastSum Then
            FindNpcPath = 4
            Exit Function
        Else
            LastSum = Sum
        End If
    End If
    
    'we increase the pointer to point to the next squares to be expanded
    tim = tim + 1
Loop

'We work backwards to find the way...
LastX = FX
LastY = FY

ReDim path(tim + 1)

'The following code may be a little bit confusing but ill try my best to explain it.
'We are working backwards to find ONE of the shortest ways back to Start.
'So we repeat the loop until the LastX and LastY arent in start. Look in the code to see
'how LastX and LasY change
Do While LastX <> sX Or LastY <> sY
    'We decrease tim by one, and then we are finding any adjacent square to the final one, that
    'has that value. So lets say the tim would be 5, because it takes 5 steps to get to the target.
    'Now everytime we decrease that, so we make it 4, and we look for any adjacent square that has
    'that value. When we find it, we just color it yellow as for the solution
    tim = tim - 1
    'reset did to false
    did = False
    
    'If we arent on edge
    If LastX < Map(MapNum).MaxX Then
        'check the square on the right of the solution. Is it a tim-1 one? or just a blank one
        If pos(LastX + 1, LastY) = 100 + tim Then
            'if it, then make it yellow, and change did to true
            LastX = LastX + 1
            did = True
        End If
    End If
    
    'This will then only work if the previous part didnt execute, and did is still false. THen
    'we want to check another square, the on left. Is it a tim-1 one ?
    If did = False Then
        If LastX > 0 Then
            If pos(LastX - 1, LastY) = 100 + tim Then
                LastX = LastX - 1
                did = True
            End If
        End If
    End If
    
    'We check the one below it
    If did = False Then
        If LastY < Map(MapNum).MaxY Then
            If pos(LastX, LastY + 1) = 100 + tim Then
                LastY = LastY + 1
                did = True
            End If
        End If
    End If
    
    'And above it. One of these have to be it, since we have found the solution, we know that already
    'there is a way back.
    If did = False Then
        If LastY > 0 Then
            If pos(LastX, LastY - 1) = 100 + tim Then
                LastY = LastY - 1
            End If
        End If
    End If
    
    path(tim).x = LastX
    path(tim).y = LastY
    
    'Now we loop back and decrease tim, and look for the next square with lower value
    DoEvents
Loop

'Ok we got a path. Now, lets look at the first step and see what direction we should take.
If path(1).x > LastX Then
    FindNpcPath = DIR_RIGHT
ElseIf path(1).y > LastY Then
    FindNpcPath = DIR_DOWN
ElseIf path(1).y < LastY Then
    FindNpcPath = DIR_UP
ElseIf path(1).x < LastX Then
    FindNpcPath = DIR_LEFT
End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindNpcPath", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub CacheMapBlocks(MapNum As Long)
Dim x As Long, y As Long

   On Error GoTo errorhandler

    ReDim MapBlocks(MapNum).Blocks(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            If NpcTileIsOpen(MapNum, x, y) = False Then
                MapBlocks(MapNum).Blocks(x, y) = 9
            End If
        Next
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CacheMapBlocks", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub UpdateMapBlock(MapNum, x, y, blocked As Boolean)

   On Error GoTo errorhandler
   
   If MapNum <= 0 Then Exit Sub

    If blocked Then
        MapBlocks(MapNum).Blocks(x, y) = 9
    Else
        MapBlocks(MapNum).Blocks(x, y) = 0
    End If
    'Player(1).Characters(1).X = 5
    'Player(1).Characters(1).y = 9


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateMapBlock", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function IsOneBlockAway(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Boolean

   On Error GoTo errorhandler

    If x1 = x2 Then
        If y1 = y2 - 1 Or y1 = y2 + 1 Then
            IsOneBlockAway = True
        Else
            IsOneBlockAway = False
        End If
    ElseIf y1 = y2 Then
        If x1 = x2 - 1 Or x1 = x2 + 1 Then
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
    HandleError "IsOneBlockAway", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetNpcDir(x As Long, y As Long, x1 As Long, y1 As Long) As Long
    Dim i As Long, distance As Long
    

   On Error GoTo errorhandler

    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_RIGHT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_LEFT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If y - y1 > 0 Then
        If y - y1 > distance Then
            i = DIR_DOWN
            distance = y - y1
        End If
    ElseIf y - y1 < 0 Then
        If ((y - y1) * -1) > distance Then
            i = DIR_UP
            distance = ((y - y1) * -1)
        End If
    End If
    
    GetNpcDir = i


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetNpcDir", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear
        
End Function
