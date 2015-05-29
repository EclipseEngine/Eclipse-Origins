Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long

   On Error GoTo errorhandler

    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindOpenPlayerSlot", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim i As Long

   On Error GoTo errorhandler

    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindOpenMapItemSlot", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long

   On Error GoTo errorhandler

    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "TotalOnlinePlayers", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(GetPlayerName(i)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindPlayer", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal itemval As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional PlayerDrop As Boolean = False)
    Dim i As Long, d As Long
    Dim shortestZone As Long
    Dim shortestNum As Long
    Dim shortestD As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, itemval, MapNum, x, y, playerName, True, PlayerDrop)
    
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


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnItem", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal itemval As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True, Optional PlayerDrop As Boolean = False)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
            MapItem(MapNum, i).playerName = playerName
            MapItem(MapNum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(MapNum, i).canDespawn = canDespawn
            MapItem(MapNum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(MapNum, i).Num = ItemNum
            MapItem(MapNum, i).Value = itemval
            MapItem(MapNum, i).x = x
            MapItem(MapNum, i).y = y
            MapItem(MapNum, i).PlayerDrop = PlayerDrop
            ' send to map
            SendSpawnItemToMap MapNum, i
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnItemSlot", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long


   On Error GoTo errorhandler

    For i = MIN_MAPS To MAX_MAPS
        SetLoadingProgress "Spawning Map Items.", 30, i / MAX_MAPS
        DoEvents
        Call SpawnMapItems(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnAllMapsItems", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(x, y).type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(x, y).Data1).Stackable = 1 And Map(MapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, Map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                End If
            End If

        Next
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnMapItems", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function Random1(ByVal Low As Long, ByVal High As Long) As Long
Dim lcount As Long
   On Error GoTo errorhandler
    'Random = Low - 1
    Do Until (Random1 >= Low And Random1 <= High) Or lcount > 10
        Random1 = ((High - Low + 1) * Rnd) + Low
        lcount = lcount + 1
    Loop
    
    If Random1 < Low Then Random1 = Low: Exit Function
    If Random1 > High Then Random1 = High: Exit Function


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "Random", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function
Function Random(ByVal Low As Long, ByVal High As Long) As Long

   On Error GoTo errorhandler

    Random = ((High - Low + 1) * Rnd) + Low


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "Random", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Sub SpawnNpc(ByVal mapnpcnum As Long, ByVal MapNum As Long, Optional ForcedSpawn As Boolean = False)
    Dim Buffer As clsBuffer
    Dim npcnum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long, n As Long, j As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range

   On Error GoTo errorhandler
   
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Then Exit Sub
    npcnum = Map(MapNum).Npc(mapnpcnum)
    If ForcedSpawn = False And Map(MapNum).NpcSpawnType(mapnpcnum) = 1 Then npcnum = 0
    If npcnum > 0 Then
    
        MapNpc(MapNum).Npc(mapnpcnum).Num = npcnum
        MapNpc(MapNum).Npc(mapnpcnum).Target = 0
        MapNpc(MapNum).Npc(mapnpcnum).TargetType = 0 ' clear
        
        MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.HP) = GetNpcMaxVital(npcnum, Vitals.HP)
        MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.MP) = GetNpcMaxVital(npcnum, Vitals.MP)
        
        MapNpc(MapNum).Npc(mapnpcnum).Dir = Int(Rnd * 4)
        MapNpc(MapNum).Npc(mapnpcnum).StunDuration = 0
        MapNpc(MapNum).Npc(mapnpcnum).StunTimer = 0
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(MapNum).MaxX
            For y = 0 To Map(MapNum).MaxY
                If Map(MapNum).Tile(x, y).type = TILE_TYPE_NPCSPAWN Then
                    If Map(MapNum).Tile(x, y).Data1 = mapnpcnum Then
                        MapNpc(MapNum).Npc(mapnpcnum).x = x
                        MapNpc(MapNum).Npc(mapnpcnum).y = y
                        MapNpc(MapNum).Npc(mapnpcnum).Dir = Map(MapNum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, Map(MapNum).MaxX)
                y = Random(0, Map(MapNum).MaxY)
    
                If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
                If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(MapNum, x, y) Then
                    MapNpc(MapNum).Npc(mapnpcnum).x = x
                    MapNpc(MapNum).Npc(mapnpcnum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(MapNum).MaxX
                For y = 0 To Map(MapNum).MaxY

                    If NpcTileIsOpen(MapNum, x, y) Then
                        MapNpc(MapNum).Npc(mapnpcnum).x = x
                        MapNpc(MapNum).Npc(mapnpcnum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong 0
            Buffer.WriteLong mapnpcnum
            Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).Num
            Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).Dir
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
            UpdateMapBlock MapNum, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y, True
        End If
        
        j = 1
        For n = 1 To MAX_NPC_DROPS
            If Npc(npcnum).DropItems(n) = 0 Then Exit For

            If Rnd <= Npc(npcnum).DropChances(n) / 100 Then
                MapNpc(MapNum).Npc(mapnpcnum).Inventory(j).Num = Npc(npcnum).DropItems(n)
                MapNpc(MapNum).Npc(mapnpcnum).Inventory(j).Value = Npc(npcnum).DropItemValues(n)
                j = j + 1
            End If
        Next
        
        SendMapNpcVitals MapNum, mapnpcnum
    Else
        MapNpc(MapNum).Npc(mapnpcnum).Num = 0
        MapNpc(MapNum).Npc(mapnpcnum).Target = 0
        MapNpc(MapNum).Npc(mapnpcnum).TargetType = 0 ' clear
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong 0
        Buffer.WriteLong mapnpcnum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnNpc", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SpawnMapEventsFor(Index As Long, MapNum As Long)
Dim i As Long, x As Long, y As Long, z As Long, spawncurrentevent As Boolean, p As Long, compare As Long
Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    TempPlayer(Index).EventMap.CurrentEvents = 0
    ReDim TempPlayer(Index).EventMap.EventPages(0)
    If Map(MapNum).EventCount > 0 Then
        ReDim TempPlayer(Index).EventProcessing(1 To Map(MapNum).EventCount)
        TempPlayer(Index).EventProcessingCount = Map(MapNum).EventCount
    Else
        ReDim TempPlayer(Index).EventProcessing(0)
        TempPlayer(Index).EventProcessingCount = 0
    End If
    
    If Map(MapNum).EventCount <= 0 Then Exit Sub
    For i = 1 To Map(MapNum).EventCount
        If Map(MapNum).Events(i).PageCount > 0 Then
            For z = Map(MapNum).Events(i).PageCount To 1 Step -1
                With Map(MapNum).Events(i).Pages(z)
                    spawncurrentevent = True
                    
                    If .chkVariable = 1 Then
                        Select Case .VariableCompare
                            Case 0
                                If Player(Index).characters(TempPlayer(Index).CurChar).Variables(.VariableIndex) <> .VariableCondition Then
                                    spawncurrentevent = False
                                End If
                            Case 1
                                If Player(Index).characters(TempPlayer(Index).CurChar).Variables(.VariableIndex) < .VariableCondition Then
                                    spawncurrentevent = False
                                End If
                            Case 2
                                If Player(Index).characters(TempPlayer(Index).CurChar).Variables(.VariableIndex) > .VariableCondition Then
                                    spawncurrentevent = False
                                End If
                            Case 3
                                If Player(Index).characters(TempPlayer(Index).CurChar).Variables(.VariableIndex) <= .VariableCondition Then
                                    spawncurrentevent = False
                                End If
                            Case 4
                                If Player(Index).characters(TempPlayer(Index).CurChar).Variables(.VariableIndex) >= .VariableCondition Then
                                    spawncurrentevent = False
                                End If
                            Case 5
                                If Player(Index).characters(TempPlayer(Index).CurChar).Variables(.VariableIndex) = .VariableCondition Then
                                    spawncurrentevent = False
                                End If
                        End Select
                    End If
                    
                    If .chkSwitch = 1 Then
                        If .SwitchCompare = 1 Then
                            If Player(Index).characters(TempPlayer(Index).CurChar).Switches(.SwitchIndex) = 1 Then
                                spawncurrentevent = False
                            End If
                        Else
                            If Player(Index).characters(TempPlayer(Index).CurChar).Switches(.SwitchIndex) = 0 Then
                                spawncurrentevent = False
                            End If
                        End If
                    End If
                    
                    If .chkHasItem = 1 Then
                        If HasItem(Index, .HasItemIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSelfSwitch = 1 Then
                        If .SelfSwitchCompare = 0 Then
                            compare = 1
                        Else
                            compare = 0
                        End If
                        If Map(MapNum).Events(i).Global = 1 Then
                            If Map(MapNum).Events(i).SelfSwitches(.SelfSwitchIndex) <> compare Then
                                spawncurrentevent = False
                            End If
                        Else
                            If compare = 1 Then
                                spawncurrentevent = False
                            End If
                        End If
                    End If
                    
                    If spawncurrentevent = True Or (spawncurrentevent = False And z = 1) Then
                        'spawn the event... send data to player
                        TempPlayer(Index).EventMap.CurrentEvents = TempPlayer(Index).EventMap.CurrentEvents + 1
                        ReDim Preserve TempPlayer(Index).EventMap.EventPages(TempPlayer(Index).EventMap.CurrentEvents)
                        With TempPlayer(Index).EventMap.EventPages(TempPlayer(Index).EventMap.CurrentEvents)
                            If Map(MapNum).Events(i).Pages(z).GraphicType = 1 Then
                                Select Case Map(MapNum).Events(i).Pages(z).GraphicY
                                    Case 0
                                        .Dir = DIR_DOWN
                                    Case 1
                                        .Dir = DIR_LEFT
                                    Case 2
                                        .Dir = DIR_RIGHT
                                    Case 3
                                        .Dir = DIR_UP
                                End Select
                            Else
                                .Dir = 0
                            End If
                            .GraphicNum = Map(MapNum).Events(i).Pages(z).Graphic
                            .GraphicType = Map(MapNum).Events(i).Pages(z).GraphicType
                            .GraphicX = Map(MapNum).Events(i).Pages(z).GraphicX
                            .GraphicY = Map(MapNum).Events(i).Pages(z).GraphicY
                            .GraphicX2 = Map(MapNum).Events(i).Pages(z).GraphicX2
                            .GraphicY2 = Map(MapNum).Events(i).Pages(z).GraphicY2
                            Select Case Map(MapNum).Events(i).Pages(z).MoveSpeed
                                Case 0
                                    .movementspeed = 2
                                Case 1
                                    .movementspeed = 3
                                Case 2
                                    .movementspeed = 4
                                Case 3
                                    .movementspeed = 6
                                Case 4
                                    .movementspeed = 12
                                Case 5
                                    .movementspeed = 24
                            End Select
                            If Map(MapNum).Events(i).Global Then
                                .x = TempEventMap(MapNum).Events(i).x
                                .y = TempEventMap(MapNum).Events(i).y
                                .Dir = TempEventMap(MapNum).Events(i).Dir
                                .MoveRouteStep = TempEventMap(MapNum).Events(i).MoveRouteStep
                            Else
                                .x = Map(MapNum).Events(i).x
                                .y = Map(MapNum).Events(i).y
                                .MoveRouteStep = 0
                            End If
                            .Position = Map(MapNum).Events(i).Pages(z).Position
                            .eventID = i
                            .pageID = z
                            If spawncurrentevent = True Then
                                .Visible = 1
                            Else
                                .Visible = 0
                            End If
                            
                            .MoveType = Map(MapNum).Events(i).Pages(z).MoveType
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(MapNum).Events(i).Pages(z).MoveRouteCount
                                ReDim .MoveRoute(0 To Map(MapNum).Events(i).Pages(z).MoveRouteCount)
                                If Map(MapNum).Events(i).Pages(z).MoveRouteCount > 0 Then
                                    For p = 0 To Map(MapNum).Events(i).Pages(z).MoveRouteCount
                                        .MoveRoute(p) = Map(MapNum).Events(i).Pages(z).MoveRoute(p)
                                    Next
                                End If
                                .MoverouteComplete = 0
                            Else
                                .MoverouteComplete = 1
                            End If
                            
                            .RepeatMoveRoute = Map(MapNum).Events(i).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(MapNum).Events(i).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(MapNum).Events(i).Pages(z).MoveFreq
                            .MoveSpeed = Map(MapNum).Events(i).Pages(z).MoveSpeed
                            
                            .WalkingAnim = Map(MapNum).Events(i).Pages(z).WalkAnim
                            .WalkThrough = Map(MapNum).Events(i).Pages(z).WalkThrough
                            .ShowName = Map(MapNum).Events(i).Pages(z).ShowName
                            .FixedDir = Map(MapNum).Events(i).Pages(z).DirFix
                            .questnum = Map(MapNum).Events(i).Pages(z).questnum
                        End With
                        GoTo nextevent
                    End If
                End With
            Next
        End If
nextevent:
    Next
    
    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
        For i = 1 To TempPlayer(Index).EventMap.CurrentEvents
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnEvent
            Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(i).eventID
            With TempPlayer(Index).EventMap.EventPages(i)
                Buffer.WriteString Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Name
                Buffer.WriteLong .Dir
                Buffer.WriteLong .GraphicNum
                Buffer.WriteLong .GraphicType
                Buffer.WriteLong .GraphicX
                Buffer.WriteLong .GraphicX2
                Buffer.WriteLong .GraphicY
                Buffer.WriteLong .GraphicY2
                Buffer.WriteLong .movementspeed
                Buffer.WriteLong .x
                Buffer.WriteLong .y
                Buffer.WriteLong .Position
                Buffer.WriteLong .Visible
                Buffer.WriteLong Map(MapNum).Events(.eventID).Pages(.pageID).WalkAnim
                Buffer.WriteLong Map(MapNum).Events(.eventID).Pages(.pageID).DirFix
                Buffer.WriteLong Map(MapNum).Events(.eventID).Pages(.pageID).WalkThrough
                Buffer.WriteLong Map(MapNum).Events(.eventID).Pages(.pageID).ShowName
                Buffer.WriteLong Map(MapNum).Events(.eventID).Pages(.pageID).questnum
            End With
            SendDataTo Index, Buffer.ToArray
            Set Buffer = Nothing
        Next
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnMapEventsFor", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long

   On Error GoTo errorhandler

    NpcTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex
            If IsPlaying(LoopI) Then
                If GetPlayerMap(LoopI) = MapNum Then
                    If GetPlayerX(LoopI) = x Then
                        If GetPlayerY(LoopI) = y Then
                            NpcTileIsOpen = False
                            Exit Function
                        End If
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum).Npc(LoopI).Num > 0 Then
            If MapNpc(MapNum).Npc(LoopI).x = x Then
                If MapNpc(MapNum).Npc(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next
    
    For LoopI = 1 To TempEventMap(MapNum).EventCount
        If TempEventMap(MapNum).Events(LoopI).Active = 1 Then
            If MapNpc(MapNum).Npc(LoopI).x = TempEventMap(MapNum).Events(LoopI).x Then
                If MapNpc(MapNum).Npc(LoopI).y = TempEventMap(MapNum).Events(LoopI).y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If
    Next

    If Map(MapNum).Tile(x, y).type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).Tile(x, y).type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).Tile(x, y).type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "NpcTileIsOpen", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next
    
    CacheMapBlocks MapNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnMapNpcs", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long


   On Error GoTo errorhandler

    For i = MIN_MAPS To MAX_MAPS
        SetLoadingProgress "Spawning Map Npcs.", 31, i / MAX_MAPS
        DoEvents
        Call SpawnMapNpcs(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnAllMapNpcs", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SpawnAllZoneNpcs()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ZONES
        Call SpawnZoneNpcs(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnAllZoneNpcs", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SpawnZoneNpcs(ByVal ZoneNum As Long)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call SpawnZoneNpc(ZoneNum, i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnZoneNpcs", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SpawnAllMapGlobalEvents()
    Dim i As Long


   On Error GoTo errorhandler

    For i = MIN_MAPS To MAX_MAPS
        SetLoadingProgress "Spawning Global Events.", 32, i / MAX_MAPS
        DoEvents
        Call SpawnGlobalEvents(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnAllMapGlobalEvents", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SpawnGlobalEvents(ByVal MapNum As Long)
    Dim i As Long, z As Long
    

   On Error GoTo errorhandler

    If Map(MapNum).EventCount > 0 Then
        TempEventMap(MapNum).EventCount = 0
        ReDim TempEventMap(MapNum).Events(0)
        For i = 1 To Map(MapNum).EventCount
            TempEventMap(MapNum).EventCount = TempEventMap(MapNum).EventCount + 1
            ReDim Preserve TempEventMap(MapNum).Events(0 To TempEventMap(MapNum).EventCount)
            If Map(MapNum).Events(i).PageCount > 0 Then
                If Map(MapNum).Events(i).Global = 1 Then
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).x = Map(MapNum).Events(i).x
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).y = Map(MapNum).Events(i).y
                    If Map(MapNum).Events(i).Pages(1).GraphicType = 1 Then
                        Select Case Map(MapNum).Events(i).Pages(1).GraphicY
                            Case 0
                                TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Dir = DIR_DOWN
                            Case 1
                                TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Dir = DIR_LEFT
                            Case 2
                                TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Dir = DIR_RIGHT
                            Case 3
                                TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Dir = DIR_UP
                        End Select
                    Else
                        TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Dir = DIR_DOWN
                    End If
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Active = 1
                    
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveType = Map(MapNum).Events(i).Pages(1).MoveType
                    
                    If TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveType = 2 Then
                        TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveRouteCount = Map(MapNum).Events(i).Pages(1).MoveRouteCount
                        ReDim TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveRoute(0 To Map(MapNum).Events(i).Pages(1).MoveRouteCount)
                        For z = 0 To Map(MapNum).Events(i).Pages(1).MoveRouteCount
                            TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveRoute(z) = Map(MapNum).Events(i).Pages(1).MoveRoute(z)
                        Next
                        TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoverouteComplete = 0
                    Else
                        TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoverouteComplete = 1
                    End If
                    
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).RepeatMoveRoute = Map(MapNum).Events(i).Pages(1).RepeatMoveRoute
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).IgnoreIfCannotMove = Map(MapNum).Events(i).Pages(1).IgnoreMoveRoute
                    
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveFreq = Map(MapNum).Events(i).Pages(1).MoveFreq
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveSpeed = Map(MapNum).Events(i).Pages(1).MoveSpeed
                    
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).WalkThrough = Map(MapNum).Events(i).Pages(1).WalkThrough
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).FixedDir = Map(MapNum).Events(i).Pages(1).DirFix
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).WalkingAnim = Map(MapNum).Events(i).Pages(1).WalkAnim
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).ShowName = Map(MapNum).Events(i).Pages(1).ShowName
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).questnum = Map(MapNum).Events(i).Pages(1).questnum
                End If
            End If
        Next
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnGlobalEvents", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal Dir As Byte, IsZoneNpc As Boolean, ZoneNum As Long) As Boolean
    Dim i As Long, Buffer As clsBuffer
    Dim n As Long
    Dim x As Long
    Dim y As Long, z As Long, j As Long


   On Error GoTo errorhandler

    If IsZoneNpc Then
        ' Check for subscript out of range
        If MapNum <= 0 Or MapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS * 2 Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Exit Function
        End If
        x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x
        y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y
    Else
        ' Check for subscript out of range
        If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Exit Function
        End If
        x = MapNpc(MapNum).Npc(mapnpcnum).x
        y = MapNpc(MapNum).Npc(mapnpcnum).y
    End If
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y - 1 > 0 Then
                n = Map(MapNum).Tile(x, y - 1).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If IsZoneNpc Then
                            If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = ZoneNpc(ZoneNum).Npc(mapnpcnum).x) And (GetPlayerY(i) = ZoneNpc(ZoneNum).Npc(mapnpcnum).y - 1) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        Else
                            If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapnpcnum).x) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapnpcnum).y - 1) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        End If
                        If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive Then
                            If IsZoneNpc Then
                                If (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y - 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            Else
                                If (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = MapNpc(MapNum).Npc(mapnpcnum).x) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = MapNpc(MapNum).Npc(mapnpcnum).y - 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next

                If IsZoneNpc Then
                    ' Check to make sure that there is not another npc in the way
                    For i = 1 To MAX_MAP_NPCS
                        If (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x) And (MapNpc(MapNum).Npc(i).y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    Next
                Else
                    ' Check to make sure that there is not another npc in the way
                    For i = 1 To MAX_MAP_NPCS
                        If (i <> mapnpcnum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(mapnpcnum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(mapnpcnum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    Next
                End If
                
                For i = 1 To MAX_ZONES
                    For j = 1 To MAX_MAP_NPCS * 2
                        If IsZoneNpc And ZoneNum = i And mapnpcnum = j Then
                            
                        Else
                            If IsZoneNpc Then
                                If ZoneNpc(i).Npc(j).Num > 0 And ZoneNpc(i).Npc(j).Map = MapNum And ZoneNpc(i).Npc(j).x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x And ZoneNpc(ZoneNum).Npc(j).y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y - 1 Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                Next
            Else
                If IsZoneNpc Then
                    If Map(MapNum).Up > 0 Then
                        If MapZones(ZoneNum).MapCount > 0 Then
                            For j = 1 To MapZones(ZoneNum).MapCount
                                If Map(MapNum).Up = MapZones(ZoneNum).Maps(j) Then
                                    'Can Move
                                    Set Buffer = New clsBuffer
                                    Buffer.WriteLong SNpcDead
                                    Buffer.WriteLong ZoneNum
                                    Buffer.WriteLong mapnpcnum
                                    SendDataToMap MapNum, Buffer.ToArray()
                                    Set Buffer = Nothing
                                    ZoneNpc(ZoneNum).Npc(mapnpcnum).Map = Map(MapNum).Up
                                    ZoneNpc(ZoneNum).Npc(mapnpcnum).y = Map(Map(MapNum).Up).MaxY
                                    Set Buffer = New clsBuffer
                                    Buffer.WriteLong SSpawnNpc
                                    Buffer.WriteLong ZoneNum
                                    Buffer.WriteLong mapnpcnum
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Num
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).x
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).y
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Dir
                                    SendDataToMap Map(MapNum).Up, Buffer.ToArray()
                                    Set Buffer = Nothing
                                    SendZoneNpcVitals ZoneNum, mapnpcnum
                                    CanNpcMove = False
                                Else
                                    If j = MapZones(ZoneNum).MapCount Then CanNpcMove = False
                                End If
                            Next
                        Else
                            CanNpcMove = False
                        End If
                    Else
                        CanNpcMove = False
                    End If
                Else
                    CanNpcMove = False
                End If
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y + 1 < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(x, y + 1).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If IsZoneNpc Then
                            If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = ZoneNpc(ZoneNum).Npc(mapnpcnum).x) And (GetPlayerY(i) = ZoneNpc(ZoneNum).Npc(mapnpcnum).y + 1) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        Else
                            If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapnpcnum).x) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapnpcnum).y + 1) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        End If
                        If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive Then
                            If IsZoneNpc Then
                                If (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y + 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            Else
                                If (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = MapNpc(MapNum).Npc(mapnpcnum).x) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = MapNpc(MapNum).Npc(mapnpcnum).y + 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next

                If IsZoneNpc Then
                    ' Check to make sure that there is not another npc in the way
                    For i = 1 To MAX_MAP_NPCS
                        If (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x) And (MapNpc(MapNum).Npc(i).y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    Next
                Else
                    ' Check to make sure that there is not another npc in the way
                    For i = 1 To MAX_MAP_NPCS
                        If (i <> mapnpcnum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(mapnpcnum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(mapnpcnum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    Next
                End If
                
                For i = 1 To MAX_ZONES
                    For j = 1 To MAX_MAP_NPCS * 2
                        If IsZoneNpc And ZoneNum = i And mapnpcnum = j Then
                            
                        Else
                            If IsZoneNpc Then
                                If ZoneNpc(i).Npc(j).Num > 0 And ZoneNpc(i).Npc(j).Map = MapNum And ZoneNpc(i).Npc(j).x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x And ZoneNpc(ZoneNum).Npc(j).y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y + 1 Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                Next
            Else
                If IsZoneNpc Then
                    If Map(MapNum).Down > 0 Then
                        If MapZones(ZoneNum).MapCount > 0 Then
                            For j = 1 To MapZones(ZoneNum).MapCount
                                If Map(MapNum).Down = MapZones(ZoneNum).Maps(j) Then
                                    Set Buffer = New clsBuffer
                                    Buffer.WriteLong SNpcDead
                                    Buffer.WriteLong ZoneNum
                                    Buffer.WriteLong mapnpcnum
                                    SendDataToMap MapNum, Buffer.ToArray()
                                    Set Buffer = Nothing
                                    ZoneNpc(ZoneNum).Npc(mapnpcnum).Map = Map(MapNum).Down
                                    ZoneNpc(ZoneNum).Npc(mapnpcnum).y = 0
                                    Set Buffer = New clsBuffer
                                    Buffer.WriteLong SSpawnNpc
                                    Buffer.WriteLong ZoneNum
                                    Buffer.WriteLong mapnpcnum
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Num
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).x
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).y
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Dir
                                    SendDataToMap Map(MapNum).Down, Buffer.ToArray()
                                    Set Buffer = Nothing
                                    SendZoneNpcVitals ZoneNum, mapnpcnum
                                    CanNpcMove = False
                                Else
                                    If j = MapZones(ZoneNum).MapCount Then CanNpcMove = False
                                End If
                            Next
                        Else
                            CanNpcMove = False
                        End If
                    Else
                        CanNpcMove = False
                    End If
                Else
                    CanNpcMove = False
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x - 1 > 0 Then
                n = Map(MapNum).Tile(x - 1, y).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If IsZoneNpc Then
                            If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = ZoneNpc(ZoneNum).Npc(mapnpcnum).x - 1) And (GetPlayerY(i) = ZoneNpc(ZoneNum).Npc(mapnpcnum).y) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        Else
                            If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapnpcnum).x - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapnpcnum).y) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        End If
                        If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive Then
                            If IsZoneNpc Then
                                If (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x - 1) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            Else
                                If (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = MapNpc(MapNum).Npc(mapnpcnum).x - 1) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = MapNpc(MapNum).Npc(mapnpcnum).y) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
                
                If IsZoneNpc Then
                    ' Check to make sure that there is not another npc in the way
                    For i = 1 To MAX_MAP_NPCS
                        If (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x - 1) And (MapNpc(MapNum).Npc(i).y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    Next
                Else
                    ' Check to make sure that there is not another npc in the way
                    For i = 1 To MAX_MAP_NPCS
                        If (i <> mapnpcnum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(mapnpcnum).x - 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(mapnpcnum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    Next
                End If
                
                For i = 1 To MAX_ZONES
                    For j = 1 To MAX_MAP_NPCS * 2
                        If IsZoneNpc And ZoneNum = i And mapnpcnum = j Then
                            
                        Else
                            If IsZoneNpc Then
                                If ZoneNpc(i).Npc(j).Num > 0 And ZoneNpc(i).Npc(j).Map = MapNum And ZoneNpc(i).Npc(j).x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x - 1 And ZoneNpc(ZoneNum).Npc(j).y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                Next
            Else
                If IsZoneNpc Then
                    If Map(MapNum).Left > 0 Then
                        If MapZones(ZoneNum).MapCount > 0 Then
                            For j = 1 To MapZones(ZoneNum).MapCount
                                If Map(MapNum).Left = MapZones(ZoneNum).Maps(j) Then
                                    Set Buffer = New clsBuffer
                                    Buffer.WriteLong SNpcDead
                                    Buffer.WriteLong ZoneNum
                                    Buffer.WriteLong mapnpcnum
                                    SendDataToMap MapNum, Buffer.ToArray()
                                    Set Buffer = Nothing
                                    ZoneNpc(ZoneNum).Npc(mapnpcnum).Map = Map(MapNum).Left
                                    ZoneNpc(ZoneNum).Npc(mapnpcnum).x = Map(Map(MapNum).Left).MaxX
                                    Set Buffer = New clsBuffer
                                    Buffer.WriteLong SSpawnNpc
                                    Buffer.WriteLong ZoneNum
                                    Buffer.WriteLong mapnpcnum
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Num
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).x
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).y
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Dir
                                    SendDataToMap Map(MapNum).Left, Buffer.ToArray()
                                    Set Buffer = Nothing
                                    SendZoneNpcVitals ZoneNum, mapnpcnum
                                    CanNpcMove = False
                                Else
                                    If j = MapZones(ZoneNum).MapCount Then CanNpcMove = False
                                End If
                            Next
                        Else
                            CanNpcMove = False
                        End If
                    Else
                        CanNpcMove = False
                    End If
                Else
                    CanNpcMove = False
                End If
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x + 1 < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(x + 1, y).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If IsZoneNpc Then
                            If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = ZoneNpc(ZoneNum).Npc(mapnpcnum).x + 1) And (GetPlayerY(i) = ZoneNpc(ZoneNum).Npc(mapnpcnum).y) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        Else
                            If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapnpcnum).x + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapnpcnum).y) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        End If
                        If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive Then
                            If IsZoneNpc Then
                                If (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x + 1) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            Else
                                If (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = MapNpc(MapNum).Npc(mapnpcnum).x + 1) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = MapNpc(MapNum).Npc(mapnpcnum).y) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next

                If IsZoneNpc Then
                    ' Check to make sure that there is not another npc in the way
                    For i = 1 To MAX_MAP_NPCS
                        If (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x + 1) And (MapNpc(MapNum).Npc(i).y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    Next
                Else
                    ' Check to make sure that there is not another npc in the way
                    For i = 1 To MAX_MAP_NPCS
                        If (i <> mapnpcnum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(mapnpcnum).x + 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(mapnpcnum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    Next
                End If
                
                For i = 1 To MAX_ZONES
                    For j = 1 To MAX_MAP_NPCS * 2
                        If IsZoneNpc And ZoneNum = i And mapnpcnum = j Then
                            
                        Else
                            If IsZoneNpc Then
                                If ZoneNpc(i).Npc(j).Num > 0 And ZoneNpc(i).Npc(j).Map = MapNum And ZoneNpc(i).Npc(j).x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x + 1 And ZoneNpc(ZoneNum).Npc(j).y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                Next
            Else
                If IsZoneNpc Then
                    If Map(MapNum).Right > 0 Then
                        If MapZones(ZoneNum).MapCount > 0 Then
                            For j = 1 To MapZones(ZoneNum).MapCount
                                If Map(MapNum).Right = MapZones(ZoneNum).Maps(j) Then
                                    Set Buffer = New clsBuffer
                                    Buffer.WriteLong SNpcDead
                                    Buffer.WriteLong ZoneNum
                                    Buffer.WriteLong mapnpcnum
                                    SendDataToMap MapNum, Buffer.ToArray()
                                    Set Buffer = Nothing
                                    ZoneNpc(ZoneNum).Npc(mapnpcnum).Map = Map(MapNum).Right
                                    ZoneNpc(ZoneNum).Npc(mapnpcnum).x = 0
                                    Set Buffer = New clsBuffer
                                    Buffer.WriteLong SSpawnNpc
                                    Buffer.WriteLong ZoneNum
                                    Buffer.WriteLong mapnpcnum
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Num
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).x
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).y
                                    Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Dir
                                    SendDataToMap Map(MapNum).Right, Buffer.ToArray()
                                    Set Buffer = Nothing
                                    SendZoneNpcVitals ZoneNum, mapnpcnum
                                    CanNpcMove = False
                                Else
                                    If j = MapZones(ZoneNum).MapCount Then CanNpcMove = False
                                End If
                            Next
                        Else
                            CanNpcMove = False
                        End If
                    Else
                        CanNpcMove = False
                    End If
                Else
                    CanNpcMove = False
                End If
            End If

    End Select


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanNpcMove", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub NpcMove(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal Dir As Long, ByVal movement As Long, IsZoneNpc As Boolean, ZoneNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    If IsZoneNpc Then
        ' Check for subscript out of range
        If MapNum <= 0 Or MapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS * 2 Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
            Exit Sub
        End If
    
        ZoneNpc(ZoneNum).Npc(mapnpcnum).Dir = Dir
        UpdateMapBlock MapNum, ZoneNpc(ZoneNum).Npc(mapnpcnum).x, ZoneNpc(ZoneNum).Npc(mapnpcnum).y, False
    
        Select Case Dir
            Case DIR_UP
                ZoneNpc(ZoneNum).Npc(mapnpcnum).y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SNpcMove
                Buffer.WriteLong ZoneNum
                Buffer.WriteLong mapnpcnum
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).x
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).y
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Dir
                Buffer.WriteLong movement
                SendDataToMap MapNum, Buffer.ToArray()
                Set Buffer = Nothing
            Case DIR_DOWN
                ZoneNpc(ZoneNum).Npc(mapnpcnum).y = ZoneNpc(ZoneNum).Npc(mapnpcnum).y + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SNpcMove
                Buffer.WriteLong ZoneNum
                Buffer.WriteLong mapnpcnum
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).x
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).y
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Dir
                Buffer.WriteLong movement
                SendDataToMap MapNum, Buffer.ToArray()
                Set Buffer = Nothing
            Case DIR_LEFT
                ZoneNpc(ZoneNum).Npc(mapnpcnum).x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SNpcMove
                Buffer.WriteLong ZoneNum
                Buffer.WriteLong mapnpcnum
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).x
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).y
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Dir
                Buffer.WriteLong movement
                SendDataToMap MapNum, Buffer.ToArray()
                Set Buffer = Nothing
            Case DIR_RIGHT
                ZoneNpc(ZoneNum).Npc(mapnpcnum).x = ZoneNpc(ZoneNum).Npc(mapnpcnum).x + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SNpcMove
                Buffer.WriteLong ZoneNum
                Buffer.WriteLong mapnpcnum
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).x
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).y
                Buffer.WriteLong ZoneNpc(ZoneNum).Npc(mapnpcnum).Dir
                Buffer.WriteLong movement
                SendDataToMap MapNum, Buffer.ToArray()
                Set Buffer = Nothing
        End Select
        
        UpdateMapBlock MapNum, ZoneNpc(ZoneNum).Npc(mapnpcnum).x, ZoneNpc(ZoneNum).Npc(mapnpcnum).y, True
    Else
        ' Check for subscript out of range
        If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
            Exit Sub
        End If
    
        MapNpc(MapNum).Npc(mapnpcnum).Dir = Dir
        UpdateMapBlock MapNum, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y, False
    
        Select Case Dir
            Case DIR_UP
                MapNpc(MapNum).Npc(mapnpcnum).y = MapNpc(MapNum).Npc(mapnpcnum).y - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SNpcMove
                Buffer.WriteLong ZoneNum
                Buffer.WriteLong mapnpcnum
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).x
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).y
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).Dir
                Buffer.WriteLong movement
                SendDataToMap MapNum, Buffer.ToArray()
                Set Buffer = Nothing
            Case DIR_DOWN
                MapNpc(MapNum).Npc(mapnpcnum).y = MapNpc(MapNum).Npc(mapnpcnum).y + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SNpcMove
                Buffer.WriteLong ZoneNum
                Buffer.WriteLong mapnpcnum
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).x
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).y
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).Dir
                Buffer.WriteLong movement
                SendDataToMap MapNum, Buffer.ToArray()
                Set Buffer = Nothing
            Case DIR_LEFT
                MapNpc(MapNum).Npc(mapnpcnum).x = MapNpc(MapNum).Npc(mapnpcnum).x - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SNpcMove
                Buffer.WriteLong ZoneNum
                Buffer.WriteLong mapnpcnum
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).x
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).y
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).Dir
                Buffer.WriteLong movement
                SendDataToMap MapNum, Buffer.ToArray()
                Set Buffer = Nothing
            Case DIR_RIGHT
                MapNpc(MapNum).Npc(mapnpcnum).x = MapNpc(MapNum).Npc(mapnpcnum).x + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SNpcMove
                Buffer.WriteLong ZoneNum
                Buffer.WriteLong mapnpcnum
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).x
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).y
                Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).Dir
                Buffer.WriteLong movement
                SendDataToMap MapNum, Buffer.ToArray()
                Set Buffer = Nothing
        End Select
        
        UpdateMapBlock MapNum, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y, True
    
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "NpcMove", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal Dir As Long, IsZoneNpc As Boolean, ZoneNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
     ' Check if have player on map

   On Error GoTo errorhandler

    If PlayersOnMap(MapNum) = NO Then Exit Sub
    
    If IsZoneNpc Then
        ' Check for subscript out of range
        If MapNum <= 0 Or MapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS * 2 Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Exit Sub
        End If
    
        ZoneNpc(ZoneNum).Npc(mapnpcnum).Dir = Dir
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDir
        Buffer.WriteLong ZoneNum
        Buffer.WriteLong mapnpcnum
        Buffer.WriteLong Dir
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing

    Else
        ' Check for subscript out of range
        If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Exit Sub
        End If
    
        MapNpc(MapNum).Npc(mapnpcnum).Dir = Dir
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDir
        Buffer.WriteLong ZoneNum
        Buffer.WriteLong mapnpcnum
        Buffer.WriteLong Dir
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "NpcDir", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim n As Long

   On Error GoTo errorhandler

    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                n = n + 1
            End If
        End If

    Next

    GetTotalMapPlayers = n


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetTotalMapPlayers", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub ClearTempTiles()
    Dim i As Long


   On Error GoTo errorhandler

    For i = MIN_MAPS To MAX_MAPS
        ClearTempTile i
        SetLoadingProgress "Clearing Temporary Tiles.", 4, i / MAX_MAPS
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearTempTiles", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearTempTile(ByVal MapNum As Long)
    Dim y As Long
    Dim x As Long

   On Error GoTo errorhandler

    temptile(MapNum).DoorTimer = 0
    ReDim temptile(MapNum).DoorOpen(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            temptile(MapNum).DoorOpen(x, y) = NO
        Next
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearTempTile", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CacheResources(ByVal MapNum As Long)
    Dim x As Long, y As Long, Resource_Count As Long

   On Error GoTo errorhandler

    Resource_Count = 0

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            If Map(MapNum).Tile(x, y).type = TILE_TYPE_RESOURCE Then
                If Map(MapNum).Tile(x, y).Data1 > 0 Then
                    Resource_Count = Resource_Count + 1
                    ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                    ResourceCache(MapNum).ResourceData(Resource_Count).x = x
                    ResourceCache(MapNum).ResourceData(Resource_Count).y = y
                    ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(Map(MapNum).Tile(x, y).Data1).Health
                End If
            End If

        Next
    Next

    ResourceCache(MapNum).Resource_Count = Resource_Count


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub PlayerSwitchBankSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long


   On Error GoTo errorhandler

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(Index, oldSlot)
    OldValue = GetPlayerBankItemValue(Index, oldSlot)
    NewNum = GetPlayerBankItemNum(Index, newSlot)
    NewValue = GetPlayerBankItemValue(Index, newSlot)
    
    SetPlayerBankItemNum Index, newSlot, OldNum
    SetPlayerBankItemValue Index, newSlot, OldValue
    
    SetPlayerBankItemNum Index, oldSlot, NewNum
    SetPlayerBankItemValue Index, oldSlot, NewValue
        
    SendBank Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerSwitchBankSlots", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long


   On Error GoTo errorhandler

    OldNum = GetPlayerInvItemNum(Index, oldSlot)
    OldValue = GetPlayerInvItemValue(Index, oldSlot)
    NewNum = GetPlayerInvItemNum(Index, newSlot)
    NewValue = GetPlayerInvItemValue(Index, newSlot)
    SetPlayerInvItemNum Index, newSlot, OldNum
    SetPlayerInvItemValue Index, newSlot, OldValue
    SetPlayerInvItemNum Index, oldSlot, NewNum
    SetPlayerInvItemValue Index, oldSlot, NewValue
    SendInventory Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerSwitchInvSlots", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long


   On Error GoTo errorhandler

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(Index, oldSlot)
    NewNum = GetPlayerSpell(Index, newSlot)
    SetPlayerSpell Index, oldSlot, NewNum
    SetPlayerSpell Index, newSlot, OldNum
    SendPlayerSpells Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerSwitchSpellSlots", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long)


   On Error GoTo errorhandler

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(Index, GetPlayerEquipment(Index, EqSlot)) > 0 Then
        GiveInvItem Index, GetPlayerEquipment(Index, EqSlot), 0
        PlayerMsg Index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(Index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipment(Index, EqSlot)
        ' remove equipment
        SetPlayerEquipment Index, 0, EqSlot
        SendWornEquipment Index
        SendMapEquipment Index
        SendStats Index
        ' send vitals
        Call SendVital(Index, Vitals.HP)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
    Else
        PlayerMsg Index, "Your inventory is full.", BrightRed
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerUnequipItem", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   

   On Error GoTo errorhandler

    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CheckGrammar", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long

   On Error GoTo errorhandler

    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "isInRange", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function
Function dist(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Dim nVal As Long

   On Error GoTo errorhandler

    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    dist = nVal


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "dist", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean

   On Error GoTo errorhandler

    If Not blockvar And (2 ^ Dir) Then
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

Public Function rand(ByVal Low As Long, ByVal High As Long) As Long

   On Error GoTo errorhandler

    'rand = Random(Low, High)
    rand = Int((High - Low + 1) * Rnd) + Low


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "rand", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal Index As Long)
Dim partyNum As Long, i As Long


   On Error GoTo errorhandler

    partyNum = TempPlayer(Index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
        
            ' check if leader
            If Party(partyNum).Leader = Index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> Index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(i) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(Index) & " left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = Index Then
                        Party(partyNum).Member(i) = 0
                        TempPlayer(Index).inParty = 0
                        TempPlayer(Index).partyInvite = 0
                        Exit For
                        End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo Index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(Index) & " left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = Index Then
                        Party(partyNum).Member(i) = 0
                        TempPlayer(Index).inParty = 0
                        TempPlayer(Index).partyInvite = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo Index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                Index = Party(partyNum).Member(i)
                ' player exist?
                If Index > 0 Then
                    ' remove them
                    TempPlayer(Index).partyInvite = 0
                    TempPlayer(Index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo Index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Party_PlayerLeave", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub Party_Invite(ByVal Index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if the person is a valid target

   On Error GoTo errorhandler

    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg Index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg Index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(Index).inParty > 0 Then
        partyNum = TempPlayer(Index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = Index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, Index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = Index
                    ' let them know
                    PlayerMsg Index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg Index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg Index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, Index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = Index
        ' let them know
        PlayerMsg Index, "Invitation sent.", Pink
        Exit Sub
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Party_Invite", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub Party_InviteAccept(ByVal Index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if already in a party

   On Error GoTo errorhandler

    If TempPlayer(Index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(Index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) = 0 Then
                'add to the party
                Party(partyNum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                SendPartyVitals partyNum, targetPlayer
                ' let everyone know they've joined
                PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg Index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = Index
        Party(partyNum).Member(1) = Index
        Party(partyNum).Member(2) = targetPlayer
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, Index
        SendPartyVitals partyNum, targetPlayer
        ' let them know it's created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(Index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(Index).inParty = partyNum
        TempPlayer(targetPlayer).inParty = partyNum
        Exit Sub
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Party_InviteAccept", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub Party_InviteDecline(ByVal Index As Long, ByVal targetPlayer As Long)

   On Error GoTo errorhandler

    PlayerMsg Index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Party_InviteDecline", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, x As Long
    ' find the high index

   On Error GoTo errorhandler

    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(x) = Party(partyNum).Member(x + 1)
                    Party(partyNum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partyNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Party_CountMembers", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal Exp As Long, ByVal Index As Long, ByVal MapNum As Long)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long, LoseMemberCount As Byte

    ' check if it's worth sharing

   On Error GoTo errorhandler

    If Not Exp >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP Index, Exp
        Exit Sub
    End If
    
    ' check members in outhers maps
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        If tmpIndex > 0 Then
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) <> MapNum Then
                    LoseMemberCount = LoseMemberCount + 1
                End If
            End If
        End If
    Next i
    
    ' find out the equal share
    expShare = Exp \ (Party(partyNum).MemberCount - LoseMemberCount)
    leftOver = Exp Mod (Party(partyNum).MemberCount - LoseMemberCount)
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) = MapNum Then
                    ' give them their share
                    GivePlayerEXP tmpIndex, expShare
                End If
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(rand(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Party_ShareExp", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub GivePlayerEXP(ByVal Index As Long, ByVal Exp As Long)
    ' give the exp

   On Error GoTo errorhandler

    Call SetPlayerExp(Index, GetPlayerExp(Index) + Exp)
    SendActionMsg GetPlayerMap(Index), "+" & Exp & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
    ' check if we've leveled
    CheckPlayerLevelUp Index
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive Then
        If Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).LevelingType = 0 Then
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.Exp = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Exp + (Exp * (Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).ExpGain / 100))
            SendActionMsg GetPlayerMap(Index), "+" & (Exp * (Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).ExpGain / 100)) & " EXP", White, 1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32)
            CheckPetLevelUp Index
        End If
    End If
    SendEXP Index
    SendPlayerData Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "GivePlayerEXP", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function CanEventMove(Index As Long, ByVal MapNum As Long, x As Long, y As Long, eventID As Long, WalkThrough As Long, ByVal Dir As Byte, Optional globalevent As Boolean = False) As Boolean
    Dim i As Long
    Dim n As Long, z As Long, p As Long, begineventprocessing As Boolean

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    CanEventMove = True
    
    

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(MapNum).Tile(x, y - 1).type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If
                
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = y - 1) Then
                            CanEventMove = False
                            'There IS a player in the way. But now maybe we can call the event touch thingy!
                            If Map(MapNum).Events(eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).Trigger = 1 Then
                                begineventprocessing = True
                                If begineventprocessing = True Then
                                    'Process this event, it is on-touch and everything checks out.
                                    If Map(MapNum).Events(eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).CommandListCount > 0 Then
                                        TempPlayer(Index).EventProcessing(eventID).Active = 1
                                        TempPlayer(Index).EventProcessing(eventID).ActionTimer = GetTickCount
                                        TempPlayer(Index).EventProcessing(eventID).CurList = 1
                                        TempPlayer(Index).EventProcessing(eventID).CurSlot = 1
                                        TempPlayer(Index).EventProcessing(eventID).eventID = eventID
                                        TempPlayer(Index).EventProcessing(eventID).pageID = TempPlayer(Index).EventMap.EventPages(eventID).pageID
                                        TempPlayer(Index).EventProcessing(eventID).WaitingForResponse = 0
                                        ReDim TempPlayer(Index).EventProcessing(eventID).ListLeftOff(0 To Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).CommandListCount)
                                    End If
                                    begineventprocessing = False
                                End If
                            End If
                        End If
                    End If
                Next
                
                If CanEventMove = False Then Exit Function
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(MapNum).Npc(i).x = x) And (MapNpc(MapNum).Npc(i).y = y - 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(MapNum).EventCount > 0 Then
                        For z = 1 To TempEventMap(MapNum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(MapNum).Events(z).x = x) And (TempEventMap(MapNum).Events(z).y = y - 1) And (TempPlayer(Index).EventMap.EventPages(z).WalkThrough = 0) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
                            If (TempPlayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(Index).EventMap.EventPages(z).x = TempPlayer(Index).EventMap.EventPages(eventID).x) And (TempPlayer(Index).EventMap.EventPages(z).y = TempPlayer(Index).EventMap.EventPages(eventID).y - 1) And (TempPlayer(Index).EventMap.EventPages(z).WalkThrough = 0) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(x, y).DirBlock, DIR_UP + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(x, y + 1).type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = y + 1) Then
                            CanEventMove = False
                            'There IS a player in the way. But now maybe we can call the event touch thingy!
                            If Map(MapNum).Events(eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).Trigger = 1 Then
                                begineventprocessing = True
                                If begineventprocessing = True Then
                                    'Process this event, it is on-touch and everything checks out.
                                    If Map(MapNum).Events(eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).CommandListCount > 0 Then
                                        TempPlayer(Index).EventProcessing(eventID).Active = 1
                                        TempPlayer(Index).EventProcessing(eventID).ActionTimer = GetTickCount
                                        TempPlayer(Index).EventProcessing(eventID).CurList = 1
                                        TempPlayer(Index).EventProcessing(eventID).CurSlot = 1
                                        TempPlayer(Index).EventProcessing(eventID).eventID = eventID
                                        TempPlayer(Index).EventProcessing(eventID).pageID = TempPlayer(Index).EventMap.EventPages(eventID).pageID
                                        TempPlayer(Index).EventProcessing(eventID).WaitingForResponse = 0
                                        ReDim TempPlayer(Index).EventProcessing(eventID).ListLeftOff(0 To Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).CommandListCount)
                                    End If
                                    begineventprocessing = False
                                End If
                            End If
                        End If
                    End If
                Next
                
                If CanEventMove = False Then Exit Function

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(MapNum).Npc(i).x = x) And (MapNpc(MapNum).Npc(i).y = y + 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(MapNum).EventCount > 0 Then
                        For z = 1 To TempEventMap(MapNum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(MapNum).Events(z).x = x) And (TempEventMap(MapNum).Events(z).y = y + 1) And (TempPlayer(Index).EventMap.EventPages(z).WalkThrough = 0) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
                            If (TempPlayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(Index).EventMap.EventPages(z).x = TempPlayer(Index).EventMap.EventPages(eventID).x) And (TempPlayer(Index).EventMap.EventPages(z).y = TempPlayer(Index).EventMap.EventPages(eventID).y + 1) And (TempPlayer(Index).EventMap.EventPages(z).WalkThrough = 0) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(x, y).DirBlock, DIR_DOWN + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(MapNum).Tile(x - 1, y).type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = x - 1) And (GetPlayerY(i) = y) Then
                            CanEventMove = False
                            'There IS a player in the way. But now maybe we can call the event touch thingy!
                            If Map(MapNum).Events(eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).Trigger = 1 Then
                                begineventprocessing = True
                                If begineventprocessing = True Then
                                    'Process this event, it is on-touch and everything checks out.
                                    If Map(MapNum).Events(eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).CommandListCount > 0 Then
                                        TempPlayer(Index).EventProcessing(eventID).Active = 1
                                        TempPlayer(Index).EventProcessing(eventID).ActionTimer = GetTickCount
                                        TempPlayer(Index).EventProcessing(eventID).CurList = 1
                                        TempPlayer(Index).EventProcessing(eventID).CurSlot = 1
                                        TempPlayer(Index).EventProcessing(eventID).eventID = eventID
                                        TempPlayer(Index).EventProcessing(eventID).pageID = TempPlayer(Index).EventMap.EventPages(eventID).pageID
                                        TempPlayer(Index).EventProcessing(eventID).WaitingForResponse = 0
                                        ReDim TempPlayer(Index).EventProcessing(eventID).ListLeftOff(0 To Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).CommandListCount)
                                    End If
                                    begineventprocessing = False
                                End If
                            End If
                        End If
                    End If
                Next
                
                If CanEventMove = False Then Exit Function

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(MapNum).Npc(i).x = x - 1) And (MapNpc(MapNum).Npc(i).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(MapNum).EventCount > 0 Then
                        For z = 1 To TempEventMap(MapNum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(MapNum).Events(z).x = x - 1) And (TempEventMap(MapNum).Events(z).y = y) And (TempPlayer(Index).EventMap.EventPages(z).WalkThrough = 0) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
                            If (TempPlayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(Index).EventMap.EventPages(z).x = TempPlayer(Index).EventMap.EventPages(eventID).x - 1) And (TempPlayer(Index).EventMap.EventPages(z).y = TempPlayer(Index).EventMap.EventPages(eventID).y) And (TempPlayer(Index).EventMap.EventPages(z).WalkThrough = 0) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(x, y).DirBlock, DIR_LEFT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(x + 1, y).type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = x + 1) And (GetPlayerY(i) = y) Then
                            CanEventMove = False
                            'There IS a player in the way. But now maybe we can call the event touch thingy!
                            If Map(MapNum).Events(eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).Trigger = 1 Then
                                begineventprocessing = True
                                If begineventprocessing = True Then
                                    'Process this event, it is on-touch and everything checks out.
                                    If Map(MapNum).Events(eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).CommandListCount > 0 Then
                                        TempPlayer(Index).EventProcessing(eventID).Active = 1
                                        TempPlayer(Index).EventProcessing(eventID).ActionTimer = GetTickCount
                                        TempPlayer(Index).EventProcessing(eventID).CurList = 1
                                        TempPlayer(Index).EventProcessing(eventID).CurSlot = 1
                                        TempPlayer(Index).EventProcessing(eventID).eventID = eventID
                                        TempPlayer(Index).EventProcessing(eventID).pageID = TempPlayer(Index).EventMap.EventPages(eventID).pageID
                                        TempPlayer(Index).EventProcessing(eventID).WaitingForResponse = 0
                                        ReDim TempPlayer(Index).EventProcessing(eventID).ListLeftOff(0 To Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventID).pageID).CommandListCount)
                                    End If
                                    begineventprocessing = False
                                End If
                            End If
                        End If
                    End If
                Next
                
                If CanEventMove = False Then Exit Function

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(MapNum).Npc(i).x = x + 1) And (MapNpc(MapNum).Npc(i).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(MapNum).EventCount > 0 Then
                        For z = 1 To TempEventMap(MapNum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(MapNum).Events(z).x = x + 1) And (TempEventMap(MapNum).Events(z).y = y) And (TempPlayer(Index).EventMap.EventPages(z).WalkThrough = 0) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
                            If (TempPlayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(Index).EventMap.EventPages(z).x = TempPlayer(Index).EventMap.EventPages(eventID).x + 1) And (TempPlayer(Index).EventMap.EventPages(z).y = TempPlayer(Index).EventMap.EventPages(eventID).y) And (TempPlayer(Index).EventMap.EventPages(z).WalkThrough = 0) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(x, y).DirBlock, DIR_RIGHT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

    End Select


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanEventMove", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub EventDir(PlayerIndex As Long, ByVal MapNum As Long, ByVal eventID As Long, ByVal Dir As Long, Optional globalevent As Boolean = False)
    Dim Buffer As clsBuffer
    Dim eventIndex As Long, i As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent = False Then
        If TempPlayer(PlayerIndex).EventMap.CurrentEvents > 0 Then
            For i = 1 To TempPlayer(PlayerIndex).EventMap.CurrentEvents
                If eventID = i Then
                    eventIndex = eventID
                    eventID = TempPlayer(PlayerIndex).EventMap.EventPages(i).eventID
                    Exit For
                End If
            Next
        End If
        
        If eventIndex = 0 Or eventID = 0 Then Exit Sub
    End If
    
    If globalevent Then
        If Map(MapNum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(MapNum).Events(eventID).Dir = Dir
    Else
        If Map(MapNum).Events(eventID).Pages(TempPlayer(PlayerIndex).EventMap.EventPages(eventIndex).pageID).DirFix = 0 Then TempPlayer(PlayerIndex).EventMap.EventPages(eventIndex).Dir = Dir
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventDir
    Buffer.WriteLong eventID
    If globalevent Then
        Buffer.WriteLong TempEventMap(MapNum).Events(eventID).Dir
    Else
        Buffer.WriteLong TempPlayer(PlayerIndex).EventMap.EventPages(eventIndex).Dir
    End If
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EventDir", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub EventMove(Index As Long, MapNum As Long, ByVal eventID As Long, ByVal Dir As Long, movementspeed As Long, Optional globalevent As Boolean = False)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim eventIndex As Long, i As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent = False Then
        If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
            For i = 1 To TempPlayer(Index).EventMap.CurrentEvents
                If eventID = i Then
                    eventIndex = eventID
                    eventID = TempPlayer(Index).EventMap.EventPages(i).eventID
                    Exit For
                End If
            Next
        End If
        
        If eventIndex = 0 Or eventID = 0 Then Exit Sub
    Else
        eventIndex = eventID
        If eventIndex = 0 Then Exit Sub
    End If
    
    If globalevent Then
        If Map(MapNum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(MapNum).Events(eventID).Dir = Dir
        UpdateMapBlock MapNum, TempEventMap(MapNum).Events(eventID).x, TempEventMap(MapNum).Events(eventID).y, False
    Else
        If Map(MapNum).Events(eventID).Pages(TempPlayer(Index).EventMap.EventPages(eventIndex).pageID).DirFix = 0 Then TempPlayer(Index).EventMap.EventPages(eventIndex).Dir = Dir
    End If

    Select Case Dir
        Case DIR_UP
            If globalevent Then
                TempEventMap(MapNum).Events(eventIndex).y = TempEventMap(MapNum).Events(eventIndex).y - 1
                UpdateMapBlock MapNum, TempEventMap(MapNum).Events(eventIndex).x, TempEventMap(MapNum).Events(eventIndex).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).x
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap MapNum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(Index).EventMap.EventPages(eventIndex).y = TempPlayer(Index).EventMap.EventPages(eventIndex).y - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).x
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap MapNum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
            
        Case DIR_DOWN
            If globalevent Then
                TempEventMap(MapNum).Events(eventIndex).y = TempEventMap(MapNum).Events(eventIndex).y + 1
                UpdateMapBlock MapNum, TempEventMap(MapNum).Events(eventIndex).x, TempEventMap(MapNum).Events(eventIndex).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).x
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap MapNum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(Index).EventMap.EventPages(eventIndex).y = TempPlayer(Index).EventMap.EventPages(eventIndex).y + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).x
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap MapNum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
        Case DIR_LEFT
            If globalevent Then
                TempEventMap(MapNum).Events(eventIndex).x = TempEventMap(MapNum).Events(eventIndex).x - 1
                UpdateMapBlock MapNum, TempEventMap(MapNum).Events(eventIndex).x, TempEventMap(MapNum).Events(eventIndex).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).x
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap MapNum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(Index).EventMap.EventPages(eventIndex).x = TempPlayer(Index).EventMap.EventPages(eventIndex).x - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).x
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap MapNum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
        Case DIR_RIGHT
            If globalevent Then
                TempEventMap(MapNum).Events(eventIndex).x = TempEventMap(MapNum).Events(eventIndex).x + 1
                UpdateMapBlock MapNum, TempEventMap(MapNum).Events(eventIndex).x, TempEventMap(MapNum).Events(eventIndex).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).x
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(MapNum).Events(eventIndex).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap MapNum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(Index).EventMap.EventPages(eventIndex).x = TempPlayer(Index).EventMap.EventPages(eventIndex).x + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).x
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(Index).EventMap.EventPages(eventIndex).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap MapNum, Buffer.ToArray()
                Else
                    SendDataTo Index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EventMove", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SpawnZoneNpc(ByVal ZoneNum As Long, ByVal ZoneNPCNum As Long)
    Dim Buffer As clsBuffer
    Dim npcnum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long, n As Long, j As Long
    Dim Spawned As Boolean, MapNum As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If ZoneNPCNum <= 0 Or ZoneNPCNum > MAX_MAP_NPCS * 2 Then Exit Sub
    npcnum = MapZones(ZoneNum).NPCs(ZoneNPCNum)
    If npcnum > 0 Then
    
        'We need to find a map
        If MapZones(ZoneNum).MapCount > 0 Then
        
            i = Random(1, MapZones(ZoneNum).MapCount)
            If i > MapZones(ZoneNum).MapCount Then i = MapZones(ZoneNum).MapCount
            If i < 1 Then i = 1
            MapNum = MapZones(ZoneNum).Maps(i)
        End If
        
        If MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num = npcnum
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Target = 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).TargetType = 0 ' clear
        
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP) = GetNpcMaxVital(npcnum, Vitals.HP)
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.MP) = GetNpcMaxVital(npcnum, Vitals.MP)
        
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Dir = Int(Rnd * 4)
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Map = MapNum
        
        'ZoneNpc(ZoneNum).NPC(ZoneNPCNum).StunDuration = 0
        'ZoneNpc(ZoneNum).NPC(ZoneNPCNum).StunTimer = 0
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, Map(MapNum).MaxX)
                y = Random(0, Map(MapNum).MaxY)
    
                If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
                If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(MapNum, x, y) Then
                    ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x = x
                    ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(MapNum).MaxX
                For y = 0 To Map(MapNum).MaxY

                    If NpcTileIsOpen(MapNum, x, y) Then
                        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x = x
                        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong ZoneNum
            Buffer.WriteLong ZoneNPCNum
            Buffer.WriteLong ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
            Buffer.WriteLong ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x
            Buffer.WriteLong ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y
            Buffer.WriteLong ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Dir
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
            UpdateMapBlock MapNum, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y, True
        End If
        
        j = 1
        For n = 1 To MAX_NPC_DROPS
            If Npc(npcnum).DropItems(n) = 0 Then Exit For

            If Rnd <= Npc(npcnum).DropChances(n) / 100 Then
                ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(j).Num = Npc(npcnum).DropItems(n)
                ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(j).Value = Npc(npcnum).DropItemValues(n)
                j = j + 1
            End If
        Next
        
        SendZoneNpcVitals ZoneNum, ZoneNPCNum
    Else
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num = 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Target = 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).TargetType = 0 ' clear
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong ZoneNum
        Buffer.WriteLong ZoneNPCNum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnZoneNpc", "modGameLogic", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
