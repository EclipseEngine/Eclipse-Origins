Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal Index As Long, ByVal slot As Long)

   On Error GoTo errorhandler

    If Not IsPlaying(Index) Then
        TempPlayer(Index).CurChar = slot
        If Options.StaffOnly = 1 Then
            If GetPlayerAccess(Index) = 0 Then
                AlertMsg Index, "Server server is in staff-only mode. Please check back later!"
                Exit Sub
            End If
        End If
        Call JoinGame(Index)
        Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUseChar", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub JoinGame(ByVal Index As Long)
    Dim i As Long
    
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    SendMaxes Index
    

    
    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
    
    ' send the login ok
    SendLoginOk Index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendAnimations(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendResources(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    Call SendPlayerSpells(Index)
    Call SendHotbar(Index)
    Call SendHouseConfigs(Index)
    Call SendUnreadMail(Index)
    Call SendPlayerFriends(Index)
    Call SendQuests(Index)
    Call SendPlayerQuests(Index)
    Call SendPets(Index)
    Call SendProjectiles(Index)
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, i)
    Next
    SendEXP Index
    Call SendStats(Index)
    
    If Player(Index).characters(TempPlayer(Index).CurChar).InHouse Then
        Player(Index).characters(TempPlayer(Index).CurChar).InHouse = 0
        Player(Index).characters(TempPlayer(Index).CurChar).x = Player(Index).characters(TempPlayer(Index).CurChar).LastX
        Player(Index).characters(TempPlayer(Index).CurChar).y = Player(Index).characters(TempPlayer(Index).CurChar).LastY
        Player(Index).characters(TempPlayer(Index).CurChar).Map = Player(Index).characters(TempPlayer(Index).CurChar).LastMap
    End If
    
    If GetPlayerMap(Index) < 0 Then
        Call SetPlayerMap(Index, GetPlayerMap(Index) * -1)
    End If
    
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & Options.Game_Name & "!", Cyan)
    Else
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & Options.Game_Name & "!", BrightRed)
    End If
    
    For i = 1 To Player_HighIndex
        If i <> Index Then
            If IsPlaying(i) Then
                SendPlayerFriends i
            End If
        End If
    Next
    
    ' Send welcome messages
    Call SendWelcome(Index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next
    
    ' Send the flag so they know they can start doing stuff
    SendInGame Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "JoinGame", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LeftGame(ByVal Index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    

   On Error GoTo errorhandler

    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(Index).InTrade > 0 Then
            tradeTarget = TempPlayer(Index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave Index

        ' save and clear data.
        Call SavePlayer(Index)
        Call SaveBank(Index)
        Call ClearBank(Index)

        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(Index)
        
        For i = 1 To Player_HighIndex
            If i <> Index Then
                If IsPlaying(i) Then
                    SendPlayerFriends i
                End If
            End If
        Next
        
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LeftGame", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerProtection(ByVal Index As Long) As Long
    Dim armor As Long
    Dim Helm As Long

   On Error GoTo errorhandler

    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then
        Exit Function
    End If

    armor = GetPlayerEquipment(Index, armor)
    Helm = GetPlayerEquipment(Index, Helmet)
    GetPlayerProtection = (GetPlayerStat(Index, Stats.Endurance) \ 5)

    If armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerProtection", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean

   On Error GoTo errorhandler

    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(Index, Weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Strength) \ 2) + (GetPlayerLevel(Index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPlayerCriticalHit", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long

   On Error GoTo errorhandler

    ShieldSlot = GetPlayerEquipment(Index, Shield)

    If ShieldSlot > 0 Then
        n = Random(1, 2)

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Endurance) / 2) + (GetPlayerLevel(Index) / 2)
            n = rand(1, 100)

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPlayerBlockHit", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional HouseTeleport As Boolean = False)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= MIN_MAPS Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    If MapNum = GetPlayerMap(Index) Then
        SendPlayerXYToMap Index
        Call CheckTasks(Index, TASK_GOTOMAP, MapNum)
    End If
    
    TempPlayer(Index).EventProcessingCount = 0
    TempPlayer(Index).EventMap.CurrentEvents = 0
    
    If HouseTeleport = False Then
        Player(Index).characters(TempPlayer(Index).CurChar).InHouse = 0
    End If
    
    If Player(Index).characters(TempPlayer(Index).CurChar).InHouse > 0 Then
        If IsPlaying(Player(Index).characters(TempPlayer(Index).CurChar).InHouse) Then
            If Player(Player(Index).characters(TempPlayer(Index).CurChar).InHouse).characters(TempPlayer(Index).CurChar).InHouse <> Player(Index).characters(TempPlayer(Index).CurChar).InHouse Then
                Player(Index).characters(TempPlayer(Index).CurChar).InHouse = 0
                PlayerWarp Index, Player(Index).characters(TempPlayer(Index).CurChar).LastMap, Player(Index).characters(TempPlayer(Index).CurChar).LastX, Player(Index).characters(TempPlayer(Index).CurChar).LastY
                Exit Sub
            Else
                SendFurnitureToHouse Player(Index).characters(TempPlayer(Index).CurChar).InHouse
            End If
        End If
    End If
    
    ' clear target
    TempPlayer(Index).Target = 0
    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
    TempPlayer(Index).TargetZone = 0
    SendTarget Index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If OldMap <> MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If
    
    UpdateMapBlock OldMap, GetPlayerX(Index), GetPlayerY(Index), False
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive Then
        UpdateMapBlock GetPlayerMap(Index), Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, False
    End If
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.x = x
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.y = y
    TempPlayer(Index).PetTarget = 0
    TempPlayer(Index).PetTargetType = 0
    SendPlayerXY Index
    UpdateMapBlock MapNum, x, y, True
    
    ' send player's equipment to new map
    SendMapEquipment Index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(MapNum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    SendMapEquipmentTo i, Index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS

            If MapNpc(OldMap).Npc(i).Num > 0 Then
                MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HP)
            End If

        Next

    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    TempPlayer(Index).GettingMap = YES
    Call CheckTasks(Index, TASK_GOTOMAP, MapNum)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong MapNum
    Buffer.WriteLong Map(MapNum).Revision
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerWarp", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, MapNum As Long, i As Long
    Dim x As Long, y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, Amount As Long, begineventprocessing As Boolean

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    Moved = NO
    MapNum = GetPlayerMap(Index)
    
    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_UP + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_KEY And temptile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(Index)).Up).MaxY
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                    TempPlayer(Index).TargetZone = 0
                    SendTarget Index
                End If
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < Map(MapNum).MaxY Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_DOWN + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_KEY And temptile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Down > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                    TempPlayer(Index).TargetZone = 0
                    SendTarget Index
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_KEY And temptile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(Index)).Left).MaxX
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, NewMapX, GetPlayerY(Index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                    TempPlayer(Index).TargetZone = 0
                    SendTarget Index
                End If
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < Map(MapNum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_KEY And temptile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Right > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                    TempPlayer(Index).TargetZone = 0
                    SendTarget Index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .type = TILE_TYPE_WARP Then
            MapNum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(Index, MapNum, x, y)
            Moved = YES
        End If
    
        ' Check to see if the tile is a door tile, and if so warp them
        If .type = TILE_TYPE_DOOR Then
            MapNum = .Data1
            x = .Data2
            y = .Data3
            ' send the animation to the map
            SendDoorAnimation GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
            Call PlayerWarp(Index, MapNum, x, y)
            Moved = YES
        End If
    
        ' Check for key trigger open
        If .type = TILE_TYPE_KEYOPEN Then
            x = .Data1
            y = .Data2
    
            If Map(GetPlayerMap(Index)).Tile(x, y).type = TILE_TYPE_KEY And temptile(GetPlayerMap(Index)).DoorOpen(x, y) = NO Then
                temptile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                temptile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                SendMapKey Index, x, y, 1
                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop Index, x
                    TempPlayer(Index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .type = TILE_TYPE_BANK Then
            SendBank Index
            TempPlayer(Index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            Amount = .Data2
            If Not GetPlayerVital(Index, VitalType) = GetPlayerMaxVital(Index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = BrightGreen
                Else
                    Colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(Index), "+" & Amount, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
                SetPlayerVital Index, VitalType, GetPlayerVital(Index, VitalType) + Amount
                PlayerMsg Index, "You feel rejuvinating forces flowing through your body.", BrightGreen
                Call SendVital(Index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .type = TILE_TYPE_TRAP Then
            Amount = .Data1
            If NewOptions.CombatMode = 1 Then
                Amount = Amount - GetPlayerStat(Index, Willpower)
                If Amount < 0 Then Amount = 0
            End If
            SendActionMsg GetPlayerMap(Index), "-" & Amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
            If GetPlayerVital(Index, HP) - Amount <= 0 Then
                KillPlayer Index
                PlayerMsg Index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital Index, HP, GetPlayerVital(Index, HP) - Amount
                PlayerMsg Index, "You're injured by a trap.", BrightRed
                Call SendVital(Index, HP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            Moved = YES
        End If
        
        ' Slide
        If .type = TILE_TYPE_SLIDE Then
            ForcePlayerMove Index, MOVING_WALKING, .Data1
            Moved = YES
        End If
        
        ' Instance
        If .type = TILE_TYPE_INSTANCE Then
            'Call PlayerWarpToInstance(Index, .Data1, .Data2, .Data3)
        End If
        
        ' Random Dungeon
        If .type = TILE_TYPE_RANDOMDUNGEON Then
            'Call GenerateRandomDungeon(Index, .Data1, .Data2)
        End If
        
        If .type = TILE_TYPE_HOUSE Then
            If Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex = .Data1 Then
                'Do warping and such to the player's house :/
                Player(Index).characters(TempPlayer(Index).CurChar).LastMap = GetPlayerMap(Index)
                Player(Index).characters(TempPlayer(Index).CurChar).LastX = GetPlayerX(Index)
                Player(Index).characters(TempPlayer(Index).CurChar).LastY = GetPlayerY(Index)
                Player(Index).characters(TempPlayer(Index).CurChar).InHouse = Index
                SendDataTo Index, PlayerData(Index)
                Call PlayerWarp(Index, HouseConfig(Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex).BaseMap, HouseConfig(Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex).x, HouseConfig(Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex).y, True)
                Exit Sub
            Else
                'Send the buy sequence and see what happens. (To be recreated in events.)
                Set Buffer = New clsBuffer
                Buffer.WriteLong SBuyHouse
                Buffer.WriteLong .Data1
                SendDataTo Index, Buffer.ToArray
                Set Buffer = Nothing
                TempPlayer(Index).BuyHouseIndex = .Data1
            End If
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
    End If
    
    x = GetPlayerX(Index)
    y = GetPlayerY(Index)
    
    If Moved = YES Then
        If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
            For i = 1 To TempPlayer(Index).EventMap.CurrentEvents
                If Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Global = 1 Then
                    If Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).x = x And Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).y = y And Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).Trigger = 1 And TempPlayer(Index).EventMap.EventPages(i).Visible = 1 Then begineventprocessing = True
                Else
                    If TempPlayer(Index).EventMap.EventPages(i).x = x And TempPlayer(Index).EventMap.EventPages(i).y = y And Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).Trigger = 1 And TempPlayer(Index).EventMap.EventPages(i).Visible = 1 Then begineventprocessing = True
                End If
                begineventprocessing = False
                If begineventprocessing = True Then
                    'Process this event, it is on-touch and everything checks out.
                    If Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).CommandListCount > 0 Then
                        TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).Active = 1
                        TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).ActionTimer = GetTickCount
                        TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).CurList = 1
                        TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).CurSlot = 1
                        TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).eventID = TempPlayer(Index).EventMap.EventPages(i).eventID
                        TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).pageID = TempPlayer(Index).EventMap.EventPages(i).pageID
                        TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).WaitingForResponse = 0
                        ReDim TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).ListLeftOff(0 To Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).CommandListCount)
                    End If
                    begineventprocessing = False
                End If
            Next
        End If
        
        If Player(Index).characters(TempPlayer(Index).CurChar).InHouse > 0 Then
            If Player(Index).characters(TempPlayer(Index).CurChar).x = HouseConfig(Player(Player(Index).characters(TempPlayer(Index).CurChar).InHouse).characters(TempPlayer(Index).CurChar).House.HouseIndex).x Then
                If Player(Index).characters(TempPlayer(Index).CurChar).y = HouseConfig(Player(Player(Index).characters(TempPlayer(Index).CurChar).InHouse).characters(TempPlayer(Index).CurChar).House.HouseIndex).y Then
                    Call PlayerWarp(Index, Player(Index).characters(TempPlayer(Index).CurChar).LastMap, Player(Index).characters(TempPlayer(Index).CurChar).LastX, Player(Index).characters(TempPlayer(Index).CurChar).LastY)
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerMove", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ForcePlayerMove(ByVal Index As Long, ByVal movement As Long, ByVal Direction As Long)

   On Error GoTo errorhandler

    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(Index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(Index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove Index, Direction, movement, True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ForcePlayerMove", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub CheckEquippedItems(ByVal Index As Long)
    Dim slot As Long
    Dim ItemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped

   On Error GoTo errorhandler

    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(Index, i)

        If ItemNum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(ItemNum).type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment Index, 0, i
                Case Equipment.armor

                    If Item(ItemNum).type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment Index, 0, i
                Case Equipment.Helmet

                    If Item(ItemNum).type <> ITEM_TYPE_HELMET Then SetPlayerEquipment Index, 0, i
                Case Equipment.Shield

                    If Item(ItemNum).type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment Index, 0, i
            End Select

        Else
            SetPlayerEquipment Index, 0, i
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckEquippedItems", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(ItemNum).Stackable = 1 Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindOpenInvSlot", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function FindOpenBankSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long


   On Error GoTo errorhandler

    If Not IsPlaying(Index) Then Exit Function
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindOpenBankSlot", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Stackable = 1 Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HasItem", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function FindItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            FindItem = i
            Exit Function
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindItem", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function TakeInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal itemval As Long) As Boolean
    Dim i As Long
    Dim n As Long, tookitem As Boolean
    

   On Error GoTo errorhandler

    If itemval = 0 Then itemval = 1

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV
        If itemval > 0 Then
            ' Check to see if the player has the item
            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                If Item(ItemNum).Stackable = 1 Then
    
                    ' Is what we are trying to take away more then what they have?  If so just set it to zero
                    If itemval >= GetPlayerInvItemValue(Index, i) Then
                        itemval = itemval - GetPlayerInvItemValue(Index, i)
                        Call SetPlayerInvItemNum(Index, i, 0)
                        Call SetPlayerInvItemValue(Index, i, 0)
                        ' Send the inventory update
                        Call SendInventoryUpdate(Index, i)
                    Else
                        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - itemval)
                        Call SendInventoryUpdate(Index, i)
                        itemval = 0
                    End If
                Else
                    itemval = itemval - 1
                    Call SetPlayerInvItemNum(Index, i, 0)
                    Call SetPlayerInvItemValue(Index, i, 0)
                    ' Send the inventory update
                    Call SendInventoryUpdate(Index, i)
                End If
    
                If itemval = 0 Then
                    TakeInvItem = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    TakeInvItem = False


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "TakeInvItem", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function TakeInvSlot(ByVal Index As Long, ByVal invslot As Long, ByVal itemval As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ItemNum
    

   On Error GoTo errorhandler

    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or invslot <= 0 Or invslot > MAX_ITEMS Then
        Exit Function
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, invslot)

    If Item(ItemNum).Stackable = 1 Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If itemval >= GetPlayerInvItemValue(Index, invslot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(Index, invslot, GetPlayerInvItemValue(Index, invslot) - itemval)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(Index, invslot, 0)
        Call SetPlayerInvItemValue(Index, invslot, 0)
        Exit Function
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "TakeInvSlot", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function GiveInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal itemval As Long, Optional ByVal sendupdate As Boolean = True, Optional ByVal sendmsg As Boolean = True, Optional ByVal skipTaskCheck As Boolean = False) As Boolean
    Dim i As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(Index, ItemNum)

    ' Check to see if inventory is full
    If i <> 0 Then
        If Item(ItemNum).Stackable = 1 Then
            Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + itemval)
            Call SetPlayerInvItemNum(Index, i, ItemNum)
            If skipTaskCheck = False Then
                Call CheckTasks(Index, TASK_AQUIREITEMS, GetPlayerInvItemNum(Index, i))
                Call CheckTasks(Index, TASK_FETCHRETURN, GetPlayerInvItemNum(Index, i))
            End If
        Else
            Call SetPlayerInvItemNum(Index, i, ItemNum)
            If skipTaskCheck = False Then
                Call CheckTasks(Index, TASK_AQUIREITEMS, GetPlayerInvItemNum(Index, i))
                Call CheckTasks(Index, TASK_FETCHRETURN, GetPlayerInvItemNum(Index, i))
            End If
        End If
        If sendupdate Then Call SendInventoryUpdate(Index, i)
        GiveInvItem = True
    Else
        If sendmsg Then
            Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
        End If
        GiveInvItem = False
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GiveInvItem", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function HasSpell(ByVal Index As Long, ByVal Spellnum As Long) As Boolean
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = Spellnum Then
            HasSpell = True
            Exit Function
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HasSpell", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindOpenSpellSlot", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub PlayerMapGetItem(ByVal Index As Long)
    Dim i As Long
    Dim n As Long
    Dim MapNum As Long
    Dim Msg As String


   On Error GoTo errorhandler

    If Not IsPlaying(Index) Then Exit Sub
    MapNum = GetPlayerMap(Index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(Index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(MapNum, i).x = GetPlayerX(Index)) Then
                    If (MapItem(MapNum, i).y = GetPlayerY(Index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(Index, MapItem(MapNum, i).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, i).Num)
    
                            If Item(GetPlayerInvItemNum(Index, n)).Stackable = 1 Then
                                Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, i).Value)
                                Msg = MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(Index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                            End If
    
                            ' Erase item from the map
                            ClearMapItem i, MapNum
                            
                            Call SendInventoryUpdate(Index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), 0, 0)
                            SendActionMsg GetPlayerMap(Index), Msg, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                            Call CheckTasks(Index, TASK_AQUIREITEMS, GetPlayerInvItemNum(Index, n))
                            Call CheckTasks(Index, TASK_FETCHRETURN, GetPlayerInvItemNum(Index, n))
                            Exit For
                        Else
                            Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerMapGetItem", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function CanPlayerPickupItem(ByVal Index As Long, ByVal mapItemNum As Long)
Dim MapNum As Long


   On Error GoTo errorhandler

    MapNum = GetPlayerMap(Index)
    
    ' no lock or locked to player?
    If MapItem(MapNum, mapItemNum).playerName = vbNullString Or MapItem(MapNum, mapItemNum).playerName = Trim$(GetPlayerName(Index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPlayerPickupItem", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub PlayerMapDropItem(ByVal Index As Long, ByVal invNum As Long, ByVal Amount As Long)
    Dim i As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Or TempPlayer(Index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(Index, invNum) > 0) Then
        If (GetPlayerInvItemNum(Index, invNum) <= MAX_ITEMS) Then
            i = FindOpenMapItemSlot(GetPlayerMap(Index))

            If i <> 0 Then
                MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, invNum)
                MapItem(GetPlayerMap(Index), i).x = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), i).y = GetPlayerY(Index)
                MapItem(GetPlayerMap(Index), i).playerName = Trim$(GetPlayerName(Index))
                MapItem(GetPlayerMap(Index), i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(Index), i).canDespawn = True
                MapItem(GetPlayerMap(Index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME

                If Item(GetPlayerInvItemNum(Index, invNum)).Stackable = 1 Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, invNum) Then
                        MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, invNum)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, invNum, 0)
                        Call SetPlayerInvItemValue(Index, invNum, 0)
                    Else
                        MapItem(GetPlayerMap(Index), i).Value = Amount
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, invNum, GetPlayerInvItemValue(Index, invNum) - Amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), i).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, invNum, 0)
                    Call SetPlayerInvItemValue(Index, invNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(Index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, Amount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), Trim$(GetPlayerName(Index)), MapItem(GetPlayerMap(Index), i).canDespawn)
            Else
                Call PlayerMsg(Index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerMapDropItem", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    

   On Error GoTo errorhandler

    level_count = 0
    
    Do While GetPlayerExp(Index) >= GetPlayerNextLevel(Index)
        expRollover = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
        
        ' can level up?
        If Not SetPlayerLevel(Index, GetPlayerLevel(Index) + 1) Then
            Exit Sub
        End If
        
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + 3)
        Call SetPlayerExp(Index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " level!", Brown
        Else
            'plural
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " levels!", Brown
        End If
        SendEXP Index
        SendPlayerData Index
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckPlayerLevelUp", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal Index As Long) As String

   On Error GoTo errorhandler

    GetPlayerLogin = Trim$(Player(Index).login)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerLogin", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal login As String)

   On Error GoTo errorhandler

    Player(Index).login = login


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerLogin", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String

   On Error GoTo errorhandler

    GetPlayerPassword = Trim$(Player(Index).Password)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerPassword", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)

   On Error GoTo errorhandler

    Player(Index).Password = Password


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerPassword", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerName(ByVal Index As Long) As String


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).characters(TempPlayer(Index).CurChar).Name)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerName", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).Name = Name


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long

   On Error GoTo errorhandler

    GetPlayerClass = Player(Index).characters(TempPlayer(Index).CurChar).Class


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).Class = ClassNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).characters(TempPlayer(Index).CurChar).Level


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function SetPlayerLevel(ByVal Index As Long, ByVal Level As Long) As Boolean

   On Error GoTo errorhandler

    SetPlayerLevel = False
    If Level > MAX_LEVELS Then Exit Function
    Player(Index).characters(TempPlayer(Index).CurChar).Level = Level
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive Then
        If Player(Index).characters(TempPlayer(Index).CurChar).Pet.AdoptiveStats Then
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level = Level
        End If
    End If
    SetPlayerLevel = True


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "SetPlayerLevel", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long

   On Error GoTo errorhandler
    If GetPlayerLevel(Index) = MAX_LEVELS Then GetPlayerNextLevel = 0: Exit Function
    GetPlayerNextLevel = (50 / 3) * (GetPlayerLevel(Index) + 1) ^ 3


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerNextLevel", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetPlayerExp(ByVal Index As Long) As Long

   On Error GoTo errorhandler

    GetPlayerExp = Player(Index).characters(TempPlayer(Index).CurChar).Exp


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).Exp = Exp
    If GetPlayerLevel(Index) = MAX_LEVELS And Player(Index).characters(TempPlayer(Index).CurChar).Exp > GetPlayerNextLevel(Index) Then
        Player(Index).characters(TempPlayer(Index).CurChar).Exp = GetPlayerNextLevel(Index)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If Index < 0 Then Exit Function
    If TempPlayer(Index).CurChar = 0 Then Exit Function
    GetPlayerAccess = Player(Index).characters(TempPlayer(Index).CurChar).access


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal access As Long)
Dim i As Long, x As Long
   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).access = access
    i = FindAccount(Trim$(Player(Index).login))
    If i > 0 Then
        account(i).access = 0
        For x = 1 To MAX_PLAYER_CHARS
            If Player(Index).characters(x).access > account(i).access Then
                account(i).access = Player(Index).characters(x).access
            End If
        Next
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).characters(TempPlayer(Index).CurChar).PK


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerPK", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).PK = PK


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerPK", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long

   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).characters(TempPlayer(Index).CurChar).Vital(Vital)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).characters(TempPlayer(Index).CurChar).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    If GetPlayerVital(Index, Vital) < 0 Then
        Player(Index).characters(TempPlayer(Index).CurChar).Vital(Vital) = 0
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function GetPlayerStat(ByVal Index As Long, ByVal stat As Stats) As Long
    Dim x As Long, i As Long

   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    
    x = Player(Index).characters(TempPlayer(Index).CurChar).stat(stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(Index).characters(TempPlayer(Index).CurChar).Equipment(i) > 0 Then
            If Item(Player(Index).characters(TempPlayer(Index).CurChar).Equipment(i)).Add_Stat(stat) > 0 Then
                x = x + Item(Player(Index).characters(TempPlayer(Index).CurChar).Equipment(i)).Add_Stat(stat)
            End If
        End If
    Next
    
    GetPlayerStat = x


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function GetPlayerRawStat(ByVal Index As Long, ByVal stat As Stats) As Long

   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(Index).characters(TempPlayer(Index).CurChar).stat(stat)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerRawStat", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal stat As Stats, ByVal Value As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).stat(stat) = Value
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive Then
        If Player(Index).characters(TempPlayer(Index).CurChar).Pet.AdoptiveStats Then
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(stat) = Value
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).characters(TempPlayer(Index).CurChar).Points


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal Points As Long)

   On Error GoTo errorhandler

    If Points <= 0 Then Points = 0
    Player(Index).characters(TempPlayer(Index).CurChar).Points = Points


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If IsPlaying(Index) Then
        GetPlayerMap = Player(Index).characters(TempPlayer(Index).CurChar).Map
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)


   On Error GoTo errorhandler

    If MapNum > MIN_MAPS And MapNum <= MAX_MAPS Then
        Player(Index).characters(TempPlayer(Index).CurChar).Map = MapNum
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function GetPlayerX(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).characters(TempPlayer(Index).CurChar).x


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerX", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).x = x


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerY(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).characters(TempPlayer(Index).CurChar).y


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerY", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).y = y


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).characters(TempPlayer(Index).CurChar).Dir


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).Dir = Dir


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerIP(ByVal Index As Long) As String


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerIP", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long) As Long

   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(Index).characters(TempPlayer(Index).CurChar).Inv(invslot).Num


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long, ByVal ItemNum As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).Inv(invslot).Num = ItemNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(Index).characters(TempPlayer(Index).CurChar).Inv(invslot).Value


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long, ByVal itemvalue As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).Inv(invslot).Value = itemvalue


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(Index).characters(TempPlayer(Index).CurChar).Spell(spellslot)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerSpell", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal Spellnum As Long)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).Spell(spellslot) = Spellnum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerSpell", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(Index).characters(TempPlayer(Index).CurChar).Equipment(EquipmentSlot)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)

   On Error GoTo errorhandler

    Player(Index).characters(TempPlayer(Index).CurChar).Equipment(EquipmentSlot) = invNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerEquipment", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ToDo
Sub OnDeath(ByVal Index As Long)
    Dim i As Long
    
    ' Set HP to nothing

   On Error GoTo errorhandler

    Call SetPlayerVital(Index, Vitals.HP, 0)

    ' Warp player away
    Call SetPlayerDir(Index, DIR_DOWN)
    
    With Map(GetPlayerMap(Index))
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp Index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(Index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(Index).spellBuffer.Spell = 0
    TempPlayer(Index).spellBuffer.Timer = 0
    TempPlayer(Index).spellBuffer.Target = 0
    TempPlayer(Index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(Index)
    
    TempPlayer(Index).InBank = False
    TempPlayer(Index).InShop = 0
    If TempPlayer(Index).InTrade > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).Num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).Num = 0
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).Value = 0
        Next
        
        TempPlayer(Index).InTrade = 0
        TempPlayer(TempPlayer(Index).InTrade).InTrade = 0
        
        SendCloseTrade Index
        SendCloseTrade TempPlayer(Index).InTrade
    End If
    
    ' Restore vitals
    Call SetPlayerVital(Index, Vitals.HP, GetPlayerMaxVital(Index, Vitals.HP))
    Call SetPlayerVital(Index, Vitals.MP, GetPlayerMaxVital(Index, Vitals.MP))
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "OnDeath", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub CheckResource(ByVal Index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    
     ' Check attack timer

   On Error GoTo errorhandler

    If GetPlayerEquipment(Index, Weapon) > 0 Then
        If GetTickCount < TempPlayer(Index).AttackTimer + Item(GetPlayerEquipment(Index, Weapon)).speed Then Exit Sub
    Else
        If GetTickCount < TempPlayer(Index).AttackTimer + 1000 Then Exit Sub
    End If
    
    If Map(GetPlayerMap(Index)).Tile(x, y).type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(Index)).Tile(x, y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count

            If ResourceCache(GetPlayerMap(Index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(Index)).ResourceData(i).y = y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(Index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(Index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Then

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(Index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).x
                        rY = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).y
                        
                        Damage = Item(GetPlayerEquipment(Index, Weapon)).Data2
                    
                        ' check if damage is more than health
                        If Damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                SendActionMsg GetPlayerMap(Index), "-" & ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                SendResourceCacheToMap GetPlayerMap(Index), Resource_num
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                End If
                                ' carry on
                                GiveInvItem Index, Resource(Resource_index).ItemReward, 1
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                                Call CheckTasks(Index, TASK_GETRESOURCES, Resource_index)
                                Call CheckTasks(Index, TASK_AQUIREITEMS, Resource(Resource_index).ItemReward)
                                Call CheckTasks(Index, TASK_FETCHRETURN, Resource(Resource_index).ItemReward)
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage
                                SendActionMsg GetPlayerMap(Index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            End If
                            ' send the sound
                            SendMapSound Index, rX, rY, SoundEntity.seResource, Resource_index
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(Index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                        End If
                    End If
                    ' Reset attack timer
                    TempPlayer(Index).AttackTimer = GetTickCount

                Else
                    PlayerMsg Index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                If Resource(Resource_index).ToolRequired = 0 Then
                
                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(Index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).x
                        rY = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).y
                        
                        Damage = GetPlayerDamage(Index)
                    
                        ' check if damage is more than health
                        If Damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                SendActionMsg GetPlayerMap(Index), "-" & ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                SendResourceCacheToMap GetPlayerMap(Index), Resource_num
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                End If
                                ' carry on
                                GiveInvItem Index, Resource(Resource_index).ItemReward, 1
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                                Call CheckTasks(Index, TASK_GETRESOURCES, Resource_index)
                                Call CheckTasks(Index, TASK_AQUIREITEMS, Resource(Resource_index).ItemReward)
                                Call CheckTasks(Index, TASK_FETCHRETURN, Resource(Resource_index).ItemReward)
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage
                                SendActionMsg GetPlayerMap(Index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            End If
                            ' send the sound
                            SendMapSound Index, rX, rY, SoundEntity.seResource, Resource_index
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(Index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                        End If
                    End If
                    ' Reset attack timer
                    TempPlayer(Index).AttackTimer = GetTickCount

                Else
                    PlayerMsg Index, "You have the wrong type of tool equiped.", BrightRed
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckResource", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long

   On Error GoTo errorhandler

    GetPlayerBankItemNum = Bank(Index).Item(BankSlot).Num


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerBankItemNum", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)

   On Error GoTo errorhandler

    Bank(Index).Item(BankSlot).Num = ItemNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerBankItemNum", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long

   On Error GoTo errorhandler

    GetPlayerBankItemValue = Bank(Index).Item(BankSlot).Value


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerBankItemValue", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal itemvalue As Long)

   On Error GoTo errorhandler

    Bank(Index).Item(BankSlot).Value = itemvalue


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerBankItemValue", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal invslot As Long, ByVal Amount As Long)
Dim BankSlot


   On Error GoTo errorhandler

    If invslot < 0 Or invslot > MAX_INV Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invslot) Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(Index, GetPlayerInvItemNum(Index, invslot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, invslot)).Stackable = 1 Then
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, invslot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invslot), Amount)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, invslot))
                Call SetPlayerBankItemValue(Index, BankSlot, Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invslot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, invslot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invslot), 0)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, invslot))
                Call SetPlayerBankItemValue(Index, BankSlot, 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invslot), 0)
            End If
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "GiveBankItem", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
Dim invslot


   On Error GoTo errorhandler

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerBankItemValue(Index, BankSlot) Then
        Exit Sub
    End If
    
    invslot = FindOpenInvSlot(Index, GetPlayerBankItemNum(Index, BankSlot))
        
    If invslot > 0 Then
        If Item(GetPlayerBankItemNum(Index, BankSlot)).Stackable = 1 Then
            Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), Amount)
            Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - Amount)
            If GetPlayerBankItemValue(Index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(Index, BankSlot) > 1 Then
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - 1)
            Else
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TakeBankItem", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub KillPlayer(ByVal Index As Long)
Dim Exp As Long

    ' Calculate exp to give attacker

   On Error GoTo errorhandler
    
    If NewOptions.ExpLoss = 0 Then
        Exp = GetPlayerExp(Index) \ 3
    
        ' Make sure we dont get less then 0
        If Exp < 0 Then Exp = 0
        If Exp = 0 Then
            Call PlayerMsg(Index, "You lost no exp.", BrightRed)
        Else
            Call SetPlayerExp(Index, GetPlayerExp(Index) - Exp)
            SendEXP Index
            Call PlayerMsg(Index, "You lost " & Exp & " exp.", BrightRed)
        End If
    End If
    
    Call OnDeath(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "KillPlayer", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub UseItem(ByVal Index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, x As Long, y As Long, ItemNum As Long, x1 As Long, y1 As Long

    ' Prevent hacking

   On Error GoTo errorhandler

    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(Index, invNum) > 0) And (GetPlayerInvItemNum(Index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(Index, invNum)).Data2
        ItemNum = GetPlayerInvItemNum(Index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(ItemNum).type
            Case ITEM_TYPE_ARMOR
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, armor) > 0 Then
                    tempItem = GetPlayerEquipment(Index, armor)
                End If

                SetPlayerEquipment Index, ItemNum, armor
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 0

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_WEAPON
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Weapon) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Weapon)
                End If

                SetPlayerEquipment Index, ItemNum, Weapon
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_HELMET
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Helmet)
                End If

                SetPlayerEquipment Index, ItemNum, Helmet
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_SHIELD
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Shield)
                End If

                SetPlayerEquipment Index, ItemNum, Shield
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_RING
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, ring) > 0 Then
                    tempItem = GetPlayerEquipment(Index, ring)
                End If

                SetPlayerEquipment Index, ItemNum, ring
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
        Case ITEM_TYPE_NECKLACE
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, NECKLACE) > 0 Then
                    tempItem = GetPlayerEquipment(Index, NECKLACE)
                End If

                SetPlayerEquipment Index, ItemNum, NECKLACE
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_GLOVES
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, GLOVES) > 0 Then
                    tempItem = GetPlayerEquipment(Index, GLOVES)
                End If

                SetPlayerEquipment Index, ItemNum, GLOVES
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_COAT
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, COAT) > 0 Then
                    tempItem = GetPlayerEquipment(Index, COAT)
                End If

                SetPlayerEquipment Index, ItemNum, COAT
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(ItemNum).AddHP > 0 Then
                    Player(Index).characters(TempPlayer(Index).CurChar).Vital(Vitals.HP) = Player(Index).characters(TempPlayer(Index).CurChar).Vital(Vitals.HP) + Item(ItemNum).AddHP
                    If GetPlayerVital(Index, HP) > GetPlayerMaxVital(Index, HP) Then Call SetPlayerVital(Index, HP, GetPlayerMaxVital(Index, HP))
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, HP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                ' add mp
                If Item(ItemNum).AddMP > 0 Then
                    Player(Index).characters(TempPlayer(Index).CurChar).Vital(Vitals.MP) = Player(Index).characters(TempPlayer(Index).CurChar).Vital(Vitals.MP) + Item(ItemNum).AddMP
                    If GetPlayerVital(Index, MP) > GetPlayerMaxVital(Index, MP) Then Call SetPlayerVital(Index, MP, GetPlayerMaxVital(Index, MP))
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, MP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                ' add exp
                If Item(ItemNum).AddEXP > 0 Then
                    SetPlayerExp Index, GetPlayerExp(Index) + Item(ItemNum).AddEXP
                    CheckPlayerLevelUp Index
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendEXP Index
                End If
                Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                Call TakeInvItem(Index, Player(Index).characters(TempPlayer(Index).CurChar).Inv(invNum).Num, 0)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_KEY
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If

                Select Case GetPlayerDir(Index)
                    Case DIR_UP

                        If GetPlayerY(Index) > 0 Then
                            x = GetPlayerX(Index)
                            y = GetPlayerY(Index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(Index) < Map(GetPlayerMap(Index)).MaxY Then
                            x = GetPlayerX(Index)
                            y = GetPlayerY(Index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT

                        If GetPlayerX(Index) > 0 Then
                            x = GetPlayerX(Index) - 1
                            y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT

                        If GetPlayerX(Index) < Map(GetPlayerMap(Index)).MaxX Then
                            x = GetPlayerX(Index) + 1
                            y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(Index)).Tile(x, y).type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If ItemNum = Map(GetPlayerMap(Index)).Tile(x, y).Data1 Then
                        temptile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                        temptile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                        SendMapKey Index, x, y, 1
                        Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(Index)).Tile(x, y).Data2 = 1 Then
                            Call TakeInvItem(Index, ItemNum, 0)
                            Call PlayerMsg(Index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_SPELL
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(ItemNum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq

                        If i <= GetPlayerLevel(Index) Then
                            i = FindOpenSpellSlot(Index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(Index, n) Then
                                    Call SetPlayerSpell(Index, i, n)
                                    Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                                    Call TakeInvItem(Index, ItemNum, 0)
                                    Call PlayerMsg(Index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                    SendPlayerSpells Index
                                Else
                                    Call PlayerMsg(Index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(Index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(Index, "You must be level " & i & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(Index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_FURNITURE
                PlayerMsg Index, "To place furniture, simply drag it into your house where you want it.", White
            Case ITEM_TYPE_PET
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the pet num
                n = Item(ItemNum).Data1

                If n > 0 Then
                    Call SummonPet(Index, n)
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
                
        End Select
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UseItem", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SetPlayerSprite(Index As Long, Sprite As Long)

   On Error GoTo errorhandler

    If Index > 0 And Index <= MAX_PLAYERS And TempPlayer(Index).CurChar > 0 Then
        Player(Index).characters(TempPlayer(Index).CurChar).Face(FaceEnum.Head) = Sprite
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SetPlayerSprite", "modPlayer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
