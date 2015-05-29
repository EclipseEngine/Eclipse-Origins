Attribute VB_Name = "modEventLogic"
Option Explicit
Public Sub RemoveDeadEvents()
Dim i As Long, MapNum As Long, Buffer As clsBuffer, x As Long, id As Long, page As Long, compare As Long
    'See if we should remove any events....

   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex
        If IsPlaying(i) = False Then TempPlayer(i).EventMap.CurrentEvents = 0: Exit Sub
        If TempPlayer(i).EventMap.CurrentEvents > 0 Then
            MapNum = GetPlayerMap(i)
            For x = 1 To TempPlayer(i).EventMap.CurrentEvents
                id = TempPlayer(i).EventMap.EventPages(x).eventID
                page = TempPlayer(i).EventMap.EventPages(x).pageID
                If Map(MapNum).Events(id).PageCount >= page Then
                
                    'See if there is any reason to delete this event....
                    'In other words, go back through conditions and make sure they all check up.
                    If TempPlayer(i).EventMap.EventPages(x).Visible = 1 Then
                        If Map(MapNum).Events(id).Pages(page).chkHasItem = 1 Then
                            If HasItem(i, Map(MapNum).Events(id).Pages(page).HasItemIndex) = 0 Then
                                TempPlayer(i).EventMap.EventPages(x).Visible = 0
                            End If
                        End If
                        
                        
                        
                        If Map(MapNum).Events(id).Pages(page).chkSelfSwitch = 1 Then
                            If Map(MapNum).Events(id).Pages(page).SelfSwitchCompare = 0 Then
                                compare = 1
                            Else
                                compare = 0
                            End If
                            If Map(MapNum).Events(id).Global = 1 Then
                                If Map(MapNum).Events(id).SelfSwitches(Map(MapNum).Events(id).Pages(page).SelfSwitchIndex) <> compare Then
                                    TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                End If
                            Else
                                If TempPlayer(i).EventMap.EventPages(id).SelfSwitches(Map(MapNum).Events(id).Pages(page).SelfSwitchIndex) <> compare Then
                                    TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                End If
                            End If
                        End If
                        
                        If Map(MapNum).Events(id).Pages(page).chkVariable = 1 Then
                            Select Case Map(MapNum).Events(id).Pages(page).VariableCompare
                                Case 0
                                    If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(page).VariableIndex) <> Map(MapNum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 1
                                    If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(page).VariableIndex) < Map(MapNum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 2
                                    If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(page).VariableIndex) > Map(MapNum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 3
                                    If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(page).VariableIndex) <= Map(MapNum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 4
                                    If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(page).VariableIndex) >= Map(MapNum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 5
                                    If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(page).VariableIndex) = Map(MapNum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                            End Select
                        End If
                        
                        If Map(MapNum).Events(id).Pages(page).chkSwitch = 1 Then
                            If Map(MapNum).Events(id).Pages(page).SwitchCompare = 1 Then
                                If Player(i).characters(TempPlayer(i).CurChar).Switches(Map(MapNum).Events(id).Pages(page).SwitchIndex) = 1 Then
                                    TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                End If
                            Else
                                If Player(i).characters(TempPlayer(i).CurChar).Switches(Map(MapNum).Events(id).Pages(page).SwitchIndex) = 0 Then
                                    TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                End If
                            End If
                        End If
                        
                        If Map(MapNum).Events(id).Global = 1 And TempPlayer(i).EventMap.EventPages(x).Visible = 0 Then TempEventMap(MapNum).Events(id).Active = 0
                        
                        If TempPlayer(i).EventMap.EventPages(x).Visible = 0 Then
                            Set Buffer = New clsBuffer
                            Buffer.WriteLong SSpawnEvent
                            Buffer.WriteLong id
                            With TempPlayer(i).EventMap.EventPages(x)
                                Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Name
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
                                Buffer.WriteLong Map(MapNum).Events(id).Pages(page).WalkAnim
                                Buffer.WriteLong Map(MapNum).Events(id).Pages(page).DirFix
                                Buffer.WriteLong Map(MapNum).Events(id).Pages(page).WalkThrough
                                Buffer.WriteLong Map(MapNum).Events(id).Pages(page).ShowName
                                Buffer.WriteLong .questnum
                            End With
                            SendDataTo i, Buffer.ToArray
                            Set Buffer = Nothing
                        End If
                    End If
                End If
            Next
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "RemoveDeadEvents", "modEventLogic", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Public Sub SpawnNewEvents()
Dim Buffer As clsBuffer, pageID As Long, id As Long, compare As Long, i As Long, MapNum As Long, x As Long, z As Long, spawnevent As Boolean, p As Long
Dim n As Long, q As Long, r As Long
    'That was only removing events... now we gotta worry about spawning them again, luckily, it is almost the same exact thing, but backwards!

   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex
        If TempPlayer(i).EventMap.CurrentEvents > 0 Then
            MapNum = GetPlayerMap(i)
            If MapNum = 0 Then Exit Sub
            For x = 1 To TempPlayer(i).EventMap.CurrentEvents
                id = TempPlayer(i).EventMap.EventPages(x).eventID
                pageID = TempPlayer(i).EventMap.EventPages(x).pageID
                If TempPlayer(i).EventMap.EventPages(x).Visible = 0 Then pageID = 0
                
                For z = Map(MapNum).Events(id).PageCount To 1 Step -1
                        
                    spawnevent = True
                        
                    If Map(MapNum).Events(id).Pages(z).chkHasItem = 1 Then
                        If HasItem(i, Map(MapNum).Events(id).Pages(z).HasItemIndex) = 0 Then
                            spawnevent = False
                        End If
                    End If
                        
                    If Map(MapNum).Events(id).Pages(z).chkSelfSwitch = 1 Then
                        If Map(MapNum).Events(id).Pages(z).SelfSwitchCompare = 0 Then
                            compare = 1
                        Else
                            compare = 0
                        End If
                        If Map(MapNum).Events(id).Global = 1 Then
                            If Map(MapNum).Events(id).SelfSwitches(Map(MapNum).Events(id).Pages(z).SelfSwitchIndex) <> compare Then
                                spawnevent = False
                            End If
                        Else
                            If TempPlayer(i).EventMap.EventPages(id).SelfSwitches(Map(MapNum).Events(id).Pages(z).SelfSwitchIndex) <> compare Then
                                spawnevent = False
                            End If
                        End If
                    End If
                
                        
                    If Map(MapNum).Events(id).Pages(z).chkVariable = 1 Then
                        Select Case Map(MapNum).Events(id).Pages(z).VariableCompare
                            Case 0
                                If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(z).VariableIndex) <> Map(MapNum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 1
                                If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(z).VariableIndex) < Map(MapNum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 2
                                If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(z).VariableIndex) > Map(MapNum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 3
                                If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(z).VariableIndex) <= Map(MapNum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 4
                                If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(z).VariableIndex) >= Map(MapNum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 5
                                If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(MapNum).Events(id).Pages(z).VariableIndex) = Map(MapNum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                        End Select
                    End If
                        
                    If Map(MapNum).Events(id).Pages(z).chkSwitch = 1 Then
                        If Map(MapNum).Events(id).Pages(z).SwitchCompare = 0 Then
                            If Player(i).characters(TempPlayer(i).CurChar).Switches(Map(MapNum).Events(id).Pages(z).SwitchIndex) = 0 Then
                                spawnevent = False
                            End If
                        Else
                            If Player(i).characters(TempPlayer(i).CurChar).Switches(Map(MapNum).Events(id).Pages(z).SwitchIndex) = 1 Then
                                spawnevent = False
                            Else
                                spawnevent = True
                            End If
                        End If
                    End If
                        
                    If spawnevent = True Then
                        If TempPlayer(i).EventMap.EventPages(x).Visible = 1 Then
                            If z <= pageID Then
                                spawnevent = False
                            End If
                        End If
                    End If
                        
                    If spawnevent = True Then
                        
                        If TempPlayer(i).EventProcessingCount > 0 Then
                            For n = 1 To UBound(TempPlayer(i).EventProcessing)
                                If TempPlayer(i).EventProcessing(n).eventID = id Then
                                    TempPlayer(i).EventProcessing(n).Active = 0
                                End If
                            Next
                        End If
                    
                    
                        With TempPlayer(i).EventMap.EventPages(id)
                            If Map(MapNum).Events(id).Pages(z).GraphicType = 1 Then
                                Select Case Map(MapNum).Events(id).Pages(z).GraphicY
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
                            .GraphicNum = Map(MapNum).Events(id).Pages(z).Graphic
                            .GraphicType = Map(MapNum).Events(id).Pages(z).GraphicType
                            .GraphicX = Map(MapNum).Events(id).Pages(z).GraphicX
                            .GraphicY = Map(MapNum).Events(id).Pages(z).GraphicY
                            .GraphicX2 = Map(MapNum).Events(id).Pages(z).GraphicX2
                            .GraphicY2 = Map(MapNum).Events(id).Pages(z).GraphicY2
                            .questnum = Map(MapNum).Events(id).Pages(z).questnum
                            Select Case Map(MapNum).Events(id).Pages(z).MoveSpeed
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
                            .Position = Map(MapNum).Events(id).Pages(z).Position
                            .eventID = id
                            .pageID = z
                            .Visible = 1
                                
                            .MoveType = Map(MapNum).Events(id).Pages(z).MoveType
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(MapNum).Events(id).Pages(z).MoveRouteCount
                                If .MoveRouteCount > 0 Then
                                    ReDim .MoveRoute(0 To Map(MapNum).Events(id).Pages(z).MoveRouteCount)
                                    For p = 0 To Map(MapNum).Events(id).Pages(z).MoveRouteCount
                                        .MoveRoute(p) = Map(MapNum).Events(id).Pages(z).MoveRoute(p)
                                    Next
                                    .MoverouteComplete = 0
                                Else
                                    .MoverouteComplete = 1
                                End If
                            Else
                                .MoverouteComplete = 1
                            End If
                                
                            .RepeatMoveRoute = Map(MapNum).Events(id).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(MapNum).Events(id).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(MapNum).Events(id).Pages(z).MoveFreq
                            .MoveSpeed = Map(MapNum).Events(id).Pages(z).MoveSpeed
                                
                            .WalkThrough = Map(MapNum).Events(id).Pages(z).WalkThrough
                            .ShowName = Map(MapNum).Events(id).Pages(z).ShowName
                            .WalkingAnim = Map(MapNum).Events(id).Pages(z).WalkAnim
                            .FixedDir = Map(MapNum).Events(id).Pages(z).DirFix
                            
                            
                        End With
                        
                        If Map(MapNum).Events(id).Global = 1 Then
                            If spawnevent Then TempEventMap(MapNum).Events(id).Active = z: TempEventMap(MapNum).Events(id).Position = Map(MapNum).Events(id).Pages(z).Position
                        End If
                        
                        
                        
                        Set Buffer = New clsBuffer
                        Buffer.WriteLong SSpawnEvent
                        Buffer.WriteLong id
                        With TempPlayer(i).EventMap.EventPages(x)
                            Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Name
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
                            Buffer.WriteLong Map(MapNum).Events(id).Pages(z).WalkAnim
                            Buffer.WriteLong Map(MapNum).Events(id).Pages(z).DirFix
                            Buffer.WriteLong Map(MapNum).Events(id).Pages(z).WalkThrough
                            Buffer.WriteLong Map(MapNum).Events(id).Pages(z).ShowName
                            Buffer.WriteLong Map(MapNum).Events(id).Pages(z).questnum
                            Buffer.WriteLong .questnum
                        End With
                        SendDataTo i, Buffer.ToArray
                        Set Buffer = Nothing
                        z = 1
                    End If
                Next
            Next
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpawnNewEvents", "modEventLogic", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Public Sub ProcessEventMovement()
Dim rand As Long, x As Long, i As Long, playerID As Long, eventID As Long, WalkThrough As Long, isglobal As Boolean, MapNum As Long, actualmovespeed As Long, Buffer As clsBuffer, z As Long, sendupdate As Boolean
Dim donotprocessmoveroute As Boolean, pageNum As Long
    'Process Movement if needed for each player/each map/each event....

   On Error GoTo errorhandler

    For i = MIN_MAPS To MAX_MAPS
        If PlayersOnMap(i) Then
            'Manage Global Events First, then all the others.....
            If TempEventMap(i).EventCount > 0 Then
                For x = 1 To TempEventMap(i).EventCount
                    If TempEventMap(i).Events(x).Active > 0 Then
                        pageNum = 1
                        If TempEventMap(i).Events(x).MoveTimer <= GetTickCount Then
                            'Real event! Lets process it!
                            Select Case TempEventMap(i).Events(x).MoveType
                                Case 0
                                    'Nothing, fixed position
                                Case 1 'Random, move randomly if possible...
                                    rand = Random(0, 3)
                                    If CanEventMove(0, i, TempEventMap(i).Events(x).x, TempEventMap(i).Events(x).y, x, TempEventMap(i).Events(x).WalkThrough, rand, True) Then
                                        Select Case TempEventMap(i).Events(x).MoveSpeed
                                            Case 0
                                                EventMove 0, i, x, rand, 2, True
                                            Case 1
                                                EventMove 0, i, x, rand, 3, True
                                            Case 2
                                                EventMove 0, i, x, rand, 4, True
                                            Case 3
                                                EventMove 0, i, x, rand, 6, True
                                            Case 4
                                                EventMove 0, i, x, rand, 12, True
                                            Case 5
                                                EventMove 0, i, x, rand, 24, True
                                        End Select
                                    Else
                                        EventDir 0, i, x, rand, True
                                    End If
                                Case 2 'Move Route
                                    With TempEventMap(i).Events(x)
                                        isglobal = True
                                        MapNum = i
                                        playerID = 0
                                        eventID = x
                                        WalkThrough = TempEventMap(i).Events(x).WalkThrough
                                        If .MoveRouteCount > 0 Then
                                            If .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 1 Then
                                                .MoveRouteStep = 0
                                                .MoverouteComplete = 1
                                            ElseIf .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 0 Then
                                                donotprocessmoveroute = True
                                                .MoverouteComplete = 1
                                            Else
                                                .MoverouteComplete = 0
                                            End If
                                            If donotprocessmoveroute = False Then
                                                .MoveRouteStep = .MoveRouteStep + 1
                                                Select Case .MoveSpeed
                                                    Case 0
                                                        actualmovespeed = 2
                                                    Case 1
                                                        actualmovespeed = 3
                                                    Case 2
                                                        actualmovespeed = 4
                                                    Case 3
                                                        actualmovespeed = 6
                                                    Case 4
                                                        actualmovespeed = 12
                                                    Case 5
                                                        actualmovespeed = 24
                                                End Select
                                                Select Case .MoveRoute(.MoveRouteStep).Index
                                                    Case 1
                                                        If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, DIR_UP, isglobal) Then
                                                            EventMove playerID, MapNum, eventID, DIR_UP, actualmovespeed, isglobal
                                                        Else
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        End If
                                                    Case 2
                                                        If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, DIR_DOWN, isglobal) Then
                                                            EventMove playerID, MapNum, eventID, DIR_DOWN, actualmovespeed, isglobal
                                                        Else
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        End If
                                                    Case 3
                                                        If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, DIR_LEFT, isglobal) Then
                                                            EventMove playerID, MapNum, eventID, DIR_LEFT, actualmovespeed, isglobal
                                                        Else
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        End If
                                                    Case 4
                                                        If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, DIR_RIGHT, isglobal) Then
                                                            EventMove playerID, MapNum, eventID, DIR_RIGHT, actualmovespeed, isglobal
                                                        Else
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        End If
                                                    Case 5
                                                        z = Random(0, 3)
                                                        If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                            EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                        Else
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        End If
                                                    Case 6
                                                        If isglobal = False Then
                                                            If IsOneBlockAway(.x, .y, GetPlayerX(playerID), GetPlayerY(playerID)) = True Then
                                                                EventDir playerID, GetPlayerMap(playerID), eventID, GetDirToPlayer(playerID, GetPlayerMap(playerID), eventID), False
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            Else
                                                                z = CanEventMoveTowardsPlayer(playerID, MapNum, eventID)
                                                                If z >= 4 Then
                                                                    'No
                                                                    If .IgnoreIfCannotMove = 0 Then
                                                                        .MoveRouteStep = .MoveRouteStep - 1
                                                                    End If
                                                                Else
                                                                    'i is the direct, lets go...
                                                                    If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                        EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                                    Else
                                                                        If .IgnoreIfCannotMove = 0 Then
                                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Case 7
                                                        If isglobal = False Then
                                                            z = CanEventMoveAwayFromPlayer(playerID, MapNum, eventID)
                                                            If z >= 5 Then
                                                                'No
                                                            Else
                                                                'i is the direct, lets go...
                                                                If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                    EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                                Else
                                                                    If .IgnoreIfCannotMove = 0 Then
                                                                        .MoveRouteStep = .MoveRouteStep - 1
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Case 8
                                                        If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, .Dir, isglobal) Then
                                                            EventMove playerID, MapNum, eventID, .Dir, actualmovespeed, isglobal
                                                        Else
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        End If
                                                    Case 9
                                                        Select Case .Dir
                                                            Case DIR_UP
                                                                z = DIR_DOWN
                                                            Case DIR_DOWN
                                                                z = DIR_UP
                                                            Case DIR_LEFT
                                                                z = DIR_RIGHT
                                                            Case DIR_RIGHT
                                                                z = DIR_LEFT
                                                        End Select
                                                        If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                            EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                        Else
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        End If
                                                    Case 10
                                                        .MoveTimer = GetTickCount + 100
                                                    Case 11
                                                        .MoveTimer = GetTickCount + 500
                                                    Case 12
                                                        .MoveTimer = GetTickCount + 1000
                                                    Case 13
                                                        EventDir playerID, MapNum, eventID, DIR_UP, isglobal
                                                    Case 14
                                                        EventDir playerID, MapNum, eventID, DIR_DOWN, isglobal
                                                    Case 15
                                                        EventDir playerID, MapNum, eventID, DIR_LEFT, isglobal
                                                    Case 16
                                                        EventDir playerID, MapNum, eventID, DIR_RIGHT, isglobal
                                                    Case 17
                                                        Select Case .Dir
                                                            Case DIR_UP
                                                                z = DIR_RIGHT
                                                            Case DIR_RIGHT
                                                                z = DIR_DOWN
                                                            Case DIR_LEFT
                                                                z = DIR_UP
                                                            Case DIR_DOWN
                                                                z = DIR_LEFT
                                                        End Select
                                                        EventDir playerID, MapNum, eventID, z, isglobal
                                                    Case 18
                                                        Select Case .Dir
                                                            Case DIR_UP
                                                                z = DIR_LEFT
                                                            Case DIR_RIGHT
                                                                z = DIR_UP
                                                            Case DIR_LEFT
                                                                z = DIR_DOWN
                                                            Case DIR_DOWN
                                                                z = DIR_RIGHT
                                                        End Select
                                                        EventDir playerID, MapNum, eventID, z, isglobal
                                                    Case 19
                                                        Select Case .Dir
                                                            Case DIR_UP
                                                                z = DIR_DOWN
                                                            Case DIR_RIGHT
                                                                z = DIR_LEFT
                                                            Case DIR_LEFT
                                                                z = DIR_RIGHT
                                                            Case DIR_DOWN
                                                                z = DIR_UP
                                                        End Select
                                                        EventDir playerID, MapNum, eventID, z, isglobal
                                                    Case 20
                                                        z = Random(0, 3)
                                                        EventDir playerID, MapNum, eventID, z, isglobal
                                                    Case 21
                                                        If isglobal = False Then
                                                            z = GetDirToPlayer(playerID, MapNum, eventID)
                                                            EventDir playerID, MapNum, eventID, z, isglobal
                                                        End If
                                                    Case 22
                                                        If isglobal = False Then
                                                            z = GetDirAwayFromPlayer(playerID, MapNum, eventID)
                                                            EventDir playerID, MapNum, eventID, z, isglobal
                                                        End If
                                                    Case 23
                                                        .MoveSpeed = 0
                                                    Case 24
                                                        .MoveSpeed = 1
                                                    Case 25
                                                        .MoveSpeed = 2
                                                    Case 26
                                                        .MoveSpeed = 3
                                                    Case 27
                                                        .MoveSpeed = 4
                                                    Case 28
                                                        .MoveSpeed = 5
                                                    Case 29
                                                        .MoveFreq = 0
                                                    Case 30
                                                        .MoveFreq = 1
                                                    Case 31
                                                        .MoveFreq = 2
                                                    Case 32
                                                        .MoveFreq = 3
                                                    Case 33
                                                        .MoveFreq = 4
                                                    Case 34
                                                        .WalkingAnim = 1
                                                        'Need to send update to client
                                                        sendupdate = True
                                                    Case 35
                                                        .WalkingAnim = 0
                                                        'Need to send update to client
                                                        sendupdate = True
                                                    Case 36
                                                        .FixedDir = 1
                                                        'Need to send update to client
                                                        sendupdate = True
                                                    Case 37
                                                        .FixedDir = 0
                                                        'Need to send update to client
                                                        sendupdate = True
                                                    Case 38
                                                        .WalkThrough = 1
                                                    Case 39
                                                        .WalkThrough = 0
                                                    Case 40
                                                        .Position = 0
                                                        'Need to send update to client
                                                        sendupdate = True
                                                    Case 41
                                                        .Position = 1
                                                        'Need to send update to client
                                                        sendupdate = True
                                                    Case 42
                                                        .Position = 2
                                                        'Need to send update to client
                                                        sendupdate = True
                                                    Case 43
                                                        .GraphicType = .MoveRoute(.MoveRouteStep).Data1
                                                        .GraphicNum = .MoveRoute(.MoveRouteStep).Data2
                                                        .GraphicX = .MoveRoute(.MoveRouteStep).Data3
                                                        .GraphicX2 = .MoveRoute(.MoveRouteStep).Data4
                                                        .GraphicY = .MoveRoute(.MoveRouteStep).data5
                                                        .GraphicY2 = .MoveRoute(.MoveRouteStep).data6
                                                        If .GraphicType = 1 Then
                                                            Select Case .GraphicY
                                                                Case 0
                                                                    .Dir = DIR_DOWN
                                                                Case 1
                                                                    .Dir = DIR_LEFT
                                                                Case 2
                                                                    .Dir = DIR_RIGHT
                                                                Case 3
                                                                    .Dir = DIR_UP
                                                            End Select
                                                        End If
                                                        'Need to Send Update to client
                                                        sendupdate = True
                                                End Select
                                                
                                                If sendupdate Then
                                                    Set Buffer = New clsBuffer
                                                    Buffer.WriteLong SSpawnEvent
                                                    Buffer.WriteLong eventID
                                                    With TempEventMap(i).Events(x)
                                                        Buffer.WriteString Map(i).Events(x).Name
                                                        Buffer.WriteLong .Dir
                                                        Buffer.WriteLong .GraphicNum
                                                        Buffer.WriteLong .GraphicType
                                                        Buffer.WriteLong .GraphicX
                                                        Buffer.WriteLong .GraphicX2
                                                        Buffer.WriteLong .GraphicY
                                                        Buffer.WriteLong .GraphicY2
                                                        Buffer.WriteLong .MoveSpeed
                                                        Buffer.WriteLong .x
                                                        Buffer.WriteLong .y
                                                        Buffer.WriteLong .Position
                                                        Buffer.WriteLong .Active
                                                        Buffer.WriteLong .WalkingAnim
                                                        Buffer.WriteLong .FixedDir
                                                        Buffer.WriteLong .WalkThrough
                                                        Buffer.WriteLong .ShowName
                                                        Buffer.WriteLong .questnum
                                                    End With
                                                    SendDataToMap i, Buffer.ToArray
                                                    Set Buffer = Nothing
                                                End If
                                            End If
                                            donotprocessmoveroute = False
                                        End If
                                    End With
                            End Select
                            
                            Select Case TempEventMap(i).Events(x).MoveFreq
                                Case 0
                                    TempEventMap(i).Events(x).MoveTimer = GetTickCount + 4000
                                Case 1
                                    TempEventMap(i).Events(x).MoveTimer = GetTickCount + 2000
                                Case 2
                                    TempEventMap(i).Events(x).MoveTimer = GetTickCount + 1000
                                Case 3
                                    TempEventMap(i).Events(x).MoveTimer = GetTickCount + 500
                                Case 4
                                    TempEventMap(i).Events(x).MoveTimer = GetTickCount + 250
                            End Select
                        End If
                    End If
                Next
            End If
        End If
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProcessEventMovement", "modEventLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub ProcessLocalEventMovement()
Dim rand As Long, x As Long, i As Long, playerID As Long, eventID As Long, WalkThrough As Long, isglobal As Boolean, MapNum As Long, actualmovespeed As Long, Buffer As clsBuffer, z As Long, sendupdate As Boolean
Dim donotprocessmoveroute As Boolean

   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            playerID = i
            If TempPlayer(i).EventMap.CurrentEvents > 0 Then
                For x = 1 To TempPlayer(i).EventMap.CurrentEvents
                    If Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Global = 0 Then
                        If TempPlayer(i).EventMap.EventPages(x).Visible = 1 Then
                            If TempPlayer(i).EventMap.EventPages(x).MoveTimer <= GetTickCount Then
                                'Real event! Lets process it!
                                Select Case TempPlayer(i).EventMap.EventPages(x).MoveType
                                    Case 0
                                        'Nothing, fixed position
                                    Case 1 'Random, move randomly if possible...
                                        rand = Random(0, 3)
                                        playerID = i
                                        If CanEventMove(i, GetPlayerMap(i), TempPlayer(i).EventMap.EventPages(x).x, TempPlayer(i).EventMap.EventPages(x).y, x, TempPlayer(i).EventMap.EventPages(x).WalkThrough, rand, False) Then
                                            Select Case TempPlayer(i).EventMap.EventPages(x).MoveSpeed
                                                Case 0
                                                    EventMove i, GetPlayerMap(i), x, rand, 2, False
                                                Case 1
                                                    EventMove i, GetPlayerMap(i), x, rand, 3, False
                                                Case 2
                                                    EventMove i, GetPlayerMap(i), x, rand, 4, False
                                                Case 3
                                                    EventMove i, GetPlayerMap(i), x, rand, 6, False
                                                Case 4
                                                    EventMove i, GetPlayerMap(i), x, rand, 12, False
                                                Case 5
                                                    EventMove i, GetPlayerMap(i), x, rand, 24, False
                                            End Select
                                        Else
                                            EventDir 0, GetPlayerMap(i), x, rand, True
                                        End If
                                    Case 2 'Move Route - later!
                                        With TempPlayer(i).EventMap.EventPages(x)
                                            isglobal = False
                                            sendupdate = False
                                            MapNum = GetPlayerMap(i)
                                            playerID = i
                                            eventID = x
                                            WalkThrough = .WalkThrough
                                            If .MoveRouteCount > 0 Then
                                                If .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 1 Then
                                                    .MoveRouteStep = 0
                                                    .MoverouteComplete = 1
                                                ElseIf .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 0 Then
                                                    donotprocessmoveroute = True
                                                    .MoverouteComplete = 1
                                                Else
                                                    .MoverouteComplete = 0
                                                End If
                                                If donotprocessmoveroute = False Then
                                                    .MoveRouteStep = .MoveRouteStep + 1
                                                    Select Case .MoveSpeed
                                                        Case 0
                                                            actualmovespeed = 2
                                                        Case 1
                                                            actualmovespeed = 3
                                                        Case 2
                                                            actualmovespeed = 4
                                                        Case 3
                                                            actualmovespeed = 6
                                                        Case 4
                                                            actualmovespeed = 12
                                                        Case 5
                                                            actualmovespeed = 24
                                                    End Select
                                                    Select Case .MoveRoute(.MoveRouteStep).Index
                                                        Case 1
                                                            If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, DIR_UP, isglobal) Then
                                                                EventMove playerID, MapNum, eventID, DIR_UP, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        Case 2
                                                            If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, DIR_DOWN, isglobal) Then
                                                                EventMove playerID, MapNum, eventID, DIR_DOWN, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        Case 3
                                                            If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, DIR_LEFT, isglobal) Then
                                                                EventMove playerID, MapNum, eventID, DIR_LEFT, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        Case 4
                                                            If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, DIR_RIGHT, isglobal) Then
                                                                EventMove playerID, MapNum, eventID, DIR_RIGHT, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        Case 5
                                                            z = Random(0, 3)
                                                            If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        Case 6
                                                            If isglobal = False Then
                                                                If IsOneBlockAway(.x, .y, GetPlayerX(playerID), GetPlayerY(playerID)) = True Then
                                                                    EventDir playerID, GetPlayerMap(playerID), eventID, GetDirToPlayer(playerID, GetPlayerMap(playerID), eventID), False
                                                                    'Lets do cool stuff!
                                                                    If Map(GetPlayerMap(playerID)).Events(eventID).Pages(TempPlayer(playerID).EventMap.EventPages(eventID).pageID).Trigger = 1 Then
                                                                        If Map(MapNum).Events(eventID).Pages(TempPlayer(playerID).EventMap.EventPages(eventID).pageID).CommandListCount > 0 Then
                                                                            TempPlayer(playerID).EventProcessing(eventID).Active = 1
                                                                            TempPlayer(playerID).EventProcessing(eventID).ActionTimer = GetTickCount
                                                                            TempPlayer(playerID).EventProcessing(eventID).CurList = 1
                                                                            TempPlayer(playerID).EventProcessing(eventID).CurSlot = 1
                                                                            TempPlayer(playerID).EventProcessing(eventID).eventID = eventID
                                                                            TempPlayer(playerID).EventProcessing(eventID).pageID = TempPlayer(playerID).EventMap.EventPages(eventID).pageID
                                                                            TempPlayer(playerID).EventProcessing(eventID).WaitingForResponse = 0
                                                                            ReDim TempPlayer(playerID).EventProcessing(eventID).ListLeftOff(0 To Map(GetPlayerMap(playerID)).Events(TempPlayer(playerID).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(playerID).EventMap.EventPages(eventID).pageID).CommandListCount)
                                                                        End If
                                                                    End If
                                                                    If .IgnoreIfCannotMove = 0 Then
                                                                        .MoveRouteStep = .MoveRouteStep - 1
                                                                    End If
                                                                Else
                                                                    z = CanEventMoveTowardsPlayer(playerID, MapNum, eventID)
                                                                    If z >= 4 Then
                                                                        'No
                                                                        If .IgnoreIfCannotMove = 0 Then
                                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                                        End If
                                                                    Else
                                                                        'i is the direct, lets go...
                                                                        If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                            EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                                        Else
                                                                            If .IgnoreIfCannotMove = 0 Then
                                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Case 7
                                                            If isglobal = False Then
                                                                z = CanEventMoveAwayFromPlayer(playerID, MapNum, eventID)
                                                                If z >= 5 Then
                                                                    'No
                                                                Else
                                                                    'i is the direct, lets go...
                                                                    If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                        EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                                    Else
                                                                        If .IgnoreIfCannotMove = 0 Then
                                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Case 8
                                                            If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, .Dir, isglobal) Then
                                                                EventMove playerID, MapNum, eventID, .Dir, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        Case 9
                                                            Select Case .Dir
                                                                Case DIR_UP
                                                                    z = DIR_DOWN
                                                                Case DIR_DOWN
                                                                    z = DIR_UP
                                                                Case DIR_LEFT
                                                                    z = DIR_RIGHT
                                                                Case DIR_RIGHT
                                                                    z = DIR_LEFT
                                                            End Select
                                                            If CanEventMove(playerID, MapNum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        Case 10
                                                            .MoveTimer = GetTickCount + 100
                                                        Case 11
                                                            .MoveTimer = GetTickCount + 500
                                                        Case 12
                                                            .MoveTimer = GetTickCount + 1000
                                                        Case 13
                                                            EventDir playerID, MapNum, eventID, DIR_UP, isglobal
                                                        Case 14
                                                            EventDir playerID, MapNum, eventID, DIR_DOWN, isglobal
                                                        Case 15
                                                            EventDir playerID, MapNum, eventID, DIR_LEFT, isglobal
                                                        Case 16
                                                            EventDir playerID, MapNum, eventID, DIR_RIGHT, isglobal
                                                        Case 17
                                                            Select Case .Dir
                                                                Case DIR_UP
                                                                    z = DIR_RIGHT
                                                                Case DIR_RIGHT
                                                                    z = DIR_DOWN
                                                                Case DIR_LEFT
                                                                    z = DIR_UP
                                                                Case DIR_DOWN
                                                                    z = DIR_LEFT
                                                            End Select
                                                            EventDir playerID, MapNum, eventID, z, isglobal
                                                        Case 18
                                                            Select Case .Dir
                                                                Case DIR_UP
                                                                    z = DIR_LEFT
                                                                Case DIR_RIGHT
                                                                    z = DIR_UP
                                                                Case DIR_LEFT
                                                                    z = DIR_DOWN
                                                                Case DIR_DOWN
                                                                    z = DIR_RIGHT
                                                            End Select
                                                            EventDir playerID, MapNum, eventID, z, isglobal
                                                        Case 19
                                                            Select Case .Dir
                                                                Case DIR_UP
                                                                    z = DIR_DOWN
                                                                Case DIR_RIGHT
                                                                    z = DIR_LEFT
                                                                Case DIR_LEFT
                                                                    z = DIR_RIGHT
                                                                Case DIR_DOWN
                                                                    z = DIR_UP
                                                            End Select
                                                            EventDir playerID, MapNum, eventID, z, isglobal
                                                        Case 20
                                                            z = Random(0, 3)
                                                            EventDir playerID, MapNum, eventID, z, isglobal
                                                        Case 21
                                                            If isglobal = False Then
                                                                z = GetDirToPlayer(playerID, MapNum, eventID)
                                                                EventDir playerID, MapNum, eventID, z, isglobal
                                                            End If
                                                        Case 22
                                                            If isglobal = False Then
                                                                z = GetDirAwayFromPlayer(playerID, MapNum, eventID)
                                                                EventDir playerID, MapNum, eventID, z, isglobal
                                                            End If
                                                        Case 23
                                                            .MoveSpeed = 0
                                                        Case 24
                                                            .MoveSpeed = 1
                                                        Case 25
                                                            .MoveSpeed = 2
                                                        Case 26
                                                            .MoveSpeed = 3
                                                        Case 27
                                                            .MoveSpeed = 4
                                                        Case 28
                                                            .MoveSpeed = 5
                                                        Case 29
                                                            .MoveFreq = 0
                                                        Case 30
                                                            .MoveFreq = 1
                                                        Case 31
                                                            .MoveFreq = 2
                                                        Case 32
                                                            .MoveFreq = 3
                                                        Case 33
                                                            .MoveFreq = 4
                                                        Case 34
                                                            .WalkingAnim = 1
                                                            'Need to send update to client
                                                            sendupdate = True
                                                        Case 35
                                                            .WalkingAnim = 0
                                                            'Need to send update to client
                                                            sendupdate = True
                                                        Case 36
                                                            .FixedDir = 1
                                                            'Need to send update to client
                                                            sendupdate = True
                                                        Case 37
                                                            .FixedDir = 0
                                                            'Need to send update to client
                                                            sendupdate = True
                                                        Case 38
                                                            .WalkThrough = 1
                                                        Case 39
                                                            .WalkThrough = 0
                                                        Case 40
                                                            .Position = 0
                                                            'Need to send update to client
                                                            sendupdate = True
                                                        Case 41
                                                            .Position = 1
                                                            'Need to send update to client
                                                            sendupdate = True
                                                        Case 42
                                                            .Position = 2
                                                            'Need to send update to client
                                                            sendupdate = True
                                                        Case 43
                                                            .GraphicType = .MoveRoute(.MoveRouteStep).Data1
                                                            .GraphicNum = .MoveRoute(.MoveRouteStep).Data2
                                                            .GraphicX = .MoveRoute(.MoveRouteStep).Data3
                                                            .GraphicX2 = .MoveRoute(.MoveRouteStep).Data4
                                                            .GraphicY = .MoveRoute(.MoveRouteStep).data5
                                                            .GraphicY2 = .MoveRoute(.MoveRouteStep).data6
                                                            If .GraphicType = 1 Then
                                                                Select Case .GraphicY
                                                                    Case 0
                                                                        .Dir = DIR_DOWN
                                                                    Case 1
                                                                        .Dir = DIR_LEFT
                                                                    Case 2
                                                                        .Dir = DIR_RIGHT
                                                                    Case 3
                                                                        .Dir = DIR_UP
                                                                End Select
                                                            End If
                                                            'Need to Send Update to client
                                                            sendupdate = True
                                                    End Select
                                                    
                                                    If sendupdate Then
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SSpawnEvent
                                                        Buffer.WriteLong TempPlayer(playerID).EventMap.EventPages(eventID).eventID
                                                        With TempPlayer(playerID).EventMap.EventPages(eventID)
                                                            Buffer.WriteString Map(GetPlayerMap(playerID)).Events(TempPlayer(playerID).EventMap.EventPages(eventID).eventID).Name
                                                            Buffer.WriteLong .Dir
                                                            Buffer.WriteLong .GraphicNum
                                                            Buffer.WriteLong .GraphicType
                                                            Buffer.WriteLong .GraphicX
                                                            Buffer.WriteLong .GraphicX2
                                                            Buffer.WriteLong .GraphicY
                                                            Buffer.WriteLong .GraphicY2
                                                            Buffer.WriteLong .MoveSpeed
                                                            Buffer.WriteLong .x
                                                            Buffer.WriteLong .y
                                                            Buffer.WriteLong .Position
                                                            Buffer.WriteLong .Visible
                                                            Buffer.WriteLong .WalkingAnim
                                                            Buffer.WriteLong .FixedDir
                                                            Buffer.WriteLong .WalkThrough
                                                            Buffer.WriteLong .ShowName
                                                        End With
                                                        SendDataTo playerID, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                    End If
                                                End If
                                                donotprocessmoveroute = False
                                            End If
                                        End With
                                End Select
                                Select Case TempPlayer(playerID).EventMap.EventPages(x).MoveFreq
                                    Case 0
                                        TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 4000
                                    Case 1
                                        TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 2000
                                    Case 2
                                        TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 1000
                                    Case 3
                                        TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 500
                                    Case 4
                                        TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 250
                                End Select
                            End If
                        End If
                    End If
                Next
            End If
        End If
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProcessLocalEventMovement", "modEventLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub ProcessEventCommands()
Dim Buffer As clsBuffer, i As Long, x As Long, z As Long, removeEventProcess As Boolean, w As Long, v As Long, p As Long
Dim restartlist As Boolean, restartloop As Boolean, endprocess As Boolean
    'Now, we process the damn things for commands :P

   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            For x = 1 To TempPlayer(i).EventMap.CurrentEvents
                If TempPlayer(i).EventMap.EventPages(x).Visible Then
                    If Map(Player(i).characters(TempPlayer(i).CurChar).Map).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Pages(TempPlayer(i).EventMap.EventPages(x).pageID).Trigger = 2 Then 'Parallel Process baby!
                        If TempPlayer(i).EventProcessingCount > 0 Then
                            If TempPlayer(i).EventProcessing(x).Active = 0 Then
                                If Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Pages(TempPlayer(i).EventMap.EventPages(x).pageID).CommandListCount > 0 Then
                                    'start new event processing
                                    TempPlayer(i).EventProcessing(TempPlayer(i).EventMap.EventPages(x).eventID).Active = 1
                                    With TempPlayer(i).EventProcessing(TempPlayer(i).EventMap.EventPages(x).eventID)
                                        .ActionTimer = GetTickCount
                                        .CurList = 1
                                        .CurSlot = 1
                                        .eventID = TempPlayer(i).EventMap.EventPages(x).eventID
                                        .pageID = TempPlayer(i).EventMap.EventPages(x).pageID
                                        .WaitingForResponse = 0
                                        ReDim .ListLeftOff(0 To Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Pages(TempPlayer(i).EventMap.EventPages(x).pageID).CommandListCount)
                                    End With
                                End If
                            End If
                        Else
                            If Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Pages(TempPlayer(i).EventMap.EventPages(x).pageID).CommandListCount > 0 Then
                                'Clearly need to start it!
                                TempPlayer(i).EventProcessing(TempPlayer(i).EventMap.EventPages(x).eventID).Active = 1
                                With TempPlayer(i).EventProcessing(TempPlayer(i).EventMap.EventPages(x).eventID)
                                    .ActionTimer = GetTickCount
                                    .CurList = 1
                                    .CurSlot = 1
                                    .eventID = TempPlayer(i).EventMap.EventPages(x).eventID
                                    .pageID = TempPlayer(i).EventMap.EventPages(x).pageID
                                    .WaitingForResponse = 0
                                    ReDim .ListLeftOff(0 To Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Pages(TempPlayer(i).EventMap.EventPages(x).pageID).CommandListCount)
                                End With
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
    
    'That is it for starting parallel processes :D now we just have to make the code that actually processes the events to their fullest
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).EventProcessingCount > 0 Then
                restartloop = True
                Do While restartloop = True
                    restartloop = False
                    For x = 1 To TempPlayer(i).EventProcessingCount
                        If TempPlayer(i).EventProcessing(x).Active = 1 Then
                            If x > TempPlayer(i).EventProcessingCount Then Exit For
                            With TempPlayer(i).EventProcessing(x)
                                If TempPlayer(i).EventProcessingCount = 0 Then Exit Sub
                                removeEventProcess = False
                                If .WaitingForResponse = 2 Then
                                    If TempPlayer(i).InShop = 0 Then
                                        .WaitingForResponse = 0
                                    End If
                                End If
                                If .WaitingForResponse = 3 Then
                                    If TempPlayer(i).InBank = False Then
                                        .WaitingForResponse = 0
                                    End If
                                End If
                                If .WaitingForResponse = 4 Then
                                    'waiting for eventmovement to complete
                                    If .EventMovingType = 0 Then
                                        If TempPlayer(i).EventMap.EventPages(.EventMovingID).MoverouteComplete = 1 Then
                                            .WaitingForResponse = 0
                                        End If
                                    Else
                                        If TempEventMap(GetPlayerMap(i)).Events(.EventMovingID).MoverouteComplete = 1 Then
                                            .WaitingForResponse = 0
                                        End If
                                    End If
                                End If
                                If .WaitingForResponse = 0 Then
                                    If .ActionTimer <= GetTickCount Then
                                        restartlist = True
                                        endprocess = False
                                        Do While restartlist = True And endprocess = False And .WaitingForResponse = 0
                                            restartlist = False
                                            If .ListLeftOff(.CurList) > 0 Then
                                                .CurSlot = .ListLeftOff(.CurList) + 1
                                                .ListLeftOff(.CurList) = 0
                                            End If
                                            If .CurList > Map(Player(i).characters(TempPlayer(i).CurChar).Map).Events(.eventID).Pages(.pageID).CommandListCount Then
                                                'Get rid of this event, it is bad
                                                removeEventProcess = True
                                                endprocess = True
                                            End If
                                            If .CurSlot > Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).CommandCount Then
                                                If .CurList = 1 Then
                                                    'Get rid of this event, it is bad
                                                    removeEventProcess = True
                                                    endprocess = True
                                                Else
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).ParentList
                                                    .CurSlot = 1
                                                    restartlist = True
                                                End If
                                            End If
                                            If restartlist = False And endprocess = False Then
                                                'If we are still here, then we are good to process shit :D
                                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Index
                                                    Case EventType.evAddText
                                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                            Case 0
                                                                PlayerMsg i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                            Case 1
                                                                MapMsg GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                            Case 2
                                                                GlobalMsg Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                        End Select
                                                    Case EventType.evShowText
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SEventChat
                                                        Buffer.WriteLong .eventID
                                                        Buffer.WriteLong .pageID
                                                        Buffer.WriteLong Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                        Buffer.WriteString ParseEventText(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1)
                                                        Buffer.WriteLong 0
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).CommandCount > .CurSlot Then
                                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evShowText Or Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evShowChoices Then
                                                                Buffer.WriteLong 1
                                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evCondition Then
                                                                Buffer.WriteLong 2
                                                            Else
                                                                Buffer.WriteLong 0
                                                            End If
                                                        Else
                                                            Buffer.WriteLong 2
                                                        End If
                                                        SendDataTo i, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                        .WaitingForResponse = 1
                                                    Case EventType.evShowChoices
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SEventChat
                                                        Buffer.WriteLong .eventID
                                                        Buffer.WriteLong .pageID
                                                        Buffer.WriteLong Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).data5
                                                        Buffer.WriteString ParseEventText(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1)
                                                        If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text2)) > 0 Then
                                                            w = 1
                                                            If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text3)) > 0 Then
                                                                w = 2
                                                                If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text4)) > 0 Then
                                                                    w = 3
                                                                    If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text5)) > 0 Then
                                                                        w = 4
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        Buffer.WriteLong w
                                                        For v = 1 To w
                                                            Select Case v
                                                                Case 1
                                                                    Buffer.WriteString ParseEventText(i, Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text2))
                                                                Case 2
                                                                    Buffer.WriteString ParseEventText(i, Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text3))
                                                                Case 3
                                                                    Buffer.WriteString ParseEventText(i, Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text4))
                                                                Case 4
                                                                    Buffer.WriteString ParseEventText(i, Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text5))
                                                            End Select
                                                        Next
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).CommandCount > .CurSlot Then
                                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evShowText Or Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evShowChoices Then
                                                                Buffer.WriteLong 1
                                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evCondition Then
                                                                Buffer.WriteLong 2
                                                            Else
                                                                Buffer.WriteLong 0
                                                            End If
                                                        Else
                                                            Buffer.WriteLong 2
                                                        End If
                                                        SendDataTo i, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                        .WaitingForResponse = 1
                                                    Case EventType.evPlayerVar
                                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                            Case 0
                                                                Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                            Case 1
                                                                Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) + Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                            Case 2
                                                                Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) - Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                            Case 3
                                                                Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Random(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4)
                                                        End Select
                                                    Case EventType.evPlayerSwitch
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                            Player(i).characters(TempPlayer(i).CurChar).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = 1
                                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                            Player(i).characters(TempPlayer(i).CurChar).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = 0
                                                        End If
                                                    Case EventType.evSelfSwitch
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Global = 1 Then
                                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                                Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 1
                                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                                Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 0
                                                            End If
                                                        Else
                                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                                TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 1
                                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                                TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 0
                                                            End If
                                                        End If
                                                    Case EventType.evCondition
                                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Condition
                                                            Case 0
                                                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                                    Case 0
                                                                        If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 1
                                                                        If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) >= Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 2
                                                                        If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) <= Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 3
                                                                        If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) > Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 4
                                                                        If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) < Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 5
                                                                        If Player(i).characters(TempPlayer(i).CurChar).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) <> Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                End Select
                                                            Case 1
                                                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                                    Case 0
                                                                        If Player(i).characters(TempPlayer(i).CurChar).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = 1 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 1
                                                                        If Player(i).characters(TempPlayer(i).CurChar).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = 0 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                End Select
                                                            Case 2
                                                                If HasItem(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) >= Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2 Then
                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                    .CurSlot = 1
                                                                Else
                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                    .CurSlot = 1
                                                                End If
                                                            Case 3
                                                                If Player(i).characters(TempPlayer(i).CurChar).Class = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                    .CurSlot = 1
                                                                Else
                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                    .CurSlot = 1
                                                                End If
                                                            Case 4
                                                                If HasSpell(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = True Then
                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                    .CurSlot = 1
                                                                Else
                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                    .CurSlot = 1
                                                                End If
                                                            Case 5
                                                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                                    Case 0
                                                                        If GetPlayerLevel(i) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 1
                                                                        If GetPlayerLevel(i) >= Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 2
                                                                        If GetPlayerLevel(i) <= Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 3
                                                                        If GetPlayerLevel(i) > Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 4
                                                                        If GetPlayerLevel(i) < Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    Case 5
                                                                        If GetPlayerLevel(i) <> Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                End Select
                                                            Case 6
                                                                If Map(GetPlayerMap(i)).Events(.eventID).Global = 1 Then
                                                                    Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                                        Case 0 'Self Switch is true
                                                                            If Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 1 Then
                                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                                .CurSlot = 1
                                                                            Else
                                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                                .CurSlot = 1
                                                                            End If
                                                                        Case 1  'self switch is false
                                                                            If Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 0 Then
                                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                                .CurSlot = 1
                                                                            Else
                                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                                .CurSlot = 1
                                                                            End If
                                                                    End Select
                                                                Else
                                                                    Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                                        Case 0 'Self Switch is true
                                                                            If TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 1 Then
                                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                                .CurSlot = 1
                                                                            Else
                                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                                .CurSlot = 1
                                                                            End If
                                                                        Case 1  'self switch is false
                                                                            If TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 0 Then
                                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                                .CurSlot = 1
                                                                            Else
                                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                                .CurSlot = 1
                                                                            End If
                                                                    End Select
                                                                End If
                                                            Case 7
                                                                If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 > 0 And Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 <= MAX_QUESTS Then
                                                                    If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2 = 0 Then
                                                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3
                                                                            Case 0
                                                                                If Player(i).characters(TempPlayer(i).CurChar).PlayerQuest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1).State = 0 Then
                                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                                    .CurSlot = 1
                                                                                Else
                                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                                    .CurSlot = 1
                                                                                End If
                                                                            Case 1
                                                                                If Player(i).characters(TempPlayer(i).CurChar).PlayerQuest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1).State = 1 Then
                                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                                    .CurSlot = 1
                                                                                Else
                                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                                    .CurSlot = 1
                                                                                End If
                                                                            Case 2
                                                                                If Player(i).characters(TempPlayer(i).CurChar).PlayerQuest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1).State = 2 Or Player(i).characters(TempPlayer(i).CurChar).PlayerQuest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1).State = 3 Then
                                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                                    .CurSlot = 1
                                                                                Else
                                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                                    .CurSlot = 1
                                                                                End If
                                                                            Case 3
                                                                                If CanStartQuest(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1, False) Then
                                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                                    .CurSlot = 1
                                                                                Else
                                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                                    .CurSlot = 1
                                                                                End If
                                                                            Case 4
                                                                                If CanEndQuest(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1, False) Then
                                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                                    .CurSlot = 1
                                                                                Else
                                                                                    .ListLeftOff(.CurList) = .CurSlot
                                                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                                    .CurSlot = 1
                                                                                End If
                                                                        End Select
                                                                    ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2 = 1 Then
                                                                        If Player(i).characters(TempPlayer(i).CurChar).PlayerQuest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1).CurrentTask = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                            .CurSlot = 1
                                                                        Else
                                                                            .ListLeftOff(.CurList) = .CurSlot
                                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                            .CurSlot = 1
                                                                        End If
                                                                    End If
                                                                End If
                                                        End Select
                                                        endprocess = True
                                                    Case EventType.evExitProcess
                                                        removeEventProcess = True
                                                        endprocess = True
                                                    Case EventType.evChangeItems
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                            If FindItem(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) > 0 Then
                                                                Call SetPlayerInvItemValue(i, FindItem(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3)
                                                                Call CheckTasks(i, TASK_AQUIREITEMS, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                                                Call CheckTasks(i, TASK_FETCHRETURN, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                                            End If
                                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                            GiveInvItem i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, True
                                                            Call CheckTasks(i, TASK_AQUIREITEMS, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                                            Call CheckTasks(i, TASK_FETCHRETURN, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 2 Then
                                                            Dim itemAmount As Long
                                                            itemAmount = HasItem(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                                            ' Check Amount
                                                            If itemAmount >= Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3 Then
                                                                TakeInvItem i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                                Call CheckTasks(i, TASK_AQUIREITEMS, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                                                Call CheckTasks(i, TASK_FETCHRETURN, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                                            End If
                                                        End If
                                                        SendInventory i
                                                    Case EventType.evRestoreHP
                                                        SetPlayerVital i, HP, GetPlayerMaxVital(i, HP)
                                                        SendVital i, HP
                                                    Case EventType.evRestoreMP
                                                        SetPlayerVital i, MP, GetPlayerMaxVital(i, MP)
                                                        SendVital i, MP
                                                    Case EventType.evLevelUp
                                                        SetPlayerExp i, GetPlayerNextLevel(i)
                                                        CheckPlayerLevelUp i
                                                        SendEXP i
                                                        SendPlayerData i
                                                    Case EventType.evChangeLevel
                                                        SetPlayerLevel i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                        SetPlayerExp i, 0
                                                        SendEXP i
                                                        SendPlayerData i
                                                    Case EventType.evChangeSkills
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                            If FindOpenSpellSlot(i) > 0 Then
                                                                If HasSpell(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = False Then
                                                                    SetPlayerSpell i, FindOpenSpellSlot(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                                Else
                                                                    'Error, already knows spell
                                                                End If
                                                            Else
                                                                'Error, no room for spells
                                                            End If
                                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                            If HasSpell(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = True Then
                                                                For p = 1 To MAX_PLAYER_SPELLS
                                                                    If Player(i).characters(TempPlayer(i).CurChar).Spell(p) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 Then
                                                                        SetPlayerSpell i, p, 0
                                                                    End If
                                                                Next
                                                            End If
                                                        End If
                                                        SendPlayerSpells i
                                                    Case EventType.evChangeClass
                                                        Player(i).characters(TempPlayer(i).CurChar).Class = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                        SendPlayerData i
                                                    Case EventType.evChangeSprite
                                                        SetPlayerSprite i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                        SendPlayerData i
                                                    Case EventType.evChangeSex
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 0 Then
                                                            Player(i).characters(TempPlayer(i).CurChar).Sex = SEX_MALE
                                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 1 Then
                                                            Player(i).characters(TempPlayer(i).CurChar).Sex = SEX_FEMALE
                                                        End If
                                                        SendPlayerData i
                                                    Case EventType.evChangePK
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 0 Then
                                                            Player(i).characters(TempPlayer(i).CurChar).PK = NO
                                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 1 Then
                                                            Player(i).characters(TempPlayer(i).CurChar).PK = YES
                                                        End If
                                                        SendPlayerData i
                                                    Case EventType.evWarpPlayer
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4 = 0 Then
                                                            PlayerWarp i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                        Else
                                                            Player(i).characters(TempPlayer(i).CurChar).Dir = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4 - 1
                                                            PlayerWarp i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                        End If
                                                        
                                                    Case EventType.evSetMoveRoute
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 <= Map(GetPlayerMap(i)).EventCount Then
                                                            If Map(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Global = 1 Then
                                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveType = 2
                                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).IgnoreIfCannotMove = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).RepeatMoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteCount = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRouteCount
                                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRoute
                                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteStep = 0
                                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoverouteComplete = 0
                                                            Else
                                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveType = 2
                                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).IgnoreIfCannotMove = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).RepeatMoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteCount = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRouteCount
                                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRoute
                                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteStep = 0
                                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoverouteComplete = 0
                                                            End If
                                                        End If
                                                    Case EventType.evPlayAnimation
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                            SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, GetPlayerX(i), GetPlayerY(i), TARGET_TYPE_PLAYER, i, i
                                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                            If Map(GetPlayerMap(i)).Events(.eventID).Global = 1 Then
                                                                SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).x, Map(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).y
                                                            Else
                                                                SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).x, TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).y, 0, 0, i
                                                            End If
                                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 2 Then
                                                            SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4, 0, 0, i
                                                        End If
                                                    Case EventType.evCustomScript
                                                        'Runs Through Cases for a script
                                                        Call CustomScript(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                                    Case EventType.evPlayBGM
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SPlayBGM
                                                        Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1
                                                        SendDataTo i, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                    Case EventType.evFadeoutBGM
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SFadeoutBGM
                                                        SendDataTo i, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                    Case EventType.evPlaySound
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SPlaySound
                                                        Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1
                                                        SendDataTo i, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                    Case EventType.evStopSound
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SStopSound
                                                        SendDataTo i, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                    Case EventType.evSetAccess
                                                        Player(i).characters(TempPlayer(i).CurChar).access = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                        SendPlayerData i
                                                    Case EventType.evOpenShop
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 > 0 Then ' shop exists?
                                                            If Len(Trim$(Shop(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Name)) > 0 Then ' name exists?
                                                                SendOpenShop i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                                TempPlayer(i).InShop = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 ' stops movement and the like
                                                                .WaitingForResponse = 2
                                                            End If
                                                        End If
                                                    Case EventType.evOpenBank
                                                        SendBank i
                                                        TempPlayer(i).InBank = True
                                                        .WaitingForResponse = 3
                                                    Case EventType.evGiveExp
                                                        GivePlayerEXP i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                    Case EventType.evShowChatBubble
                                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                            Case TARGET_TYPE_PLAYER
                                                                SendChatBubble GetPlayerMap(i), i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, DarkBrown
                                                            Case TARGET_TYPE_NPC
                                                                SendChatBubble GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, DarkBrown
                                                            Case TARGET_TYPE_EVENT
                                                                SendChatBubble GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, DarkBrown
                                                        End Select
                                                    Case EventType.evLabel
                                                        'Do nothing, just a label
                                                    Case EventType.evGotoLabel
                                                        'Find the label's list of commands and slot
                                                        FindEventLabel Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1), GetPlayerMap(i), .eventID, .pageID, .CurSlot, .CurList, .ListLeftOff
                                                    Case EventType.evSpawnNpc
                                                        If Map(GetPlayerMap(i)).Npc(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) > 0 Then
                                                            SpawnNpc Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, GetPlayerMap(i), True
                                                        End If
                                                    Case EventType.evFadeIn
                                                        SendSpecialEffect i, EFFECT_TYPE_FADEIN
                                                    Case EventType.evFadeOut
                                                        SendSpecialEffect i, EFFECT_TYPE_FADEOUT
                                                    Case EventType.evFlashWhite
                                                        SendSpecialEffect i, EFFECT_TYPE_FLASH
                                                    Case EventType.evSetFog
                                                        SendSpecialEffect i, EFFECT_TYPE_FOG, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                    Case EventType.evSetweather
                                                        SendSpecialEffect i, EFFECT_TYPE_WEATHER, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                    Case EventType.evSetTint
                                                        SendSpecialEffect i, EFFECT_TYPE_TINT, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4
                                                    Case EventType.evWait
                                                        .ActionTimer = GetTickCount + Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                    Case EventType.evOpenMail
                                                        SendMailBox i
                                                    Case EventType.evBeginQuest
                                                        If CanStartQuest(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, False) Then
                                                            QuestMessage i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Quest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).QuestDesc, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                        End If
                                                    Case EventType.evEndQuest
                                                        If CanEndQuest(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, False) Then
                                                            EndQuest i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                        End If
                                                    Case EventType.evQuestTask
                                                        If QuestInProgress(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) Then
                                                            If Player(i).characters(TempPlayer(i).CurChar).PlayerQuest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).CurrentTask = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 Then
                                                                If Quest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Task(Player(i).characters(TempPlayer(i).CurChar).PlayerQuest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).CurrentTask).type = TASK_TALKEVENT Or Quest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Task(Player(i).characters(TempPlayer(i).CurChar).PlayerQuest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).CurrentTask).type = TASK_FETCHRETURN Then
                                                                    CheckTask i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Quest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Task(Player(i).characters(TempPlayer(i).CurChar).PlayerQuest(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).CurrentTask).type, -1, True
                                                                End If
                                                            End If
                                                        End If
                                                    Case EventType.evShowPicture
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SPic
                                                        Buffer.WriteLong 0
                                                        Buffer.WriteLong Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1
                                                        Buffer.WriteLong Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                        Buffer.WriteLong Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                        Buffer.WriteLong Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data4
                                                        Buffer.WriteLong Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).data5
                                                        SendDataTo i, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                    Case EventType.evHidePicture
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SPic
                                                        Buffer.WriteLong 1
                                                        Buffer.WriteLong Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1
                                                        SendDataTo i, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                    Case EventType.evWaitMovement
                                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 <= Map(GetPlayerMap(i)).EventCount Then
                                                            If Map(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Global = 1 Then
                                                                .WaitingForResponse = 4
                                                                .EventMovingID = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                                .EventMovingType = 1
                                                            Else
                                                                .WaitingForResponse = 4
                                                                .EventMovingID = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                                .EventMovingType = 0
                                                            End If
                                                        End If
                                                    Case EventType.evHoldPlayer
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SHoldPlayer
                                                        Buffer.WriteLong 0
                                                        SendDataTo i, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                    Case EventType.evReleasePlayer
                                                        Set Buffer = New clsBuffer
                                                        Buffer.WriteLong SHoldPlayer
                                                        Buffer.WriteLong 1
                                                        SendDataTo i, Buffer.ToArray
                                                        Set Buffer = Nothing
                                                End Select
                                            End If
                                        Loop
                                        If endprocess = False Then
                                            .CurSlot = .CurSlot + 1
                                        End If
                                    End If
                                End If
                            End With
                        End If
                        If removeEventProcess = True Then
                            TempPlayer(i).EventProcessing(x).Active = 0
                            restartloop = True
                            removeEventProcess = False
                        End If
                    Next
                Loop
            End If
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProcessEventCommands", "modEventLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub UpdateEventLogic()
Dim i As Long, x As Long, y As Long, z As Long, MapNum As Long, id As Long
Dim page As Long, Buffer As clsBuffer, spawnevent As Boolean, p As Long, rand As Long, isglobal As Boolean, actualmovespeed As Long, playerID As Long, WalkThrough As Long, eventID As Long, sendupdate As Boolean, removeEventProcess As Boolean, w As Long, v As Long
    'Check Removing and Adding of Events (Did switches change or something?)

   On Error GoTo errorhandler

    RemoveDeadEvents
    SpawnNewEvents
    ProcessEventMovement
    ProcessLocalEventMovement
    ProcessEventCommands


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateEventLogic", "modEventLogic", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Sub SendSwitchesAndVariables(Index As Long, Optional everyone As Boolean = False)
Dim Buffer As clsBuffer, i As Long


   On Error GoTo errorhandler

Set Buffer = New clsBuffer
Buffer.WriteLong SSwitchesAndVariables

For i = 1 To MAX_SWITCHES
    Buffer.WriteString Switches(i)
Next

For i = 1 To MAX_VARIABLES
    Buffer.WriteString Variables(i)
Next

If everyone Then
    SendDataToAll Buffer.ToArray
Else
    SendDataTo Index, Buffer.ToArray
End If

Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSwitchesAndVariables", "modEventLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapEventData(Index As Long)
Dim Buffer As clsBuffer, i As Long, x As Long, y As Long, z As Long, MapNum As Long, w As Long

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapEventData
    MapNum = GetPlayerMap(Index)
    'Event Data
    Buffer.WriteLong Map(MapNum).EventCount
        
    If Map(MapNum).EventCount > 0 Then
        For i = 1 To Map(MapNum).EventCount
            With Map(MapNum).Events(i)
                Buffer.WriteString .Name
                Buffer.WriteLong .Global
                Buffer.WriteLong .x
                Buffer.WriteLong .y
                Buffer.WriteLong .PageCount
            End With
            If Map(MapNum).Events(i).PageCount > 0 Then
                For x = 1 To Map(MapNum).Events(i).PageCount
                    With Map(MapNum).Events(i).Pages(x)
                        Buffer.WriteLong .chkVariable
                        Buffer.WriteLong .VariableIndex
                        Buffer.WriteLong .VariableCondition
                        Buffer.WriteLong .VariableCompare
                            
                        Buffer.WriteLong .chkSwitch
                        Buffer.WriteLong .SwitchIndex
                        Buffer.WriteLong .SwitchCompare
                        
                        Buffer.WriteLong .chkHasItem
                        Buffer.WriteLong .HasItemIndex
                        Buffer.WriteLong .HasItemAmount
                            
                        Buffer.WriteLong .chkSelfSwitch
                        Buffer.WriteLong .SelfSwitchIndex
                        Buffer.WriteLong .SelfSwitchCompare
                            
                        Buffer.WriteLong .GraphicType
                        Buffer.WriteLong .Graphic
                        Buffer.WriteLong .GraphicX
                        Buffer.WriteLong .GraphicY
                        Buffer.WriteLong .GraphicX2
                        Buffer.WriteLong .GraphicY2
                        
                        Buffer.WriteLong .MoveType
                        Buffer.WriteLong .MoveSpeed
                        Buffer.WriteLong .MoveFreq
                        Buffer.WriteLong .MoveRouteCount
                        
                        Buffer.WriteLong .IgnoreMoveRoute
                        Buffer.WriteLong .RepeatMoveRoute
                            
                        If .MoveRouteCount > 0 Then
                            For y = 1 To .MoveRouteCount
                                Buffer.WriteLong .MoveRoute(y).Index
                                Buffer.WriteLong .MoveRoute(y).Data1
                                Buffer.WriteLong .MoveRoute(y).Data2
                                Buffer.WriteLong .MoveRoute(y).Data3
                                Buffer.WriteLong .MoveRoute(y).Data4
                                Buffer.WriteLong .MoveRoute(y).data5
                                Buffer.WriteLong .MoveRoute(y).data6
                            Next
                        End If
                            
                        Buffer.WriteLong .WalkAnim
                        Buffer.WriteLong .DirFix
                        Buffer.WriteLong .WalkThrough
                        Buffer.WriteLong .ShowName
                        Buffer.WriteLong .Trigger
                        Buffer.WriteLong .CommandListCount
                        
                        Buffer.WriteLong .Position
                        Buffer.WriteLong .questnum
                    End With
                        
                    If Map(MapNum).Events(i).Pages(x).CommandListCount > 0 Then
                        For y = 1 To Map(MapNum).Events(i).Pages(x).CommandListCount
                            Buffer.WriteLong Map(MapNum).Events(i).Pages(x).CommandList(y).CommandCount
                            Buffer.WriteLong Map(MapNum).Events(i).Pages(x).CommandList(y).ParentList
                            If Map(MapNum).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                For z = 1 To Map(MapNum).Events(i).Pages(x).CommandList(y).CommandCount
                                    With Map(MapNum).Events(i).Pages(x).CommandList(y).Commands(z)
                                        Buffer.WriteLong .Index
                                        Buffer.WriteString .Text1
                                        Buffer.WriteString .Text2
                                        Buffer.WriteString .Text3
                                        Buffer.WriteString .Text4
                                        Buffer.WriteString .Text5
                                        Buffer.WriteLong .Data1
                                        Buffer.WriteLong .Data2
                                        Buffer.WriteLong .Data3
                                        Buffer.WriteLong .Data4
                                        Buffer.WriteLong .data5
                                        Buffer.WriteLong .data6
                                        Buffer.WriteLong .ConditionalBranch.CommandList
                                        Buffer.WriteLong .ConditionalBranch.Condition
                                        Buffer.WriteLong .ConditionalBranch.Data1
                                        Buffer.WriteLong .ConditionalBranch.Data2
                                        Buffer.WriteLong .ConditionalBranch.Data3
                                        Buffer.WriteLong .ConditionalBranch.ElseCommandList
                                        Buffer.WriteLong .MoveRouteCount
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                Buffer.WriteLong .MoveRoute(w).Index
                                                Buffer.WriteLong .MoveRoute(w).Data1
                                                Buffer.WriteLong .MoveRoute(w).Data2
                                                Buffer.WriteLong .MoveRoute(w).Data3
                                                Buffer.WriteLong .MoveRoute(w).Data4
                                                Buffer.WriteLong .MoveRoute(w).data5
                                                Buffer.WriteLong .MoveRoute(w).data6
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    'End Event Data
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing
    Call SendSwitchesAndVariables(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapEventData", "modEventLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function ParseEventText(ByVal Index As Long, ByVal txt As String) As String
Dim i As Long, x As Long, newtxt As String, parsestring As String, z As Long

   On Error GoTo errorhandler

    txt = Replace(txt, "/name", Trim$(Player(Index).characters(TempPlayer(Index).CurChar).Name))
    txt = Replace(txt, "/p", Trim$(Player(Index).characters(TempPlayer(Index).CurChar).Name))
    Do While InStr(1, txt, "/v") > 0
        x = InStr(1, txt, "/v")
        If x > 0 Then
            i = 0
            Do Until IsNumeric(Mid(txt, x + 2 + i, 1)) = False
                i = i + 1
            Loop
            newtxt = Mid(txt, 1, x - 1)
            parsestring = Mid(txt, x + 2, i)
            z = Player(Index).characters(TempPlayer(Index).CurChar).Variables(Val(parsestring))
            newtxt = newtxt & CStr(z)
            newtxt = newtxt & Mid(txt, x + 2 + i, Len(txt) - (x + i))
            txt = newtxt
        End If
    Loop
    ParseEventText = txt


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ParseEventText", "modEventLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function FindEventLabel(ByVal Label As String, MapNum As Long, eventID As Long, pageID As Long, CurSlot As Long, CurList As Long, ListLeftOff() As Long)
Dim Stack() As Long, StackCount As Long, tmpCurSlot As Long, tmpCurList As Long, CurrentListOption() As Long
Dim removeEventProcess As Boolean, tmpListLeftOff() As Long, restartlist As Boolean, w As Long
    'Store the Old data, just in case

   On Error GoTo errorhandler

    tmpCurSlot = CurSlot
    tmpCurList = CurList
    tmpListLeftOff = ListLeftOff
    
    ReDim ListLeftOff(Map(MapNum).Events(eventID).Pages(pageID).CommandListCount)
    ReDim CurrentListOption(Map(MapNum).Events(eventID).Pages(pageID).CommandListCount)
    CurList = 1
    CurSlot = 1
    
    Do Until removeEventProcess = True
        If ListLeftOff(CurList) > 0 Then
            CurSlot = ListLeftOff(CurList)
            ListLeftOff(CurList) = 0
        End If
        If CurList > Map(MapNum).Events(eventID).Pages(pageID).CommandListCount Then
            'Get rid of this event, it is bad
            removeEventProcess = True
        End If
        
        If CurSlot > Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).CommandCount Then
            If CurList = 1 Then
                removeEventProcess = True
            Else
                CurList = Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).ParentList
                CurSlot = 1
                restartlist = True
            End If
        End If
        
        If restartlist = False Then
            If removeEventProcess = False Then
                'If we are still here, then we are good to process shit :D
                Select Case Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Index
                    Case EventType.evShowChoices
                        If Len(Trim$(Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Text2)) > 0 Then
                            w = 1
                            If Len(Trim$(Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Text3)) > 0 Then
                                w = 2
                                If Len(Trim$(Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Text4)) > 0 Then
                                    w = 3
                                    If Len(Trim$(Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Text5)) > 0 Then
                                        w = 4
                                    End If
                                End If
                            End If
                        End If
                        If w > 0 Then
                            If CurrentListOption(CurList) < w Then
                                CurrentListOption(CurList) = CurrentListOption(CurList) + 1
                                'Process
                                ListLeftOff(CurList) = CurSlot
                                Select Case CurrentListOption(CurList)
                                    Case 1
                                        CurList = Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Data1
                                    Case 2
                                        CurList = Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Data2
                                    Case 3
                                        CurList = Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Data3
                                    Case 4
                                        CurList = Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Data4
                                End Select
                                CurSlot = 0
                            Else
                                CurrentListOption(CurList) = 0
                                'continue on
                            End If
                        End If
                        w = 0
                    Case EventType.evCondition
                        If CurrentListOption(CurList) = 0 Then
                            CurrentListOption(CurList) = 1
                            ListLeftOff(CurList) = CurSlot
                            CurList = Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).ConditionalBranch.CommandList
                            CurSlot = 0
                        ElseIf CurrentListOption(CurList) = 1 Then
                            CurrentListOption(CurList) = 2
                            ListLeftOff(CurList) = CurSlot
                            CurList = Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).ConditionalBranch.ElseCommandList
                            CurSlot = 0
                        ElseIf CurrentListOption(CurList) = 2 Then
                            CurrentListOption(CurList) = 0
                        End If
                    Case EventType.evLabel
                        'Do nothing, just a label
                        If Trim$(Map(MapNum).Events(eventID).Pages(pageID).CommandList(CurList).Commands(CurSlot).Text1) = Trim$(Label) Then
                            Exit Function
                        End If
                End Select
                CurSlot = CurSlot + 1
            End If
        End If
        restartlist = False
    Loop
    
    ListLeftOff = tmpListLeftOff
    CurList = tmpCurList
    CurSlot = tmpCurSlot


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindEventLabel", "modEventLogic", Err.Number, Err.Description, Erl
    Err.Clear
End Function
