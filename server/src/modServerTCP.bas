Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()

   On Error GoTo errorhandler
    
    If DebugMode Then
        frmServer.Caption = Options.Game_Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")" & " - Debug Mode: " & ErrorCount & " errors."
    Else
        frmServer.Caption = Options.Game_Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateCaption", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub CreateFullMapCache()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_MAPS
        SetLoadingProgress "Creating map cache.", 33, i / MAX_MAPS
        DoEvents
        Call MapCache_Create(i)
        DoEvents
    Next

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CreateFullMapCache", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function IsConnected(ByVal Index As Long) As Boolean


   On Error GoTo errorhandler

    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsConnected", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function IsPlaying(ByVal Index As Long) As Boolean


   On Error GoTo errorhandler
    If Index = 0 Then IsPlaying = True: Exit Function
    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsPlaying", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean


   On Error GoTo errorhandler

    If IsConnected(Index) Then
        If LenB(Trim$(Player(Index).login)) > 0 Then
            IsLoggedIn = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsLoggedIn", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function IsMultiAccounts(ByVal login As String) As Boolean
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).login)) = LCase$(login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsMultiAccounts", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function IsMultiIPOnline(ByVal ip As String) As Boolean
    Dim i As Long
    Dim n As Long


   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = ip Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsMultiIPOnline", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function IsBanned(ByVal iporaccount As String, Optional account As Boolean) As Boolean
    Dim i As Long

   On Error GoTo errorhandler
    If BanCount > 0 Then
        If account Then
            For i = 1 To BanCount
                If Trim$(Bans(i).BanName) = Trim$(iporaccount) Then
                    IsBanned = True
                    Exit Function
                End If
            Next
        Else
            For i = 1 To BanCount
                If Trim$(Bans(i).IPAddress) = Trim$(iporaccount) Then
                    IsBanned = True
                    Exit Function
                End If
            Next
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsBanned", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SendDataTo(ByVal Index As Long, ByRef data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte


   On Error GoTo errorhandler

    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        TempData = data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.Socket(Index).SendData Buffer.ToArray()
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDataTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendDataToAll(ByRef data() As Byte)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SendDataTo(i, data)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDataToAll", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef data() As Byte)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                Call SendDataTo(i, data)
            End If
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDataToAllBut", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByRef data() As Byte)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, data)
            End If
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDataToMap", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByRef data() As Byte)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If i <> Index Then
                    Call SendDataTo(i, data)
                End If
            End If
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDataToMapBut", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendDataToParty(ByVal partyNum As Long, ByRef data() As Byte)
Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To Party(partyNum).MemberCount
        If Party(partyNum).Member(i) > 0 Then
            Call SendDataTo(Party(partyNum).Member(i), data)
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDataToParty", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataToAll Buffer.ToArray
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "GlobalMsg", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, Buffer.ToArray
        End If
    Next
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AdminMsg", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataTo Index, Buffer.ToArray
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerMsg", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataToMap MapNum, Buffer.ToArray
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapMsg", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    SendDataTo Index, Buffer.ToArray
    DoEvents
    Call CloseSocket(Index)
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AlertMsg", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal Msg As String, ByVal color As Byte)
Dim i As Long
    ' send message to all people

   On Error GoTo errorhandler

    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
                PlayerMsg Party(partyNum).Member(i), Msg, color
            End If
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PartyMsg", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal reason As String)


   On Error GoTo errorhandler

    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & reason & ")", White)
        End If

        Call AlertMsg(Index, "You have lost your connection with " & Options.Game_Name & ".")
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HackingAttempt", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long


   On Error GoTo errorhandler

    If (Index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AcceptConnection", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SocketConnected(ByVal Index As Long)
Dim i As Long


   On Error GoTo errorhandler

    If Index <> 0 Then
        ' make sure they're not banned
        If Not IsBanned(GetPlayerIP(Index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(Index) & ".")
        Else
            Call AlertMsg(Index, "You have been banned from " & Options.Game_Name & ", and can no longer play.")
        End If
        ' re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        SendHighIndex
        
        'Send Basic Server Information
        Dim Buffer As clsBuffer
        Set Buffer = New clsBuffer
        Buffer.WriteLong SServerInfo
        Buffer.WriteString Trim$(Options.Game_Name)
        Buffer.WriteString News
        Buffer.WriteString Credits
        Buffer.WriteLong CharMode
        Buffer.WriteByte 255
        Buffer.WriteString Options.DataFolder
        Buffer.WriteString Options.UpdateURL
        Buffer.WriteLong App.Major
        Buffer.WriteLong App.Minor
        Buffer.WriteLong App.Revision
        Buffer.WriteString NewOptions.MainMenuMusic
        SendDataTo Index, Buffer.ToArray
        Set Buffer = Nothing
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SocketConnected", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long


   On Error GoTo errorhandler

    If GetPlayerAccess(Index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 1000 Then
            If GetTickCount < TempPlayer(Index).DataTimer Then
                Exit Sub
            End If
        End If

        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 25 Then
            If GetTickCount < TempPlayer(Index).DataTimer Then
                Exit Sub
            End If
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = GetTickCount + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If
    
    If Index = 0 Then
    ' Get the data from the socket now
  
    TempPlayer(Index).Buffer.WriteBytes Buffer()
    Else
        frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
        TempPlayer(Index).Buffer.WriteBytes Buffer()
    End If
    
    If TempPlayer(Index).Buffer.Length >= 4 Then
        pLength = TempPlayer(Index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).Buffer.Length - 4
        If pLength <= TempPlayer(Index).Buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).Buffer.ReadLong
            HandleData Index, TempPlayer(Index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(Index).Buffer.Length >= 4 Then
            pLength = TempPlayer(Index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(Index).Buffer.Trim


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "IncomingData", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub CloseSocket(ByVal Index As Long)


   On Error GoTo errorhandler

    If Index > 0 Then
        Call LeftGame(Index)
        Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
        frmServer.Socket(Index).Close
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CloseSocket", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim i As Long, z As Long, w As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong MapNum
    Buffer.WriteString Trim$(Map(MapNum).Name)
    Buffer.WriteString Trim$(Map(MapNum).Music)
    Buffer.WriteString Trim$(Map(MapNum).BGS)
    Buffer.WriteLong Map(MapNum).Revision
    Buffer.WriteByte Map(MapNum).Moral
    Buffer.WriteLong Map(MapNum).Up
    Buffer.WriteLong Map(MapNum).Down
    Buffer.WriteLong Map(MapNum).Left
    Buffer.WriteLong Map(MapNum).Right
    Buffer.WriteLong Map(MapNum).BootMap
    Buffer.WriteByte Map(MapNum).BootX
    Buffer.WriteByte Map(MapNum).BootY
    
    Buffer.WriteLong Map(MapNum).Weather
    Buffer.WriteLong Map(MapNum).WeatherIntensity
    
    Buffer.WriteLong Map(MapNum).Fog
    Buffer.WriteLong Map(MapNum).FogSpeed
    Buffer.WriteLong Map(MapNum).FogOpacity
    
    Buffer.WriteLong Map(MapNum).Red
    Buffer.WriteLong Map(MapNum).Green
    Buffer.WriteLong Map(MapNum).Blue
    Buffer.WriteLong Map(MapNum).Alpha
    
    Buffer.WriteByte Map(MapNum).MaxX
    Buffer.WriteByte Map(MapNum).MaxY

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            With Map(MapNum).Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).x
                    Buffer.WriteLong .Layer(i).y
                    Buffer.WriteLong .Layer(i).Tileset
                Next
                For z = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Autotile(z)
                Next
                Buffer.WriteByte .type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteString .Data4
                Buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(MapNum).Npc(x)
        Buffer.WriteLong Map(MapNum).NpcSpawnType(x)
    Next
    
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            With Map(MapNum).ExTile(x, y)
                For i = 1 To ExMapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).x
                    Buffer.WriteLong .Layer(i).y
                    Buffer.WriteLong .Layer(i).Tileset
                Next
                For z = 1 To ExMapLayer.Layer_Count - 1
                    Buffer.WriteLong .Autotile(z)
                Next
            End With

        Next
    Next

    MapCache(MapNum).data = Buffer.ToArray()
    
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MapCache_Create", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(Index, s, WhoColor)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendWhosOnline", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
    Dim Buffer As clsBuffer, i As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerLevel(Index)
    Buffer.WriteLong GetPlayerPOINTS(Index)
    For i = 1 To FaceEnum.Face_Count - 1
        Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Face(i)
    Next
    For i = 1 To SpriteEnum.Sprite_Count - 1
        Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Sprite(i)
    Next
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Sex
    Buffer.WriteLong GetPlayerMap(Index)
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteLong GetPlayerClass(Index)
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).InHouse
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, i)
    Next
    
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(i)
    Next
    
    For i = 1 To 4
        Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Spell(i)
    Next
    
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.x
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.y
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Dir
    
    Buffer.WriteLong GetPetMaxVital(Index, HP)
    Buffer.WriteLong GetPetMaxVital(Index, MP)
    
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = True Then
        Buffer.WriteLong 1
    Else
        Buffer.WriteLong 0
    End If
    
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.AttackBehaviour
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Points
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Exp
    Buffer.WriteLong GetPetNextLevel(Index)
    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "PlayerData", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub SendJoinMap(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> Index Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    If Player(i).characters(TempPlayer(Index).CurChar).InHouse = Player(Index).characters(TempPlayer(Index).CurChar).InHouse Then
                        SendDataTo Index, PlayerData(i)
                    End If
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendJoinMap", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong Index
    SendDataToMapBut Index, MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendLeaveMap", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String

   On Error GoTo errorhandler

    SendDataToMap GetPlayerMap(Index), PlayerData(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerData", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(MapNum).data) - LBound(MapCache(MapNum).data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(MapNum).data()
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMap", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(MapNum, i).playerName
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).x
        Buffer.WriteLong MapItem(MapNum, i).y
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapItemsTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(MapNum, i).playerName
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).x
        Buffer.WriteLong MapItem(MapNum, i).y
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapItemsToAll", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapNpcVitals(ByVal MapNum As Long, ByVal mapnpcnum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub
    
    Buffer.WriteLong SMapNpcVitals
    Buffer.WriteLong 0
    Buffer.WriteLong mapnpcnum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNpc(MapNum).Npc(mapnpcnum).Vital(i)
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapNpcVitals", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendZoneNpcVitals(ByVal ZoneNum As Long, ByVal ZoneNPCNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcVitals
    Buffer.WriteLong ZoneNum
    Buffer.WriteLong ZoneNPCNum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(i)
    Next

    SendDataToMap ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Map, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendZoneNpcVitals", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long, x As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).x
        Buffer.WriteLong MapNpc(MapNum).Npc(i).y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(HP)
    Next
    
    For i = 1 To MAX_ZONES
        For x = 1 To MAX_MAP_NPCS * 2
            If ZoneNpc(i).Npc(x).Num > 0 Then
                If ZoneNpc(i).Npc(x).Vital(Vitals.HP) > 0 Then
                    If ZoneNpc(i).Npc(x).Map = MapNum Then
                        Buffer.WriteLong i
                        Buffer.WriteLong x
                        Buffer.WriteLong ZoneNpc(i).Npc(x).Num
                        Buffer.WriteLong ZoneNpc(i).Npc(x).x
                        Buffer.WriteLong ZoneNpc(i).Npc(x).y
                        Buffer.WriteLong ZoneNpc(i).Npc(x).Dir
                        Buffer.WriteLong ZoneNpc(i).Npc(x).Vital(HP)
                        Buffer.WriteLong ZoneNpc(i).Npc(x).Map
                    End If
                End If
            End If
        Next
    Next
    
    Buffer.WriteLong 0

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapNpcsTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).x
        Buffer.WriteLong MapNpc(MapNum).Npc(i).y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(HP)
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapNpcsToMap", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendItems(ByVal Index As Long)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendItems", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendAnimations(ByVal Index As Long)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(Index, i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendAnimations", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_NPCS

        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendNpcs", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendResources(ByVal Index As Long)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(Index, i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendResources", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(Index, i)
        Buffer.WriteLong GetPlayerInvItemValue(Index, i)
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendInventory", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal invslot As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong invslot
    Buffer.WriteLong GetPlayerInvItemNum(Index, invslot)
    Buffer.WriteLong GetPlayerInvItemValue(Index, invslot)
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendInventoryUpdate", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerWornEq
    Buffer.WriteLong GetPlayerEquipment(Index, armor)
    Buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(Index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(Index, Shield)
    Buffer.WriteLong GetPlayerEquipment(Index, gloves)
    Buffer.WriteLong GetPlayerEquipment(Index, necklace)
    Buffer.WriteLong GetPlayerEquipment(Index, ring)
    Buffer.WriteLong GetPlayerEquipment(Index, coat)
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendWornEquipment", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapEquipment(ByVal Index As Long)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerEquipment(Index, armor)
    Buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(Index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(Index, Shield)
    Buffer.WriteLong GetPlayerEquipment(Index, gloves)
    Buffer.WriteLong GetPlayerEquipment(Index, necklace)
    Buffer.WriteLong GetPlayerEquipment(Index, ring)
    Buffer.WriteLong GetPlayerEquipment(Index, coat)
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapEquipment", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal Index As Long)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, armor)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Weapon)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Helmet)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, gloves)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, necklace)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, ring)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, coat)
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapEquipmentTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
            
            If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive Then
                Buffer.WriteLong 1
                Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health
                Buffer.WriteLong GetPetMaxVital(Index, HP)
            Else
                Buffer.WriteLong 0
            End If
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
            
            If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive Then
                Buffer.WriteLong 1
                Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana
                Buffer.WriteLong GetPetMaxVital(Index, MP)
            Else
                Buffer.WriteLong 0
            End If
    End Select


    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendVital", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendEXP(ByVal Index As Long)
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(Index)
    Buffer.WriteLong GetPlayerNextLevel(Index)
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendEXP", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendStats(ByVal Index As Long)
Dim i As Long
Dim packet As String
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, i)
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendStats", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendWelcome(ByVal Index As Long)

    ' Send them MOTD

   On Error GoTo errorhandler

    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(Index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendWelcome", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SClassesData
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        Buffer.WriteString Class(i).MaleFaceParts.FHair
        Buffer.WriteString Class(i).MaleFaceParts.FHeads
        Buffer.WriteString Class(i).MaleFaceParts.FEyes
        Buffer.WriteString Class(i).MaleFaceParts.FEyebrows
        Buffer.WriteString Class(i).MaleFaceParts.FEars
        Buffer.WriteString Class(i).MaleFaceParts.FMouth
        Buffer.WriteString Class(i).MaleFaceParts.FNose
        Buffer.WriteString Class(i).MaleFaceParts.FCloth
        Buffer.WriteString Class(i).MaleFaceParts.FEtc
        Buffer.WriteString Class(i).MaleFaceParts.FFace
        
        Buffer.WriteString Class(i).FemaleFaceParts.FHair
        Buffer.WriteString Class(i).FemaleFaceParts.FHeads
        Buffer.WriteString Class(i).FemaleFaceParts.FEyes
        Buffer.WriteString Class(i).FemaleFaceParts.FEyebrows
        Buffer.WriteString Class(i).FemaleFaceParts.FEars
        Buffer.WriteString Class(i).FemaleFaceParts.FMouth
        Buffer.WriteString Class(i).FemaleFaceParts.FNose
        Buffer.WriteString Class(i).FemaleFaceParts.FCloth
        Buffer.WriteString Class(i).FemaleFaceParts.FEtc
        Buffer.WriteString Class(i).FemaleFaceParts.FFace
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).stat(q)
        Next
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendClasses", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        Buffer.WriteString Class(i).MaleFaceParts.FHair
        Buffer.WriteString Class(i).MaleFaceParts.FHeads
        Buffer.WriteString Class(i).MaleFaceParts.FEyes
        Buffer.WriteString Class(i).MaleFaceParts.FEyebrows
        Buffer.WriteString Class(i).MaleFaceParts.FEars
        Buffer.WriteString Class(i).MaleFaceParts.FMouth
        Buffer.WriteString Class(i).MaleFaceParts.FNose
        Buffer.WriteString Class(i).MaleFaceParts.FCloth
        Buffer.WriteString Class(i).MaleFaceParts.FEtc
        Buffer.WriteString Class(i).MaleFaceParts.FFace
        
        Buffer.WriteString Class(i).FemaleFaceParts.FHair
        Buffer.WriteString Class(i).FemaleFaceParts.FHeads
        Buffer.WriteString Class(i).FemaleFaceParts.FEyes
        Buffer.WriteString Class(i).FemaleFaceParts.FEyebrows
        Buffer.WriteString Class(i).FemaleFaceParts.FEars
        Buffer.WriteString Class(i).FemaleFaceParts.FMouth
        Buffer.WriteString Class(i).FemaleFaceParts.FNose
        Buffer.WriteString Class(i).FemaleFaceParts.FCloth
        Buffer.WriteString Class(i).FemaleFaceParts.FEtc
        Buffer.WriteString Class(i).FemaleFaceParts.FFace
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).stat(q)
        Next
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendNewCharClasses", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendLeftGame(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    SendDataToAllBut Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendLeftGame", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerXY", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendPlayerXYToMap(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerXYToMap", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateItemToAll", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateItemTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateAnimationToAll", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateAnimationTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateNpcToAll(ByVal npcnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    NPCSize = LenB(Npc(npcnum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(npcnum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong npcnum
    Buffer.WriteBytes NPCData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateNpcToAll", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal npcnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    NPCSize = LenB(Npc(npcnum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(npcnum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong npcnum
    Buffer.WriteBytes NPCData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateNpcTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateResourceToAll", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateResourceTo(ByVal Index As Long, ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateResourceTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendShops(ByVal Index As Long)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendShops", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateShopToAll", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateShopTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSpells", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendUpdateSpellToAll(ByVal Spellnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(Spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(Spellnum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong Spellnum
    Buffer.WriteBytes SpellData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateSpellToAll", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal Spellnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(Spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(Spellnum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong Spellnum
    Buffer.WriteBytes SpellData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateSpellTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong GetPlayerSpell(Index, i)
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerSpells", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendResourceCacheTo(ByVal Index As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' Check if have player on map

   On Error GoTo errorhandler

    If PlayersOnMap(GetPlayerMap(Index)) = NO Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).Resource_Count

    If ResourceCache(GetPlayerMap(Index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            Buffer.WriteByte ResourceCache(GetPlayerMap(Index)).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).x
            Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).y
        Next

    End If

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendResourceCacheTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendResourceCacheToMap(ByVal MapNum As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(MapNum).Resource_Count

    If ResourceCache(MapNum).Resource_Count > 0 Then

        For i = 0 To ResourceCache(MapNum).Resource_Count
            Buffer.WriteByte ResourceCache(MapNum).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).x
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).y
        Next

    End If

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendResourceCacheToMap", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendDoorAnimation(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SDoorAnimation
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDoorAnimation", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendActionMsg(ByVal MapNum As Long, ByVal message As String, ByVal color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString message
    Buffer.WriteLong color
    Buffer.WriteLong MsgType
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendActionMsg", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendBlood(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendBlood", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendAnimation(ByVal MapNum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0, Optional ByVal OnlyTo As Long = 0, Optional ZoneNum As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    Buffer.WriteLong ZoneNum
    
    If OnlyTo > 0 Then
        SendDataTo OnlyTo, Buffer.ToArray
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendAnimation", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendCooldown(ByVal Index As Long, ByVal slot As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong slot
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendCooldown", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendClearSpellBuffer(ByVal Index As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendClearSpellBuffer", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SayMsg_Map(ByVal MapNum As Long, ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SayMsg_Map", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SayMsg_Global(ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SayMsg_Global", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ResetShopAction(ByVal Index As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ResetShopAction", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendStunned(ByVal Index As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong TempPlayer(Index).StunDuration
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendStunned", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendBank(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Bank(Index).Item(i).Num
        Buffer.WriteLong Bank(Index).Item(i).Value
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendBank", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapKey(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte Value
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapKey", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapKeyToMap(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte Value
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapKeyToMap", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendOpenShop(ByVal Index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopNum
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendOpenShop", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendPlayerMove(ByVal Index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    End If
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerMove", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendTrade(ByVal Index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendTrade", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendCloseTrade(ByVal Index As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendCloseTrade", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendTradeUpdate(ByVal Index As Long, ByVal dataType As Byte)
Dim Buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    

   On Error GoTo errorhandler

    tradeTarget = TempPlayer(Index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(Index).TradeOffer(i).Num).Stackable = 1 Then
                    If TempPlayer(Index).TradeOffer(i).Value = 0 Then TempPlayer(Index).TradeOffer(i).Value = 1
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).price * TempPlayer(Index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Stackable = 1 Then
                    If TempPlayer(tradeTarget).TradeOffer(i).Value = 0 Then TempPlayer(tradeTarget).TradeOffer(i).Value = 1
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    Buffer.WriteLong totalWorth
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendTradeUpdate", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendTradeStatus", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendTarget(ByVal Index As Long)
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    Buffer.WriteLong TempPlayer(Index).Target
    Buffer.WriteLong TempPlayer(Index).TargetType
    Buffer.WriteLong TempPlayer(Index).TargetZone
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendTarget", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendHotbar(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(i).slot
        Buffer.WriteByte Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(i).sType
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendHotbar", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendLoginOk(ByVal Index As Long)
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong Index
    Buffer.WriteLong Player_HighIndex
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendLoginOk", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendInGame(ByVal Index As Long)
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendInGame", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendHighIndex()
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighIndex
    Buffer.WriteLong Player_HighIndex
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendHighIndex", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendPlayerSound(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerSound", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMapSound(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer
    
    ' Check if have player on map

   On Error GoTo errorhandler

    If PlayersOnMap(GetPlayerMap(Index)) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapSound", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).characters(TempPlayer(Index).CurChar).Name)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendTradeRequest", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendPartyInvite(ByVal Index As Long, ByVal targetPlayer As Long)
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(targetPlayer).characters(TempPlayer(Index).CurChar).Name)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPartyInvite", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
Dim Buffer As clsBuffer, i As Long


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partyNum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partyNum).Member(i)
    Next
    Buffer.WriteLong Party(partyNum).MemberCount
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPartyUpdate", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendPartyUpdateTo(ByVal Index As Long)
Dim Buffer As clsBuffer, i As Long, partyNum As Long


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partyNum = TempPlayer(Index).inParty
    If partyNum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partyNum).Member(i)
        Next
        Buffer.WriteLong Party(partyNum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPartyUpdateTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal Index As Long)
Dim Buffer As clsBuffer, i As Long


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong Index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(Index, i)
        Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Vital(i)
    Next
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPartyVitals", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendSpawnItemToMap(ByVal MapNum As Long, ByVal Index As Long)
Dim Buffer As clsBuffer
    
    ' Check if have player on map

   On Error GoTo errorhandler

    If PlayersOnMap(MapNum) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong Index
    Buffer.WriteString MapItem(MapNum, Index).playerName
    Buffer.WriteLong MapItem(MapNum, Index).Num
    Buffer.WriteLong MapItem(MapNum, Index).Value
    Buffer.WriteLong MapItem(MapNum, Index).x
    Buffer.WriteLong MapItem(MapNum, Index).y
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSpawnItemToMap", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendChatBubble(ByVal MapNum As Long, ByVal Target As Long, ByVal TargetType As Long, ByVal message As String, ByVal Colour As Long)
Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatBubble
    Buffer.WriteLong Target
    Buffer.WriteLong TargetType
    Buffer.WriteString message
    Buffer.WriteLong Colour
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendChatBubble", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendSpecialEffect(ByVal Index As Long, EffectType As Long, Optional Data1 As Long = 0, Optional Data2 As Long = 0, Optional Data3 As Long = 0, Optional Data4 As Long = 0)
Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpecialEffect
    
    Select Case EffectType
        Case EFFECT_TYPE_FADEIN
            Buffer.WriteLong EffectType
        Case EFFECT_TYPE_FADEOUT
            Buffer.WriteLong EffectType
        Case EFFECT_TYPE_FLASH
            Buffer.WriteLong EffectType
        Case EFFECT_TYPE_FOG
            Buffer.WriteLong EffectType
            Buffer.WriteLong Data1 'fognum
            Buffer.WriteLong Data2 'fog movement speed
            Buffer.WriteLong Data3 'opacity
        Case EFFECT_TYPE_WEATHER
            Buffer.WriteLong EffectType
            Buffer.WriteLong Data1 'weather type
            Buffer.WriteLong Data2 'weather intensity
        Case EFFECT_TYPE_TINT
            Buffer.WriteLong EffectType
            Buffer.WriteLong Data1 'red
            Buffer.WriteLong Data2 'green
            Buffer.WriteLong Data3 'blue
            Buffer.WriteLong Data4 'alpha
    End Select
    
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSpecialEffect", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendHouseConfigs(ByVal Index As Long)
Dim Buffer As clsBuffer, i As Long

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHouseConfigs
    
    For i = 1 To MAX_HOUSES
        Buffer.WriteString HouseConfig(i).ConfigName
        Buffer.WriteLong HouseConfig(i).BaseMap
        Buffer.WriteLong HouseConfig(i).MaxFurniture
        Buffer.WriteLong HouseConfig(i).price
    Next
    
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendHouseConfigs", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendFurnitureToHouse(HouseIndex As Long)
    Dim Buffer As clsBuffer, i As Long

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SFurniture
    Buffer.WriteLong HouseIndex
    Buffer.WriteLong Player(HouseIndex).characters(TempPlayer(HouseIndex).CurChar).House.FurnitureCount
    If Player(HouseIndex).characters(TempPlayer(HouseIndex).CurChar).House.FurnitureCount > 0 Then
        For i = 1 To Player(HouseIndex).characters(TempPlayer(HouseIndex).CurChar).House.FurnitureCount
            Buffer.WriteLong Player(HouseIndex).characters(TempPlayer(HouseIndex).CurChar).House.Furniture(i).ItemNum
            Buffer.WriteLong Player(HouseIndex).characters(TempPlayer(HouseIndex).CurChar).House.Furniture(i).x
            Buffer.WriteLong Player(HouseIndex).characters(TempPlayer(HouseIndex).CurChar).House.Furniture(i).y
        Next
    End If
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).characters(TempPlayer(i).CurChar).InHouse = HouseIndex Then
                SendDataTo i, Buffer.ToArray
            End If
        End If
    Next
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendFurnitureToHouse", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUnreadMail(Index As Long)
    Dim Name As String, filename As String, i As Long, x As Long, unreadcount As Long, Buffer As clsBuffer

   On Error GoTo errorhandler

    If IsPlaying(Index) Then
        Name = Trim$(Player(Index).login)
        filename = App.path & "\data\accounts\" & Name & "\" & Name & "_char" & CStr(TempPlayer(Index).CurChar) & "_mail.ini"
        i = Val(GetVar(filename, "Mail", "MessageCount"))
        If i > 0 Then
            For x = 1 To i
                If Val(GetVar(filename, "Message" & CStr(x), "Deleted")) = 0 Then
                    If Val(GetVar(filename, "Message" & CStr(x), "Unread")) = 1 Then
                        unreadcount = unreadcount + 1
                    End If
                End If
            Next
        End If
        
        Set Buffer = New clsBuffer
        Buffer.WriteLong SMailUnread
        Buffer.WriteLong unreadcount
        SendDataTo Index, Buffer.ToArray
        Set Buffer = Nothing
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUnreadMail", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMailBox(Index As Long, Optional openmailbox As Long = 1)
    Dim Name As String, filename As String, i As Long, x As Long, unreadcount As Long, Buffer As clsBuffer, z As Long

   On Error GoTo errorhandler

    If IsPlaying(Index) Then
        Name = Trim$(Player(Index).login)
        filename = App.path & "\data\accounts\" & Name & "\" & Name & "_char" & CStr(TempPlayer(Index).CurChar) & "_mail.ini"
        i = Val(GetVar(filename, "Mail", "MessageCount"))
        'If i > 0 Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SMailBox
            If i > 0 Then
                For x = 1 To i
                    If Val(GetVar(filename, "Message" & CStr(x), "Deleted")) = 0 Then
                        z = z + 1
                    End If
                Next
                Buffer.WriteLong z
                Buffer.WriteLong openmailbox
                If z > 0 Then
                    For x = 1 To i
                        If Val(GetVar(filename, "Message" & CStr(x), "Deleted")) = 0 Then
                            Buffer.WriteLong x
                            Buffer.WriteLong Val(GetVar(filename, "Message" & CStr(x), "Unread"))
                            Buffer.WriteString Trim$(GetVar(filename, "Message" & CStr(x), "From"))
                            Buffer.WriteString Trim$(GetVar(filename, "Message" & CStr(x), "Body"))
                            Buffer.WriteLong Val(GetVar(filename, "Message" & CStr(x), "ItemNum"))
                            Buffer.WriteLong Val(GetVar(filename, "Message" & CStr(x), "ItemVal"))
                            Buffer.WriteString GetVar(filename, "Message" & CStr(x), "Date")
                        End If
                    Next
                End If
            Else
                Buffer.WriteLong 0
                Buffer.WriteLong openmailbox
            End If
            SendDataTo Index, Buffer.ToArray
            Set Buffer = Nothing
        'End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMailBox", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub SendPlayerFriends(Index As Long)
    Dim Buffer As clsBuffer, i As Long, blankname As String * ACCOUNT_LENGTH

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SFriends
    For i = 1 To 25
        If Trim$(Player(Index).characters(TempPlayer(Index).CurChar).Friends(i)) = Trim$(blankname) Then
            Buffer.WriteString ""
            Buffer.WriteLong 0
        Else
            Buffer.WriteString Trim$(Player(Index).characters(TempPlayer(Index).CurChar).Friends(i))
            If FindPlayer(Player(Index).characters(TempPlayer(Index).CurChar).Friends(i)) > 0 Then
                Buffer.WriteLong 1
            Else
                Buffer.WriteLong 0
            End If
        End If
    Next
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerFriends", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendPlayers(Index As Long)
    Dim Buffer As clsBuffer, i As Long
    

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayers
    Buffer.WriteLong MAX_PLAYERS
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Buffer.WriteLong 1
            Buffer.WriteString Trim$(Player(i).login)
            Buffer.WriteString Trim$(GetPlayerName(i))
            If GetPlayerClass(i) > 0 Then
                Buffer.WriteString Trim$(Class(GetPlayerClass(i)).Name)
            Else
                Buffer.WriteString "N/A"
            End If
            Buffer.WriteLong GetPlayerMap(i)
            Buffer.WriteLong GetPlayerLevel(i)
        Else
            Buffer.WriteLong 0
        End If
    Next
    
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayers", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendAdmin(Index As Long)
    Dim Buffer As clsBuffer, i As Long
    

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAdmin
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing
    
    SendPlayers Index
    SendMaps Index
    
    If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
        SendAccounts Index
        SendBans Index
    End If
    



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendAdmin", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendAccounts(Index As Long)
    Dim Buffer As clsBuffer, i As Long, x As Long
    

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAccounts
    Buffer.WriteLong AccountCount
    Buffer.WriteLong MAX_PLAYER_CHARS
    If AccountCount > 0 Then
        For i = 1 To AccountCount
            Buffer.WriteString Trim$(account(i).login)
            Buffer.WriteString account(i).ip
            
            If BanCount > 0 Then
                For x = 1 To BanCount
                    If Trim$(Bans(x).BanName) = Trim$(account(i).login) Or Trim$(Bans(x).IPAddress) = Trim$(account(i).ip) Then
                        Buffer.WriteLong 1
                        Exit For
                    Else
                        If x = BanCount Then
                            Buffer.WriteLong 0
                            Exit For
                        End If
                    End If
                Next
            Else
                Buffer.WriteLong 0
            End If
            
            For x = 1 To MAX_PLAYER_CHARS
                Buffer.WriteString Trim(account(i).characters(x))
            Next
        Next
    End If
    
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendAccounts", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendMaps(Index As Long)
    Dim Buffer As clsBuffer, i As Long
    

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMaps
    Buffer.WriteLong MAX_MAPS
    For i = 1 To MAX_MAPS
        Buffer.WriteString Trim$(Map(i).Name)
    Next
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendAccounts", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendBans(Index As Long)
    Dim Buffer As clsBuffer, i As Long, x As Long
    

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBans
    Buffer.WriteLong BanCount
    If BanCount > 0 Then
        For i = 1 To BanCount
            Buffer.WriteString Trim$(Bans(i).IPAddress)
            Buffer.WriteString Trim$(Bans(i).BanName)
            Buffer.WriteString Trim$(Bans(i).BanReason)
            Buffer.WriteString Trim$(Bans(i).BanChar)
        Next
    End If
    
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendAccounts", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendServerOpts(Index As Long)
    Dim Buffer As clsBuffer, i As Long
    

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Call PlayerMsg(Index, "You do not have a high enough access to edit server options!", BrightRed): Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SServerOpts
    Buffer.WriteString News
    Buffer.WriteString Credits
    Buffer.WriteString Options.MOTD
    Buffer.WriteString Options.Game_Name
    Buffer.WriteString Options.Website
    Buffer.WriteString Options.DataFolder
    Buffer.WriteString Options.UpdateURL
    Buffer.WriteLong AccountCount
    Buffer.WriteLong TotalOnlinePlayers
    Buffer.WriteLong StartTime - GetTickCount
    Buffer.WriteString App.Major & "." & App.Minor & "." & App.Revision
    Buffer.WriteLong Options.StaffOnly
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing
    



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendServerOpts", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendGameOpts(Index As Long)
    Dim Buffer As clsBuffer, i As Long
    

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Call PlayerMsg(Index, "You do not have a high enough access to edit server options!", BrightRed): Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SGameOpts
    Buffer.WriteLong NewOptions.CombatMode
    Buffer.WriteLong NewOptions.MaxLevel
    Buffer.WriteString NewOptions.MainMenuMusic
    Buffer.WriteLong NewOptions.ItemLoss
    Buffer.WriteLong NewOptions.ExpLoss
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing
    



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendGameOpts", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendMaxes(Index As Long)
Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMax
    Buffer.WriteLong MAX_MAPS
    Buffer.WriteLong MAX_LEVELS
    If Index = 0 Then
        SendDataToAll Buffer.ToArray
    Else
        SendDataTo Index, Buffer.ToArray
    End If
    Set Buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMaxes", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateProjectileToAll(ByVal ProjectileNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim Projectileize As Long
    Dim ProjectileData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Projectileize = LenB(Projectiles(ProjectileNum))
    ReDim ProjectileData(Projectileize - 1)
    CopyMemory ProjectileData(0), ByVal VarPtr(Projectiles(ProjectileNum)), Projectileize
    
    Buffer.WriteLong SUpdateProjectile
    Buffer.WriteLong ProjectileNum
    Buffer.WriteBytes ProjectileData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateProjectileToAll", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateProjectileTo(ByVal Index As Long, ByVal ProjectileNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim Projectileize As Long
    Dim ProjectileData() As Byte
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Projectileize = LenB(Projectiles(ProjectileNum))
    ReDim ProjectileData(Projectileize - 1)
    CopyMemory ProjectileData(0), ByVal VarPtr(Projectiles(ProjectileNum)), Projectileize
    
    Buffer.WriteLong SUpdateProjectile
    Buffer.WriteLong ProjectileNum
    Buffer.WriteBytes ProjectileData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateProjectileTo", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendProjectiles(ByVal Index As Long)
    Dim i As Long

   On Error GoTo errorhandler

    For i = 1 To MAX_PROJECTILES
        If LenB(Trim$(Projectiles(i).Name)) > 0 Then
            Call SendUpdateProjectileTo(Index, i)
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendProjectiles", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendProjectileToMap(ByVal MapNum As Long, ByVal ProjectileNum As Long)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapProjectile
    
    With MapProjectiles(MapNum, ProjectileNum)
        Buffer.WriteLong ProjectileNum
        Buffer.WriteLong .ProjectileNum
        Buffer.WriteLong .Owner
        Buffer.WriteByte .OwnerType
        Buffer.WriteByte .Dir
        Buffer.WriteLong .x
        Buffer.WriteLong .y
    End With

    SendDataToMap MapNum, Buffer.ToArray
    Set Buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendProjectileToMap", "modServerTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
