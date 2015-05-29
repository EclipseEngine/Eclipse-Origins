Attribute VB_Name = "modClientTCP"
Option Explicit
Private PlayerBuffer As clsBuffer

Sub TcpInit()

   On Error GoTo errorhandler

    Set PlayerBuffer = New clsBuffer

    ' connect
    frmMain.Socket.Close
    frmMain.Socket.RemoteHost = Servers(ServerIndex).ip
    frmMain.Socket.RemotePort = Servers(ServerIndex).port





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TcpInit", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub DestroyTCP()

   On Error GoTo errorhandler

    frmMain.Socket.Close




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DestroyTCP", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long



   On Error GoTo errorhandler

    frmMain.Socket.GetData buffer, vbUnicode, DataLength
    PlayerBuffer.WriteBytes buffer()
    If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.length - 4
        If pLength <= PlayerBuffer.length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "IncomingData", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function ConnectToServer(ByVal i As Long) As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit

   On Error GoTo errorhandler

    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    Wait = GetTickCount
    frmMain.Socket.Close
    frmMain.Socket.Connect
    If i = 1 Then
        SetStatus "Connecting to server..."
    End If
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop
    ConnectToServer = IsConnected




   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ConnectToServer", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function IsConnected() As Boolean

   On Error GoTo errorhandler

    If frmMain.Socket.state = sckConnected Then
        IsConnected = True
    End If




   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsConnected", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    ' if the player doesn't exist, the name will equal 0

   On Error GoTo errorhandler

    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If




   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsPlaying", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub SendData(ByRef data() As Byte)
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    If IsConnected Then
        Set buffer = New clsBuffer
                    buffer.WriteLong (UBound(data) - LBound(data)) + 1
        buffer.WriteBytes data()
        frmMain.Socket.SendData buffer.ToArray()
        DoEvents
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendData", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************
Public Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CNewAccount
    buffer.WriteString Name
    buffer.WriteString Password
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendNewAccount", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendDelAccount(ByVal Name As String)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CDelAccount
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDelAccount", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CLogin
    buffer.WriteString Name
    buffer.WriteString Password
    buffer.WriteLong App.Major
    buffer.WriteLong App.Minor
    buffer.WriteLong App.Revision
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendLogin", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long)
Dim buffer As clsBuffer, i As Long, z As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CAddChar
    buffer.WriteString Name
    buffer.WriteLong Sex
    buffer.WriteLong ClassNum
    'Gotta rewrite for faces
    buffer.WriteLong NewCharHair
    buffer.WriteLong NewCharHead
    buffer.WriteLong NewCharEye
    buffer.WriteLong NewCharEyebrow
    buffer.WriteLong NewCharEar
    buffer.WriteLong NewCharMouth
    buffer.WriteLong NewCharNose
    buffer.WriteLong NewCharShirt
    buffer.WriteLong NewCharEtc
    For i = 1 To SpriteEnum.Sprite_Count - 1
        Select Case i
            Case SpriteEnum.Body
                buffer.WriteLong NewCharHead
            Case SpriteEnum.Hair
                buffer.WriteLong NewCharHair
            Case SpriteEnum.Pants
                If Sex = SEX_MALE Then
                    If NumMaleLegs > 0 Then
                        z = Rand(1, UBound(Tex_MaleLegs))
                        If z = 0 Then z = 1
                        If z > UBound(Tex_MaleLegs) Then z = UBound(Tex_MaleLegs)
                        buffer.WriteLong z
                    Else
                        buffer.WriteLong 0
                    End If
                Else
                    If NumFemaleLegs > 0 Then
                        z = Rand(1, UBound(Tex_FemaleLegs))
                        If z = 0 Then z = 1
                        If z > UBound(Tex_FemaleLegs) Then z = UBound(Tex_FemaleLegs)
                        buffer.WriteLong z
                    Else
                        buffer.WriteLong 0
                    End If
                End If
            Case SpriteEnum.Shirt
                buffer.WriteLong NewCharShirt
            Case SpriteEnum.Shoes
                If Sex = SEX_MALE Then
                    If NumMaleShoes > 0 Then
                        z = Rand(1, UBound(Tex_MaleShoes))
                        If z = 0 Then z = 1
                        If z > UBound(Tex_MaleShoes) Then z = UBound(Tex_MaleShoes)
                        buffer.WriteLong z
                    Else
                        buffer.WriteLong 0
                    End If
                Else
                    If NumFemaleShoes > 0 Then
                        z = Rand(1, UBound(Tex_FemaleShoes))
                        If z = 0 Then z = 1
                        If z > UBound(Tex_FemaleShoes) Then z = UBound(Tex_FemaleShoes)
                        buffer.WriteLong z
                    Else
                        buffer.WriteLong 0
                    End If
                End If
        End Select
    Next
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendAddChar", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CUseChar
    buffer.WriteLong CharSlot
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUseChar", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SayMsg(ByVal Text As String)
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SayMsg", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub BroadcastMsg(ByVal Text As String)
Dim buffer As clsBuffer



   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CBroadcastMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "BroadcastMsg", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EmoteMsg(ByVal Text As String)
Dim buffer As clsBuffer



   On Error GoTo errorhandler
    Set buffer = New clsBuffer
    buffer.WriteLong CEmoteMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EmoteMsg", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMsg
    buffer.WriteString MsgTo
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerMsg", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendPlayerMove()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMove
    buffer.WriteLong GetPlayerDir(MyIndex)
    buffer.WriteLong Player(MyIndex).Moving
    buffer.WriteLong Player(MyIndex).X
    buffer.WriteLong Player(MyIndex).Y
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerMove", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendPlayerDir()
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerDir
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerDir", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendPlayerRequestNewMap()
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    If GettingMap Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNewMap
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerRequestNewMap", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendMap()
Dim packet As String
Dim X As Long
Dim Y As Long
Dim i As Long, z As Long, w As Long
Dim buffer As clsBuffer



   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    CanMoveNow = False

    With Map
        buffer.WriteLong CMapData
        buffer.WriteString Trim$(.Name)
        buffer.WriteString Trim$(.Music)
        buffer.WriteString Trim$(.BGS)
        buffer.WriteByte .Moral
        buffer.WriteLong .Up
        buffer.WriteLong .Down
        buffer.WriteLong .Left
        buffer.WriteLong .Right
        buffer.WriteLong .BootMap
        buffer.WriteByte .BootX
        buffer.WriteByte .BootY
            buffer.WriteLong Map.Weather
        buffer.WriteLong Map.WeatherIntensity
            buffer.WriteLong Map.Fog
        buffer.WriteLong Map.FogSpeed
        buffer.WriteLong Map.FogOpacity
            buffer.WriteLong Map.Red
        buffer.WriteLong Map.Green
        buffer.WriteLong Map.Blue
        buffer.WriteLong Map.Alpha
        buffer.WriteByte .MaxX
        buffer.WriteByte .MaxY
    End With

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY

            With Map.Tile(X, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(i).X
                    buffer.WriteLong .Layer(i).Y
                    buffer.WriteLong .Layer(i).Tileset
                Next
                For z = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Autotile(z)
                Next
                buffer.WriteByte .type
                buffer.WriteLong .Data1
                buffer.WriteLong .data2
                buffer.WriteLong .Data3
                buffer.WriteString .Data4
                buffer.WriteByte .DirBlock
            End With

        Next
    Next

    With Map

        For X = 1 To MAX_MAP_NPCS
            buffer.WriteLong .Npc(X)
            buffer.WriteLong .NpcSpawnType(X)
        Next

    End With
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY

            With Map.exTile(X, Y)
                For i = 1 To ExMapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(i).X
                    buffer.WriteLong .Layer(i).Y
                    buffer.WriteLong .Layer(i).Tileset
                Next
                For z = 1 To ExMapLayer.Layer_Count - 1
                    buffer.WriteLong .Autotile(z)
                Next
            End With

        Next
    Next
    
    'Event Data
    buffer.WriteLong Map.EventCount
        If Map.EventCount > 0 Then
        For i = 1 To Map.EventCount
            With Map.Events(i)
                buffer.WriteString .Name
                buffer.WriteLong .Global
                buffer.WriteLong .X
                buffer.WriteLong .Y
                buffer.WriteLong .pageCount
            End With
            If Map.Events(i).pageCount > 0 Then
                For X = 1 To Map.Events(i).pageCount
                    With Map.Events(i).Pages(X)
                        buffer.WriteLong .chkVariable
                        buffer.WriteLong .VariableIndex
                        buffer.WriteLong .VariableCondition
                        buffer.WriteLong .VariableCompare
                                                buffer.WriteLong .chkSwitch
                        buffer.WriteLong .SwitchIndex
                        buffer.WriteLong .SwitchCompare
                                            buffer.WriteLong .chkHasItem
                        buffer.WriteLong .HasItemIndex
                        buffer.WriteLong .HasItemAmount
                                                buffer.WriteLong .chkSelfSwitch
                        buffer.WriteLong .SelfSwitchIndex
                        buffer.WriteLong .SelfSwitchCompare
                                                buffer.WriteLong .GraphicType
                        buffer.WriteLong .Graphic
                        buffer.WriteLong .GraphicX
                        buffer.WriteLong .GraphicY
                        buffer.WriteLong .GraphicX2
                        buffer.WriteLong .GraphicY2
                                            buffer.WriteLong .MoveType
                        buffer.WriteLong .MoveSpeed
                        buffer.WriteLong .MoveFreq
                        buffer.WriteLong .MoveRouteCount
                                            buffer.WriteLong .IgnoreMoveRoute
                        buffer.WriteLong .RepeatMoveRoute
                                                If .MoveRouteCount > 0 Then
                            For Y = 1 To .MoveRouteCount
                                buffer.WriteLong .MoveRoute(Y).Index
                                buffer.WriteLong .MoveRoute(Y).Data1
                                buffer.WriteLong .MoveRoute(Y).data2
                                buffer.WriteLong .MoveRoute(Y).Data3
                                buffer.WriteLong .MoveRoute(Y).Data4
                                buffer.WriteLong .MoveRoute(Y).Data5
                                buffer.WriteLong .MoveRoute(Y).Data6
                            Next
                        End If
                                                buffer.WriteLong .WalkAnim
                        buffer.WriteLong .DirFix
                        buffer.WriteLong .WalkThrough
                        buffer.WriteLong .ShowName
                        buffer.WriteLong .Trigger
                        buffer.WriteLong .CommandListCount
                                            buffer.WriteLong .Position
                        buffer.WriteLong .questnum
                    End With
                                        If Map.Events(i).Pages(X).CommandListCount > 0 Then
                        For Y = 1 To Map.Events(i).Pages(X).CommandListCount
                            buffer.WriteLong Map.Events(i).Pages(X).CommandList(Y).CommandCount
                            buffer.WriteLong Map.Events(i).Pages(X).CommandList(Y).ParentList
                            If Map.Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                For z = 1 To Map.Events(i).Pages(X).CommandList(Y).CommandCount
                                    With Map.Events(i).Pages(X).CommandList(Y).Commands(z)
                                        buffer.WriteLong .Index
                                        buffer.WriteString .Text1
                                        buffer.WriteString .Text2
                                        buffer.WriteString .Text3
                                        buffer.WriteString .Text4
                                        buffer.WriteString .Text5
                                        buffer.WriteLong .Data1
                                        buffer.WriteLong .data2
                                        buffer.WriteLong .Data3
                                        buffer.WriteLong .Data4
                                        buffer.WriteLong .Data5
                                        buffer.WriteLong .Data6
                                        buffer.WriteLong .ConditionalBranch.CommandList
                                        buffer.WriteLong .ConditionalBranch.Condition
                                        buffer.WriteLong .ConditionalBranch.Data1
                                        buffer.WriteLong .ConditionalBranch.data2
                                        buffer.WriteLong .ConditionalBranch.Data3
                                        buffer.WriteLong .ConditionalBranch.ElseCommandList
                                        buffer.WriteLong .MoveRouteCount
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                buffer.WriteLong .MoveRoute(w).Index
                                                buffer.WriteLong .MoveRoute(w).Data1
                                                buffer.WriteLong .MoveRoute(w).data2
                                                buffer.WriteLong .MoveRoute(w).Data3
                                                buffer.WriteLong .MoveRoute(w).Data4
                                                buffer.WriteLong .MoveRoute(w).Data5
                                                buffer.WriteLong .MoveRoute(w).Data6
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

    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMap", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub WarpMeTo(ByVal Name As String)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CWarpMeTo
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "WarpMeTo", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub WarpToMe(ByVal Name As String)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CWarpToMe
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "WarpToMe", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub WarpTo(ByVal MapNum As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CWarpTo
    buffer.WriteLong MapNum
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "WarpTo", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSetAccess
    buffer.WriteString Name
    buffer.WriteLong Access
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSetAccess", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSetSprite
    buffer.WriteLong SpriteNum
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSetSprite", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendKick(ByVal Name As String)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CKickPlayer
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendKick", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendBan(ByVal Name As String, Optional reason As String = "", Optional account As Boolean = False)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CBanPlayer
    buffer.WriteString Name
    If account Then
        buffer.WriteLong 1
    Else
        buffer.WriteLong 0
    End If
    buffer.WriteString reason
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendBan", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendBanList()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CBanList
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendBanList", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendRequestEditItem()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditItem
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendSaveItem(ByVal ItemNum As Long)
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    buffer.WriteLong CSaveItem
    buffer.WriteLong ItemNum
    buffer.WriteBytes ItemData
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSaveItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendRequestEditAnimation()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditAnimation
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditAnimation", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    buffer.WriteLong CSaveAnimation
    buffer.WriteLong Animationnum
    buffer.WriteBytes AnimationData
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSaveAnimation", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendRequestEditNpc()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditNpc
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditNpc", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendSaveNpc(ByVal npcNum As Long)
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    NpcSize = LenB(Npc(npcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(npcNum)), NpcSize
    buffer.WriteLong CSaveNpc
    buffer.WriteLong npcNum
    buffer.WriteBytes NpcData
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSaveNpc", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendRequestEditResource()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditResource
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditResource", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    buffer.WriteLong CSaveResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSaveResource", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendMapRespawn()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CMapRespawn
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapRespawn", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendUseItem(ByVal InvNum As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CUseItem
    buffer.WriteLong InvNum
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUseItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    If InBank Or InShop Then Exit Sub
    ' do basic checks
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    If PlayerInv(InvNum).Num < 1 Or PlayerInv(InvNum).Num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
        If Amount < 0 Or Amount > PlayerInv(InvNum).Value Then Exit Sub
    End If
    If Amount = 0 Then Amount = 1
    Set buffer = New clsBuffer
    buffer.WriteLong CMapDropItem
    buffer.WriteLong InvNum
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    Player(MyIndex).MapGetTimer = GetTickCount + 500




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDropItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendWhosOnline()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CWhosOnline
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendWhosOnline", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSetMotd
    buffer.WriteString MOTD
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMOTDChange", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendRequestEditShop()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditShop
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditShop", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendSaveShop(ByVal ShopNum As Long)
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    buffer.WriteLong CSaveShop
    buffer.WriteLong ShopNum
    buffer.WriteBytes ShopData
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSaveShop", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendRequestEditSpell()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditSpell
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditSpell", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendSaveSpell(ByVal Spellnum As Long)
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    SpellSize = LenB(spell(Spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(spell(Spellnum)), SpellSize
    buffer.WriteLong CSaveSpell
    buffer.WriteLong Spellnum
    buffer.WriteBytes SpellData
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSaveSpell", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendRequestEditMap()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditMap
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditMap", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendBanDestroy(banIndex As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CBanDestroy
    buffer.WriteLong banIndex
    SendData buffer.ToArray()
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendBanDestroy", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSwapInvSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSwapSpellSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendChangeSpellSlots", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub GetPing()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    PingStart = GetTickCount
    Set buffer = New clsBuffer
    buffer.WriteLong CCheckPing
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "GetPing", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendUnequip(ByVal EqNum As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CUnequip
    buffer.WriteLong EqNum
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUnequip", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendRequestPlayerData()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestPlayerData
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestPlayerData", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendRequestItems()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestItems
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestItems", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendRequestAnimations()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestAnimations
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestAnimations", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendRequestNPCS()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNPCS
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestNPCS", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendRequestResources()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestResources
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestResources", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendRequestSpells()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestSpells
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestSpells", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendRequestShops()
    Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestShops
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestShops", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSpawnItem
    buffer.WriteLong tmpItem
    buffer.WriteLong tmpAmount
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSpawnItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendTrainStat(ByVal StatNum As Byte)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CUseStatPoint
    buffer.WriteByte StatNum
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendRequestLevelUp()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestLevelUp
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestLevelUp", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub BuyItem(ByVal Shopslot As Long)
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CBuyItem
    buffer.WriteLong Shopslot
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "BuyItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SellItem(ByVal invslot As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSellItem
    buffer.WriteLong invslot
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SellItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DepositItem(ByVal invslot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CDepositItem
    buffer.WriteLong invslot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DepositItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CWithdrawItem
    buffer.WriteLong bankslot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "WithdrawItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CloseBank()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CCloseBank
    SendData buffer.ToArray()
    Set buffer = Nothing
    InBank = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CloseBank", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CloseShop()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CCloseShop
    SendData buffer.ToArray()
    Set buffer = Nothing
    InShop = 0
    ShopAction = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CloseShop", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub



Public Sub ChangeBankSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CChangeBankSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ChangeBankSlots", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub AdminWarp(ByVal X As Long, ByVal Y As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CAdminWarp
    buffer.WriteLong X
    buffer.WriteLong Y
    SendData buffer.ToArray()
    WalkToX = -1
    WalkToY = -1
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AdminWarp", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub AcceptTrade()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTrade
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AcceptTrade", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DeclineTrade()
Dim buffer As clsBuffer



   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTrade
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DeclineTrade", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub TradeItem(ByVal invslot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CTradeItem
    buffer.WriteLong invslot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    LastItemDesc = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TradeItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub UntradeItem(ByVal invslot As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CUntradeItem
    buffer.WriteLong invslot
    SendData buffer.ToArray()
    Set buffer = Nothing
    LastItemDesc = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UntradeItem", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal slot As Long, ByVal hotbarNum As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarChange
    buffer.WriteLong sType
    buffer.WriteLong slot
    buffer.WriteLong hotbarNum
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendHotbarChange", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendHotbarUse(ByVal slot As Long)
Dim buffer As clsBuffer, X As Long

    ' check if spell

   On Error GoTo errorhandler

    If Hotbar(slot).sType = 2 Then ' spell
        For X = 1 To MAX_PLAYER_SPELLS
            ' is the spell matching the hotbar?
            If PlayerSpells(X) = Hotbar(slot).slot Then
                ' found it, cast it
                CastSpell X
                Exit Sub
            End If
        Next
        ' can't find the spell, exit out
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarUse
    buffer.WriteLong slot
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendHotbarUse", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendMapReport()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CMapReport
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMapReport", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub PlayerSearch(ByVal CurX As Long, ByVal CurY As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    If isInBounds Then
        Set buffer = New clsBuffer
        buffer.WriteLong CSearch
        buffer.WriteLong CurX
        buffer.WriteLong CurY
        SendData buffer.ToArray()
        Set buffer = Nothing
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerSearch", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendTradeRequest()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendTradeRequest", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendAcceptTradeRequest()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendAcceptTradeRequest", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendDeclineTradeRequest()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTradeRequest
    buffer.WriteLong 0
    SendData buffer.ToArray()
    Set buffer = Nothing





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendPartyLeave()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CPartyLeave
    SendData buffer.ToArray()
    Set buffer = Nothing





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPartyLeave", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendPartyRequest()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CPartyRequest
    SendData buffer.ToArray()
    Set buffer = Nothing





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPartyRequest", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendAcceptParty()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptParty
    SendData buffer.ToArray()
    Set buffer = Nothing





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendAcceptParty", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendDeclineParty()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineParty
    SendData buffer.ToArray()
    Set buffer = Nothing





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendDeclineParty", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendRequestEditZone()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditZone
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditZone", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Public Sub SendRequestEditHouse()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditHouse
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditHouse", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendRequestEditProjectiles()
Dim buffer As clsBuffer
    
   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditProjectiles
    SendData buffer.ToArray()
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditProjectiles", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendSaveProjectile(ByVal Index As Long)
    Dim buffer As clsBuffer
    Dim ProjectileSize As Long
    Dim ProjectileData() As Byte
    

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    ProjectileSize = LenB(Projectiles(Index))
    ReDim ProjectileData(ProjectileSize - 1)
    CopyMemory ProjectileData(0), ByVal VarPtr(Projectiles(Index)), ProjectileSize
    buffer.WriteLong CSaveProjectile
    buffer.WriteLong Index
    buffer.WriteBytes ProjectileData
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSaveProjectile", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendRequestProjectiles()
    Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestProjectiles
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestProjectiles", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendClearProjectile(ByVal ProjectileNum As Long, ByVal CollisionIndex As Long, ByVal CollisionType As Byte, ByVal CollisionZone As Long)
    Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CClearProjectile
    buffer.WriteLong ProjectileNum
    buffer.WriteLong CollisionIndex
    buffer.WriteByte CollisionType
    buffer.WriteLong CollisionZone
    SendData buffer.ToArray()
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestProjectiles", "modClientTCP", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
