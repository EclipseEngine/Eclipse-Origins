Attribute VB_Name = "modPets"
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public MAX_PETS As Long
Public Pet() As PetRec
Public Const ITEM_TYPE_PET As Byte = 10


Public Const TARGET_TYPE_PET As Byte = 7

' PET constants
Public Const PET_BEHAVIOUR_FOLLOW As Byte = 0 'The pet will attack all npcs around
Public Const PET_BEHAVIOUR_GOTO As Byte = 1 'If attacked, the pet will fight back
Public Const PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT As Byte = 1 'The pet will attack all npcs around
Public Const PET_ATTACK_BEHAVIOUR_GUARD As Byte = 2 'If attacked, the pet will fight back
Public Const PET_ATTACK_BEHAVIOUR_DONOTHING As Byte = 3 'The pet will not attack even if attacked

Public Type PetRec
    Num As Long
    Name As String * NAME_LENGTH
    Sprite As Long
    
    Range As Long

    Level As Long
    
    MaxLevel As Long
    ExpGain As Long
    LevelPnts As Long
    
    StatType As Byte '1 for set stats, 2 for relation to owner's stats
    LevelingType As Byte '0 for leveling on own, 1 for not leveling
    
    stat(1 To Stats.Stat_Count - 1) As Byte
    
    Spell(1 To 4) As Long
End Type

Public Type PlayerPetRec
    Num As Long
    Health As Long
    Mana As Long
    Level As Long
    stat(1 To Stats.Stat_Count - 1) As Byte
    Spell(1 To 4) As Long
    x As Long
    y As Long
    Dir As Long
    Alive As Boolean
    AttackBehaviour As Long
    AdoptiveStats As Boolean
    Points As Long
    Exp As Long
End Type


'Database
' **********
' ** pets **
' **********
Sub Savepets()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_PETS
        Call Savepet(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Savepets", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Savepet(ByVal petNum As Long)
    Dim filename As String
    Dim F As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\pets\pet" & petNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Pet(petNum)
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Savepet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub Loadpets()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    

   On Error GoTo errorhandler

    Call Checkpets

    For i = 1 To MAX_PETS
        filename = App.path & "\data\pets\pet" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Pet(i)
        Close #F
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Loadpets", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Checkpets()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\pets\pet" & i & ".dat") Then
            Call Savepet(i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Checkpets", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Clearpet(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Pet(Index)), LenB(Pet(Index)))
    Pet(Index).Name = vbNullString


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Clearpet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub Clearpets()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_PETS
        Call Clearpet(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Clearpets", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub


'ModServerTCP
Sub SendPets(ByVal Index As Long)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_PETS

        If LenB(Trim$(Pet(i).Name)) > 0 Then
            Call SendUpdatePetTo(Index, i)
        End If

    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPets", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Sub SendUpdatePetToAll(ByVal petNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdatePet
    Buffer.WriteLong petNum
    With Pet(petNum)
        Buffer.WriteLong .Num
        Buffer.WriteString .Name
        Buffer.WriteLong .Sprite
        Buffer.WriteLong .Range
        Buffer.WriteLong .Level
        Buffer.WriteLong .MaxLevel
        Buffer.WriteLong .ExpGain
        Buffer.WriteLong .LevelPnts
        Buffer.WriteByte .StatType
        Buffer.WriteByte .LevelingType
        For i = 1 To Stats.Stat_Count - 1
            Buffer.WriteByte .stat(i)
        Next
        For i = 1 To 4
            Buffer.WriteLong .Spell(i)
        Next
    End With
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdatePetToAll", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdatePetTo(ByVal Index As Long, ByVal petNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim PetSize As Long
    Dim PetData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdatePet
    Buffer.WriteLong petNum
    With Pet(petNum)
        Buffer.WriteLong .Num
        Buffer.WriteString .Name
        Buffer.WriteLong .Sprite
        Buffer.WriteLong .Range
        Buffer.WriteLong .Level
        Buffer.WriteLong .MaxLevel
        Buffer.WriteLong .ExpGain
        Buffer.WriteLong .LevelPnts
        Buffer.WriteByte .StatType
        Buffer.WriteByte .LevelingType
        For i = 1 To Stats.Stat_Count - 1
            Buffer.WriteByte .stat(i)
        Next
        For i = 1 To 4
            Buffer.WriteLong .Spell(i)
        Next
    End With
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdatePetTo", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub



'ModPets
Sub ReleasePet(Index)
Dim i As Long

   On Error GoTo errorhandler

    UpdateMapBlock GetPlayerMap(Index), Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, False
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num = 0
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.AttackBehaviour = 0
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Dir = 0
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health = 0
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level = 0
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana = 0
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.x = 0
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.y = 0
    
    TempPlayer(Index).PetTarget = 0
    TempPlayer(Index).PetTargetType = 0
    TempPlayer(Index).PetTargetZone = 0
    TempPlayer(Index).GoToX = -1
    TempPlayer(Index).GoToY = -1
    
    For i = 1 To 4
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Spell(i) = 0
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(i) = 0
    Next
    
    Call SendDataToMap(GetPlayerMap(Index), PlayerData(Index))
    
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index)).Npc(i).Vital(Vitals.HP) > 0 Then
            If MapNpc(GetPlayerMap(Index)).Npc(i).TargetType = TARGET_TYPE_PET Then
                If MapNpc(GetPlayerMap(Index)).Npc(i).Target = Index Then
                    MapNpc(GetPlayerMap(Index)).Npc(i).TargetType = TARGET_TYPE_PLAYER
                    MapNpc(GetPlayerMap(Index)).Npc(i).Target = Index
                End If
            End If
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ReleasePet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub SummonPet(Index As Long, petNum As Long)

   On Error GoTo errorhandler

    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health > 0 Then
        If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num = 0 Then
            Call PlayerMsg(Index, BrightRed, "You have summoned a " & Trim$(Pet(petNum).Name))
        Else
        End If
    End If
    
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num = petNum
    
    For i = 1 To 4
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Spell(i) = Pet(petNum).Spell(i)
    Next
    
    If Pet(petNum).StatType = 0 Then
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health = GetPlayerMaxVital(Index, HP)
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana = GetPlayerMaxVital(Index, MP)
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level = GetPlayerLevel(Index)
        For i = 1 To Stats.Stat_Count - 1
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(i) = Player(Index).characters(TempPlayer(Index).CurChar).stat(i)
        Next
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.AdoptiveStats = True
    Else
        For i = 1 To Stats.Stat_Count - 1
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(i) = Pet(petNum).stat(i)
        Next
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level = Pet(petNum).Level
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.AdoptiveStats = False
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health = GetPetMaxVital(Index, HP)
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana = GetPetMaxVital(Index, MP)
    End If
    
    
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.x = GetPlayerX(Index)
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.y = GetPlayerY(Index)
    
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = True
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Points = 0
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Exp = 0
    
    Player(Index).characters(TempPlayer(Index).CurChar).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_GUARD 'By default it will guard but this can be changed
    
    Call SendDataToMap(GetPlayerMap(Index), PlayerData(Index))


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SummonPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PetMove
' Author    : JC Snider
' Date      : 6/27/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub PetMove(Index As Long, ByVal MapNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If MapNum < MIN_MAPS Or MapNum > MAX_MAPS Or Index <= 0 Or Index > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Dir = Dir
    UpdateMapBlock GetPlayerMap(Index), Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, False

    Select Case Dir
        Case DIR_UP
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong Index
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.x
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.y
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong Index
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.x
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.y
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong Index
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.x
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.y
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPetMove
            Buffer.WriteLong Index
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.x
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.y
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select
    
    UpdateMapBlock GetPlayerMap(Index), Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetMove", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Function CanPetMove(Index As Long, ByVal MapNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If MapNum < MIN_MAPS Or MapNum > MAX_MAPS Or Index <= 0 Or mapnpcnum > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
If Index <= 0 Or Index > Player_HighIndex Then Exit Function
    x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x
    y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y
    
    If x < 0 Or x > Map(MapNum).MaxX Then Exit Function
    If y < 0 Or y > Map(MapNum).MaxY Then Exit Function
    
    CanPetMove = True
    
    If TempPlayer(Index).PetspellBuffer.Spell > 0 Then
        CanPetMove = False
        Exit Function
    End If

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(MapNum).Tile(x, y - 1).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x + 1) And (GetPlayerY(i) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y - 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).characters(TempPlayer(i).CurChar).Pet.Alive = True And (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y - 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x) And (MapNpc(MapNum).Npc(i).y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y - 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y).DirBlock, DIR_UP + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(x, y + 1).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If

                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x) And (GetPlayerY(i) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y + 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).characters(TempPlayer(i).CurChar).Pet.Alive = True And (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y + 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x) And (MapNpc(MapNum).Npc(i).y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y + 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y).DirBlock, DIR_DOWN + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(MapNum).Tile(x - 1, y).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If

                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x - 1) And (GetPlayerY(i) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).characters(TempPlayer(i).CurChar).Pet.Alive = True And (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x - 1) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x - 1) And (MapNpc(MapNum).Npc(i).y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y).DirBlock, DIR_LEFT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(x + 1, y).type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If

                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x + 1) And (GetPlayerY(i) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).characters(TempPlayer(i).CurChar).Pet.Alive = True And (GetPlayerMap(i) = MapNum) And (Player(i).characters(TempPlayer(i).CurChar).Pet.x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x + 1) And (Player(i).characters(TempPlayer(i).CurChar).Pet.y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x + 1) And (MapNpc(MapNum).Npc(i).y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y).DirBlock, DIR_RIGHT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If

    End Select


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPetMove", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Function
Sub PetDir(ByVal Index As Long, ByVal Dir As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If Index <= 0 Or Index > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If TempPlayer(Index).PetspellBuffer.Spell > 0 Then
        Exit Sub
    End If

    Player(Index).characters(TempPlayer(Index).CurChar).Pet.Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPetDir
    Buffer.WriteLong Index
    Buffer.WriteLong Dir
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetDir", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function PetTryWalk(Index As Long, targetx As Long, targety As Long) As Boolean
    Dim i As Long
    Dim x As Long
    Dim MapNum As Long

   On Error GoTo errorhandler

    MapNum = GetPlayerMap(Index)
    x = Index
    
    If IsOneBlockAway(targetx, targety, CLng(Player(Index).characters(TempPlayer(Index).CurChar).Pet.x), CLng(Player(Index).characters(TempPlayer(Index).CurChar).Pet.y)) = False Then
        If PathfindingType = 1 Then
                                            i = Int(Rnd * 5)
                                            ' Lets move the npc
                                            Select Case i
                                                Case 0
                                                    ' Up
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.y > targety And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_UP) Then
                                                            Call PetMove(x, MapNum, DIR_UP, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Down
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.y < targety And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_DOWN) Then
                                                            Call PetMove(x, MapNum, DIR_DOWN, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Left
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x > targetx And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_LEFT) Then
                                                            Call PetMove(x, MapNum, DIR_LEFT, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Right
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x < targetx And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_RIGHT) Then
                                                            Call PetMove(x, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                Case 1
                
                                                    ' Right
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x < targetx And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_RIGHT) Then
                                                            Call PetMove(x, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Left
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x > targetx And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_LEFT) Then
                                                            Call PetMove(x, MapNum, DIR_LEFT, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Down
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.y < targety And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_DOWN) Then
                                                            Call PetMove(x, MapNum, DIR_DOWN, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Up
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.y > targety And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_UP) Then
                                                            Call PetMove(x, MapNum, DIR_UP, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                Case 2
                
                                                    ' Down
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.y < targety And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_DOWN) Then
                                                            Call PetMove(x, MapNum, DIR_DOWN, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Up
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.y > targety And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_UP) Then
                                                            Call PetMove(x, MapNum, DIR_UP, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Right
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x < targetx And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_RIGHT) Then
                                                            Call PetMove(x, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Left
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x > targetx And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_LEFT) Then
                                                            Call PetMove(x, MapNum, DIR_LEFT, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                Case 3
                
                                                    ' Left
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x > targetx And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_LEFT) Then
                                                            Call PetMove(x, MapNum, DIR_LEFT, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Right
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x < targetx And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_RIGHT) Then
                                                            Call PetMove(x, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Up
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.y > targety And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_UP) Then
                                                            Call PetMove(x, MapNum, DIR_UP, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                
                                                    ' Down
                                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.y < targety And Not didwalk Then
                                                        If CanPetMove(x, MapNum, DIR_DOWN) Then
                                                            Call PetMove(x, MapNum, DIR_DOWN, MOVING_WALKING)
                                                            didwalk = True
                                                        End If
                                                    End If
                                            End Select
                                            
                                ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                If Not didwalk Then
                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x - 1 = targetx And Player(x).characters(TempPlayer(x).CurChar).Pet.y = targety Then
                                        If Player(x).characters(TempPlayer(x).CurChar).Pet.Dir <> DIR_LEFT Then
                                            Call PetDir(x, DIR_LEFT)
                                        End If
    
                                        didwalk = True
                                    End If
    
                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x + 1 = targetx And Player(x).characters(TempPlayer(x).CurChar).Pet.y = targety Then
                                        If Player(x).characters(TempPlayer(x).CurChar).Pet.Dir <> DIR_RIGHT Then
                                            Call PetDir(x, DIR_RIGHT)
                                        End If
    
                                        didwalk = True
                                    End If
    
                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x = targetx And Player(x).characters(TempPlayer(x).CurChar).Pet.y - 1 = targety Then
                                        If Player(x).characters(TempPlayer(x).CurChar).Pet.Dir <> DIR_UP Then
                                            Call PetDir(x, DIR_UP)
                                        End If
    
                                        didwalk = True
                                    End If
    
                                    If Player(x).characters(TempPlayer(x).CurChar).Pet.x = targetx And Player(x).characters(TempPlayer(x).CurChar).Pet.y + 1 = targety Then
                                        If Player(x).characters(TempPlayer(x).CurChar).Pet.Dir <> DIR_DOWN Then
                                            Call PetDir(x, DIR_DOWN)
                                        End If
    
                                        didwalk = True
                                    End If
                                End If
        Else
            'Pathfind
            i = FindPetPath(MapNum, x, targetx, targety)
            If i < 4 Then 'Returned an answer. Move the NPC
                If CanPetMove(x, MapNum, i) Then
                    PetMove x, MapNum, i, MOVING_WALKING
                    didwalk = True
                End If
            End If
        End If
    Else
        'Look to target
        If Player(Index).characters(TempPlayer(Index).CurChar).Pet.x > TempPlayer(Index).GoToX Then
            If CanPetMove(x, MapNum, DIR_LEFT) Then
                PetMove x, MapNum, DIR_LEFT, MOVING_WALKING
                didwalk = True
            Else
                PetDir x, DIR_LEFT
                didwalk = True
            End If
        ElseIf Player(Index).characters(TempPlayer(Index).CurChar).Pet.x < TempPlayer(Index).GoToX Then
            If CanPetMove(x, MapNum, DIR_RIGHT) Then
                PetMove x, MapNum, DIR_RIGHT, MOVING_WALKING
                didwalk = True
            Else
                PetDir x, DIR_RIGHT
                didwalk = True
            End If
        ElseIf Player(Index).characters(TempPlayer(Index).CurChar).Pet.y > TempPlayer(Index).GoToY Then
            If CanPetMove(x, MapNum, DIR_UP) Then
                PetMove x, MapNum, DIR_UP, MOVING_WALKING
                didwalk = True
            Else
                PetDir x, DIR_UP
                didwalk = True
            End If
        ElseIf Player(Index).characters(TempPlayer(Index).CurChar).Pet.y < TempPlayer(Index).GoToY Then
            If CanPetMove(x, MapNum, DIR_DOWN) Then
                PetMove x, MapNum, DIR_DOWN, MOVING_WALKING
                didwalk = True
            Else
                PetDir x, DIR_DOWN
                didwalk = True
            End If
        End If
    End If
                                
                                        ' We could not move so Target must be behind something, walk randomly.
                                        If Not didwalk Then
                                            i = Int(Rnd * 2)
        
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
        
                                                If CanPetMove(x, MapNum, i) Then
                                                    Call PetMove(x, MapNum, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
 PetTryWalk = didwalk


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "PetTryWalk", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function FindPetPath(MapNum As Long, Index As Long, targetx As Long, targety As Long) As Long
Dim tim As Long, sX As Long, sY As Long, pos() As Long, reachable As Boolean, x As Long, y As Long, j As Long, LastSum As Long, Sum As Long, FX As Long, FY As Long, i As Long
Dim path() As Vector, LastX As Long, LastY As Long, did As Boolean

'Initialization phase

   On Error GoTo errorhandler

tim = 0
sX = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x
sY = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y

FX = targetx
FY = targety

If FX = -1 Then Exit Function
If FY = -1 Then Exit Function

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
            FindPetPath = 4
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
    FindPetPath = DIR_RIGHT
ElseIf path(1).y > LastY Then
    FindPetPath = DIR_DOWN
ElseIf path(1).y < LastY Then
    FindPetPath = DIR_UP
ElseIf path(1).x < LastX Then
    FindPetPath = DIR_LEFT
End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "FindPetPath", "modServerLoop", Err.Number, Err.Description, Erl
    Err.Clear

End Function


' ###################################
' ##      Pet Attacking NPC        ##
' ###################################

Public Sub TryPetAttackNpc(ByVal Index As Long, ByVal mapnpcnum As Long)
Dim blockAmount As Long
Dim npcnum As Long
Dim MapNum As Long
Dim Damage As Long


   On Error GoTo errorhandler

    Damage = 0

    ' Can we attack the npc?
    If CanPetAttackNpc(Index, mapnpcnum) Then
    
        MapNum = GetPlayerMap(Index)
        npcnum = MapNpc(MapNum).Npc(mapnpcnum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(npcnum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (MapNpc(MapNum).Npc(mapnpcnum).x * 32), (MapNpc(MapNum).Npc(mapnpcnum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(npcnum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (MapNpc(MapNum).Npc(mapnpcnum).x * 32), (MapNpc(MapNum).Npc(mapnpcnum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(mapnpcnum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (Npc(npcnum).stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPetCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
            
        If Damage > 0 Then
            Call PetAttackNpc(Index, mapnpcnum, Damage)
        Else
            Call PlayerMsg(Index, "Your pet's attack does nothing.", BrightRed)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryPetAttackNpc", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function CanPetCrit(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long
    

   On Error GoTo errorhandler

    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Then Exit Function

    CanPetCrit = False
    If NewOptions.CombatMode = 1 Then
        rate = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Agility) / 3
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPetCrit = True
        End If
    Else
        rate = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Agility) / 52.08
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPetCrit = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPetCrit", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetPetDamage(ByVal Index As Long) As Long
    Dim weaponNum As Long
    

   On Error GoTo errorhandler

    GetPetDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Or Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Then
        Exit Function
    End If

    If NewOptions.CombatMode = 1 Then
        GetPetDamage = (Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Strength) * 2) + (Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level * 3) + Random(0, 20)
    Else
        'OLD GetPetDamage = 0.085 * 5 * Player(index).characters(TempPlayer(index).CurChar).Pet.Stat(Stats.Strength) + (Player(index).characters(TempPlayer(index).CurChar).Pet.Level / 5)
        GetPetDamage = 0.085 * 5 * Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Strength) + (Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level / 5)
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPetDamage", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function CanPetAttackNpc(ByVal Attacker As Long, ByVal mapnpcnum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim npcnum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Alive = False Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).Npc(mapnpcnum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Attacker)
    npcnum = MapNpc(MapNum).Npc(mapnpcnum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If


    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
    
    If TempPlayer(Attacker).PetspellBuffer.Spell > 0 And IsSpell = False Then Exit Function
    
        ' exit out early
        If IsSpell Then
             If npcnum > 0 Then
                If Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    CanPetAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        attackspeed = 1000 'Pet cannot weild a weapon

        If npcnum > 0 And GetTickCount > TempPlayer(Attacker).PetAttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Dir
                Case DIR_UP
                    NpcX = MapNpc(MapNum).Npc(mapnpcnum).x
                    NpcY = MapNpc(MapNum).Npc(mapnpcnum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(MapNum).Npc(mapnpcnum).x
                    NpcY = MapNpc(MapNum).Npc(mapnpcnum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(MapNum).Npc(mapnpcnum).x + 1
                    NpcY = MapNpc(MapNum).Npc(mapnpcnum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(MapNum).Npc(mapnpcnum).x - 1
                    NpcY = MapNpc(MapNum).Npc(mapnpcnum).y
            End Select

            If NpcX = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.x Then
                If NpcY = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.y Then
                    If Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        CanPetAttackNpc = True
                    Else
                        CanPetAttackNpc = False
                    End If
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPetAttackNpc", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub PetAttackNpc(ByVal Attacker As Long, ByVal mapnpcnum As Long, ByVal Damage As Long, Optional ByVal Spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim DEF As Long
    Dim MapNum As Long
    Dim npcnum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Damage < 0 Or Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Alive = False Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Attacker)
    npcnum = MapNpc(MapNum).Npc(mapnpcnum).Num
    Name = Trim$(Npc(npcnum).Name)
    
    If Spellnum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong Attacker
        Buffer.WriteLong 1
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
    
    ' Check for weapon
    n = 0 'no weapon, pet :P
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = GetTickCount

    If Damage >= MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(MapNum).Npc(mapnpcnum).x * 32), (MapNpc(MapNum).Npc(mapnpcnum).y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Attacker, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y, SoundEntity.seSpell, Spellnum

        ' Calculate exp to give attacker
        Exp = Npc(npcnum).Exp

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, MapNum
        Else
            ' no party - keep exp for self
            GivePlayerEXP Attacker, Exp
        End If
        
        For n = 1 To 20
            If MapNpc(MapNum).Npc(mapnpcnum).Num > 0 Then
                SpawnItem MapNpc(MapNum).Npc(mapnpcnum).Inventory(n).Num, MapNpc(MapNum).Npc(mapnpcnum).Inventory(n).Value, MapNum, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y
                MapNpc(MapNum).Npc(mapnpcnum).Inventory(n).Value = 0
                MapNpc(MapNum).Npc(mapnpcnum).Inventory(n).Num = 0
            End If
        Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(mapnpcnum).Num = 0
        MapNpc(MapNum).Npc(mapnpcnum).SpawnWait = GetTickCount
        MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.HP) = 0
        MapNpc(MapNum).Npc(mapnpcnum).TargetType = 0
        MapNpc(MapNum).Npc(mapnpcnum).Target = 0
        UpdateMapBlock MapNum, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y, False
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).Npc(mapnpcnum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).Npc(mapnpcnum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong 0
        Buffer.WriteLong mapnpcnum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = MapNum Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = mapnpcnum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                    If TempPlayer(i).PetTargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).PetTarget = mapnpcnum Then
                            TempPlayer(i).PetTarget = 0
                            TempPlayer(i).PetTargetType = TARGET_TYPE_NONE
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.HP) = MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).Npc(mapnpcnum).x * 32), (MapNpc(MapNum).Npc(mapnpcnum).y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Attacker, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y, SoundEntity.seSpell, Spellnum

        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(mapnpcnum).TargetType = TARGET_TYPE_PET ' player's pet
        MapNpc(MapNum).Npc(mapnpcnum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(mapnpcnum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(mapnpcnum).Num Then
                    MapNpc(MapNum).Npc(i).Target = Attacker
                    MapNpc(MapNum).Npc(i).TargetType = TARGET_TYPE_PET ' pet
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(MapNum).Npc(mapnpcnum).stopRegen = True
        MapNpc(MapNum).Npc(mapnpcnum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If Spellnum > 0 Then
            If Spell(Spellnum).StunDuration > 0 Then StunNPC mapnpcnum, MapNum, Spellnum
            ' DoT
            If Spell(Spellnum).Duration > 0 Then
                AddDoT_Npc MapNum, mapnpcnum, Spellnum, Attacker, 3
            End If
        End If
        
        SendMapNpcVitals MapNum, mapnpcnum
    End If

    If Spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).PetAttackTimer = GetTickCount
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetAttackNpc", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub



' ###################################
' ##      NPC Attacking Pet        ##
' ###################################

Public Sub TryNpcAttackPet(ByVal mapnpcnum As Long, ByVal Index As Long)
Dim MapNum As Long, npcnum As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?

   On Error GoTo errorhandler

    If CanNpcAttackPet(mapnpcnum, Index) Then
        MapNum = GetPlayerMap(Index)
        npcnum = MapNpc(MapNum).Npc(mapnpcnum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPetDodge(Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcnum)
        
        If NewOptions.CombatMode = 1 Then
            ' take away armour
            Damage = Damage - ((Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Willpower) * 2) + (Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level * 3))
            ' * 1.5 if crit hit
            If CanNpcCrit(npcnum) Then
                Damage = Damage * 1.5
                SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(mapnpcnum).x * 32), (MapNpc(MapNum).Npc(mapnpcnum).y * 32)
            End If
        Else
            ' if the player blocks, take away the block amount
            blockAmount = 0
            Damage = Damage - blockAmount
            
            ' take away armour
            Damage = Damage - rand(1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Agility) * 2))
            
            ' randomise for up to 10% lower than max hit
            Damage = rand(1, Damage)
            
            ' * 1.5 if crit hit
            If CanNpcCrit(npcnum) Then
                Damage = Damage * 1.5
                SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(mapnpcnum).x * 32), (MapNpc(MapNum).Npc(mapnpcnum).y * 32)
            End If
        End If

        If Damage > 0 Then
            Call NpcAttackPet(mapnpcnum, Index, Damage)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryNpcAttackPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function CanNpcAttackPet(ByVal mapnpcnum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long
    Dim npcnum As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Not IsPlaying(Index) Or Not Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = True Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).Npc(mapnpcnum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Index)
    npcnum = MapNpc(MapNum).Npc(mapnpcnum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum).Npc(mapnpcnum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(MapNum).Npc(mapnpcnum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(Index) And Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = True Then
        If npcnum > 0 Then

            ' Check if at same coordinates
            If (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y + 1 = MapNpc(MapNum).Npc(mapnpcnum).y) And (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x = MapNpc(MapNum).Npc(mapnpcnum).x) Then
                CanNpcAttackPet = True
            Else
                If (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y - 1 = MapNpc(MapNum).Npc(mapnpcnum).y) And (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x = MapNpc(MapNum).Npc(mapnpcnum).x) Then
                    CanNpcAttackPet = True
                Else
                    If (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y = MapNpc(MapNum).Npc(mapnpcnum).y) And (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x + 1 = MapNpc(MapNum).Npc(mapnpcnum).x) Then
                        CanNpcAttackPet = True
                    Else
                        If (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y = MapNpc(MapNum).Npc(mapnpcnum).y) And (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x - 1 = MapNpc(MapNum).Npc(mapnpcnum).x) Then
                            CanNpcAttackPet = True
                        End If
                    End If
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanNpcAttackPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub NpcAttackPet(ByVal mapnpcnum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim MapNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).Npc(mapnpcnum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Victim)
    Name = Trim$(Npc(MapNpc(MapNum).Npc(mapnpcnum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong mapnpcnum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(MapNum).Npc(mapnpcnum).stopRegen = True
    MapNpc(MapNum).Npc(mapnpcnum).stopRegenTimer = GetTickCount

    If Damage >= Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health, BrightRed, 1, (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x * 32), (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y * 32)
        
        ' send the sound
        SendMapSound Victim, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y, SoundEntity.seNpc, MapNpc(MapNum).Npc(mapnpcnum).Num
        
        ' kill player
        Call PlayerMsg(Victim, "Your " & Trim$(Pet(Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Num).Name) & " was killed by a " & Trim$(Npc(MapNpc(MapNum).Npc(mapnpcnum).Num).Name) & ".", BrightRed)
        ReleasePet (Victim)

        ' Now that pet is dead, go for owner
        MapNpc(MapNum).Npc(mapnpcnum).Target = Victim
        MapNpc(MapNum).Npc(mapnpcnum).TargetType = TARGET_TYPE_PLAYER
    Else
        ' Player not dead, just do the damage
        Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health = Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health - Damage
        Call SendPetVital(Victim, Vitals.HP)
        Call SendAnimation(MapNum, Npc(MapNpc(GetPlayerMap(Victim)).Npc(mapnpcnum).Num).Animation, 0, 0, TARGET_TYPE_PET, Victim)
        
        ' send the sound
        SendMapSound Victim, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y, SoundEntity.seNpc, MapNpc(MapNum).Npc(mapnpcnum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x * 32), (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y * 32)
        SendBlood GetPlayerMap(Victim), Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "NpcAttackPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function CanPetAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean


   On Error GoTo errorhandler

    If Not IsSpell Then
        If GetTickCount < TempPlayer(Attacker).PetAttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function
    

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function
    
    If TempPlayer(Attacker).PetspellBuffer.Spell > 0 And IsSpell = False Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Dir
            Case DIR_UP
    
                If Not ((GetPlayerY(Victim) + 1 = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.y) And (GetPlayerX(Victim) = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.x)) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(Victim) - 1 = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.y) And (GetPlayerX(Victim) = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.x)) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(Victim) = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.y) And (GetPlayerX(Victim) + 1 = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.x)) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(Victim) = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.y) And (GetPlayerX(Victim) - 1 = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.x)) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If
    
    ' Don't attack a party member
    If TempPlayer(Attacker).inParty > 0 And TempPlayer(Victim).inParty > 0 Then
        If TempPlayer(Attacker).inParty = TempPlayer(Victim).inParty Then
            Call PlayerMsg(Attacker, "You can't attack another party member!", BrightRed)
            Exit Function
        End If
    End If
  

    CanPetAttackPlayer = True


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPetAttackPlayer", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Function
'Pet Vital Stuffs
Sub SendPetVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPetVital
    
    Buffer.WriteLong Index
    
    If Vital = Vitals.HP Then
        Buffer.WriteLong 1
    ElseIf Vital = Vitals.MP Then
        Buffer.WriteLong 2
    End If

    Select Case Vital
        Case HP
            Buffer.WriteLong GetPetMaxVital(Index, HP)
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health
        Case MP
            Buffer.WriteLong GetPetMaxVital(Index, MP)
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana
    End Select

    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPetVital", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub




' ################
' ## Pet Spells ##
' ################

Public Sub BufferPetSpell(ByVal Index As Long, ByVal spellslot As Long)
    Dim Spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim TargetType As Byte
    Dim Target As Long
    Dim TargetZone As Long
    
    ' Prevent subscript out of range

   On Error GoTo errorhandler

    If spellslot <= 0 Or spellslot > 4 Then Exit Sub
    
    Spellnum = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Spell(spellslot)
    MapNum = GetPlayerMap(Index)
    
    If Spellnum <= 0 Or Spellnum > MAX_SPELLS Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).PetSpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & "'s Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(Spellnum).MPCost

    ' Check if they have enough MP
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana < MPCost Then
        Call PlayerMsg(Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & " does not have enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(Spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level Then
        Call PlayerMsg(Index, Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & " must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(Spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell, even as a pet owner.", BrightRed)
        Exit Sub
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(Spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(Spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(Spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    TargetType = TempPlayer(Index).PetTargetType
    Target = TempPlayer(Index).PetTarget
    TargetZone = TempPlayer(Index).PetTargetZone
    Range = Spell(Spellnum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        'PET
        Case 0, 1, SPELL_TYPE_PET ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                If SpellCastType = SPELL_TYPE_HEALHP Or SpellCastType = SPELL_TYPE_HEALMP Then
                    Target = Index
                    TargetType = TARGET_TYPE_PET
                Else
                    PlayerMsg Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & " does not have a target.", BrightRed
                End If
            End If
            If TargetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg Index, "Target not in range of " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & ".", BrightRed
                Else
                    ' go through spell types
                    If Spell(Spellnum).type <> SPELL_TYPE_DAMAGEHP And Spell(Spellnum).type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPetAttackPlayer(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf TargetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, MapNpc(MapNum).Npc(Target).x, MapNpc(MapNum).Npc(Target).y) Then
                    PlayerMsg Index, "Target not in range of " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & ".", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(Spellnum).type <> SPELL_TYPE_DAMAGEHP And Spell(Spellnum).type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPetAttackNpc(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                 ' if have target, check in range
                If Not isInRange(Range, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, ZoneNpc(TargetZone).Npc(Target).x, ZoneNpc(TargetZone).Npc(Target).y) Then
                    PlayerMsg Index, "Target not in range of " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & ".", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(Spellnum).type <> SPELL_TYPE_DAMAGEHP And Spell(Spellnum).type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPetAttackZoneNpc(Index, TargetZone, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            'PET
            ElseIf TargetType = TARGET_TYPE_PET Then
                ' if have target, check in range
                If Not isInRange(Range, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, Player(Target).characters(TempPlayer(Target).CurChar).Pet.x, Player(Target).characters(TempPlayer(Target).CurChar).Pet.y) Then
                    PlayerMsg Index, "Target not in range of " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & ".", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(Spellnum).type <> SPELL_TYPE_DAMAGEHP And Spell(Spellnum).type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPetAttackPet(Index, Target, Spellnum) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(Spellnum).CastAnim, 0, 0, TARGET_TYPE_PET, Index
        SendActionMsg MapNum, "Casting " & Trim$(Spell(Spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32
        TempPlayer(Index).PetspellBuffer.Spell = spellslot
        TempPlayer(Index).PetspellBuffer.Timer = GetTickCount
        TempPlayer(Index).PetspellBuffer.Target = Target
        TempPlayer(Index).PetspellBuffer.tType = TargetType
        TempPlayer(Index).PetspellBuffer.TargetZone = TargetZone
        Exit Sub
    Else
        SendClearPetSpellBuffer Index
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "BufferPetSpell", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub SendClearPetSpellBuffer(ByVal Index As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearPetSpellBuffer
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendClearPetSpellBuffer", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub


Public Sub PetCastSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal TargetType As Byte, Optional TakeMana As Boolean = True, Optional ZoneNum As Long = 0)
    Dim Spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long, z As Long
    
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    

   On Error GoTo errorhandler

    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > 4 Then Exit Sub

    Spellnum = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Spell(spellslot)
    MapNum = GetPlayerMap(Index)

    MPCost = Spell(Spellnum).MPCost

    ' Check if they have enough MP
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana < MPCost Then
        Call PlayerMsg(Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & " does not have enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(Spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level Then
        Call PlayerMsg(Index, Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & " must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(Spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator for even your pet to cast this spell.", BrightRed)
        Exit Sub
    End If

    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(Spellnum).IsProjectile = True Then
        SpellCastType = 4 ' Projectile
    ElseIf Spell(Spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(Spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(Spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    ' set the vital
    Vital = Spell(Spellnum).Vital
    AoE = Spell(Spellnum).AoE
    Range = Spell(Spellnum).Range
    
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(Spellnum).type
                Case SPELL_TYPE_HEALHP
                    SpellPet_Effect Vitals.HP, True, Index, Vital, Spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPet_Effect Vitals.MP, True, Index, Vital, Spellnum
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x
                y = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y
            ElseIf SpellCastType = 3 Then
                If TargetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
                
                If TargetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                ElseIf TargetType = TARGET_TYPE_NPC Then
                    x = MapNpc(MapNum).Npc(Target).x
                    y = MapNpc(MapNum).Npc(Target).y
                ElseIf TargetType = TARGET_TYPE_PET Then
                    x = Player(Target).characters(TempPlayer(Target).CurChar).Pet.x
                    y = Player(Target).characters(TempPlayer(Target).CurChar).Pet.y
                ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                    x = ZoneNpc(ZoneNum).Npc(Target).x
                    y = ZoneNpc(ZoneNum).Npc(Target).y
                End If
                
                If Not isInRange(Range, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, x, y) Then
                    PlayerMsg Index, Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & "'s target not in range.", BrightRed
                    SendClearPetSpellBuffer Index
                End If
            End If
            Select Case Spell(Spellnum).type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> Index Then
                                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPetAttackPlayer(Index, i, True) And Index <> Target Then
                                            SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PetAttackPlayer Index, i, Vital, Spellnum
                                        End If
                                    End If
                                    If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive = True Then
                                        If isInRange(AoE, x, y, Player(i).characters(TempPlayer(i).CurChar).Pet.x, Player(i).characters(TempPlayer(i).CurChar).Pet.y) Then
                                            If CanPetAttackPet(Index, i, Spellnum) Then
                                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PET, i
                                                PetAttackPet Index, i, Vital, Spellnum
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                    If CanPetAttackNpc(Index, i, True) Then
                                        SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PetAttackNpc Index, i, Vital, Spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_ZONES
                        For z = 1 To MAX_MAP_NPCS * 2
                            If ZoneNpc(i).Npc(z).Num > 0 Then
                                If ZoneNpc(i).Npc(z).Vital(HP) > 0 Then
                                    If ZoneNpc(i).Npc(z).Map = MapNum Then
                                        If isInRange(AoE, x, y, ZoneNpc(i).Npc(z).x, ZoneNpc(i).Npc(z).y) Then
                                            If CanPetAttackZoneNpc(Index, i, z, True) Then
                                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_ZONENPC, z, 0, i
                                                PetAttackZoneNpc Index, i, z, Vital, Spellnum
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(Spellnum).type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(Spellnum).type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(Spellnum).type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                    
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, increment, i, Vital, Spellnum
                                End If
                                If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive Then
                                    If isInRange(AoE, x, y, Player(i).characters(TempPlayer(i).CurChar).Pet.x, Player(i).characters(TempPlayer(i).CurChar).Pet.y) Then
                                        SpellPet_Effect VitalType, increment, i, Vital, Spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If TargetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
            
            If TargetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            ElseIf TargetType = TARGET_TYPE_NPC Then
                x = MapNpc(MapNum).Npc(Target).x
                y = MapNpc(MapNum).Npc(Target).y
            ElseIf TargetType = TARGET_TYPE_PET Then
                x = Player(Target).characters(TempPlayer(Target).CurChar).Pet.x
                y = Player(Target).characters(TempPlayer(Target).CurChar).Pet.y
            ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                x = ZoneNpc(ZoneNum).Npc(Target).x
                y = ZoneNpc(ZoneNum).Npc(Target).y
            End If
                
            If Not isInRange(Range, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, x, y) Then
                PlayerMsg Index, "Target is not in range of your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & "!", BrightRed
                SendClearPetSpellBuffer Index
                Exit Sub
            End If
            
            Select Case Spell(Spellnum).type
                Case SPELL_TYPE_DAMAGEHP
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If CanPetAttackPlayer(Index, Target, True) And Index <> Target Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                PetAttackPlayer Index, Target, Vital, Spellnum
                                DidCast = True
                            End If
                        End If
                    ElseIf TargetType = TARGET_TYPE_NPC Then
                        If CanPetAttackNpc(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                PetAttackNpc Index, Target, Vital, Spellnum
                                DidCast = True
                            End If
                        End If
                    ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                        If CanPetAttackZoneNpc(Index, ZoneNum, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_ZONENPC, Target, 0, ZoneNum
                                PetAttackZoneNpc Index, ZoneNum, Target, Vital, Spellnum
                                DidCast = True
                            End If
                        End If
                    ElseIf TargetType = TARGET_TYPE_PET Then
                        If CanPetAttackPet(Index, Target, Spellnum) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PET, Target
                                PetAttackPet Index, Target, Vital, Spellnum
                                DidCast = True
                            End If
                        End If
                    End If
                    
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(Spellnum).type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    ElseIf Spell(Spellnum).type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(Spellnum).type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    End If
                    
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If Spell(Spellnum).type = SPELL_TYPE_DAMAGEMP Then
                            If CanPetAttackPlayer(Index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, Spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, Spellnum
                        End If
                    ElseIf TargetType = TARGET_TYPE_NPC Then
                        If Spell(Spellnum).type = SPELL_TYPE_DAMAGEMP Then
                            If CanPetAttackNpc(Index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, Spellnum, MapNum
                            End If
                        Else
                            If Spell(Spellnum).type = SPELL_TYPE_HEALHP Or Spell(Spellnum).type = SPELL_TYPE_HEALMP Then
                                SpellPet_Effect VitalType, increment, Index, Vital, Spellnum
                            Else
                                SpellNpc_Effect VitalType, increment, Target, Vital, Spellnum, MapNum
                            End If
                        End If
                    ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                        If Spell(Spellnum).type = SPELL_TYPE_DAMAGEMP Then
                            If CanPetAttackZoneNpc(Index, ZoneNum, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, Spellnum, MapNum, ZoneNum
                            End If
                        Else
                            If Spell(Spellnum).type = SPELL_TYPE_HEALHP Or Spell(Spellnum).type = SPELL_TYPE_HEALMP Then
                                SpellPet_Effect VitalType, increment, Index, Vital, Spellnum
                            Else
                                SpellNpc_Effect VitalType, increment, Target, Vital, Spellnum, MapNum
                            End If
                        End If
                    ElseIf TargetType = TARGET_TYPE_PET Then
                        If Spell(Spellnum).type = SPELL_TYPE_DAMAGEMP Then
                            If CanPetAttackPet(Index, Target, Spellnum) Then
                                SpellPet_Effect VitalType, increment, Target, Vital, Spellnum
                            End If
                        Else
                            SpellPet_Effect VitalType, increment, Target, Vital, Spellnum
                            Call SendPetVital(Target, Vital)
                        End If
                    End If
            End Select
        Case 4 ' Projectile
            Call PetFireProjectile(Index, Spellnum)
            DidCast = True
    End Select
    
    If DidCast Then
        If TakeMana Then Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana - MPCost
        Call SendPetVital(Index, Vitals.MP)
        Call SendPetVital(Index, Vitals.HP)
        
        TempPlayer(Index).PetSpellCD(spellslot) = GetTickCount + (Spell(Spellnum).CDTime * 1000)

        SendActionMsg MapNum, Trim$(Spell(Spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetCastSpell", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SpellPet_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal Spellnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long


   On Error GoTo errorhandler

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(Index), Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PET, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32
        
        ' send the sound
        SendMapSound Index, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y, SoundEntity.seSpell, Spellnum
        
        If increment Then
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health + Damage
            If Spell(Spellnum).Duration > 0 Then
                AddHoT_Pet Index, Spellnum
            End If
        ElseIf Not increment Then
            If Vital = Vitals.HP Then
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health - Damage
            ElseIf Vital = Vitals.MP Then
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana - Damage
            End If
        End If
    End If
    
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health > GetPetMaxVital(Index, HP) Then Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health = GetPetMaxVital(Index, HP)
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana > GetPetMaxVital(Index, MP) Then Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana = GetPetMaxVital(Index, MP)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpellPet_Effect", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Public Sub AddHoT_Pet(ByVal Index As Long, ByVal Spellnum As Long)
Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).PetHoT(i)
            If .Spell = Spellnum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = Spellnum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddHoT_Pet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Public Sub AddDoT_Pet(ByVal Index As Long, ByVal Spellnum As Long, ByVal Caster As Long, AttackerType As Long)
Dim i As Long


   On Error GoTo errorhandler

    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Then Exit Sub

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).PetDoT(i)
            If .Spell = Spellnum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                .AttackerType = AttackerType
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = Spellnum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                .AttackerType = AttackerType
                Exit Sub
            End If
        End With
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddDoT_Pet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub PetAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal Spellnum As Long = 0)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Or Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0 'No Weapon, PET!
    
    If Spellnum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong Attacker
        Buffer.WriteLong 1
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, Spellnum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker) & "'s " & Trim$(Pet(Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Num).Name) & ".", BrightRed)
        If NewOptions.ExpLoss = 0 Then
            ' Calculate exp to give attacker
            Exp = (GetPlayerExp(Victim) \ 10)
    
            ' Make sure we dont get less then 0
            If Exp < 0 Then
                Exp = 0
            End If
    
            If Exp = 0 Then
                Call PlayerMsg(Victim, "You lost no exp.", BrightRed)
                Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
            Else
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                SendEXP Victim
                Call PlayerMsg(Victim, "You lost " & Exp & " exp.", BrightRed)
                
                ' check if we're in a party
                If TempPlayer(Attacker).inParty > 0 Then
                    ' pass through party exp share function
                    Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, MapNum
                Else
                    ' not in party, get exp for self
                    GivePlayerEXP Attacker, Exp
                End If
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                    If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive = True Then
                        If TempPlayer(i).PetTargetType = TARGET_TYPE_PLAYER Then
                            If TempPlayer(i).PetTarget = Victim Then
                                TempPlayer(i).PetTarget = 0
                                TempPlayer(i).PetTargetType = TARGET_TYPE_NONE
                            End If
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If

        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, Spellnum
        
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If Spellnum > 0 Then
            If Spell(Spellnum).StunDuration > 0 Then StunPlayer Victim, Spellnum
            ' DoT
            If Spell(Spellnum).Duration > 0 Then
                AddDoT_Player Victim, Spellnum, Attacker
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).PetAttackTimer = GetTickCount


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetAttackPlayer", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function CanPetAttackPet(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Long = 0) As Boolean


   On Error GoTo errorhandler

    If Not IsSpell Then
        If GetTickCount < TempPlayer(Attacker).PetAttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Or Not IsPlaying(Attacker) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function
    
    If TempPlayer(Attacker).PetspellBuffer.Spell > 0 And IsSpell = False Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Dir
            Case DIR_UP
    
                If Not ((Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y - 1 = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.y) And (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.x)) Then Exit Function
            Case DIR_DOWN
    
                If Not ((Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y + 1 = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.y) And (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.x)) Then Exit Function
            Case DIR_LEFT
    
                If Not ((Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.y) And (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x + 1 = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.x)) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.y) And (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x - 1 = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.x)) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If
    
    ' Don't attack a party member
    If TempPlayer(Attacker).inParty > 0 And TempPlayer(Victim).inParty > 0 Then
        If TempPlayer(Attacker).inParty = TempPlayer(Victim).inParty Then
            Call PlayerMsg(Attacker, "You can't attack another party member!", BrightRed)
            Exit Function
        End If
    End If
    
    If TempPlayer(Attacker).inParty > 0 And TempPlayer(Victim).inParty > 0 And TempPlayer(Attacker).inParty = TempPlayer(Victim).inParty Then
        If IsSpell > 0 Then
            If Spell(IsSpell).type = SPELL_TYPE_HEALMP Or Spell(IsSpell).type = SPELL_TYPE_HEALHP Then
                'Carry On :D
            Else
               'Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & " because you are in a party with him!", BrightRed)
               Exit Function
            End If
        Else
            'Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & " because you are in a party with him!", BrightRed)
            Exit Function
        End If
    End If

    CanPetAttackPet = True


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPetAttackPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Function
Sub PetAttackPet(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal Spellnum As Long = 0)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Or Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Alive = False Or Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0 'No Weapon, PET!
    
    If Spellnum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong Attacker
        Buffer.WriteLong 1
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = GetTickCount

    If Damage >= Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health Then
        SendActionMsg GetPlayerMap(Victim), "-" & Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health, BrightRed, 1, (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x * 32), (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y * 32)
        UpdateMapBlock GetPlayerMap(Victim), Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y, False
        ' send the sound
        If Spellnum > 0 Then SendMapSound Victim, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y, SoundEntity.seSpell, Spellnum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker) & "'s " & Trim$(Pet(Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Num).Name) & ".", BrightRed)
        ' Calculate exp to give attacker
        Exp = (GetPlayerExp(Victim) \ 10)

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If

        If Exp = 0 Then
            Call PlayerMsg(Victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            SendEXP Victim
            Call PlayerMsg(Victim, "You lost " & Exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, MapNum
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, Exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                    If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive = True Then
                        If TempPlayer(i).PetTargetType = TARGET_TYPE_PLAYER Then
                            If TempPlayer(i).PetTarget = Victim Then
                                TempPlayer(i).PetTarget = 0
                                TempPlayer(i).PetTargetType = TARGET_TYPE_NONE
                            End If
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
        
        ' kill player
        Call PlayerMsg(Victim, "Your " & Trim$(Pet(Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Num).Name) & " was killed by " & Trim$(GetPlayerName(Attacker)) & "'s " & Trim$(Pet(Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Num).Name) & "!", BrightRed)
        ReleasePet (Victim)
    Else
        ' Player not dead, just do the damage
        Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health = Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health - Damage
        Call SendPetVital(Victim, Vitals.HP)
        
        'Set pet to begin attacking the other pet if it isn't dead or dosent have another target
        If TempPlayer(Victim).PetTarget <= 0 And TempPlayer(Victim).PetBehavior <> PET_BEHAVIOUR_GOTO Then
            TempPlayer(Victim).PetTarget = Attacker
            TempPlayer(Victim).PetTargetType = TARGET_TYPE_PET
        End If
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Victim, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y, SoundEntity.seSpell, Spellnum
        
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x * 32), (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y * 32)
        SendBlood GetPlayerMap(Victim), Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y
        
        ' set the regen timer
        TempPlayer(Victim).PetstopRegen = True
        TempPlayer(Victim).PetstopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If Spellnum > 0 Then
            If Spell(Spellnum).StunDuration > 0 Then StunPet Victim, Spellnum
            ' DoT
            If Spell(Spellnum).Duration > 0 Then
                AddDoT_Pet Victim, Spellnum, Attacker, TARGET_TYPE_PET
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).PetAttackTimer = GetTickCount


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetAttackPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Public Sub StunPet(ByVal Index As Long, ByVal Spellnum As Long)
    ' check if it's a stunning spell

   On Error GoTo errorhandler

    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = True Then
        If Spell(Spellnum).StunDuration > 0 Then
            ' set the values on index
            TempPlayer(Index).PetStunDuration = Spell(Spellnum).StunDuration
            TempPlayer(Index).PetStunTimer = GetTickCount
            ' tell him he's stunned
            PlayerMsg Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & " has been stunned.", BrightRed
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "StunPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub HandleDoT_Pet(ByVal Index As Long, ByVal dotNum As Long)

   On Error GoTo errorhandler

    With TempPlayer(Index).PetDoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If .AttackerType = TARGET_TYPE_PET Then
                    If CanPetAttackPet(.Caster, Index, .Spell) Then
                        PetAttackPet .Caster, Index, Spell(.Spell).Vital
                        Call SendPetVital(Index, HP)
                        Call SendPetVital(Index, MP)
                    End If
                ElseIf .AttackerType = TARGET_TYPE_PLAYER Then
                    If CanPlayerAttackPet(.Caster, Index, .Spell) Then
                        PlayerAttackPet .Caster, Index, Spell(.Spell).Vital
                        Call SendPetVital(Index, HP)
                        Call SendPetVital(Index, MP)
                    End If
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleDoT_Pet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub HandleHoT_Pet(ByVal Index As Long, ByVal hotNum As Long)

   On Error GoTo errorhandler

    With TempPlayer(Index).PetHoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg GetPlayerMap(Index), "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health + Spell(.Spell).Vital
                If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health > GetPetMaxVital(Index, HP) Then Player(Index).characters(TempPlayer(Index).CurChar).Pet.Health = GetPetMaxVital(Index, HP)
                If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana > GetPetMaxVital(Index, MP) Then Player(Index).characters(TempPlayer(Index).CurChar).Pet.Mana = GetPetMaxVital(Index, MP)
                Call SendPetVital(Index, HP)
                Call SendPetVital(Index, MP)
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleHoT_Pet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub TryPetAttackPlayer(ByVal Index As Long, Victim As Long)
Dim MapNum As Long, npcnum As Long, blockAmount As Long, Damage As Long
    

   On Error GoTo errorhandler

    If GetPlayerMap(Index) <> GetPlayerMap(Victim) Then Exit Sub
    
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Then Exit Sub

    ' Can the npc attack the player?
    If CanPetAttackPlayer(Index, Victim) Then
        MapNum = GetPlayerMap(Index)
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(Victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(Victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if crit hit
        If CanPetCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32)
        End If

        If Damage > 0 Then
            Call PetAttackPlayer(Index, Victim, Damage)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryPetAttackPlayer", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function CanPetDodge(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long
    

   On Error GoTo errorhandler

    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Then Exit Function

    CanPetDodge = False
    If NewOptions.CombatMode = 1 Then
        rate = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Agility) / 4
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPetDodge = True
        End If
    Else
        rate = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Agility) / 83.3
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPetDodge = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPetDodge", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function CanPetParry(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long


   On Error GoTo errorhandler

    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Then Exit Function
    
    CanPetParry = False
    If NewOptions.CombatMode = 1 Then
        rate = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Agility) / 6
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPetParry = True
        End If
    Else
        rate = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Strength) * 0.25
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPetParry = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPetParry", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Sub TryPetAttackPet(ByVal Index As Long, Victim As Long)
Dim MapNum As Long, npcnum As Long, blockAmount As Long, Damage As Long
    

   On Error GoTo errorhandler

    If GetPlayerMap(Index) <> GetPlayerMap(Victim) Then Exit Sub
    
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Or Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Alive = False Then Exit Sub

    ' Can the npc attack the player?
    If CanPetAttackPet(Index, Victim) Then
        MapNum = GetPlayerMap(Index)
    
        ' check if PLAYER can avoid the attack
        If CanPetDodge(Victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x * 32), (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y * 32)
            Exit Sub
        End If
        If CanPetParry(Victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x * 32), (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
        ' if the player blocks, take away the block amount
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if crit hit
        If CanPetCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32)
        End If

        If Damage > 0 Then
            Call PetAttackPet(Index, Victim, Damage)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryPetAttackPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function CanPlayerAttackPet(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

   On Error GoTo errorhandler

    If IsSpell = False Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function
    
    If Not Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Alive Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function

    If IsSpell = False Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
    
                If Not ((Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y + 1 = GetPlayerY(Attacker)) And (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y - 1 = GetPlayerY(Attacker)) And (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y = GetPlayerY(Attacker)) And (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y = GetPlayerY(Attacker)) And (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "s " & Trim$(Pet(Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Num).Name) & "!", BrightRed)
        Exit Function
    End If
    
    ' Don't attack a party member
    If TempPlayer(Attacker).inParty > 0 And TempPlayer(Victim).inParty > 0 Then
        If TempPlayer(Attacker).inParty = TempPlayer(Victim).inParty Then
            Call PlayerMsg(Attacker, "You can't attack another party member!", BrightRed)
            Exit Function
        End If
    End If

    
    
    If TempPlayer(Attacker).inParty > 0 And TempPlayer(Victim).inParty > 0 And TempPlayer(Attacker).inParty = TempPlayer(Victim).inParty Then
        If IsSpell > 0 Then
            If Spell(IsSpell).type = SPELL_TYPE_HEALMP Or Spell(IsSpell).type = SPELL_TYPE_HEALHP Then
                'Carry On :D
            Else
               'Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & " because you are in a party with him!", BrightRed)
               Exit Function
            End If
        Else
            'Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & " because you are in a party with him!", BrightRed)
            Exit Function
        End If
    End If

    CanPlayerAttackPet = True


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPlayerAttackPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Function
Sub PlayerAttackPet(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal Spellnum As Long = 0)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Or Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Alive = False Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health Then
        SendActionMsg GetPlayerMap(Victim), "-" & Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health, BrightRed, 1, (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x * 32), (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y * 32)
        UpdateMapBlock GetPlayerMap(Victim), Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y, False
        ' send the sound
        If Spellnum > 0 Then SendMapSound Victim, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y, SoundEntity.seSpell, Spellnum
        
        ' Calculate exp to give attacker
        Exp = (GetPlayerExp(Victim) \ 10)

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If

        If Exp = 0 Then
            Call PlayerMsg(Victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            SendEXP Victim
            Call PlayerMsg(Victim, "You lost " & Exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, MapNum
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, Exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PET Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        Call PlayerMsg(Victim, "Your " & Trim$(Pet(Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Num).Name) & " was killed by  " & Trim$(GetPlayerName(Attacker)) & ".", BrightRed)
        ReleasePet (Victim)
    Else
        ' Player not dead, just do the damage
        Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health = Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health - Damage
        Call SendPetVital(Victim, Vitals.HP)
        
        'Set pet to begin attacking the other pet if it isn't dead or dosent have another target
        If TempPlayer(Victim).PetTarget <= 0 And TempPlayer(Victim).PetBehavior <> PET_BEHAVIOUR_GOTO Then
            TempPlayer(Victim).PetTarget = Attacker
            TempPlayer(Victim).PetTargetType = TARGET_TYPE_PLAYER
        End If
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Victim, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y, SoundEntity.seSpell, Spellnum
        
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x * 32), (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y * 32)
        SendBlood GetPlayerMap(Victim), Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y
        
        ' set the regen timer
        TempPlayer(Victim).PetstopRegen = True
        TempPlayer(Victim).PetstopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If Spellnum > 0 Then
            If Spell(Spellnum).StunDuration > 0 Then StunPet Victim, Spellnum
            ' DoT
            If Spell(Spellnum).Duration > 0 Then
                AddDoT_Pet Victim, Spellnum, Attacker, TARGET_TYPE_PLAYER
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerAttackPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function IsPetByPlayer(Index) As Boolean
    Dim x As Long, y As Long, x1 As Long, y1 As Long

   On Error GoTo errorhandler

    If Index <= 0 Or Index > MAX_PLAYERS Or Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Then Exit Function
    
    IsPetByPlayer = False
    
    x = GetPlayerX(Index)
    y = GetPlayerY(Index)
    x1 = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x
    y1 = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y
    
    If x = x1 Then
        If y = y1 + 1 Or y = y1 - 1 Then
            IsPetByPlayer = True
        End If
    ElseIf y = y1 Then
        If x = x1 - 1 Or x = x1 + 1 Then
            IsPetByPlayer = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsPetByPlayer", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Function
Function GetPetVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range

   On Error GoTo errorhandler

    If Index <= 0 Or Index > MAX_PLAYERS Or Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Then
        GetPetVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetPetVitalRegen = i


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPetVitalRegen", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub TryPlayerAttackPet(ByVal Attacker As Long, ByVal Victim As Long)
Dim blockAmount As Long
Dim MapNum As Long
Dim Damage As Long, x As Long, i As Long


   On Error GoTo errorhandler

    Damage = 0
    If Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Alive = False Then Exit Sub
    ' Can we attack the npc?
    If CanPlayerAttackPet(Attacker, Victim) Then
    
        MapNum = GetPlayerMap(Attacker)
    
        TempPlayer(Attacker).Target = Victim
        TempPlayer(Attacker).TargetType = TARGET_TYPE_PET
    
        ' check if NPC can avoid the attack
        If CanPetDodge(Victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If
        If CanPetParry(Victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = 0
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (GetPlayerStat(Victim, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        End If
        
        If Damage > 0 Then
            Call PlayerAttackPet(Attacker, Victim, Damage)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryPlayerAttackPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function GetPetMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long


   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If NewOptions.CombatMode = 1 Then
        Select Case Vital
            Case HP
                GetPetMaxVital = ((Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level * 4) + (Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Endurance) * 10)) + 150
            Case MP
                GetPetMaxVital = ((Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level * 4) + (Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Willpower) / 2)) * 5 + 50
        End Select
    Else
        Select Case Vital
            Case HP
                GetPetMaxVital = ((Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level / 2) + (Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Endurance) / 2)) * 15 + 100
            Case MP
                GetPetMaxVital = ((Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level / 2) + (Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Willpower) / 2)) * 5 + 50
        End Select
    End If



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPetMaxVital", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function GetPetNextLevel(ByVal Index As Long) As Long

   On Error GoTo errorhandler
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = True Then
        If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level = Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).MaxLevel Then GetPetNextLevel = 0: Exit Function
        GetPetNextLevel = (50 / 3) * ((Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level + 1) ^ 3 - (6 * (Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level + 1) ^ 2) + 17 * (Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level + 1) - 12)
    End If

   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPetNextLevel", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub CheckPetLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    

   On Error GoTo errorhandler

    level_count = 0
    
    Do While Player(Index).characters(TempPlayer(Index).CurChar).Pet.Exp >= GetPetNextLevel(Index)
        expRollover = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Exp - GetPetNextLevel(Index)
        
        ' can level up?
        If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level < 99 And Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level < Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).MaxLevel Then
            Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level + 1
        End If
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Points = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Points + Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).LevelPnts
        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Exp = expRollover
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            PlayerMsg Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & " has gained " & level_count & " level!", Brown
        Else
            'plural
            PlayerMsg Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & " has gained " & level_count & " levels!", Brown
        End If
        SendPlayerData Index
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckPetLevelUp", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub TryPetAttackZoneNpc(ByVal Index As Long, ZoneNum As Long, ZoneNPCNum As Long)
Dim blockAmount As Long
Dim npcnum As Long
Dim MapNum As Long
Dim Damage As Long


   On Error GoTo errorhandler

    Damage = 0

    ' Can we attack the npc?
    If CanPetAttackZoneNpc(Index, ZoneNum, ZoneNPCNum) Then
    
        MapNum = GetPlayerMap(Index)
        npcnum = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(npcnum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32), (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(npcnum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32), (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
       If NewOptions.CombatMode = 1 Then
            If CanNpcBlock(npcnum) Then
                SendActionMsg MapNum, "Block!", BrightCyan, 1, (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32), (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32)
                Damage = 0
                Exit Sub
            Else
                Damage = Damage - ((Npc(npcnum).stat(Stats.Willpower) * 2) + (Npc(npcnum).Level * 3))
                ' * 1.5 if it's a crit!
                If CanPetCrit(Index) Then
                    Damage = Damage * 1.5
                    SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32)
                End If
            End If
        Else
            ' if the npc blocks, take away the block amount
            blockAmount = CanNpcBlock(npcnum)
            Damage = Damage - blockAmount
            
            ' take away armour
            Damage = Damage - rand(1, (Npc(npcnum).stat(Stats.Agility) * 2))
            ' randomise from 1 to max hit
            Damage = rand(1, Damage)
            
            ' * 1.5 if it's a crit!
            If CanPetCrit(Index) Then
                Damage = Damage * 1.5
                SendActionMsg MapNum, "Critical!", BrightCyan, 1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32)
            End If
        End If
            
        If Damage > 0 Then
            Call PetAttackZoneNpc(Index, ZoneNum, ZoneNPCNum, Damage)
        Else
            Call PlayerMsg(Index, "Your pet's attack does nothing.", BrightRed)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryPetAttackZoneNpc", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Public Function CanPetAttackZoneNpc(ByVal Attacker As Long, ZoneNum As Long, ZoneNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim npcnum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Attacker) = False Or ZoneNPCNum <= 0 Or ZoneNPCNum > MAX_MAP_NPCS * 2 Then
        Exit Function
    End If
    
    If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Map <> GetPlayerMap(Attacker) Then Exit Function
    

    ' Check for subscript out of range
    If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Attacker)
    npcnum = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
    
    ' Make sure the npc isn't already dead
    If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP) <= 0 Then
        If Npc(ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If Npc(ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
    
        ' exit out early
        If IsSpell Then
             If npcnum > 0 Then
                If Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    CanPetAttackZoneNpc = True
                    Exit Function
                End If
            End If
        End If


        attackspeed = 1000

        If npcnum > 0 And GetTickCount > TempPlayer(Attacker).PetAttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Dir
                Case DIR_UP
                    NpcX = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x
                    NpcY = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y - 1
                Case DIR_DOWN
                    NpcX = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x
                    NpcY = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y + 1
                Case DIR_LEFT
                    NpcX = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x + 1
                    NpcY = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y
                Case DIR_RIGHT
                    NpcX = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x - 1
                    NpcY = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y
            End Select

            If NpcX = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.x Then
                If NpcY = Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.y Then
                    If Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        CanPetAttackZoneNpc = True
                    ElseIf Npc(npcnum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or Npc(npcnum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                        If Len(Trim$(Npc(npcnum).AttackSay)) > 0 Then
                            PlayerMsg Attacker, Trim$(Npc(npcnum).Name) & ": " & Trim$(Npc(npcnum).AttackSay), White
                        End If
                    End If
                End If
            End If
        End If
    End If
    


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPetAttackZoneNpc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub PetAttackZoneNpc(ByVal Attacker As Long, ByVal ZoneNum As Long, ByVal ZoneNPCNum As Long, ByVal Damage As Long, Optional ByVal Spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim DEF As Long
    Dim MapNum As Long
    Dim npcnum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Attacker) = False Or ZoneNPCNum <= 0 Or ZoneNPCNum > MAX_MAP_NPCS * 2 Or Damage < 0 Or Player(Attacker).characters(TempPlayer(Attacker).CurChar).Pet.Alive = False Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Attacker)
    npcnum = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
    Name = Trim$(Npc(npcnum).Name)
    
    If Spellnum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SAttack
        Buffer.WriteLong Attacker
        Buffer.WriteLong 1
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
    
    ' Check for weapon
    n = 0 'no weapon, pet :P
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = GetTickCount

    If Damage >= ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(Attacker), "-" & ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP), BrightRed, 1, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32
        SendBlood GetPlayerMap(Attacker), ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Attacker, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y, SoundEntity.seSpell, Spellnum

        ' Calculate exp to give attacker
        Exp = Npc(npcnum).Exp

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, MapNum
        Else
            ' no party - keep exp for self
            GivePlayerEXP Attacker, Exp
        End If
        
        For n = 1 To 20
            If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num > 0 Then
                SpawnItem ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(n).Num, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(n).Value, MapNum, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y
                ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(n).Value = 0
                ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(n).Num = 0
            End If
        Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num = 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).SpawnWait = GetTickCount
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP) = 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).TargetType = 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Target = 0
        UpdateMapBlock MapNum, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y, False
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With ZoneNpc(ZoneNum).Npc(ZoneNPCNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With ZoneNpc(ZoneNum).Npc(ZoneNPCNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong ZoneNum
        Buffer.WriteLong ZoneNPCNum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = MapNum Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_ZONENPC Then
                        If TempPlayer(i).Target = ZoneNPCNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            TempPlayer(i).TargetZone = 0
                            SendTarget i
                        End If
                    End If
                    If TempPlayer(i).PetTargetType = TARGET_TYPE_ZONENPC Then
                        If TempPlayer(i).PetTarget = ZoneNPCNum Then
                            TempPlayer(i).PetTarget = 0
                            TempPlayer(i).PetTargetType = TARGET_TYPE_NONE
                            TempPlayer(i).PetTargetZone = 0
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP) = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32), (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32)
        SendBlood GetPlayerMap(Attacker), ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Attacker, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y, SoundEntity.seSpell, Spellnum

        ' Set the NPC target to the player
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).TargetType = TARGET_TYPE_PET ' player's pet
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Target = Attacker

        ' set the regen timer
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).stopRegen = True
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If Spellnum > 0 Then
            If Spell(Spellnum).StunDuration > 0 Then StunNPC ZoneNPCNum, MapNum, Spellnum, ZoneNum
            ' DoT
            If Spell(Spellnum).Duration > 0 Then
                AddDoT_Npc MapNum, ZoneNPCNum, Spellnum, Attacker, ZoneNum
            End If
        End If
        
        SendZoneNpcVitals ZoneNum, ZoneNPCNum
    End If

    If Spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).PetAttackTimer = GetTickCount
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetAttackZoneNpc", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub PetFireProjectile(ByVal Index As Long, ByVal Spellnum As Long)
Dim ProjectileSlot As Long
Dim ProjectileNum As Long
Dim MapNum As Long
Dim i As Long

    ' Prevent subscript out of range
   On Error GoTo errorhandler
   
    MapNum = GetPlayerMap(Index)
    
    'Find a free projectile
    For i = 1 To MAX_PROJECTILES
        If MapProjectiles(MapNum, i).ProjectileNum = 0 Then ' Free Projectile
            ProjectileSlot = i
            Exit For
        End If
    Next
    
    'Check for no projectile, if so just overwrite the first slot
    If ProjectileSlot = 0 Then ProjectileSlot = 1
    
    If Spellnum < 1 Or Spellnum > MAX_SPELLS Then Exit Sub

    ProjectileNum = Spell(Spellnum).Projectile
    
    With MapProjectiles(MapNum, ProjectileSlot)
        .ProjectileNum = ProjectileNum
        .Owner = Index
        .OwnerType = TARGET_TYPE_PET
        .Dir = Player(i).characters(TempPlayer(i).CurChar).Pet.Dir
        .x = Player(i).characters(TempPlayer(i).CurChar).Pet.x
        .y = Player(i).characters(TempPlayer(i).CurChar).Pet.y
        .Timer = GetTickCount + 60000
    End With
   
   Call SendProjectileToMap(MapNum, ProjectileSlot)
   
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetFireProjectile", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
