Attribute VB_Name = "modPets"
Public MAX_PETS As Long
Public Pet() As PetRec
Public Const EDITOR_PET As Byte = 7
Public Pet_Changed() As Boolean
Public Const ITEM_TYPE_PET As Byte = 10
Public Const TARGET_TYPE_PET = 7

Public Const PetHpBarWidth As Long = 129
Public Const PetMpBarWidth As Long = 129

Public PetSpellBuffer As Long
Public PetSpellBufferTimer As Long
Public PetSpellCD(1 To 4) As Long

'Pet Constants
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
    
    spell(1 To 4) As Long
End Type

Public Type PlayerPetRec
    Num As Long
    Health As Long
    Mana As Long
    Level As Long
    stat(1 To Stats.Stat_Count - 1) As Byte
    spell(1 To 4) As Long
    Points As Long
    X As Long
    Y As Long
    dir As Long
    MaxHp As Long
    MaxMP As Long
    Alive As Boolean
    AttackBehaviour As Long
    Exp As Long
    TNL As Long
    
    'Client Use Only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    Damage As Long
End Type

'ClientTCP
Public Sub SendPetBehaviour(Index As Long)
    Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong csetbehaviour
    buffer.WriteLong Index
    SendData buffer.ToArray
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPetBehaviour", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Public Sub SendRequestEditPet()
Dim buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditPet
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub
Sub SendRequestPets()
Dim buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestPets
    SendData buffer.ToArray()
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestPets", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Public Sub SendSavePet(ByVal petNum As Long)
Dim buffer As clsBuffer
Dim i As Long
    

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CSavePet
    buffer.WriteLong petNum
    With Pet(petNum)
        buffer.WriteLong .Num
        buffer.WriteString .Name
        buffer.WriteLong .Sprite
        buffer.WriteLong .Range
        buffer.WriteLong .Level
        buffer.WriteLong .MaxLevel
        buffer.WriteLong .ExpGain
        buffer.WriteLong .LevelPnts
        buffer.WriteByte .StatType
        buffer.WriteByte .LevelingType
        For i = 1 To Stats.Stat_Count - 1
            buffer.WriteByte .stat(i)
        Next
        For i = 1 To 4
            buffer.WriteLong .spell(i)
        Next
    End With
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSavePet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub



'ModGameEditors
' ////////////////
' // Pet Editor //
' ////////////////
Public Sub PetEditorInit()
Dim i As Long, prefix As String


   On Error GoTo errorhandler

    If frmEditor_Pet.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Pet.lstIndex.ListIndex + 1
    
    With frmEditor_Pet
        .txtName.Text = Trim$(Pet(EditorIndex).Name)
        If Pet(EditorIndex).Sprite < 0 Or Pet(EditorIndex).Sprite > .scrlSprite.max Then Pet(EditorIndex).Sprite = 0
        .scrlSprite.Value = Pet(EditorIndex).Sprite
        .scrlRange.Value = Pet(EditorIndex).Range
        
        
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).Value = Pet(EditorIndex).stat(i)
        Next
        
        If Pet(EditorIndex).StatType = 1 Then
            .optCustomStats.Value = True
            .picCustomStats.Visible = True
        Else
            .optAdoptStats.Value = True
            .picCustomStats.Visible = False
        End If
        
        .scrlStat(6).Value = Pet(EditorIndex).Level
        
        For i = 1 To 4
            .scrlSpell(i) = Pet(EditorIndex).spell(i)
            prefix = "Spell " & Index & ": "
    
            If .scrlSpell(i).Value = 0 Then
                .lblSpell(i).Caption = prefix & "None"
            Else
                .lblSpell(i).Caption = prefix & Trim$(spell(.scrlSpell(i).Value).Name)
            End If
        Next
        
        If Pet(EditorIndex).LevelingType = 0 Then
            .optLevel.Value = True
            .picPetlevel.Visible = True
            .scrlPetExp.Value = Pet(EditorIndex).ExpGain
            If Pet(EditorIndex).MaxLevel > 0 Then .scrlMaxLevel.Value = Pet(EditorIndex).MaxLevel
            .scrlPetPnts.Value = Pet(EditorIndex).LevelPnts
        Else
            .optDoNotLevel.Value = True
            .picPetlevel.Visible = flase
            .scrlPetExp.Value = Pet(EditorIndex).ExpGain
            .scrlMaxLevel.Value = Pet(EditorIndex).MaxLevel
            .scrlPetPnts.Value = Pet(EditorIndex).LevelPnts
        End If
    End With
    
    Pet_Changed(EditorIndex) = True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetEditorInit", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub
Public Sub PetEditorOk()
Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_PETS
        If Pet_Changed(i) Then
            Call SendSavePet(i)
        End If
    Next
    
    Unload frmEditor_Pet
    Editor = 0
    ClearChanged_Pet


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetEditorOk", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Public Sub PetEditorCancel()



   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Pet
    ClearChanged_Pet
    ClearPets
    SendRequestPets


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetEditorCancel", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Public Sub ClearChanged_Pet()


   On Error GoTo errorhandler

    ReDim Pet_Changed(MAX_PETS)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_Pet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub



'ModDatabase
Sub ClearPet(ByVal Index As Long)


   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Pet(Index)), LenB(Pet(Index)))
    Pet(Index).Name = vbNullString


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearPet", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ClearPets()
Dim i As Long



   On Error GoTo errorhandler

    For i = 1 To MAX_PETS
        Call ClearPet(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearPets", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ProcessPetMovement(ByVal Index As Long)

    ' Check if NPC is walking, and if so process moving them over

   On Error GoTo errorhandler

    If Player(Index).Pet.Moving = MOVING_WALKING Then
        
        Select Case Player(Index).Pet.dir
            Case DIR_UP
                Player(Index).Pet.YOffset = Player(Index).Pet.YOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If Player(Index).Pet.YOffset < 0 Then Player(Index).Pet.YOffset = 0
                
            Case DIR_DOWN
                Player(Index).Pet.YOffset = Player(Index).Pet.YOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If Player(Index).Pet.YOffset > 0 Then Player(Index).Pet.YOffset = 0
                
            Case DIR_LEFT
                Player(Index).Pet.XOffset = Player(Index).Pet.XOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If Player(Index).Pet.XOffset < 0 Then Player(Index).Pet.XOffset = 0
                
            Case DIR_RIGHT
                Player(Index).Pet.XOffset = Player(Index).Pet.XOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If Player(Index).Pet.XOffset > 0 Then Player(Index).Pet.XOffset = 0
                
        End Select
    
        ' Check if completed walking over to the next tile
        If Player(Index).Pet.Moving > 0 Then
            If Player(Index).Pet.dir = DIR_RIGHT Or Player(Index).Pet.dir = DIR_DOWN Then
                If (Player(Index).Pet.XOffset >= 0) And (Player(Index).Pet.YOffset >= 0) Then
                    Player(Index).Pet.Moving = 0
                    If Player(Index).Pet.Step = 1 Then
                        Player(Index).Pet.Step = 2
                    Else
                        Player(Index).Pet.Step = 1
                    End If
                End If
            Else
                If (Player(Index).Pet.XOffset <= 0) And (Player(Index).Pet.YOffset <= 0) Then
                    Player(Index).Pet.Moving = 0
                    If Player(Index).Pet.Step = 1 Then
                        Player(Index).Pet.Step = 2
                    Else
                        Player(Index).Pet.Step = 1
                    End If
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ProcessPetMovement", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub


'frmMain
Public Sub PetMove(ByVal X As Long, ByVal Y As Long)
    Dim buffer As clsBuffer

    

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CPetMove
    buffer.WriteLong X
    buffer.WriteLong Y
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PetMove", "modPets", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Sub SendTrainPetStat(ByVal StatNum As Byte)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CPetUseStatPoint
    buffer.WriteByte StatNum
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPetTrainStat", "modPets", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

