Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long

   On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If NewOptions.CombatMode = 1 Then
        Select Case Vital
            Case HP
                GetPlayerMaxVital = ((GetPlayerLevel(Index) * 4) + (GetPlayerStat(Index, Endurance) * 10)) + 150
            Case MP
                GetPlayerMaxVital = ((GetPlayerLevel(Index) * 4) + (GetPlayerStat(Index, Intelligence) * 10)) + 150
        End Select
    Else
        Select Case Vital
            Case HP
                GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 15 + 100
            Case MP
                GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Willpower) / 2)) * 5 + 50
        End Select
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If
    
    If NewOptions.CombatMode = 1 Then
        Select Case Vital
            Case HP
                i = (GetPlayerStat(Index, Stats.Endurance) / 2)
            Case MP
                i = (GetPlayerStat(Index, Stats.Intelligence) / 2)
        End Select
    Else
        Select Case Vital
            Case HP
                i = (GetPlayerStat(Index, Stats.Willpower) * 0.8) + 6
            Case MP
                i = (GetPlayerStat(Index, Stats.Willpower) / 4) + 12.5
        End Select
    End If

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerVitalRegen", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim weaponNum As Long
    

   On Error GoTo errorhandler

    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    If NewOptions.CombatMode = 1 Then
        If GetPlayerEquipment(Index, Weapon) > 0 Then
            weaponNum = GetPlayerEquipment(Index, Weapon)
            GetPlayerDamage = (GetPlayerStat(Index, Strength) * 2) + (Item(weaponNum).Data2 * 2) + (GetPlayerLevel(Index) * 3) + Random(0, 20)
        Else
            GetPlayerDamage = (GetPlayerStat(Index, Strength) * 2) + (GetPlayerLevel(Index) * 3) + Random(0, 20)
        End If
    Else
        If GetPlayerEquipment(Index, Weapon) > 0 Then
            weaponNum = GetPlayerEquipment(Index, Weapon)
            GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) * Item(weaponNum).Data2 + (GetPlayerLevel(Index) / 5)
        Else
            GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) + (GetPlayerLevel(Index) / 5)
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetPlayerDamage", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function GetNpcMaxVital(ByVal npcnum As Long, ByVal Vital As Vitals) As Long
    Dim x As Long

    ' Prevent subscript out of range

   On Error GoTo errorhandler

    If npcnum <= 0 Or npcnum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = Npc(npcnum).HP
        Case MP
            GetNpcMaxVital = 30 + (Npc(npcnum).stat(Intelligence) * 10) + 2
    End Select


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetNpcMaxVital", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function GetNpcVitalRegen(ByVal npcnum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range

   On Error GoTo errorhandler

    If npcnum <= 0 Or npcnum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If
    
    If NewOptions.CombatMode = 1 Then
        Select Case Vital
            Case HP
                i = (Npc(npcnum).stat(Stats.Endurance) / 2)
            Case MP
                i = (Npc(npcnum).stat(Stats.Intelligence) / 2)
        End Select
    Else
        Select Case Vital
            Case HP
                i = (Npc(npcnum).stat(Stats.Willpower) * 0.8) + 6
            Case MP
                i = (Npc(npcnum).stat(Stats.Willpower) / 4) + 12.5
        End Select
    End If
    
    GetNpcVitalRegen = i


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetNpcVitalRegen", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Function GetNpcDamage(ByVal npcnum As Long) As Long

   On Error GoTo errorhandler

    If NewOptions.CombatMode = 1 Then
        GetNpcDamage = (Npc(npcnum).stat(Stats.Strength) * 2) + (Npc(npcnum).Damage * 2) + (Npc(npcnum).Level * 3) + rand(1, 20)
    Else
        GetNpcDamage = 0.085 * 5 * Npc(npcnum).stat(Stats.Strength) * Npc(npcnum).Damage + (Npc(npcnum).Level / 5)
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetNpcDamage", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long


   On Error GoTo errorhandler

    CanPlayerBlock = False
    If NewOptions.CombatMode = 1 Then
        If GetPlayerEquipment(Index, Shield) > 0 Then
            rate = GetPlayerStat(Index, Strength) / 3
            If rate >= rand(1, 100) Then CanPlayerBlock = True
        End If
    Else
        rate = 0
        ' TODO : make it based on shield lulz
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPlayerBlock", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function CanPlayerCrit(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long


   On Error GoTo errorhandler

    CanPlayerCrit = False
    
    If NewOptions.CombatMode = 1 Then
        rate = GetPlayerStat(Index, Agility) / 3
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPlayerCrit = True
        End If
    Else
        rate = GetPlayerStat(Index, Agility) / 52.08
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPlayerCrit = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPlayerCrit", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function CanPlayerDodge(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long


   On Error GoTo errorhandler

    CanPlayerDodge = False
    
    If NewOptions.CombatMode = 1 Then
        rate = GetPlayerStat(Index, Agility) / 4
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPlayerDodge = True
        End If
    Else
        rate = GetPlayerStat(Index, Agility) / 83.3
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPlayerDodge = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPlayerDodge", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function CanPlayerParry(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long


   On Error GoTo errorhandler

    CanPlayerParry = False
    
    If NewOptions.CombatMode = 1 Then
        rate = GetPlayerStat(Index, Agility) / 6
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPlayerParry = True
        End If
    Else
        rate = GetPlayerStat(Index, Strength) * 0.25
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanPlayerParry = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPlayerParry", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function CanNpcBlock(ByVal npcnum As Long) As Boolean
Dim rate As Long
Dim stat As Long
Dim rndNum As Long


   On Error GoTo errorhandler

    CanNpcBlock = False
    If NewOptions.CombatMode = 1 Then
        'No Shield, No Block
    Else
        stat = Npc(npcnum).stat(Stats.Agility) / 5  'guessed shield agility
        rate = stat / 12.08
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanNpcBlock = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanNpcBlock", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
    
End Function

Public Function CanNpcCrit(ByVal npcnum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long


   On Error GoTo errorhandler

    CanNpcCrit = False
    If NewOptions.CombatMode = 1 Then
        rate = Npc(npcnum).stat(Stats.Agility) / 3
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanNpcCrit = True
        End If
    Else
        rate = Npc(npcnum).stat(Stats.Agility) / 52.08
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanNpcCrit = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanNpcCrit", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function CanNpcDodge(ByVal npcnum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long


   On Error GoTo errorhandler

    CanNpcDodge = False
    If NewOptions.CombatMode = 1 Then
        rate = Npc(npcnum).stat(Stats.Agility) / 4
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanNpcDodge = True
        End If
    Else
        rate = Npc(npcnum).stat(Stats.Agility) / 83.3
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanNpcDodge = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanNpcDodge", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function CanNpcParry(ByVal npcnum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long


   On Error GoTo errorhandler

    CanNpcParry = False
    If NewOptions.CombatMode = 1 Then
        rate = Npc(npcnum).stat(Stats.Agility) / 6
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanNpcParry = True
        End If
    Else
        rate = Npc(npcnum).stat(Stats.Strength) * 0.25
        rndNum = rand(1, 100)
        If rndNum <= rate Then
            CanNpcParry = True
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanNpcParry", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal mapnpcnum As Long)
Dim blockAmount As Long
Dim npcnum As Long
Dim MapNum As Long
Dim Damage As Long, i As Long, armor As Long


   On Error GoTo errorhandler

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, mapnpcnum) Then
    
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
        Damage = GetPlayerDamage(Index)
        If NewOptions.CombatMode = 1 Then
            If CanNpcBlock(npcnum) Then
                SendActionMsg MapNum, "Block!", BrightCyan, 1, (MapNpc(MapNum).Npc(mapnpcnum).x * 32), (MapNpc(MapNum).Npc(mapnpcnum).y * 32)
                Damage = 0
                TempPlayer(Index).Target = mapnpcnum
                TempPlayer(Index).TargetType = TARGET_TYPE_NPC
                TempPlayer(Index).TargetZone = 0
                SendTarget Index
                Exit Sub
            Else
                Damage = Damage - ((Npc(npcnum).stat(Stats.Willpower) * 2) + (Npc(npcnum).Level * 3))
                ' * 1.5 if it's a crit!
                If CanPlayerCrit(Index) Then
                    Damage = Damage * 1.5
                    SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
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
            If CanPlayerCrit(Index) Then
                Damage = Damage * 1.5
                SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            End If
        End If
        
        TempPlayer(Index).Target = mapnpcnum
        TempPlayer(Index).TargetType = TARGET_TYPE_NPC
        TempPlayer(Index).TargetZone = 0
        SendTarget Index
            
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, mapnpcnum, Damage)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryPlayerAttackNpc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub TryPlayerAttackZoneNpc(ByVal Index As Long, ZoneNum As Long, ZoneNPCNum As Long)
Dim blockAmount As Long
Dim npcnum As Long
Dim MapNum As Long
Dim Damage As Long


   On Error GoTo errorhandler

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackZoneNpc(Index, ZoneNum, ZoneNPCNum) Then
    
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
        Damage = GetPlayerDamage(Index)
        
        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        If NewOptions.CombatMode = 1 Then
            If CanNpcBlock(npcnum) Then
                SendActionMsg MapNum, "Block!", BrightCyan, 1, (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32), (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32)
                Damage = 0
                TempPlayer(Index).Target = ZoneNPCNum
                TempPlayer(Index).TargetType = TARGET_TYPE_NPC
                TempPlayer(Index).TargetZone = ZoneNum
                SendTarget Index
                Exit Sub
            Else
                Damage = Damage - ((Npc(npcnum).stat(Stats.Willpower) * 2) + (Npc(npcnum).Level * 3))
                ' * 1.5 if it's a crit!
                If CanPlayerCrit(Index) Then
                    Damage = Damage * 1.5
                    SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
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
            If CanPlayerCrit(Index) Then
                Damage = Damage * 1.5
                SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            End If
        End If
        
        
        TempPlayer(Index).Target = ZoneNPCNum
        TempPlayer(Index).TargetType = TARGET_TYPE_NPC
        TempPlayer(Index).TargetZone = ZoneNum
            
        If Damage > 0 Then
            Call PlayerAttackZoneNpc(Index, ZoneNum, ZoneNPCNum, Damage)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryPlayerAttackZoneNpc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal mapnpcnum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim npcnum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Then
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
        If Npc(MapNpc(MapNum).Npc(mapnpcnum).Num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If Npc(MapNpc(MapNum).Npc(mapnpcnum).Num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
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
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).speed
        Else
            attackspeed = 1000
        End If

        If npcnum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
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

            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then
                    If Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        CanPlayerAttackNpc = True
                    ElseIf Npc(npcnum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or Npc(npcnum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                        If Len(Trim$(Npc(npcnum).AttackSay)) > 0 Then
                            PlayerMsg Attacker, Trim$(Npc(npcnum).Name) & ": " & Trim$(Npc(npcnum).AttackSay), White
                        End If
                        ' Reset attack timer
                        TempPlayer(Attacker).AttackTimer = GetTickCount
                    End If
                End If
            End If
        End If
    End If
    


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPlayerAttackNpc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Function CanPlayerAttackZoneNpc(ByVal Attacker As Long, ZoneNum As Long, ZoneNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
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
                    CanPlayerAttackZoneNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).speed
        Else
            attackspeed = 1000
        End If

        If npcnum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    NpcX = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x
                    NpcY = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y + 1
                Case DIR_DOWN
                    NpcX = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x
                    NpcY = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y - 1
                Case DIR_LEFT
                    NpcX = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x + 1
                    NpcY = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y
                Case DIR_RIGHT
                    NpcX = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x - 1
                    NpcY = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y
            End Select

            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then
                    If Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcnum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        CanPlayerAttackZoneNpc = True
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
    HandleError "CanPlayerAttackZoneNpc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal mapnpcnum As Long, ByVal Damage As Long, Optional ByVal Spellnum As Long, Optional ByVal overTime As Boolean = False)
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

    If IsPlaying(Attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Attacker)
    npcnum = MapNpc(MapNum).Npc(mapnpcnum).Num
    Name = Trim$(Npc(npcnum).Name)
    
    If NewOptions.CombatMode = 1 Then
        If Spellnum > 0 Then
            'Magic Resist
            Damage = Damage - ((Npc(npcnum).stat(Willpower) * 2) + (Npc(npcnum).Level * 3))
            If Damage <= 0 Then Exit Sub
        End If
    End If
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    If n > 0 Then
    
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(MapNum).Npc(mapnpcnum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(MapNum).Npc(mapnpcnum).x * 32), (MapNpc(MapNum).Npc(mapnpcnum).y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Attacker, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y, SoundEntity.seSpell, Spellnum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If Spellnum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y)
                If Spellnum = 0 Then SendMapSound Attacker, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y, SoundEntity.seItem, GetPlayerEquipment(Attacker, Weapon)
            End If
        End If

        ' Calculate exp to give attacker
        Exp = Npc(npcnum).Exp

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, GetPlayerMap(Attacker)
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
        MapNpc(MapNum).Npc(mapnpcnum).Target = 0
        MapNpc(MapNum).Npc(mapnpcnum).TargetType = 0
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
        
        Call CheckTasks(Attacker, TASK_KILLNPCS, npcnum)
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
                If Player(i).characters(TempPlayer(i).CurChar).Map = MapNum Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = mapnpcnum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            TempPlayer(i).TargetZone = 0
                            SendTarget i
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
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If Spellnum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapnpcnum)
                If Spellnum = 0 Then SendMapSound Attacker, MapNpc(MapNum).Npc(mapnpcnum).x, MapNpc(MapNum).Npc(mapnpcnum).y, SoundEntity.seItem, GetPlayerEquipment(Attacker, Weapon)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(mapnpcnum).TargetType = 1 ' player
        MapNpc(MapNum).Npc(mapnpcnum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(mapnpcnum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(mapnpcnum).Num Then
                    MapNpc(MapNum).Npc(i).Target = Attacker
                    MapNpc(MapNum).Npc(i).TargetType = 1 ' player
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
                AddDoT_Npc MapNum, mapnpcnum, Spellnum, Attacker
            End If
        End If
        
        SendMapNpcVitals MapNum, mapnpcnum
    End If

    If Spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = GetTickCount
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerAttackNpc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub PlayerAttackZoneNpc(ByVal Attacker As Long, ZoneNum As Long, ZoneNPCNum As Long, ByVal Damage As Long, Optional ByVal Spellnum As Long, Optional ByVal overTime As Boolean = False)
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

    If IsPlaying(Attacker) = False Or ZoneNPCNum <= 0 Or ZoneNPCNum > MAX_MAP_NPCS * 2 Or Damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Attacker)
    npcnum = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
    Name = Trim$(Npc(npcnum).Name)
    
    If NewOptions.CombatMode = 1 Then
        If Spellnum > 0 Then
            'Magic Resist
            Damage = Damage - ((Npc(npcnum).stat(Willpower) * 2) + (Npc(npcnum).Level * 3))
            If Damage <= 0 Then Exit Sub
        End If
    End If
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(Attacker), "-" & ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP), BrightRed, 1, (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32), (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32)
        SendBlood GetPlayerMap(Attacker), ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Attacker, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y, SoundEntity.seSpell, Spellnum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If Spellnum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y)
                If Spellnum = 0 Then SendMapSound Attacker, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y, SoundEntity.seItem, GetPlayerEquipment(Attacker, Weapon)
            End If
        End If

        ' Calculate exp to give attacker
        Exp = Npc(npcnum).Exp

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, GetPlayerMap(Attacker)
        Else
            ' no party - keep exp for self
            GivePlayerEXP Attacker, Exp
        End If
        
        For n = 1 To 20
            If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(n).Num > 0 Then
                SpawnItem ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(n).Num, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(n).Value, MapNum, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y
                ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(n).Value = 0
                ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Inventory(n).Num = 0
            End If
        Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num = 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).SpawnWait = GetTickCount
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP) = 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Target = 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).TargetType = 0
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
        
        Call CheckTasks(Attacker, TASK_KILLNPCS, npcnum)
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
                If Player(i).characters(TempPlayer(i).CurChar).Map = MapNum Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_ZONENPC Then
                        If TempPlayer(i).Target = ZoneNPCNum Then
                            If TempPlayer(i).TargetZone = ZoneNum Then
                                TempPlayer(i).Target = 0
                                TempPlayer(i).TargetType = TARGET_TYPE_NONE
                                TempPlayer(i).TargetZone = 0
                                SendTarget i
                            End If
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
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If Spellnum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_ZONENPC, ZoneNPCNum, 0, ZoneNum)
                If Spellnum = 0 Then SendMapSound Attacker, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y, SoundEntity.seItem, GetPlayerEquipment(Attacker, Weapon)
            End If
        End If

        ' Set the NPC target to the player
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).TargetType = 1 ' player
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num Then
                    MapNpc(MapNum).Npc(i).Target = Attacker
                    MapNpc(MapNum).Npc(i).TargetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).stopRegen = True
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If Spellnum > 0 Then
            If Spell(Spellnum).StunDuration > 0 Then StunNPC ZoneNPCNum, 0, Spellnum, ZoneNum
            ' DoT
            If Spell(Spellnum).Duration > 0 Then
                AddDoT_Npc MapNum, ZoneNPCNum, Spellnum, Attacker, ZoneNum
            End If
        End If
        
        SendZoneNpcVitals ZoneNum, ZoneNPCNum
    End If

    If Spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = GetTickCount
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerAttackZoneNpc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal mapnpcnum As Long, ByVal Index As Long, Optional ByVal IsSpell As Boolean = False)
Dim MapNum As Long, npcnum As Long, blockAmount As Long, Damage As Long, i As Long, armor As Long

    ' Can the npc attack the player?

   On Error GoTo errorhandler

    If CanNpcAttackPlayer(mapnpcnum, Index, IsSpell) Then
        MapNum = GetPlayerMap(Index)
        npcnum = MapNpc(MapNum).Npc(mapnpcnum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).characters(TempPlayer(Index).CurChar).x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(Index) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(Index).characters(TempPlayer(Index).CurChar).x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcnum)
        
        If NewOptions.CombatMode = 1 Then
            If CanPlayerBlock(Index) Then
                SendActionMsg MapNum, "Block!", Pink, 1, (Player(Index).characters(TempPlayer(Index).CurChar).x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).y * 32)
                Exit Sub
            Else
                For i = 1 To Equipment.Equipment_Count - 1
                    If GetPlayerEquipment(Index, i) > 0 Then
                        armor = armor + Item(GetPlayerEquipment(Index, i)).Data2
                    End If
                Next
                ' take away armour
                Damage = Damage - ((GetPlayerStat(Index, Willpower) * 2) + (GetPlayerLevel(Index) * 3) + armor)
                ' * 1.5 if crit hit
                If CanNpcCrit(npcnum) Then
                    Damage = Damage * 1.5
                    SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(mapnpcnum).x * 32), (MapNpc(MapNum).Npc(mapnpcnum).y * 32)
                End If
            End If
        Else
            ' if the player blocks, take away the block amount
            blockAmount = CanPlayerBlock(Index)
            Damage = Damage - blockAmount
            
            ' take away armour
            Damage = Damage - rand(1, (GetPlayerStat(Index, Agility) * 2))
            
            ' randomise for up to 10% lower than max hit
            Damage = rand(1, Damage)
            
            ' * 1.5 if crit hit
            If CanNpcCrit(npcnum) Then
                Damage = Damage * 1.5
                SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(mapnpcnum).x * 32), (MapNpc(MapNum).Npc(mapnpcnum).y * 32)
            End If
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(mapnpcnum, Index, Damage)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryNpcAttackPlayer", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub TryZoneNpcAttackPlayer(ZoneNum As Long, ZoneNPCNum As Long, ByVal Index As Long)
Dim MapNum As Long, npcnum As Long, blockAmount As Long, Damage As Long, i As Long, armor As Long

    ' Can the npc attack the player?

   On Error GoTo errorhandler

    If CanZoneNpcAttackPlayer(ZoneNum, ZoneNPCNum, Index) Then
        MapNum = GetPlayerMap(Index)
        npcnum = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).characters(TempPlayer(Index).CurChar).x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(Index) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(Index).characters(TempPlayer(Index).CurChar).x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcnum)
        
        If NewOptions.CombatMode = 1 Then
            If CanPlayerBlock(Index) Then
                SendActionMsg MapNum, "Block!", Pink, 1, (Player(Index).characters(TempPlayer(Index).CurChar).x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).y * 32)
                Exit Sub
            Else
                For i = 1 To Equipment.Equipment_Count - 1
                    If GetPlayerEquipment(Index, i) > 0 Then
                        armor = armor + Item(GetPlayerEquipment(Index, i)).Data2
                    End If
                Next
                ' take away armour
                Damage = Damage - ((GetPlayerStat(Index, Willpower) * 2) + (GetPlayerLevel(Index) * 3) + armor)
                ' * 1.5 if crit hit
                If CanNpcCrit(Index) Then
                    Damage = Damage * 1.5
                    SendActionMsg MapNum, "Critical!", BrightCyan, 1, (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32), (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32)
                End If
            End If
        Else
            ' if the player blocks, take away the block amount
            blockAmount = CanPlayerBlock(Index)
            Damage = Damage - blockAmount
            
            ' take away armour
            Damage = Damage - rand(1, (GetPlayerStat(Index, Agility) * 2))
            
            ' randomise for up to 10% lower than max hit
            Damage = rand(1, Damage)
            
            ' * 1.5 if crit hit
            If CanNpcCrit(Index) Then
                Damage = Damage * 1.5
                SendActionMsg MapNum, "Critical!", BrightCyan, 1, (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32), (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32)
            End If
        End If
        


        If Damage > 0 Then
            Call ZoneNpcAttackPlayer(ZoneNum, ZoneNPCNum, Index, Damage)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryZoneNpcAttackPlayer", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function CanZoneNpcAttackPlayer(ZoneNum As Long, ZoneNPCNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long
    Dim npcnum As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If ZoneNPCNum <= 0 Or ZoneNPCNum > MAX_MAP_NPCS * 2 Or Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Index)
    npcnum = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num

    ' Make sure the npc isn't already dead
    If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < ZoneNpc(ZoneNum).Npc(ZoneNPCNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    ZoneNpc(ZoneNum).Npc(ZoneNPCNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If npcnum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y) And (GetPlayerX(Index) = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x) Then
                CanZoneNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y) And (GetPlayerX(Index) = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x) Then
                    CanZoneNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y) And (GetPlayerX(Index) + 1 = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x) Then
                        CanZoneNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y) And (GetPlayerX(Index) - 1 = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x) Then
                            CanZoneNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanZoneNpcAttackPlayer", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function CanNpcAttackPlayer(ByVal mapnpcnum As Long, ByVal Index As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim npcnum As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then
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
    If IsPlaying(Index) Then
        If npcnum > 0 Then
            If IsSpell = False Then

                ' Check if at same coordinates
                If (GetPlayerY(Index) + 1 = MapNpc(MapNum).Npc(mapnpcnum).y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(mapnpcnum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) - 1 = MapNpc(MapNum).Npc(mapnpcnum).y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(mapnpcnum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum).Npc(mapnpcnum).y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum).Npc(mapnpcnum).x) Then
                            CanNpcAttackPlayer = True
                        Else
                            If (GetPlayerY(Index) = MapNpc(MapNum).Npc(mapnpcnum).y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum).Npc(mapnpcnum).x) Then
                                CanNpcAttackPlayer = True
                            End If
                        End If
                    End If
                End If
            
            Else
                CanNpcAttackPlayer = True
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanNpcAttackPlayer", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub NpcAttackPlayer(ByVal mapnpcnum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim MapNum As Long
    Dim i As Long, z As Long, InvCount As Long, EqCount As Long, j As Long, x As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
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
    Buffer.WriteLong 0
    Buffer.WriteLong mapnpcnum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(MapNum).Npc(mapnpcnum).stopRegen = True
    MapNpc(MapNum).Npc(mapnpcnum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(mapnpcnum).Num
        
        ' Set NPC target to 0
        MapNpc(MapNum).Npc(mapnpcnum).Target = 0
        MapNpc(MapNum).Npc(mapnpcnum).TargetType = 0
        
        If NewOptions.ItemLoss = 0 Then
            If GetPlayerLevel(Victim) >= 10 Then
                For z = 1 To MAX_INV
                    If GetPlayerInvItemNum(Victim, z) > 0 Then
                        InvCount = InvCount + 1
                    End If
                Next
                For z = 1 To Equipment.Equipment_Count - 1
                    If GetPlayerEquipment(Victim, z) > 0 Then
                        EqCount = EqCount + 1
                    End If
                Next
                z = Random(1, InvCount + EqCount)
                If z = 0 Then z = 1
                If z > InvCount + EqCount Then z = InvCount + EqCount
                If z > InvCount Then
                    z = z - InvCount
                    For x = 1 To Equipment.Equipment_Count - 1
                        If GetPlayerEquipment(Victim, x) > 0 Then
                            j = j + 1
                            If j = z Then
                                'Here it is, drop this piece of equipment!
                                PlayerMsg Victim, "In death you lost grip on your " & Trim$(Item(GetPlayerEquipment(Victim, x)).Name), BrightRed
                                SpawnItem GetPlayerEquipment(Victim, x), 1, GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
                                SetPlayerEquipment Victim, 0, x
                                SendWornEquipment Victim
                                SendMapEquipment Victim
                            End If
                        End If
                    Next
                Else
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(Victim, x) > 0 Then
                            j = j + 1
                            If j = z Then
                                'Here it is, drop this item!
                                PlayerMsg Victim, "In death you lost grip on your " & Trim$(Item(GetPlayerInvItemNum(Victim, x)).Name), BrightRed
                                SpawnItem GetPlayerInvItemNum(Victim, x), GetPlayerInvItemValue(Victim, x), GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
                                SetPlayerInvItemNum Victim, x, 0
                                SetPlayerInvItemValue Victim, x, 0
                                SendInventory Victim
                            End If
                        End If
                    Next
                End If
            End If
        End If
        
        ' kill player
        KillPlayer Victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & Name, BrightRed)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        Call SendAnimation(MapNum, Npc(MapNpc(GetPlayerMap(Victim)).Npc(mapnpcnum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(mapnpcnum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "NpcAttackPlayer", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ZoneNpcAttackPlayer(ZoneNum As Long, ZoneNPCNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim MapNum As Long
    Dim i As Long, z As Long, InvCount As Long, EqCount As Long, j As Long, x As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If ZoneNPCNum <= 0 Or ZoneNPCNum > MAX_MAP_NPCS * 2 Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Victim)
    Name = Trim$(Npc(ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong ZoneNum
    Buffer.WriteLong ZoneNPCNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    ZoneNpc(ZoneNum).Npc(ZoneNPCNum).stopRegen = True
    ZoneNpc(ZoneNum).Npc(ZoneNPCNum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
        
        ' Set NPC target to 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Target = 0
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).TargetType = 0
        
        If NewOptions.ItemLoss = 0 Then
            If GetPlayerLevel(Victim) >= 10 Then
                For z = 1 To MAX_INV
                    If GetPlayerInvItemNum(Victim, z) > 0 Then
                        InvCount = InvCount + 1
                    End If
                Next
                For z = 1 To Equipment.Equipment_Count - 1
                    If GetPlayerEquipment(Victim, z) > 0 Then
                        EqCount = EqCount + 1
                    End If
                Next
                z = Random(1, InvCount + EqCount)
                If z = 0 Then z = 1
                If z > InvCount + EqCount Then z = InvCount + EqCount
                If z > InvCount Then
                    z = z - InvCount
                    For x = 1 To Equipment.Equipment_Count - 1
                        If GetPlayerEquipment(Victim, x) > 0 Then
                            j = j + 1
                            If j = z Then
                                'Here it is, drop this piece of equipment!
                                PlayerMsg Victim, "In death you lost grip on your " & Trim$(Item(GetPlayerEquipment(Victim, x)).Name), BrightRed
                                SpawnItem GetPlayerEquipment(Victim, x), 1, GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
                                SetPlayerEquipment Victim, 0, x
                                SendWornEquipment Victim
                                SendMapEquipment Victim
                                x = Equipment.Equipment_Count - 1
                            End If
                        End If
                    Next
                Else
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(Victim, x) > 0 Then
                            j = j + 1
                            If j = z Then
                                'Here it is, drop this item!
                                PlayerMsg Victim, "In death you lost grip on your " & Trim$(Item(GetPlayerInvItemNum(Victim, x)).Name), BrightRed
                                SpawnItem GetPlayerInvItemNum(Victim, x), GetPlayerInvItemValue(Victim, x), GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
                                SetPlayerInvItemNum Victim, x, 0
                                SetPlayerInvItemValue Victim, x, 0
                                SendInventory Victim
                                x = MAX_INV
                            End If
                        End If
                    Next
                End If
            End If
        End If
        
        ' kill player
        KillPlayer Victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & Name, BrightRed)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        Call SendAnimation(MapNum, Npc(ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ZoneNpcAttackPlayer", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long)
Dim blockAmount As Long
Dim npcnum As Long
Dim MapNum As Long
Dim Damage As Long, i As Long, armor As Long


   On Error GoTo errorhandler

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, Victim) Then
    
        MapNum = GetPlayerMap(Attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(Victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(Victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        If NewOptions.CombatMode = 1 Then
            If CanPlayerBlock(Victim) Then
                SendActionMsg MapNum, "Block!", BrightCyan, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
                Damage = 0
                Exit Sub
            Else
                For i = 1 To Equipment.Equipment_Count - 1
                    If GetPlayerEquipment(Victim, i) > 0 Then
                        armor = armor + Item(GetPlayerEquipment(Victim, i)).Data2
                    End If
                Next
                ' take away armour
                Damage = Damage - ((GetPlayerStat(Victim, Willpower) * 2) + (GetPlayerLevel(Victim) * 3) + armor)
                ' * 1.5 if it's a crit!
                If CanPlayerCrit(Attacker) Then
                    Damage = Damage * 1.5
                    SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
                End If
            End If
        Else
            ' if the npc blocks, take away the block amount
            blockAmount = CanPlayerBlock(Victim)
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
        End If
    

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, Victim, Damage)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryPlayerAttackPlayer", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean


   On Error GoTo errorhandler

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
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

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPlayerAttackPlayer = True


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanPlayerAttackPlayer", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal Spellnum As Long = 0)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long, z As Long, x As Long, j As Long, InvCount As Long, EqCount As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    If NewOptions.CombatMode = 1 Then
        If Spellnum > 0 Then
            'Magic Resist
            Damage = Damage - ((GetPlayerStat(Victim, Willpower) * 2) + (GetPlayerLevel(Victim) * 3))
            If Damage <= 0 Then Exit Sub
        End If
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, Spellnum
        ' send animation
        If n > 0 Then
            If Spellnum = 0 Then Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
            If Spellnum = 0 Then Call SendMapSound(Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seItem, GetPlayerEquipment(Attacker, Weapon))
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        
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
                    Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, GetPlayerMap(Attacker)
                Else
                    ' not in party, get exp for self
                    GivePlayerEXP Attacker, Exp
                End If
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).characters(TempPlayer(i).CurChar).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            TempPlayer(i).TargetZone = 0
                            SendTarget i
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
        
        If NewOptions.ItemLoss = 0 Then
            If Abs(GetPlayerLevel(Attacker) - GetPlayerLevel(Victim)) > 5 Then
                
            Else
                If GetPlayerLevel(Victim) >= 10 Then
                    For z = 1 To MAX_INV
                        If GetPlayerInvItemNum(Victim, z) > 0 Then
                            InvCount = InvCount + 1
                        End If
                    Next
                    For z = 1 To Equipment.Equipment_Count - 1
                        If GetPlayerEquipment(Victim, z) > 0 Then
                            EqCount = EqCount + 1
                        End If
                    Next
                    z = Random(1, InvCount + EqCount)
                    If z = 0 Then z = 1
                    If z > InvCount + EqCount Then z = InvCount + EqCount
                    If z > InvCount Then
                        z = z - InvCount
                        For x = 1 To Equipment.Equipment_Count - 1
                            If GetPlayerEquipment(Victim, x) > 0 Then
                                j = j + 1
                                If j = z Then
                                    'Here it is, drop this piece of equipment!
                                    PlayerMsg Victim, "In death you lost grip on your " & Trim$(Item(GetPlayerEquipment(Victim, x)).Name), BrightRed
                                    SpawnItem GetPlayerEquipment(Victim, x), 1, GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
                                    SetPlayerEquipment Victim, 0, x
                                    SendWornEquipment Victim
                                    SendMapEquipment Victim
                                End If
                            End If
                        Next
                    Else
                        For x = 1 To MAX_INV
                            If GetPlayerInvItemNum(Victim, x) > 0 Then
                                j = j + 1
                                If j = z Then
                                    'Here it is, drop this item!
                                    PlayerMsg Victim, "In death you lost grip on your " & Trim$(Item(GetPlayerInvItemNum(Victim, x)).Name), BrightRed
                                    SpawnItem GetPlayerInvItemNum(Victim, x), GetPlayerInvItemValue(Victim, x), GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
                                    SetPlayerInvItemNum Victim, x, 0
                                    SetPlayerInvItemValue Victim, x, 0
                                    SendInventory Victim
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
        
        Call CheckTasks(Attacker, TASK_KILLPLAYERS, 0)
        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        If Spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, Spellnum
        
        ' send animation
        If n > 0 Then
            If Spellnum = 0 Then Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
            If Spellnum = 0 Then Call SendMapSound(Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seItem, GetPlayerEquipment(Attacker, Weapon))
        End If
        
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
    TempPlayer(Attacker).AttackTimer = GetTickCount


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerAttackPlayer", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub PlayerFireProjectile(ByVal Index As Long, Optional ByVal IsSpell As Long = 0)
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
    
    'Check for spell, if so then load data acordingly
    If IsSpell > 0 Then
        ProjectileNum = Spell(IsSpell).Projectile
    Else
        ProjectileNum = Item(GetPlayerEquipment(Index, Weapon)).Data1
    End If
    
    With MapProjectiles(MapNum, ProjectileSlot)
        .ProjectileNum = ProjectileNum
        .Owner = Index
        .OwnerType = TARGET_TYPE_PLAYER
        .Dir = GetPlayerDir(Index)
        .x = GetPlayerX(Index)
        .y = GetPlayerY(Index)
        .Timer = GetTickCount + 60000
    End With
   
   Call SendProjectileToMap(MapNum, ProjectileSlot)
   
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerFireProjectile", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub NPCFireProjectile(ByVal MapNum As Long, ByVal Index As Long, ByVal Projectile As Long)
Dim ProjectileSlot As Long
Dim ProjectileNum As Long
Dim i As Long

    ' Prevent subscript out of range
   On Error GoTo errorhandler
    
    'Find a free projectile
    For i = 1 To MAX_PROJECTILES
        If MapProjectiles(MapNum, i).ProjectileNum = 0 Then ' Free Projectile
            ProjectileSlot = i
            Exit For
        End If
    Next
    
    'Check for no projectile, if so just overwrite the first slot
    If ProjectileSlot = 0 Then ProjectileSlot = 1

    With MapProjectiles(MapNum, ProjectileSlot)
        .ProjectileNum = Projectile
        .Owner = Index
        .OwnerType = TARGET_TYPE_NPC
        .Dir = MapNpc(MapNum).Npc(Index).Dir
        .x = MapNpc(MapNum).Npc(Index).x
        .y = MapNpc(MapNum).Npc(Index).y
        .Timer = GetTickCount + 60000
    End With
   
   Call SendProjectileToMap(MapNum, ProjectileSlot)
   
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "NPCFireProjectile", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal Index As Long, ByVal spellslot As Long)
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

    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    Spellnum = GetPlayerSpell(Index, spellslot)
    MapNum = GetPlayerMap(Index)
    
    If Spellnum <= 0 Or Spellnum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(Index, Spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(Spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(Spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(Spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(Spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
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
    
    TargetType = TempPlayer(Index).TargetType
    Target = TempPlayer(Index).Target
    TargetZone = TempPlayer(Index).TargetZone
    Range = Spell(Spellnum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1, SPELL_TYPE_PET ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                PlayerMsg Index, "You do not have a target.", BrightRed
            End If
            If TargetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(Spellnum).type <> SPELL_TYPE_DAMAGEHP And Spell(Spellnum).type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf TargetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(MapNum).Npc(Target).x, MapNpc(MapNum).Npc(Target).y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(Spellnum).type <> SPELL_TYPE_DAMAGEHP And Spell(Spellnum).type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), ZoneNpc(TargetZone).Npc(Target).x, ZoneNpc(TargetZone).Npc(Target).y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(Spellnum).type <> SPELL_TYPE_DAMAGEHP And Spell(Spellnum).type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackZoneNpc(Index, TargetZone, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf TargetType = TARGET_TYPE_PET Then
                If Player(Target).characters(TempPlayer(Target).CurChar).Pet.Alive Then
                    ' if have target, check in range
                    If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), Player(Target).characters(TempPlayer(Target).CurChar).Pet.x, Player(Target).characters(TempPlayer(Target).CurChar).Pet.y) Then
                        PlayerMsg Index, "Target not in range.", BrightRed
                    Else
                        ' go through spell types
                        If Spell(Spellnum).type <> SPELL_TYPE_DAMAGEHP And Spell(Spellnum).type <> SPELL_TYPE_DAMAGEMP Then
                            HasBuffered = True
                        Else
                            If CanPlayerAttackPet(Index, Target, True) Then
                                HasBuffered = True
                            End If
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(Spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg MapNum, "Casting " & Trim$(Spell(Spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        TempPlayer(Index).spellBuffer.Spell = spellslot
        TempPlayer(Index).spellBuffer.Timer = GetTickCount
        TempPlayer(Index).spellBuffer.Target = TempPlayer(Index).Target
        TempPlayer(Index).spellBuffer.tType = TempPlayer(Index).TargetType
        TempPlayer(Index).spellBuffer.TargetZone = TempPlayer(Index).TargetZone
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "BufferSpell", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal TargetType As Byte, ByVal TargetZone As Long)
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
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    Spellnum = GetPlayerSpell(Index, spellslot)
    MapNum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, Spellnum) Then Exit Sub

    MPCost = Spell(Spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
   
    LevelReq = Spell(Spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
   
    AccessReq = Spell(Spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
   
    ClassReq = Spell(Spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
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
    If NewOptions.CombatMode = 1 Then
        Vital = (GetPlayerStat(Index, Strength) * 2) + (Vital * 2) + (GetPlayerLevel(Index) * 3) + rand(1, 20)
    End If
    AoE = Spell(Spellnum).AoE
    Range = Spell(Spellnum).Range
   
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(Spellnum).type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, Index, Vital, Spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, Index, Vital, Spellnum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    PlayerWarp Index, Spell(Spellnum).Map, Spell(Spellnum).x, Spell(Spellnum).y
                    SendAnimation GetPlayerMap(Index), Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    DidCast = True
                Case SPELL_TYPE_PET
                    SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    Call SummonPet(Index, Spell(Spellnum).Pet)
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(Index)
                y = GetPlayerY(Index)
            ElseIf SpellCastType = 3 Then
                If TargetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If TargetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                ElseIf TargetType = TARGET_TYPE_NPC Then
                    x = MapNpc(MapNum).Npc(Target).x
                    y = MapNpc(MapNum).Npc(Target).y
                ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                    If GetPlayerMap(Index) = ZoneNpc(TargetZone).Npc(Target).Map Then
                        x = ZoneNpc(TargetZone).Npc(Target).x
                        y = ZoneNpc(TargetZone).Npc(Target).y
                    Else
                        Exit Sub
                    End If
                ElseIf TargetType = TARGET_TYPE_PET Then
                    x = Player(Target).characters(TempPlayer(Target).CurChar).Pet.x
                    y = Player(Target).characters(TempPlayer(Target).CurChar).Pet.y
                End If
               
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    SendClearSpellBuffer Index
                End If
            End If
            Select Case Spell(Spellnum).type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> Index Then
                                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                    If Player(i).characters(TempPlayer(i).CurChar).InHouse = Player(Index).characters(TempPlayer(Index).CurChar).InHouse Then
                                        If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                            If CanPlayerAttackPlayer(Index, i, True) Then
                                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                                PlayerAttackPlayer Index, i, Vital, Spellnum
                                            End If
                                        End If
                                        If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive Then
                                            If isInRange(AoE, x, y, Player(i).characters(TempPlayer(i).CurChar).Pet.x, Player(i).characters(TempPlayer(i).CurChar).Pet.y) Then
                                                If CanPlayerAttackPlayer(Index, i, True) Then
                                                    SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                                    PlayerAttackPet Index, i, Vital, Spellnum
                                                End If
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
                                    If CanPlayerAttackNpc(Index, i, True) Then
                                        SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc Index, i, Vital, Spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_ZONES
                        For z = 1 To MAX_MAP_NPCS * 2
                            If ZoneNpc(i).Npc(z).Map = GetPlayerMap(Index) Then
                                If ZoneNpc(i).Npc(z).Num > 0 Then
                                    If ZoneNpc(i).Npc(z).Vital(HP) > 0 Then
                                        If isInRange(AoE, x, y, ZoneNpc(i).Npc(z).x, ZoneNpc(i).Npc(z).y) Then
                                            If CanPlayerAttackZoneNpc(Index, i, z, True) Then
                                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_ZONENPC, z, 0, i
                                                PlayerAttackZoneNpc Index, i, z, Vital, Spellnum
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
                                If Player(i).characters(TempPlayer(i).CurChar).InHouse = Player(Index).characters(TempPlayer(Index).CurChar).InHouse Then
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
                        End If
                    Next
                    If False Then 'Simply put, we dont want to heal the enemies!
                        For i = 1 To MAX_MAP_NPCS
                            If MapNpc(MapNum).Npc(i).Num > 0 Then
                                If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                    If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                        SpellNpc_Effect VitalType, increment, i, Vital, Spellnum, MapNum
                                    End If
                                End If
                            End If
                        Next
                        For i = 1 To MAX_ZONES
                            For z = 1 To MAX_MAP_NPCS * 2
                                If ZoneNpc(i).Npc(z).Map = GetPlayerMap(Index) Then
                                    If ZoneNpc(i).Npc(z).Num > 0 Then
                                        If ZoneNpc(i).Npc(z).Vital(HP) > 0 Then
                                            If isInRange(AoE, x, y, ZoneNpc(i).Npc(z).x, ZoneNpc(i).Npc(z).y) Then
                                                SpellNpc_Effect VitalType, increment, z, Vital, Spellnum, MapNum, i
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
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
            ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                If GetPlayerMap(Index) = ZoneNpc(TargetZone).Npc(Target).Map Then
                    x = ZoneNpc(TargetZone).Npc(Target).x
                    y = ZoneNpc(TargetZone).Npc(Target).y
                Else
                    Exit Sub
                End If
            ElseIf TargetType = TARGET_TYPE_PET Then
                x = Player(Target).characters(TempPlayer(Target).CurChar).Pet.x
                y = Player(Target).characters(TempPlayer(Target).CurChar).Pet.y
            End If
               
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
                Exit Sub
            End If
           
            Select Case Spell(Spellnum).type
                Case SPELL_TYPE_DAMAGEHP
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                PlayerAttackPlayer Index, Target, Vital, Spellnum
                                DidCast = True
                            End If
                        End If
                    ElseIf TargetType = TARGET_TYPE_NPC Then
                        If CanPlayerAttackNpc(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                PlayerAttackNpc Index, Target, Vital, Spellnum
                                DidCast = True
                            End If
                        End If
                    ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                        If CanPlayerAttackZoneNpc(Index, TargetZone, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_ZONENPC, Target, 0, TargetZone
                                PlayerAttackZoneNpc Index, TargetZone, Target, Vital, Spellnum
                                DidCast = True
                            End If
                        End If
                    ElseIf TargetType = TARGET_TYPE_PET Then
                        If Player(Target).characters(TempPlayer(Target).CurChar).Pet.Alive Then
                            If CanPlayerAttackPet(Index, Target, True) Then
                                If Vital > 0 Then
                                    SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PET, Target
                                    PlayerAttackPet Index, Target, Vital, Spellnum
                                    DidCast = True
                                End If
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(Spellnum).type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(Spellnum).type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(Spellnum).type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    End If
                   
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If Spell(Spellnum).type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(Index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, Spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, Spellnum
                        End If
                    ElseIf TargetType = TARGET_TYPE_NPC Then
                        If Spell(Spellnum).type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(Index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, Spellnum, MapNum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, Spellnum, MapNum
                        End If
                    ElseIf TargetType = TARGET_TYPE_ZONENPC Then
                        If Spell(Spellnum).type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackZoneNpc(Index, TargetZone, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, Spellnum, MapNum, TargetZone
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, Spellnum, MapNum, TargetZone
                        End If
                    End If
            End Select
        Case 4 ' Projectile
            Call PlayerFireProjectile(Index, Spellnum)
            DidCast = True
    End Select
   
    If DidCast Then
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPCost)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
       
        TempPlayer(Index).SpellCD(spellslot) = GetTickCount + (Spell(Spellnum).CDTime * 1000)
        Call SendCooldown(Index, spellslot)
        SendActionMsg MapNum, Trim$(Spell(Spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CastSpell", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal Spellnum As Long)
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
    
        SendAnimation GetPlayerMap(Index), Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        
        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, Spellnum
        
        If increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
            If Spell(Spellnum).Duration > 0 Then
                AddHoT_Player Index, Spellnum
            End If
        ElseIf Not increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
        End If
        
        SendVital Index, Vital
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpellPlayer_Effect", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal Spellnum As Long, ByVal MapNum As Long, Optional ZoneNum As Long = 0)
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
        
        If ZoneNum > 0 Then
            SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_ZONENPC, Index, 0, ZoneNum
            SendActionMsg MapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, ZoneNpc(ZoneNum).Npc(Index).x * 32, ZoneNpc(ZoneNum).Npc(Index).y * 32
            
            ' send the sound
            SendMapSound Index, ZoneNpc(ZoneNum).Npc(Index).x, ZoneNpc(ZoneNum).Npc(Index).y, SoundEntity.seSpell, Spellnum
            
            If increment Then
               ZoneNpc(ZoneNum).Npc(Index).Vital(Vital) = ZoneNpc(ZoneNum).Npc(Index).Vital(Vital) + Damage
                If Spell(Spellnum).Duration > 0 Then
                    AddHoT_Npc MapNum, Index, Spellnum, ZoneNum
                End If
            ElseIf Not increment Then
               ZoneNpc(ZoneNum).Npc(Index).Vital(Vital) = ZoneNpc(ZoneNum).Npc(Index).Vital(Vital) - Damage
            End If

        Else
            SendAnimation MapNum, Spell(Spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
            SendActionMsg MapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(Index).x * 32, MapNpc(MapNum).Npc(Index).y * 32
            
            ' send the sound
            SendMapSound Index, MapNpc(MapNum).Npc(Index).x, MapNpc(MapNum).Npc(Index).y, SoundEntity.seSpell, Spellnum
            
            If increment Then
                MapNpc(MapNum).Npc(Index).Vital(Vital) = MapNpc(MapNum).Npc(Index).Vital(Vital) + Damage
                If Spell(Spellnum).Duration > 0 Then
                    AddHoT_Npc MapNum, Index, Spellnum, 0
                End If
            ElseIf Not increment Then
                MapNpc(MapNum).Npc(Index).Vital(Vital) = MapNpc(MapNum).Npc(Index).Vital(Vital) - Damage
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SpellNpc_Effect", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal Spellnum As Long, ByVal Caster As Long)
Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            If .Spell = Spellnum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = Spellnum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddDoT_Player", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub AddHoT_Player(ByVal Index As Long, ByVal Spellnum As Long)
Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(i)
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
    HandleError "AddHoT_Player", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal Spellnum As Long, ByVal Caster As Long, Optional ZoneNum As Long = 0)
Dim i As Long

   On Error GoTo errorhandler

    If ZoneNum > 0 Then
        For i = 1 To MAX_DOTS
            With ZoneNpc(ZoneNum).Npc(Index).DoT(i)
                If .Spell = Spellnum Then
                    .Timer = GetTickCount
                    .Caster = Caster
                    .StartTime = GetTickCount
                    Exit Sub
                End If
                
                If .Used = False Then
                    .Spell = Spellnum
                    .Timer = GetTickCount
                    .Caster = Caster
                    .Used = True
                    .StartTime = GetTickCount
                    Exit Sub
                End If
            End With
        Next
    Else
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).Npc(Index).DoT(i)
                If .Spell = Spellnum Then
                    .Timer = GetTickCount
                    .Caster = Caster
                    .StartTime = GetTickCount
                    Exit Sub
                End If
                
                If .Used = False Then
                    .Spell = Spellnum
                    .Timer = GetTickCount
                    .Caster = Caster
                    .Used = True
                    .StartTime = GetTickCount
                    Exit Sub
                End If
            End With
        Next
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddDoT_Npc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal Spellnum As Long, ZoneNum As Long)
Dim i As Long


   On Error GoTo errorhandler

    If ZoneNum > 0 Then
        For i = 1 To MAX_DOTS
            With ZoneNpc(ZoneNum).Npc(Index).HoT(i)
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
    Else
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).Npc(Index).HoT(i)
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
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "AddHoT_Npc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub HandleDoT_Player(ByVal Index As Long, ByVal dotNum As Long)

   On Error GoTo errorhandler

    With TempPlayer(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, Index, True) Then
                    PlayerAttackPlayer .Caster, Index, Spell(.Spell).Vital
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
    HandleError "HandleDoT_Player", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub HandleHoT_Player(ByVal Index As Long, ByVal hotNum As Long)

   On Error GoTo errorhandler

    With TempPlayer(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).type = SPELL_TYPE_HEALHP Then
                   SendActionMsg Player(Index).characters(TempPlayer(Index).CurChar).Map, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Player(Index).characters(TempPlayer(Index).CurChar).x * 32, Player(Index).characters(TempPlayer(Index).CurChar).y * 32
                   SetPlayerVital Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + Spell(.Spell).Vital
                Else
                   SendActionMsg Player(Index).characters(TempPlayer(Index).CurChar).Map, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, Player(Index).characters(TempPlayer(Index).CurChar).x * 32, Player(Index).characters(TempPlayer(Index).CurChar).y * 32
                   SetPlayerVital Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + Spell(.Spell).Vital
                End If
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
    HandleError "HandleHoT_Player", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub HandleDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal dotNum As Long)

   On Error GoTo errorhandler

    With MapNpc(MapNum).Npc(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, Index, True) Then
                    PlayerAttackNpc .Caster, Index, Spell(.Spell).Vital, , True
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
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
    HandleError "HandleDoT_Npc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub HandleHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal hotNum As Long)

   On Error GoTo errorhandler

    With MapNpc(MapNum).Npc(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).type = SPELL_TYPE_HEALHP Then
                    SendActionMsg MapNum, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(Index).x * 32, MapNpc(MapNum).Npc(Index).y * 32
                    MapNpc(MapNum).Npc(Index).Vital(Vitals.HP) = MapNpc(MapNum).Npc(Index).Vital(Vitals.HP) + Spell(.Spell).Vital
                Else
                    SendActionMsg MapNum, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(Index).x * 32, MapNpc(MapNum).Npc(Index).y * 32
                    MapNpc(MapNum).Npc(Index).Vital(Vitals.MP) = MapNpc(MapNum).Npc(Index).Vital(Vitals.MP) + Spell(.Spell).Vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
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
    HandleError "HandleHoT_Npc", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub StunPlayer(ByVal Index As Long, ByVal Spellnum As Long)
    ' check if it's a stunning spell

   On Error GoTo errorhandler

    If Spell(Spellnum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Spell(Spellnum).StunDuration
        TempPlayer(Index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned Index
        ' tell him he's stunned
        PlayerMsg Index, "You have been stunned.", BrightRed
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "StunPlayer", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal MapNum As Long, ByVal Spellnum As Long, Optional ZoneNum As Long = 0)
    ' check if it's a stunning spell

   On Error GoTo errorhandler

    If Spell(Spellnum).StunDuration > 0 Then
        ' set the values on index
        If ZoneNum > 0 Then
            ZoneNpc(ZoneNum).Npc(Index).StunDuration = Spell(Spellnum).StunDuration
            ZoneNpc(ZoneNum).Npc(Index).StunTimer = GetTickCount
        Else
            MapNpc(MapNum).Npc(Index).StunDuration = Spell(Spellnum).StunDuration
            MapNpc(MapNum).Npc(Index).StunTimer = GetTickCount
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "StunNPC", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub TryZoneNpcAttackPet(ZoneNum As Long, ZoneNPCNum As Long, ByVal Index As Long)
Dim MapNum As Long, npcnum As Long, blockAmount As Long, Damage As Long, i As Long, armor As Long

    ' Can the npc attack the player?

   On Error GoTo errorhandler

    If CanZoneNpcAttackPet(ZoneNum, ZoneNPCNum, Index) Then
        MapNum = GetPlayerMap(Index)
        npcnum = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPetDodge(Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32)
            Exit Sub
        End If
        If CanPetParry(Index) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcnum)
        
        If NewOptions.CombatMode = 1 Then
                ' take away armour
                Damage = Damage - ((Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(Stats.Willpower) * 2) + (Player(Index).characters(TempPlayer(Index).CurChar).Pet.Level * 3) + armor)
                ' * 1.5 if crit hit
                If CanNpcCrit(Index) Then
                    Damage = Damage * 1.5
                    SendActionMsg MapNum, "Critical!", BrightCyan, 1, (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32), (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32)
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
            If CanNpcCrit(Index) Then
                Damage = Damage * 1.5
                SendActionMsg MapNum, "Critical!", BrightCyan, 1, (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x * 32), (ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y * 32)
            End If
        End If
        


        If Damage > 0 Then
            Call ZoneNpcAttackPet(ZoneNum, ZoneNPCNum, Index, Damage)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TryZoneNpcAttackPet", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function CanZoneNpcAttackPet(ZoneNum As Long, ZoneNPCNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long
    Dim npcnum As Long

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If ZoneNPCNum <= 0 Or ZoneNPCNum > MAX_MAP_NPCS * 2 Or Not IsPlaying(Index) Then
        Exit Function
    End If
    
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Then Exit Function

    ' Check for subscript out of range
    If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Index)
    npcnum = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num

    ' Make sure the npc isn't already dead
    If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < ZoneNpc(ZoneNum).Npc(ZoneNPCNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    ZoneNpc(ZoneNum).Npc(ZoneNPCNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If npcnum > 0 Then
            
            ' Check if at same coordinates
            If (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y + 1 = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y) And (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x) Then
                CanZoneNpcAttackPet = True
            Else
                If (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y - 1 = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y) And (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x) Then
                    CanZoneNpcAttackPet = True
                Else
                    If (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y) And (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x + 1 = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x) Then
                        CanZoneNpcAttackPet = True
                    Else
                        If (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).y) And (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x - 1 = ZoneNpc(ZoneNum).Npc(ZoneNPCNum).x) Then
                            CanZoneNpcAttackPet = True
                        End If
                    End If
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanZoneNpcAttackPet", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub ZoneNpcAttackPet(ZoneNum As Long, ZoneNPCNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim MapNum As Long
    Dim i As Long, z As Long, InvCount As Long, EqCount As Long, j As Long, x As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range

   On Error GoTo errorhandler

    If ZoneNPCNum <= 0 Or ZoneNPCNum > MAX_MAP_NPCS * 2 Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Victim)
    Name = Trim$(Npc(ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong ZoneNum
    Buffer.WriteLong ZoneNPCNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    ZoneNpc(ZoneNum).Npc(ZoneNPCNum).stopRegen = True
    ZoneNpc(ZoneNum).Npc(ZoneNPCNum).stopRegenTimer = GetTickCount

    If Damage >= Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health, BrightRed, 1, (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x * 32), (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y * 32)
        
        ' send the sound
        SendMapSound Victim, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y, SoundEntity.seNpc, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
        
        ' kill player
        Call PlayerMsg(Victim, "Your " & Trim$(Pet(Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Num).Name) & " was killed by a " & Trim$(Npc(ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num).Name) & ".", BrightRed)
        ReleasePet (Victim)

        ' Now that pet is dead, go for owner
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Target = Victim
        ZoneNpc(ZoneNum).Npc(ZoneNPCNum).TargetType = TARGET_TYPE_PLAYER
    Else
        ' Player not dead, just do the damage
        Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health = Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.Health - Damage
        Call SendPetVital(Victim, Vitals.HP)
        Call SendAnimation(MapNum, Npc(ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num).Animation, 0, 0, TARGET_TYPE_PET, Victim)
        
        ' send the sound
        SendMapSound Victim, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y, SoundEntity.seNpc, ZoneNpc(ZoneNum).Npc(ZoneNPCNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x * 32), (Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y * 32)
        SendBlood GetPlayerMap(Victim), Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.x, Player(Victim).characters(TempPlayer(Victim).CurChar).Pet.y
        
        ' set the regen timer
        TempPlayer(Victim).PetstopRegen = True
        TempPlayer(Victim).PetstopRegenTimer = GetTickCount
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ZoneNpcAttackPet", "modCombat", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

