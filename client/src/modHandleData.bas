Attribute VB_Name = "modHandleData"
Option Explicit
' in Pixels
Property Get GetTitleHeight()

   On Error GoTo errorhandler

    GetTitleHeight = TwipsToPixels(frmMain.Height - frmMain.ScaleHeight)


   On Error GoTo 0
   Exit Property
errorhandler:
    HandleError "GetTitleHeight", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Property
Property Get GetTitleWidth()

   On Error GoTo errorhandler

    GetTitleWidth = TwipsToPixels(frmMain.Width - frmMain.ScaleWidth)


   On Error GoTo 0
   Exit Property
errorhandler:
    HandleError "GetTitleWidth", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Property

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr


End Function

Public Sub InitMessages()

   On Error GoTo errorhandler

    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    'Events
    HandleDataSub(SSpawnEvent) = GetAddress(AddressOf HandleSpawnEventPage)
    HandleDataSub(SEventMove) = GetAddress(AddressOf HandleEventMove)
    HandleDataSub(SEventDir) = GetAddress(AddressOf HandleEventDir)
    HandleDataSub(SEventChat) = GetAddress(AddressOf HandleEventChat)
    HandleDataSub(SEventStart) = GetAddress(AddressOf HandleEventStart)
    HandleDataSub(SEventEnd) = GetAddress(AddressOf HandleEventEnd)
    HandleDataSub(SPlayBGM) = GetAddress(AddressOf HandlePlayBGM)
    HandleDataSub(SPlaySound) = GetAddress(AddressOf HandlePlaySound)
    HandleDataSub(SFadeoutBGM) = GetAddress(AddressOf HandleFadeoutBGM)
    HandleDataSub(SStopSound) = GetAddress(AddressOf HandleStopSound)
    HandleDataSub(SSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(SMapEventData) = GetAddress(AddressOf HandleMapEventData)
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    HandleDataSub(SSpecialEffect) = GetAddress(AddressOf HandleSpecialEffect)
    HandleDataSub(SHouseConfigs) = GetAddress(AddressOf HandleHouseConfigurations)
    HandleDataSub(SBuyHouse) = GetAddress(AddressOf HandleHouseOffer)
    HandleDataSub(SMax) = GetAddress(AddressOf HandleMaxes)
    HandleDataSub(SVisit) = GetAddress(AddressOf HandleVisit)
    HandleDataSub(SFurniture) = GetAddress(AddressOf HandleFurniture)
    HandleDataSub(SMailBox) = GetAddress(AddressOf HandleMailBox)
    HandleDataSub(SMailUnread) = GetAddress(AddressOf HandleUnreadMail)
    HandleDataSub(SSelChar) = GetAddress(AddressOf HandleSelectChar)
    HandleDataSub(SFriends) = GetAddress(AddressOf HandleFriends)
    HandleDataSub(SQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(SUpdateQuest) = GetAddress(AddressOf HandleUpdateQuest)
    HandleDataSub(SPlayerQuest) = GetAddress(AddressOf HandlePlayerQuest)
    HandleDataSub(SQuestMessage) = GetAddress(AddressOf HandleQuestMessage)
    HandleDataSub(SZoneEdit) = GetAddress(AddressOf HandleEditZones)

    HandleDataSub(SServerInfo) = GetAddress(AddressOf HandleServerInfo)
    HandleDataSub(SPlayers) = GetAddress(AddressOf HandlePlayers)
    HandleDataSub(SAccounts) = GetAddress(AddressOf HandleAccounts)
    HandleDataSub(SAdmin) = GetAddress(AddressOf HandleAdmin)
    HandleDataSub(SMaps) = GetAddress(AddressOf HandleMaps)
    HandleDataSub(SBans) = GetAddress(AddressOf HandleBans)
    HandleDataSub(SServerOpts) = GetAddress(AddressOf HandleServerOpts)
    HandleDataSub(SHouseEdit) = GetAddress(AddressOf HandleEditHouses)
    HandleDataSub(SEditPlayer) = GetAddress(AddressOf HandleEditPlayer)
    
    HandleDataSub(SPic) = GetAddress(AddressOf HandlePicture)
    HandleDataSub(SHoldPlayer) = GetAddress(AddressOf HandleHoldPlayer)
    HandleDataSub(SGameOpts) = GetAddress(AddressOf HandleGameOpts)
    
    HandleDataSub(SPetEditor) = GetAddress(AddressOf HandlePetEditor)
    HandleDataSub(SUpdatePet) = GetAddress(AddressOf HandleUpdatePet)
    HandleDataSub(SPetMove) = GetAddress(AddressOf HandlePetMove)
    HandleDataSub(SPetDir) = GetAddress(AddressOf HandlePetDir)
    HandleDataSub(SPetVital) = GetAddress(AddressOf HandlePetVital)
    HandleDataSub(SClearPetSpellBuffer) = GetAddress(AddressOf HandleClearPetSpellBuffer)
    
    HandleDataSub(SRandomDungeonEditor) = GetAddress(AddressOf HandleRandomDungeonEditor)
    HandleDataSub(SUpdateRandomDungeon) = GetAddress(AddressOf HandleUpdateRandomDungeon)
    HandleDataSub(SRandomDungeonMap) = GetAddress(AddressOf HandleRandomDungeonMap)
    HandleDataSub(SProjectileEditor) = GetAddress(AddressOf HandleProjectileEditor)
    HandleDataSub(SUpdateProjectile) = GetAddress(AddressOf HandleUpdateProjectile)
    HandleDataSub(SMapProjectile) = GetAddress(AddressOf HandleMapProjectile)
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandleData(ByRef data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long



   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MsgType = buffer.ReadLong
    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    CallWindowProc HandleDataSub(MsgType), 1, buffer.ReadBytes(buffer.length), 0, 0


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    frmLoad.Visible = False
    'frmMain.Visible = True
    MenuStage = 0
    Msg = buffer.ReadString 'Parse(1)
    Set buffer = Nothing
    Call MsgBox(Msg, vbOKOnly, Servers(ServerIndex).Game_Name)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ' Now we can receive game data
    MyIndex = buffer.ReadLong
    ' player high index
    Player_HighIndex = buffer.ReadLong
    Set buffer = Nothing
    frmLoad.Visible = False
    Call SetStatus("Receiving game data...")




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim z As Long, X As Long, TempStr As String, a() As String
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1
    
    ReDim NewCharMaleHairCount(Max_Classes)
    ReDim NewCharMaleHeadCount(Max_Classes)
    ReDim NewCharMaleEyeCount(Max_Classes)
    ReDim NewCharMaleEyebrowCount(Max_Classes)
    ReDim NewCharMaleEarCount(Max_Classes)
    ReDim NewCharMaleMouthCount(Max_Classes)
    ReDim NewCharMaleNoseCount(Max_Classes)
    ReDim NewCharMaleShirtCount(Max_Classes)
    ReDim NewCharMaleEtcCount(Max_Classes)
    ReDim NewCharMaleFaceCount(Max_Classes)
    
    ReDim NewCharFemaleHairCount(Max_Classes)
    ReDim NewCharFemaleHeadCount(Max_Classes)
    ReDim NewCharFemaleEyeCount(Max_Classes)
    ReDim NewCharFemaleEyebrowCount(Max_Classes)
    ReDim NewCharFemaleEarCount(Max_Classes)
    ReDim NewCharFemaleMouthCount(Max_Classes)
    ReDim NewCharFemaleNoseCount(Max_Classes)
    ReDim NewCharFemaleShirtCount(Max_Classes)
    ReDim NewCharFemaleEtcCount(Max_Classes)
    ReDim NewCharFemaleFaceCount(Max_Classes)

    For i = 1 To Max_Classes

        With Class(i)
            .Name = buffer.ReadString
            .Vital(Vitals.HP) = buffer.ReadLong
            .Vital(Vitals.MP) = buffer.ReadLong
            
            ReDim Class(i).MaleFaceParts.FHair(0)
            ReDim Class(i).MaleFaceParts.FHeads(0)
            ReDim Class(i).MaleFaceParts.FEyes(0)
            ReDim Class(i).MaleFaceParts.FEyebrows(0)
            ReDim Class(i).MaleFaceParts.FEars(0)
            ReDim Class(i).MaleFaceParts.FMouth(0)
            ReDim Class(i).MaleFaceParts.FNose(0)
            ReDim Class(i).MaleFaceParts.FCloth(0)
            ReDim Class(i).MaleFaceParts.FEtc(0)
            ReDim Class(i).MaleFaceParts.FFace(0)
            
            ReDim Class(i).FemaleFaceParts.FHeads(0)
            ReDim Class(i).FemaleFaceParts.FEyes(0)
            ReDim Class(i).FemaleFaceParts.FEyebrows(0)
            ReDim Class(i).FemaleFaceParts.FEars(0)
            ReDim Class(i).FemaleFaceParts.FMouth(0)
            ReDim Class(i).FemaleFaceParts.FNose(0)
            ReDim Class(i).FemaleFaceParts.FCloth(0)
            ReDim Class(i).FemaleFaceParts.FEtc(0)
            ReDim Class(i).FemaleFaceParts.FHair(0)
            ReDim Class(i).FemaleFaceParts.FFace(0)
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FHair(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FHair(X) = a(X)
                Next
                NewCharMaleHairCount(i) = UBound(a) + 1
            Else
                NewCharMaleHairCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FHeads(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FHeads(X) = a(X)
                Next
                NewCharMaleHeadCount(i) = UBound(a) + 1
            Else
                NewCharMaleHeadCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FEyes(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FEyes(X) = a(X)
                Next
                NewCharMaleEyeCount(i) = UBound(a) + 1
            Else
                NewCharMaleEyeCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FEyebrows(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FEyebrows(X) = a(X)
                Next
                NewCharMaleEyebrowCount(i) = UBound(a) + 1
            Else
                NewCharMaleEyebrowCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FEars(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FEars(X) = a(X)
                Next
                NewCharMaleEarCount(i) = UBound(a) + 1
            Else
                NewCharMaleEarCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FMouth(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FMouth(X) = a(X)
                Next
                NewCharMaleMouthCount(i) = UBound(a) + 1
            Else
                NewCharMaleMouthCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FNose(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FNose(X) = a(X)
                Next
                NewCharMaleNoseCount(i) = UBound(a) + 1
            Else
                NewCharMaleNoseCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FCloth(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FCloth(X) = a(X)
                Next
                NewCharMaleShirtCount(i) = UBound(a) + 1
            Else
                NewCharMaleShirtCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FEtc(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FEtc(X) = a(X)
                Next
                NewCharMaleEtcCount(i) = UBound(a) + 1
            Else
                NewCharMaleEtcCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FFace(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FFace(X) = a(X)
                Next
                NewCharMaleFaceCount(i) = UBound(a) + 1
            Else
                NewCharMaleFaceCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FHair(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FHair(X) = a(X)
                Next
                NewCharFemaleHairCount(i) = UBound(a) + 1
            Else
                NewCharFemaleHairCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FHeads(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FHeads(X) = a(X)
                Next
                NewCharFemaleHeadCount(i) = UBound(a) + 1
            Else
                NewCharFemaleHeadCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FEyes(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FEyes(X) = a(X)
                Next
                NewCharFemaleEyeCount(i) = UBound(a) + 1
            Else
                NewCharFemaleEyeCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FEyebrows(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FEyebrows(X) = a(X)
                Next
                NewCharFemaleEyebrowCount(i) = UBound(a) + 1
            Else
                NewCharFemaleEyebrowCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FEars(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FEars(X) = a(X)
                Next
                NewCharFemaleEarCount(i) = UBound(a) + 1
            Else
                NewCharFemaleEarCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FMouth(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FMouth(X) = a(X)
                Next
                NewCharFemaleMouthCount(i) = UBound(a) + 1
            Else
                NewCharFemaleMouthCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FNose(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FNose(X) = a(X)
                Next
                NewCharFemaleNoseCount(i) = UBound(a) + 1
            Else
                NewCharFemaleNoseCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FCloth(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FCloth(X) = a(X)
                Next
                NewCharFemaleShirtCount(i) = UBound(a) + 1
            Else
                NewCharFemaleShirtCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FEtc(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FEtc(X) = a(X)
                Next
                NewCharFemaleEtcCount(i) = UBound(a) + 1
            Else
                NewCharFemaleEtcCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FFace(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FFace(X) = a(X)
                Next
                NewCharFemaleFaceCount(i) = UBound(a) + 1
            Else
                NewCharFemaleFaceCount(i) = 0
            End If
            
            For X = 1 To Stats.Stat_Count - 1
                .stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing
    ' Used for if the player is creating a new character
    'frmMain.Visible = True
    MenuStage = 5
    frmLoad.Visible = False
    NewCharSex = SEX_MALE
    newCharClass = 1
    ResetNewChar
    TxtUsername = ""
    SelTextbox = 1
    newCharSprite = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleNewCharClasses", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim z As Long, X As Long, TempStr As String, a() As String
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1
    
    ReDim NewCharMaleHairCount(Max_Classes)
    ReDim NewCharMaleHeadCount(Max_Classes)
    ReDim NewCharMaleEyeCount(Max_Classes)
    ReDim NewCharMaleEyebrowCount(Max_Classes)
    ReDim NewCharMaleEarCount(Max_Classes)
    ReDim NewCharMaleMouthCount(Max_Classes)
    ReDim NewCharMaleNoseCount(Max_Classes)
    ReDim NewCharMaleShirtCount(Max_Classes)
    ReDim NewCharMaleEtcCount(Max_Classes)
    ReDim NewCharMaleFaceCount(Max_Classes)
    
    ReDim NewCharFemaleHairCount(Max_Classes)
    ReDim NewCharFemaleHeadCount(Max_Classes)
    ReDim NewCharFemaleEyeCount(Max_Classes)
    ReDim NewCharFemaleEyebrowCount(Max_Classes)
    ReDim NewCharFemaleEarCount(Max_Classes)
    ReDim NewCharFemaleMouthCount(Max_Classes)
    ReDim NewCharFemaleNoseCount(Max_Classes)
    ReDim NewCharFemaleShirtCount(Max_Classes)
    ReDim NewCharFemaleEtcCount(Max_Classes)
    ReDim NewCharFemaleFaceCount(Max_Classes)

    For i = 1 To Max_Classes

        With Class(i)
            .Name = buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = buffer.ReadLong 'CLng(Parse(n + 2))
                    ReDim Class(i).MaleFaceParts.FHair(0)
            ReDim Class(i).MaleFaceParts.FHeads(0)
            ReDim Class(i).MaleFaceParts.FEyes(0)
            ReDim Class(i).MaleFaceParts.FEyebrows(0)
            ReDim Class(i).MaleFaceParts.FEars(0)
            ReDim Class(i).MaleFaceParts.FMouth(0)
            ReDim Class(i).MaleFaceParts.FNose(0)
            ReDim Class(i).MaleFaceParts.FCloth(0)
            ReDim Class(i).MaleFaceParts.FEtc(0)
            ReDim Class(i).FemaleFaceParts.FHair(0)
            ReDim Class(i).FemaleFaceParts.FHeads(0)
            ReDim Class(i).FemaleFaceParts.FEyes(0)
            ReDim Class(i).FemaleFaceParts.FEyebrows(0)
            ReDim Class(i).FemaleFaceParts.FEars(0)
            ReDim Class(i).FemaleFaceParts.FMouth(0)
            ReDim Class(i).FemaleFaceParts.FNose(0)
            ReDim Class(i).FemaleFaceParts.FCloth(0)
            ReDim Class(i).FemaleFaceParts.FEtc(0)
            
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FHair(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FHair(X) = a(X)
                Next
                NewCharMaleHairCount(i) = UBound(a) + 1
            Else
                NewCharMaleHairCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FHeads(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FHeads(X) = a(X)
                Next
                NewCharMaleHeadCount(i) = UBound(a) + 1
            Else
                NewCharMaleHeadCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FEyes(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FEyes(X) = a(X)
                Next
                NewCharMaleEyeCount(i) = UBound(a) + 1
            Else
                NewCharMaleEyeCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FEyebrows(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FEyebrows(X) = a(X)
                Next
                NewCharMaleEyebrowCount(i) = UBound(a) + 1
            Else
                NewCharMaleEyebrowCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FEars(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FEars(X) = a(X)
                Next
                NewCharMaleEarCount(i) = UBound(a) + 1
            Else
                NewCharMaleEarCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FMouth(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FMouth(X) = a(X)
                Next
                NewCharMaleMouthCount(i) = UBound(a) + 1
            Else
                NewCharMaleMouthCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FNose(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FNose(X) = a(X)
                Next
                NewCharMaleNoseCount(i) = UBound(a) + 1
            Else
                NewCharMaleNoseCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FCloth(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FCloth(X) = a(X)
                Next
                NewCharMaleShirtCount(i) = UBound(a) + 1
            Else
                NewCharMaleShirtCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FEtc(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FEtc(X) = a(X)
                Next
                NewCharMaleEtcCount(i) = UBound(a) + 1
            Else
                NewCharMaleEtcCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).MaleFaceParts.FFace(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).MaleFaceParts.FFace(X) = a(X)
                Next
                NewCharMaleFaceCount(i) = UBound(a) + 1
            Else
                NewCharMaleFaceCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FHair(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FHair(X) = a(X)
                Next
                NewCharFemaleHairCount(i) = UBound(a) + 1
            Else
                NewCharFemaleHairCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FHeads(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FHeads(X) = a(X)
                Next
                NewCharFemaleHeadCount(i) = UBound(a) + 1
            Else
                NewCharFemaleHeadCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FEyes(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FEyes(X) = a(X)
                Next
                NewCharFemaleEyeCount(i) = UBound(a) + 1
            Else
                NewCharFemaleEyeCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FEyebrows(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FEyebrows(X) = a(X)
                Next
                NewCharFemaleEyebrowCount(i) = UBound(a) + 1
            Else
                NewCharFemaleEyebrowCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FEars(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FEars(X) = a(X)
                Next
                NewCharFemaleEarCount(i) = UBound(a) + 1
            Else
                NewCharFemaleEarCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FMouth(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FMouth(X) = a(X)
                Next
                NewCharFemaleMouthCount(i) = UBound(a) + 1
            Else
                NewCharFemaleMouthCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FNose(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FNose(X) = a(X)
                Next
                NewCharFemaleNoseCount(i) = UBound(a) + 1
            Else
                NewCharFemaleNoseCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FCloth(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FCloth(X) = a(X)
                Next
                NewCharFemaleShirtCount(i) = UBound(a) + 1
            Else
                NewCharFemaleShirtCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FEtc(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FEtc(X) = a(X)
                Next
                NewCharFemaleEtcCount(i) = UBound(a) + 1
            Else
                NewCharFemaleEtcCount(i) = 0
            End If
            
            TempStr = buffer.ReadString
            If Len(TempStr) > 0 Then
                a = Split(TempStr, ",")
                ReDim Class(i).FemaleFaceParts.FFace(UBound(a))
                For X = 0 To UBound(a)
                    Class(i).FemaleFaceParts.FFace(X) = a(X)
                Next
                NewCharFemaleFaceCount(i) = UBound(a) + 1
            Else
                NewCharFemaleFaceCount(i) = 0
            End If
            
            For X = 1 To Stats.Stat_Count - 1
                .stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandleInGame(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
   On Error GoTo errorhandler

    InGame = True
    'If Options.FullScreen = 1 Then
    '    frmMain.BorderStyle = 1
    '    frmMain.WindowState = 0
    '    frmMain.ClipControls = True
    '    UpdateDebugCaption
    'End If
    'frmMain.Width = PixelsToTwips(CInt(GameWindowWidth) + GetTitleWidth)
    'frmMain.Height = PixelsToTwips(CInt(GameWindowHeight) + GetTitleHeight)
    'frmMain.Left = (Screen.Width - frmMain.Width) / 2
    'frmMain.Top = (Screen.Height - frmMain.Height) / 2
   ' If Options.FullScreen = 1 Then
   '     frmMain.WindowState = 2
   '     frmMain.BorderStyle = 0
   '     frmMain.ClipControls = False
   '     Call SetWindowLong(frmMain.hwnd, GWL_STYLE, GetWindowLong(frmMain.hwnd, GWL_STYLE) Xor WS_CAPTION Xor WS_BORDER)
   '     Call SetWindowPos(frmMain.hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
   ' End If
   
   
    WalkToX = GetPlayerX(MyIndex)
    WalkToY = GetPlayerY(MyIndex)
    HoldPlayer = False
    For i = 1 To 10
        Pictures(i).pic = 0
    Next
    Call GameInit
    Call GameLoop




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, buffer.ReadLong)
        n = n + 2
    Next
    ' changes to inventory, need to clear any drop menu
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, buffer.ReadLong) 'CLng(Parse(3)))
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Shield)
    ' changes to inventory, need to clear any drop menu
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim playerNum As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    playerNum = buffer.ReadLong
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Shield)
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.HP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, buffer.ReadLong)
    If buffer.ReadLong = 1 Then
        Player(Index).Pet.Health = buffer.ReadLong
        Player(Index).Pet.MaxHp = buffer.ReadLong
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerHp", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.MP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, buffer.ReadLong)
    If buffer.ReadLong = 1 Then
        Player(Index).Pet.Mana = buffer.ReadLong
        Player(Index).Pet.MaxMP = buffer.ReadLong
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerMp", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, i, buffer.ReadLong
    Next




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long
Dim TNL As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    SetPlayerExp MyIndex, buffer.ReadLong
    TNL = buffer.ReadLong




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerExp", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long, X As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    Call SetPlayerName(i, buffer.ReadString)
    Call SetPlayerLevel(i, buffer.ReadLong)
    Call SetPlayerPOINTS(i, buffer.ReadLong)
    For X = 1 To FaceEnum.Face_Count - 1
        Player(i).Face(X) = buffer.ReadLong
    Next
    For X = 1 To SpriteEnum.Sprite_Count - 1
        Player(i).Sprite(X) = buffer.ReadLong
    Next
    Player(i).Sex = buffer.ReadLong
    Call SetPlayerMap(i, buffer.ReadLong)
    Call SetPlayerX(i, buffer.ReadLong)
    Call SetPlayerY(i, buffer.ReadLong)
    Call SetPlayerDir(i, buffer.ReadLong)
    Call SetPlayerAccess(i, buffer.ReadLong)
    Call SetPlayerPK(i, buffer.ReadLong)
    Call SetPlayerClass(i, buffer.ReadLong)
    Player(i).InHouse = buffer.ReadLong
    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, buffer.ReadLong
    Next

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
    End If

    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    
    Player(i).Pet.Num = buffer.ReadLong
    Player(i).Pet.Health = buffer.ReadLong
    Player(i).Pet.Mana = buffer.ReadLong
    Player(i).Pet.Level = buffer.ReadLong
    
    For X = 1 To Stats.Stat_Count - 1
        Player(i).Pet.stat(X) = buffer.ReadLong
    Next
    
    For X = 1 To 4
       Player(i).Pet.spell(X) = buffer.ReadLong
    Next
    
    Player(i).Pet.X = buffer.ReadLong
    Player(i).Pet.Y = buffer.ReadLong
    Player(i).Pet.dir = buffer.ReadLong
    
    Player(i).Pet.MaxHp = buffer.ReadLong
    Player(i).Pet.MaxMP = buffer.ReadLong
    
    If buffer.ReadLong = 1 Then
        Player(i).Pet.Alive = True
    Else
        Player(i).Pet.Alive = False
    End If
    
    Player(i).Pet.AttackBehaviour = buffer.ReadLong
    Player(i).Pet.Points = buffer.ReadLong
    Player(i).Pet.Exp = buffer.ReadLong
    Player(i).Pet.TNL = buffer.ReadLong
    
    ' Make sure their pet isn't walking
    Player(i).Pet.Moving = 0
    Player(i).Pet.XOffset = 0
    Player(i).Pet.YOffset = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim dir As Long
Dim n As Byte
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    dir = buffer.ReadLong
    n = buffer.ReadLong
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, dir)
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)
        Case DIR_UP
            Player(i).YOffset = PIC_Y
        Case DIR_DOWN
            Player(i).YOffset = PIC_Y * -1
        Case DIR_LEFT
            Player(i).XOffset = PIC_X
        Case DIR_RIGHT
            Player(i).XOffset = PIC_X * -1
    End Select




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim X As Long
Dim Y As Long
Dim dir As Long
Dim Movement As Long
Dim buffer As clsBuffer, zonenum As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    zonenum = buffer.ReadLong
    MapNpcNum = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    dir = buffer.ReadLong
    Movement = buffer.ReadLong

    If zonenum > 0 Then
        With ZoneNPC(zonenum).Npc(MapNpcNum)
            .X = X
            .Y = Y
            .dir = dir
            .XOffset = 0
            .YOffset = 0
            .Moving = Movement
            Select Case .dir
                Case DIR_UP
                    .YOffset = PIC_Y
                Case DIR_DOWN
                    .YOffset = PIC_Y * -1
                Case DIR_LEFT
                    .XOffset = PIC_X
                Case DIR_RIGHT
                    .XOffset = PIC_X * -1
            End Select
        End With
    Else
        With MapNpc(MapNpcNum)
            .X = X
            .Y = Y
            .dir = dir
            .XOffset = 0
            .YOffset = 0
            .Moving = Movement
            Select Case .dir
                Case DIR_UP
                    .YOffset = PIC_Y
                Case DIR_DOWN
                    .YOffset = PIC_Y * -1
                Case DIR_LEFT
                    .XOffset = PIC_X
                Case DIR_RIGHT
                    .XOffset = PIC_X * -1
            End Select
        End With
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim dir As Byte
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    dir = buffer.ReadLong
    Call SetPlayerDir(i, dir)

    With Player(i)
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim dir As Byte
Dim buffer As clsBuffer, zonenum As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    zonenum = buffer.ReadLong
    i = buffer.ReadLong
    dir = buffer.ReadLong
    If zonenum = 0 Then
        With MapNpc(i)
            .dir = dir
            .XOffset = 0
            .YOffset = 0
            .Moving = 0
        End With
    Else
        With ZoneNPC(zonenum).Npc(i)
            .dir = dir
            .XOffset = 0
            .YOffset = 0
            .Moving = 0
        End With
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim dir As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    dir = buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).XOffset = 0
    Player(MyIndex).YOffset = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim dir As Long
Dim buffer As clsBuffer
Dim thePlayer As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    dir = buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).XOffset = 0
    Player(thePlayer).YOffset = 0
    GettingMap = False
    CanMoveNow = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerXYMap", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    ' Set player to attacking
    If buffer.ReadLong = 0 Then
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount
    Else
        Player(i).Pet.Attacking = 1
        Player(i).Pet.AttackTimer = GetTickCount
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer, zonenum As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    zonenum = buffer.ReadLong
    If zonenum > 0 Then
        i = buffer.ReadLong
        If i > 0 Then
            ZoneNPC(zonenum).Npc(i).Attacking = 1
            ZoneNPC(zonenum).Npc(i).AttackTimer = GetTickCount
        End If
    Else
        i = buffer.ReadLong
        ' Set player to attacking
        MapNpc(i).Attacking = 1
        MapNpc(i).AttackTimer = GetTickCount
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim i As Long
Dim NeedMap As Byte
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next
    
    ' Erase all projectiles
    For i = 1 To MAX_PROJECTILES
        ClearMapProjectile i
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    CacheMap = False
    ' clear the blood
    For i = 1 To MAX_BYTE
        Blood(i).X = 0
        Blood(i).Y = 0
        Blood(i).Sprite = 0
        Blood(i).Timer = 0
    Next
    Map.CurrentEvents = 0
    ReDim Map.MapEvents(0)
    ' Get map num
    X = buffer.ReadLong
    ' Get revision
    Y = buffer.ReadLong

    If FileExist(MAP_PATH & "map" & X & MAP_EXT, False) Then
        Call LoadMap(X)
        ' Check to see if the revisions match
        NeedMap = 1

        If Map.Revision = Y Then
            ' We do so we dont need the map
            'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
            NeedMap = 0
            CacheNewMapSounds
            initAutotiles
        End If

    Else
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong NeedMap
    SendData buffer.ToArray()
    Set buffer = Nothing
    GettingMap = True
    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
            ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandleMapData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim i As Long, z As Long, w As Long
Dim buffer As clsBuffer
Dim MapNum As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    MapNum = buffer.ReadLong
    Map.Name = buffer.ReadString
    Map.Music = buffer.ReadString
    Map.BGS = buffer.ReadString
    Map.Revision = buffer.ReadLong
    Map.Moral = buffer.ReadByte
    Map.Up = buffer.ReadLong
    Map.Down = buffer.ReadLong
    Map.Left = buffer.ReadLong
    Map.Right = buffer.ReadLong
    Map.BootMap = buffer.ReadLong
    Map.BootX = buffer.ReadByte
    Map.BootY = buffer.ReadByte
    Map.Weather = buffer.ReadLong
    Map.WeatherIntensity = buffer.ReadLong
    Map.Fog = buffer.ReadLong
    Map.FogSpeed = buffer.ReadLong
    Map.FogOpacity = buffer.ReadLong
    Map.Red = buffer.ReadLong
    Map.Green = buffer.ReadLong
    Map.Blue = buffer.ReadLong
    Map.Alpha = buffer.ReadLong
    Map.MaxX = buffer.ReadByte
    Map.MaxY = buffer.ReadByte
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    ReDim Map.exTile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Layer(i).X = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Y = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Tileset = buffer.ReadLong
            Next
            For z = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Autotile(z) = buffer.ReadLong
            Next
            Map.Tile(X, Y).type = buffer.ReadByte
            Map.Tile(X, Y).Data1 = buffer.ReadLong
            Map.Tile(X, Y).data2 = buffer.ReadLong
            Map.Tile(X, Y).Data3 = buffer.ReadLong
            Map.Tile(X, Y).Data4 = buffer.ReadString
            Map.Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map.Npc(X) = buffer.ReadLong
        Map.NpcSpawnType(X) = buffer.ReadLong
        n = n + 1
    Next
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To ExMapLayer.Layer_Count - 1
                Map.exTile(X, Y).Layer(i).X = buffer.ReadLong
                Map.exTile(X, Y).Layer(i).Y = buffer.ReadLong
                Map.exTile(X, Y).Layer(i).Tileset = buffer.ReadLong
            Next
            For z = 1 To ExMapLayer.Layer_Count - 1
                Map.exTile(X, Y).Autotile(z) = buffer.ReadLong
            Next
        Next
    Next

    ClearTempTile
    initAutotiles
    Set buffer = Nothing
    ' Save the map
    Call SaveMap(MapNum)

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
            ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If
    CacheNewMapSounds





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_MAP_ITEMS
        With MapItem(i)
            .playerName = buffer.ReadString
            .Num = buffer.ReadLong
            .Value = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
        End With
    Next




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer, stoploop As Boolean, zonenum As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(i)
            .Num = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .dir = buffer.ReadLong
            .Vital(HP) = buffer.ReadLong
        End With
    Next
    Do Until stoploop = True
        zonenum = buffer.ReadLong
        If zonenum > 0 Then
            i = buffer.ReadLong
            With ZoneNPC(zonenum).Npc(i)
                .Num = buffer.ReadLong
                .X = buffer.ReadLong
                .Y = buffer.ReadLong
                .dir = buffer.ReadLong
                .Vital(HP) = buffer.ReadLong
                .Map = buffer.ReadLong
            End With
        Else
            stoploop = True
        End If
    Loop




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleMapDone()
Dim i As Long
Dim MusicFile As String
    ' clear the action msgs

   On Error GoTo errorhandler

    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i
    Action_HighIndex = 1
    ' load tilesets we need
    LoadTilesets
            MusicFile = Trim$(Map.Music)
    If Not MusicFile = "None." Then
        PlayMusic MusicFile
    Else
        StopMusic
    End If
    WalkToX = -1
    WalkToY = -1
    ' re-position the map name
    Call UpdateDrawMapName
    Npc_HighIndex = 0
    ' Get the npc high Index
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).Num > 0 Then
            Npc_HighIndex = i
            Exit For
        End If
    Next
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    initAutotiles
    CurrentWeather = Map.Weather
    CurrentWeatherIntensity = Map.WeatherIntensity
    CurrentFog = Map.Fog
    CurrentFogSpeed = Map.FogSpeed
    CurrentFogOpacity = Map.FogOpacity
    CurrentTintR = Map.Red
    CurrentTintG = Map.Green
    CurrentTintB = Map.Blue
    CurrentTintA = Map.Alpha

    GettingMap = False
    CanMoveNow = True
    UpdateDebugCaption





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong

    With MapItem(n)
        .playerName = buffer.ReadString
        .Num = buffer.ReadLong
        .Value = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
    End With





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleItemEditor()
Dim i As Long


   On Error GoTo errorhandler

    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleItemEditor", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleAnimationEditor()
Dim i As Long


   On Error GoTo errorhandler

    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAnimationEditor", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long, i As Long, zonenum As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    zonenum = buffer.ReadLong
    n = buffer.ReadLong
    If zonenum > 0 Then
        With ZoneNPC(zonenum).Npc(n)
            .Num = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .dir = buffer.ReadLong
            .Map = GetPlayerMap(MyIndex)
            ' Client use only
            .XOffset = 0
            .YOffset = 0
            .Moving = 0
        End With
    Else
        With MapNpc(n)
            .Num = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .dir = buffer.ReadLong
            ' Client use only
            .XOffset = 0
            .YOffset = 0
            .Moving = 0
        End With
            Npc_HighIndex = 0
            ' Get the npc high Index
        For i = MAX_MAP_NPCS To 1 Step -1
            If MapNpc(i).Num > 0 Then
                Npc_HighIndex = i
                Exit For
            End If
        Next
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSpawnNpc", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer, zonenum As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    zonenum = buffer.ReadLong
    If zonenum > 0 Then
        n = buffer.ReadLong
        Call ClearZoneNpc(zonenum, n)
    Else
        n = buffer.ReadLong
        Call ClearMapNpc(n)
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleNpcDead", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleNpcEditor()
Dim i As Long


   On Error GoTo errorhandler

    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleNpcEditor", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    NpcSize = LenB(Npc(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(NpcData(0)), NpcSize
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleResourceEditor()
Dim i As Long


   On Error GoTo errorhandler

    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleResourceEditor", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ResourceNum = buffer.ReadLong
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    ClearResource ResourceNum
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    n = buffer.ReadByte
    TempTile(X, Y).DoorOpen = n




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapKey", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleEditMap()

   On Error GoTo errorhandler

    Call MapEditorInit




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEditMap", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleShopEditor()
Dim i As Long


   On Error GoTo errorhandler

    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleShopEditor", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ShopNum As Long
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ShopNum = buffer.ReadLong
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleSpellEditor()
Dim i As Long


   On Error GoTo errorhandler

    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(spell(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSpellEditor", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Spellnum As Long
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Spellnum = buffer.ReadLong
    SpellSize = LenB(spell(Spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(spell(Spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set buffer = Nothing
    ' Update the spells on the pic
    Set buffer = New clsBuffer
    buffer.WriteLong CSpells
    SendData buffer.ToArray()
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub HandleSpells(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i) = buffer.ReadLong
    Next
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Call ClearPlayer(buffer.ReadLong)
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    ' if in map editor, we cache shit ourselves

   On Error GoTo errorhandler

    If InMapEditor Then Exit Sub
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Resource_Index = buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = buffer.ReadByte
            MapResource(i).X = buffer.ReadLong
            MapResource(i).Y = buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
    Call DrawPing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleDoorAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    With TempTile(X, Y)
        .DoorFrame = 1
        .DoorAnimate = 1 ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = GetTickCount
    End With
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleDoorAnimation", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, Message As String, color As Long, tmpType As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Message = buffer.ReadString
    color = buffer.ReadLong
    tmpType = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing
    CreateActionMsg Message, color, tmpType, X, Y




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, Sprite As Long, i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing
    ' randomise sprite
    Sprite = Rand(1, BloodCount)
    ' make sure tile doesn't already have blood
    For i = 1 To MAX_BYTE
        If Blood(i).X = X And Blood(i).Y = Y Then
            ' already have blood :(
            Exit Sub
        End If
    Next
    ' carry on with the set
    BloodIndex = BloodIndex + 1
    If BloodIndex >= MAX_BYTE Then BloodIndex = 1
    With Blood(BloodIndex)
        .X = X
        .Y = Y
        .Sprite = Sprite
        .Timer = GetTickCount
    End With




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleBlood", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .LockType = buffer.ReadByte
        .lockindex = buffer.ReadLong
        .LockZone = buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With
    ' play the sound if we've got one
    PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long
Dim MapNpcNum As Byte, zonenum As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    zonenum = buffer.ReadLong
    If zonenum > 0 Then
        MapNpcNum = buffer.ReadLong
        For i = 1 To Vitals.Vital_Count - 1
            ZoneNPC(zonenum).Npc(MapNpcNum).Vital(i) = buffer.ReadLong
        Next
    Else
        MapNpcNum = buffer.ReadLong
        For i = 1 To Vitals.Vital_Count - 1
            MapNpc(MapNpcNum).Vital(i) = buffer.ReadLong
        Next
    End If
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapNpcVitals", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim slot As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    slot = buffer.ReadLong
    SpellCD(slot) = GetTickCount
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleCooldown", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SpellBuffer = 0
    SpellBufferTimer = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Access As Long
Dim Name As String
Dim Message As String
Dim colour As Long
Dim Header As String
Dim PK As Long
Dim saycolour As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Name = buffer.ReadString
    Access = buffer.ReadLong
    PK = buffer.ReadLong
    Message = buffer.ReadString
    Header = buffer.ReadString
    saycolour = buffer.ReadLong
    ' Check access level
    If PK = NO Then
        Select Case Access
            Case 0
                colour = Orange
            Case 1
                colour = DarkGrey
            Case 2
                colour = Cyan
            Case 3
                colour = BrightGreen
            Case 4
                colour = Yellow
        End Select
    Else
        colour = BrightRed
    End If
    AddText Header & Name & ": " & Message, colour
        Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ShopNum As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ShopNum = buffer.ReadLong
    Set buffer = Nothing
    OpenShop ShopNum




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    ShopAction = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleResetShopAction", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    StunDuration = buffer.ReadLong
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_BANK
        Bank.Item(i).Num = buffer.ReadLong
        Bank.Item(i).Value = buffer.ReadLong
    Next
    InBank = True
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    InTrade = buffer.ReadLong
    TradeYourWorth = 0
    TradeTheirWorth = 0
    TradeStatus = ""
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleTrade", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    InTrade = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleCloseTrade", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim dataType As Byte
Dim i As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    dataType = buffer.ReadByte
    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).Num = buffer.ReadLong
            TradeYourOffer(i).Value = buffer.ReadLong
        Next
        TradeYourWorth = buffer.ReadLong
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).Num = buffer.ReadLong
            TradeTheirOffer(i).Value = buffer.ReadLong
        Next
        TradeTheirWorth = buffer.ReadLong
    End If
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, b As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    b = buffer.ReadByte
    Set buffer = Nothing
    Select Case b
        Case 0 ' clear
            TradeStatus = vbNullString
        Case 1 ' they've accepted
             TradeStatus = "Other player has accepted."
        Case 2 ' you've accepted
             TradeStatus = "Waiting for other player to accept."
    End Select




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    myTarget = buffer.ReadLong
    myTargetType = buffer.ReadLong
    myTargetZone = buffer.ReadLong
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
        For i = 1 To MAX_HOTBAR
        Hotbar(i).slot = buffer.ReadLong
        Hotbar(i).sType = buffer.ReadByte
    Next




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player_HighIndex = buffer.ReadLong




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, entityType As Long, entityNum As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    entityType = buffer.ReadLong
    entityNum = buffer.ReadLong
    PlayMapSound X, Y, entityType, entityNum




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    theName = buffer.ReadString
    Set buffer = Nothing
    If InShop > 0 Or InTrade > 0 Or InBank > 0 Then
        Set buffer = New clsBuffer
        buffer.WriteLong CDeclineTrade
        buffer.WriteLong 1
        SendData buffer.ToArray
        Set buffer = Nothing
    Else
        dialogue "Trade Request", theName & " has requested a trade. Would you like to accept?", DIALOGUE_TYPE_TRADE, True
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleTradeRequest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    theName = buffer.ReadString
    dialogue "Party Invitation", theName & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePartyInvite", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    inParty = buffer.ReadByte
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' exit out early
        Exit Sub
    End If
    ' carry on otherwise
    Party.Leader = buffer.ReadLong
    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = buffer.ReadLong
    Next
    Party.MemberCount = buffer.ReadLong




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim playerNum As Long, partyIndex As Long
Dim buffer As clsBuffer, i As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ' which player?
    playerNum = buffer.ReadLong
    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = buffer.ReadLong
        Player(playerNum).Vital(i) = buffer.ReadLong
    Next
    ' find the party number
    For i = 1 To MAX_PARTY_MEMBERS
        If Party.Member(i) = playerNum Then
            partyIndex = i
        End If
    Next
    ' exit out if wrong data
    If partyIndex <= 0 Or partyIndex > MAX_PARTY_MEMBERS Then Exit Sub





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePartyVitals", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleSpawnEventPage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim id As Long, i As Long, z As Long, X As Long, Y As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    id = buffer.ReadLong
    If id > Map.CurrentEvents Then
        Map.CurrentEvents = id
        ReDim Preserve Map.MapEvents(Map.CurrentEvents)
    End If

    With Map.MapEvents(id)
        .Name = buffer.ReadString
        .dir = buffer.ReadLong
        .ShowDir = .dir
        .GraphicNum = buffer.ReadLong
        .GraphicType = buffer.ReadLong
        .GraphicX = buffer.ReadLong
        .GraphicX2 = buffer.ReadLong
        .GraphicY = buffer.ReadLong
        .GraphicY2 = buffer.ReadLong
        .MovementSpeed = buffer.ReadLong
        .Moving = 0
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .XOffset = 0
        .YOffset = 0
        .Position = buffer.ReadLong
        .Visible = buffer.ReadLong
        .WalkAnim = buffer.ReadLong
        .DirFix = buffer.ReadLong
        .WalkThrough = buffer.ReadLong
        .ShowName = buffer.ReadLong
        .questnum = buffer.ReadLong
    End With
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSpawnEventPage", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleEventMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim id As Long
Dim X As Long
Dim Y As Long
Dim dir As Long, ShowDir As Long
Dim Movement As Long, MovementSpeed As Long
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    id = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    dir = buffer.ReadLong
    ShowDir = buffer.ReadLong
    MovementSpeed = buffer.ReadLong
    If id > Map.CurrentEvents Then Exit Sub

    With Map.MapEvents(id)
        .X = X
        .Y = Y
        .dir = dir
        .XOffset = 0
        .YOffset = 0
        .Moving = 1
        .ShowDir = ShowDir
        .MovementSpeed = MovementSpeed
    
        Select Case dir
            Case DIR_UP
                .YOffset = PIC_Y
            Case DIR_DOWN
                .YOffset = PIC_Y * -1
            Case DIR_LEFT
                .XOffset = PIC_X
            Case DIR_RIGHT
                .XOffset = PIC_X * -1
        End Select

    End With





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEventMove", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleEventDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim dir As Byte
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    dir = buffer.ReadLong
    If i > Map.CurrentEvents Then Exit Sub

    With Map.MapEvents(i)
        .dir = dir
        .ShowDir = dir
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEventDir", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleEventChat(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim dir As Byte
Dim buffer As clsBuffer
Dim choices As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    EventReplyID = buffer.ReadLong
    EventReplyPage = buffer.ReadLong
    EventChatFace = buffer.ReadLong
    EventText = buffer.ReadString
    If EventText = "" Then EventText = " "
    EventChat = True
    ShowEventLbl = True
    choices = buffer.ReadLong
    InEvent = True
    For i = 1 To 4
        EventChoices(i) = ""
        EventChoiceVisible(i) = False
    Next
    EventChatType = 0
    If choices = 0 Then
        Else
        EventChatType = 1
        For i = 1 To choices
            EventChoices(i) = buffer.ReadString
            EventChoiceVisible(i) = True
        Next
    End If
    AnotherChat = buffer.ReadLong
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEventChat", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleEventStart(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    InEvent = True





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEventStart", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleEventEnd(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    InEvent = False





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEventEnd", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayBGM(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    str = buffer.ReadString
    StopMusic
    PlayMusic str
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayBGM", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlaySound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    str = buffer.ReadString

    PlaySound str, -1, -1
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlaySound", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleFadeoutBGM(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    'Need to learn how to fadeout :P
    'do later... way later.. like, after release, maybe never

   On Error GoTo errorhandler

    StopMusic




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleStopSound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, i As Long


   On Error GoTo errorhandler

    For i = 0 To UBound(Sounds()) - 1
        StopSound (i)
    Next




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleStopSound", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, i As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_SWITCHES
        Switches(i) = buffer.ReadString
    Next
    For i = 1 To MAX_VARIABLES
        Variables(i) = buffer.ReadString
    Next
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleMapEventData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, i As Long, X As Long, Y As Long, z As Long, w As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    'Event Data!
    Map.EventCount = buffer.ReadLong
    If Map.EventCount > 0 Then
        ReDim Map.Events(0 To Map.EventCount)
        For i = 1 To Map.EventCount
            With Map.Events(i)
                .Name = buffer.ReadString
                .Global = buffer.ReadLong
                .X = buffer.ReadLong
                .Y = buffer.ReadLong
                .pageCount = buffer.ReadLong
            End With
            If Map.Events(i).pageCount > 0 Then
                ReDim Map.Events(i).Pages(0 To Map.Events(i).pageCount)
                For X = 1 To Map.Events(i).pageCount
                    With Map.Events(i).Pages(X)
                        .chkVariable = buffer.ReadLong
                        .VariableIndex = buffer.ReadLong
                        .VariableCondition = buffer.ReadLong
                        .VariableCompare = buffer.ReadLong
                                                .chkSwitch = buffer.ReadLong
                        .SwitchIndex = buffer.ReadLong
                        .SwitchCompare = buffer.ReadLong
                                                .chkHasItem = buffer.ReadLong
                        .HasItemIndex = buffer.ReadLong
                        .HasItemAmount = buffer.ReadLong
                                                .chkSelfSwitch = buffer.ReadLong
                        .SelfSwitchIndex = buffer.ReadLong
                        .SelfSwitchCompare = buffer.ReadLong
                                                .GraphicType = buffer.ReadLong
                        .Graphic = buffer.ReadLong
                        .GraphicX = buffer.ReadLong
                        .GraphicY = buffer.ReadLong
                        .GraphicX2 = buffer.ReadLong
                        .GraphicY2 = buffer.ReadLong
                                                .MoveType = buffer.ReadLong
                        .MoveSpeed = buffer.ReadLong
                        .MoveFreq = buffer.ReadLong
                                                .MoveRouteCount = buffer.ReadLong
                                            .IgnoreMoveRoute = buffer.ReadLong
                        .RepeatMoveRoute = buffer.ReadLong
                                                If .MoveRouteCount > 0 Then
                            ReDim Map.Events(i).Pages(X).MoveRoute(0 To .MoveRouteCount)
                            For Y = 1 To .MoveRouteCount
                                .MoveRoute(Y).Index = buffer.ReadLong
                                .MoveRoute(Y).Data1 = buffer.ReadLong
                                .MoveRoute(Y).data2 = buffer.ReadLong
                                .MoveRoute(Y).Data3 = buffer.ReadLong
                                .MoveRoute(Y).Data4 = buffer.ReadLong
                                .MoveRoute(Y).Data5 = buffer.ReadLong
                                .MoveRoute(Y).Data6 = buffer.ReadLong
                            Next
                        End If
                                                .WalkAnim = buffer.ReadLong
                        .DirFix = buffer.ReadLong
                        .WalkThrough = buffer.ReadLong
                        .ShowName = buffer.ReadLong
                        .Trigger = buffer.ReadLong
                        .CommandListCount = buffer.ReadLong
                                                .Position = buffer.ReadLong
                        .questnum = buffer.ReadLong
                    End With
                                        If Map.Events(i).Pages(X).CommandListCount > 0 Then
                        ReDim Map.Events(i).Pages(X).CommandList(0 To Map.Events(i).Pages(X).CommandListCount)
                        For Y = 1 To Map.Events(i).Pages(X).CommandListCount
                            Map.Events(i).Pages(X).CommandList(Y).CommandCount = buffer.ReadLong
                            Map.Events(i).Pages(X).CommandList(Y).ParentList = buffer.ReadLong
                            If Map.Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                ReDim Map.Events(i).Pages(X).CommandList(Y).Commands(1 To Map.Events(i).Pages(X).CommandList(Y).CommandCount)
                                For z = 1 To Map.Events(i).Pages(X).CommandList(Y).CommandCount
                                    With Map.Events(i).Pages(X).CommandList(Y).Commands(z)
                                        .Index = buffer.ReadLong
                                        .Text1 = buffer.ReadString
                                        .Text2 = buffer.ReadString
                                        .Text3 = buffer.ReadString
                                        .Text4 = buffer.ReadString
                                        .Text5 = buffer.ReadString
                                        .Data1 = buffer.ReadLong
                                        .data2 = buffer.ReadLong
                                        .Data3 = buffer.ReadLong
                                        .Data4 = buffer.ReadLong
                                        .Data5 = buffer.ReadLong
                                        .Data6 = buffer.ReadLong
                                        .ConditionalBranch.CommandList = buffer.ReadLong
                                        .ConditionalBranch.Condition = buffer.ReadLong
                                        .ConditionalBranch.Data1 = buffer.ReadLong
                                        .ConditionalBranch.data2 = buffer.ReadLong
                                        .ConditionalBranch.Data3 = buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = buffer.ReadLong
                                        .MoveRouteCount = buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For w = 1 To .MoveRouteCount
                                                .MoveRoute(w).Index = buffer.ReadLong
                                                .MoveRoute(w).Data1 = buffer.ReadLong
                                                .MoveRoute(w).data2 = buffer.ReadLong
                                                .MoveRoute(w).Data3 = buffer.ReadLong
                                                .MoveRoute(w).Data4 = buffer.ReadLong
                                                .MoveRoute(w).Data5 = buffer.ReadLong
                                                .MoveRoute(w).Data6 = buffer.ReadLong
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
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapEventData", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, targetType As Long, target As Long, Message As String, colour As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    target = buffer.ReadLong
    targetType = buffer.ReadLong
    Message = buffer.ReadString
    colour = buffer.ReadLong
    AddChatBubble target, targetType, Message, colour
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleChatBubble", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleSpecialEffect(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, effectType As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    effectType = buffer.ReadLong
    Select Case effectType
        Case EFFECT_TYPE_FADEIN
            FadeType = 1
            FadeAmount = 0
        Case EFFECT_TYPE_FADEOUT
            FadeType = 0
            FadeAmount = 255
        Case EFFECT_TYPE_FLASH
            FlashTimer = GetTickCount + 150
        Case EFFECT_TYPE_FOG
            CurrentFog = buffer.ReadLong
            CurrentFogSpeed = buffer.ReadLong
            CurrentFogOpacity = buffer.ReadLong
        Case EFFECT_TYPE_WEATHER
            CurrentWeather = buffer.ReadLong
            CurrentWeatherIntensity = buffer.ReadLong
        Case EFFECT_TYPE_TINT
            CurrentTintR = buffer.ReadLong
            CurrentTintG = buffer.ReadLong
            CurrentTintB = buffer.ReadLong
            CurrentTintA = buffer.ReadLong
    End Select
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSpecialEffect", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleHouseConfigurations(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_HOUSES
        HouseConfig(i).ConfigName = buffer.ReadString
        HouseConfig(i).BaseMap = buffer.ReadLong
        HouseConfig(i).MaxFurniture = buffer.ReadLong
        HouseConfig(i).Price = buffer.ReadLong
    Next
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleHouseConfigurations", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleHouseOffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    If HouseConfig(i).MaxFurniture > 0 Then
        dialogue "Buy House?", "Would you like to buy the house: " & Trim$(HouseConfig(i).ConfigName) & vbNewLine & "Cost: " & CStr(HouseConfig(i).Price) & vbNewLine & "Furniture Limit: " & CStr(HouseConfig(i).MaxFurniture), DIALOGUE_TYPE_BUYHOUSE, True, i
    Else
        dialogue "Buy House?", "Would you like to buy the house: " & Trim$(HouseConfig(i).ConfigName) & vbNewLine & "Cost: " & CStr(HouseConfig(i).Price) & vbNewLine & "Furniture Limit: None.", DIALOGUE_TYPE_BUYHOUSE, True, i
    End If
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleHouseOffer", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleMaxes(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MAX_MAPS = buffer.ReadLong
    MAX_LEVELS = buffer.ReadLong
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMaxes", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleVisit(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    dialogue "Visitation Invitation", "You have been invited to visit " & Trim$(GetPlayerName(i)) & "'s house.", DIALOGUE_TYPE_VISIT, True, i
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleVisit", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleFurniture(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    FurnitureHouse = buffer.ReadLong
    FurnitureCount = buffer.ReadLong
    ReDim Furniture(FurnitureCount)
    If FurnitureCount > 0 Then
        For i = 1 To FurnitureCount
            Furniture(i).ItemNum = buffer.ReadLong
            Furniture(i).X = buffer.ReadLong
            Furniture(i).Y = buffer.ReadLong
        Next
    End If
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleFurniture", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleUnreadMail(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, count As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    count = buffer.ReadLong
    If count > LastMailCount Then
        LastMailCount = count
        PlaySound "Success2.wav", -1, -1
    Else
        LastMailCount = count
    End If
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUnreadMail", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleMailBox(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, count As Long, i As Long, openmailbox As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    count = buffer.ReadLong
    openmailbox = buffer.ReadLong
    If count > 0 Then
        ReDim Mail(count)
        For i = 1 To count
            Mail(i - 1).Index = buffer.ReadLong
            Mail(i - 1).Unread = buffer.ReadLong
            Mail(i - 1).From = buffer.ReadString
            Mail(i - 1).Body = buffer.ReadString
            Mail(i - 1).ItemNum = buffer.ReadLong
            Mail(i - 1).ItemVal = buffer.ReadLong
            Mail(i - 1).Date = buffer.ReadString
        Next
    Else
        ReDim Mail(0)
    End If
    MailCount = count
    For i = 0 To UBound(Mail) - 1
        If Mail(i).Unread = 1 Then
            'frmMain.lstInbox.AddItem "[UNREAD] " & "Letter from " & Mail(i).From & " on " & Mail(i).Date
        Else
            'frmMain.lstInbox.AddItem "Letter from " & Mail(i).From & " on " & Mail(i).Date
        End If
    Next
    If openmailbox = 1 Then
        InMailbox = True
        MailBoxMenu = 0
    End If
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMailBox", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleSelectChar(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, numchars As Long, CharName As String, Level As Long, Class As String, X As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    numchars = buffer.ReadLong
    ReDim CharSelection(numchars)
    SelectedChar = 1
    'frmMain.Visible = True
    ' save options
    Servers(ServerIndex).Username = Trim$(TxtUsername)

    If Servers(ServerIndex).SavePass = 0 Then
        Servers(ServerIndex).Password = vbNullString
    Else
        Servers(ServerIndex).Password = Trim$(txtPassword)
    End If
    SaveServers
    TxtUsername = ""
    txtPassword = ""
    MenuStage = 4
    For i = 1 To numchars
        CharName = Trim$(buffer.ReadString)
        Level = buffer.ReadLong
        Class = Trim$(buffer.ReadString)
            CharSelection(i).Name = Trim$(CharName)
        CharSelection(i).Level = Level
        CharSelection(i).Class = Trim$(Class)
        CharSelection(i).Sex = buffer.ReadLong
            For X = 1 To FaceEnum.Face_Count - 1
            CharSelection(i).Face(X) = buffer.ReadLong
        Next
            If Trim$(CharName) = "" Then
            CharSelection(i).Name = "Free Character Slot"
            CharSelection(i).Class = ""
        Else
            CharName = CharName & " [Lv. " & CStr(Level) & "]"
            CharSelection(i).Name = CharName
        End If
            CharSelection(i).Name = CStr(i) & ". " & CharSelection(i).Name

            If CharName <> "" Then
            'frmMenu.lstCharacters.AddItem (CStr(i) & ". " & Trim$(charname) & " a level " & CStr(level) & " " & Trim$(Class) & ".")
        Else
            'frmMenu.lstCharacters.AddItem (CStr(i) & ". Empty Slot.")
        End If
    Next
    frmLoad.Visible = False
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSelectChar", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleFriends(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, tempname As String, blankname As String * ACCOUNT_LENGTH, X As Long, Y As Long, z As Long
Dim tempOnlineFriends(1 To 25) As String, tempOfflineFriends(1 To 25) As String
Dim tempOnlineIndex(1 To 25) As Long, tempOfflineIndex(1 To 25) As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To 25
        FriendsList(i) = ""
        FriendIndex(i) = 0
        FriendOnline(i) = 0
    Next
    X = 1
    Y = 1
    For i = 1 To 25
        tempname = Trim$(buffer.ReadString)
        If Trim$(tempname) = "" Or Trim$(tempname) = Trim$(blankname) Then
            buffer.ReadLong
        Else
            If buffer.ReadLong = 1 Then
                tempOnlineFriends(X) = Trim$(tempname)
                tempOnlineIndex(X) = i
                X = X + 1
            Else
                tempOfflineFriends(Y) = Trim$(tempname)
                tempOfflineIndex(Y) = i
                Y = Y + 1
            End If
        End If
    Next
    z = 1
    For X = 1 To X
        If Trim$(tempOnlineFriends(X)) <> "" Then
            FriendsList(z) = tempOnlineFriends(X)
            FriendIndex(z) = tempOnlineIndex(X)
            FriendOnline(z) = 1
            z = z + 1
        End If
    Next
    For Y = 1 To Y
        If Trim$(tempOfflineFriends(Y)) <> "" Then
            FriendsList(z) = tempOfflineFriends(Y)
            FriendIndex(z) = tempOfflineIndex(Y)
            FriendOnline(z) = 0
            z = z + 1
        End If
    Next
    FriendCount = z
    FriendListScroll = 0
    FriendSelection = 0
    Set buffer = Nothing




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleFriends", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Private Sub HandleQuestEditor()
Dim i As Long


   On Error GoTo errorhandler

With frmEditor_Quest
    Editor = EDITOR_TASKS
    .lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_QUESTS
        .lstIndex.AddItem i & ": " & Trim$(quest(i).Name)
    Next
    .Show
    .lstIndex.ListIndex = 0
    QuestEditorInit
End With


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleQuestEditor", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleUpdateQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the Quest
    QuestSize = LenB(quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUpdateQuest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandlePlayerQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long, X As Long


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For i = 1 To MAX_QUESTS
        Player(MyIndex).PlayerQuest(i).state = buffer.ReadLong
        Player(MyIndex).PlayerQuest(i).CurrentTask = buffer.ReadLong
        For X = 1 To 5
            Player(MyIndex).PlayerQuest(i).TaskCount(X) = buffer.ReadLong
        Next
    Next
    
    RefreshQuestLog
    
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerQuest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleQuestMessage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long, questnum As Long, QuestNumForStart As Long
Dim Message As String


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    questnum = buffer.ReadLong
    If buffer.ReadLong = 0 Then
        Message = Trim$(buffer.ReadString)
        QuestNumForStart = buffer.ReadLong
        InQuestLog = True
        QuestLogMessage = Message
        QuestLogQuest = questnum
        If QuestNumForStart > 0 And QuestNumForStart <= MAX_QUESTS Then
            QuestLogFunction = 0
        Else
            QuestLogFunction = 1
        End If
    Else
        Message = Trim$(buffer.ReadString)
        QuestNumForStart = buffer.ReadLong
        InQuestLog = True
        QuestLogMessage = Message
        QuestLogQuest = questnum
        QuestLogFunction = 2
    End If
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleQuestMessage", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleEditZones(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long, X As Long



   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_ZONES
        MapZones(i).Name = Trim$(buffer.ReadString)
        MapZones(i).MapCount = buffer.ReadLong
        If MapZones(i).MapCount > 0 Then
            ReDim MapZones(i).Maps(MapZones(i).MapCount)
            For X = 1 To MapZones(i).MapCount
                MapZones(i).Maps(X) = buffer.ReadLong
            Next
        End If
        For X = 1 To MAX_MAP_NPCS * 2
            MapZones(i).NPCs(X) = buffer.ReadLong
        Next
        For X = 1 To 5
            MapZones(i).Weather(X) = buffer.ReadByte
        Next
        MapZones(i).WeatherIntensity = buffer.ReadByte
    Next
    With frmEditor_Zone
        Editor = EDITOR_ZONE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ZONES
            .lstIndex.AddItem i & ": " & Trim$(MapZones(i).Name)
        Next
            .cmbNpc.Clear
        .cmbNpc.AddItem "None"
        For i = 1 To MAX_NPCS
            .cmbNpc.AddItem CStr(i) & ". " & Trim$(Npc(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ZoneEditorInit
    End With
    'ZoneEditorInit
    ZoneEditorInit




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEditZones", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Private Sub HandleServerInfo(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Servers(ServerIndex).Game_Name = buffer.ReadString
    News = buffer.ReadString
    Credits = buffer.ReadString
    CharMode = buffer.ReadLong
    If buffer.ReadByte = 255 Then
        IsGold = True
    Else
        IsGold = False
    End If
    ServerDir = buffer.ReadString
    UpdateURL = buffer.ReadString
    If buffer.ReadLong <> App.Major Or buffer.ReadLong <> App.Minor Or buffer.ReadLong <> App.Revision Then
        MsgBox "Your client or the game server is outdated! Be sure you are launching the client from the Eclipse Origins Launcher and not eo.exe!"
        End
    End If
    Options.MenuMusic = buffer.ReadString
    If IsGold Then
        MAX_PLAYERS = 500
        MAX_ITEMS = 1000
        MAX_NPCS = 1000
        MAX_ANIMATIONS = 500
        MAX_SHOPS = 500
        MAX_SPELLS = 1000
        MAX_RESOURCES = 500
        MAX_ZONES = 255
        MAX_HOUSES = 100
        MAX_QUESTS = 250
        MAX_PETS = 1000
    Else
        MAX_PLAYERS = 10
        MAX_ITEMS = 255
        MAX_NPCS = 255
        MAX_ANIMATIONS = 255
        MAX_SHOPS = 255
        MAX_SPELLS = 255
        MAX_RESOURCES = 255
        MAX_ZONES = 1
        MAX_HOUSES = 1
        MAX_QUESTS = 5
        MAX_PETS = 1
    End If
    ReDim Player(0 To MAX_PLAYERS)
    ReDim Item(1 To MAX_ITEMS)
    ReDim Npc(1 To MAX_ITEMS)
    ReDim Shop(1 To MAX_SHOPS)
    ReDim spell(1 To MAX_SPELLS)
    ReDim Resource(1 To MAX_RESOURCES)
    ReDim Animation(1 To MAX_ANIMATIONS)
    ReDim HouseConfig(1 To MAX_HOUSES)
    ReDim MapZones(1 To MAX_ZONES)
    ReDim ZoneNPC(1 To MAX_ZONES)
    ReDim House(1 To MAX_HOUSES)
    ReDim Pet(1 To MAX_PETS)
    
    ReDim Item_Changed(1 To MAX_ITEMS)
    ReDim NPC_Changed(1 To MAX_NPCS)
    ReDim Resource_Changed(1 To MAX_RESOURCES)
    ReDim Animation_Changed(1 To MAX_ANIMATIONS)
    ReDim Spell_Changed(1 To MAX_SPELLS)
    ReDim Shop_Changed(1 To MAX_SHOPS)
    ReDim Zone_Changed(1 To MAX_ZONES)
    ReDim House_Changed(1 To MAX_HOUSES)
    ReDim Pet_Changed(1 To MAX_PETS)
    ReDim RandomDungeon_Changed(1 To MAX_RANDOMDUNGEONS)
    
    GotServerInfo = True
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleServerInfo", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleAdmin(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    If Player(MyIndex).Access > 0 Then
        frmAdmin.Visible = True
        If frmAdmin.Visible Then
            If Options.FullScreen = 0 Then
                frmAdmin.Left = frmMain.Left + frmMain.Width
                frmAdmin.Top = frmMain.Top
                frmAdmin.tabLists.Tabs(1).Selected = True
            Else
                frmAdmin.Left = 0
                frmAdmin.Top = 0
                frmAdmin.tabLists.Tabs(1).Selected = True
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAdmin", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePlayers(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, X As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    frmAdmin.lstPlayers(1).ListItems.Clear
    For i = 1 To X
        frmAdmin.lstPlayers(1).ListItems.Add (i)

        If i < 10 Then
            frmAdmin.lstPlayers(1).ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmAdmin.lstPlayers(1).ListItems(i).Text = "0" & i
        Else
            frmAdmin.lstPlayers(1).ListItems(i).Text = i
        End If
        
        If buffer.ReadLong > 0 Then
            frmAdmin.lstPlayers(1).ListItems(i).SubItems(1) = Trim$(buffer.ReadString)
            frmAdmin.lstPlayers(1).ListItems(i).SubItems(2) = Trim$(buffer.ReadString)
            frmAdmin.lstPlayers(1).ListItems(i).SubItems(3) = Trim$(buffer.ReadString)
            frmAdmin.lstPlayers(1).ListItems(i).SubItems(4) = CStr(buffer.ReadLong)
            frmAdmin.lstPlayers(1).ListItems(i).SubItems(5) = CStr(buffer.ReadLong)
        Else
            frmAdmin.lstPlayers(1).ListItems(i).SubItems(1) = vbNullString
            frmAdmin.lstPlayers(1).ListItems(i).SubItems(2) = vbNullString
            frmAdmin.lstPlayers(1).ListItems(i).SubItems(3) = vbNullString
            frmAdmin.lstPlayers(1).ListItems(i).SubItems(4) = vbNullString
            frmAdmin.lstPlayers(1).ListItems(i).SubItems(5) = vbNullString
        End If
    Next
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayers", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Private Sub HandleAccounts(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, X As Long, Y As Long, z As Long, chars As String, n As String, i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    MAX_PLAYER_CHARS = buffer.ReadLong
    frmAdmin.lstPlayers(2).ListItems.Clear
    If X > 0 Then
        For i = 1 To X
            frmAdmin.lstPlayers(2).ListItems.Add (i)
    
            If i < 10 Then
                frmAdmin.lstPlayers(2).ListItems(i).Text = "00" & i
            ElseIf i < 100 Then
                frmAdmin.lstPlayers(2).ListItems(i).Text = "0" & i
            Else
                frmAdmin.lstPlayers(2).ListItems(i).Text = i
            End If
            chars = ""
            
            frmAdmin.lstPlayers(2).ListItems(i).SubItems(1) = Trim$(buffer.ReadString)
            frmAdmin.lstPlayers(2).ListItems(i).Text = Trim$(buffer.ReadString)
            
            If buffer.ReadLong = 0 Then
                frmAdmin.lstPlayers(2).ListItems(i).SubItems(3) = "No"
            Else
                frmAdmin.lstPlayers(2).ListItems(i).SubItems(3) = "Yes"
            End If
            
            For z = 1 To MAX_PLAYER_CHARS
                n = Trim$(buffer.ReadString)
                If n <> "" Then
                    chars = chars & n & ", "
                End If
            Next
            If chars <> "" Then
                chars = Left(chars, Len(chars) - 2)
                frmAdmin.lstPlayers(2).ListItems(i).SubItems(2) = chars
            End If
        Next
    End If
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAccounts", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleMaps(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, X As Long, Name As String

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    frmAdmin.lstMaps.Clear
    For i = 1 To X
        Name = buffer.ReadString
        If Trim$(Name) = "" Then Name = "Unnamed map."
        frmAdmin.lstMaps.AddItem i & ". " & (Trim$(Name))
    Next
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMaps", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleBans(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, X As Long, Name As String

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    BanCount = X
    ReDim Bans(BanCount)
    frmAdmin.lstPlayers(3).ListItems.Clear
    If BanCount > 0 Then
        For i = 1 To X
            frmAdmin.lstPlayers(3).ListItems.Add (i)
    
            If i < 10 Then
                frmAdmin.lstPlayers(3).ListItems(i).Text = "00" & i
            ElseIf i < 100 Then
                frmAdmin.lstPlayers(3).ListItems(i).Text = "0" & i
            Else
                frmAdmin.lstPlayers(3).ListItems(i).Text = i
            End If
            
            Bans(i).IPAddress = Trim$(buffer.ReadString)
            Bans(i).BanName = Trim$(buffer.ReadString)
            Bans(i).BanReason = Trim$(buffer.ReadString)
            Bans(i).BanChar = Trim$(buffer.ReadString)
            
            frmAdmin.lstPlayers(3).ListItems(i).SubItems(1) = Bans(i).IPAddress
            frmAdmin.lstPlayers(3).ListItems(i).SubItems(2) = Bans(i).BanName
            frmAdmin.lstPlayers(3).ListItems(i).SubItems(3) = Bans(i).BanChar
            frmAdmin.lstPlayers(3).ListItems(i).SubItems(4) = Bans(i).BanReason
            
        Next
    End If
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleBans", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleServerOpts(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, X As Long, Name As String

   On Error GoTo errorhandler
   
   If frmAdmin.Visible = False Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    frmAdmin.picSAdmin.Visible = True
    frmAdmin.picSAdmin.ZOrder 0
    frmAdmin.txtNews.Text = buffer.ReadString
    frmAdmin.txtCredits.Text = buffer.ReadString
    frmAdmin.txtMOTD.Text = buffer.ReadString
    frmAdmin.txtGameName.Text = buffer.ReadString
    frmAdmin.txtGameWebsite.Text = buffer.ReadString
    frmAdmin.txtDataFolder.Text = buffer.ReadString
    frmAdmin.txtUpdateURL.Text = buffer.ReadString
    frmAdmin.lblAccounts.Caption = "Accounts: " & CStr(buffer.ReadLong)
    frmAdmin.lblPlayersOnline.Caption = "Online Players: " & CStr(buffer.ReadLong)
    frmAdmin.lblUpTime.Caption = "Uptime: " & uptimeToDHMS(buffer.ReadLong / 1000)
    frmAdmin.lblVersion.Caption = "Version: " & buffer.ReadString
    If buffer.ReadLong = 1 Then
        frmAdmin.chkStaffOnly.Value = 1
    Else
        frmAdmin.chkStaffOnly.Value = 0
    End If
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleServerOpts", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleEditHouses(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long, X As Long



   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_HOUSES
        With House(i)
            .ConfigName = Trim$(buffer.ReadString)
            .BaseMap = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .Price = buffer.ReadLong
            .MaxFurniture = buffer.ReadLong
        End With
    Next
    
    With frmEditor_House
        Editor = EDITOR_HOUSE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_HOUSES
            .lstIndex.AddItem i & ": " & Trim$(House(i).ConfigName)
        Next

        .Show
        .lstIndex.ListIndex = 0
    End With
    
    HouseEditorInit

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEditHouses", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleEditPlayer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long, X As Long, char As String



   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    Select Case i
        Case 0
            'Ready to edit player data
            frmAdmin.EditingAccount = buffer.ReadString
            frmAdmin.EditingPlayer = buffer.ReadString
            frmAdmin.picEditPlayer.ZOrder 0
            frmAdmin.picEditPlayer.Visible = True
            frmAdmin.fraCharList.Visible = False
            frmAdmin.fraEditPlayer.Visible = True
            
        Case 1
            frmAdmin.EditingAccount = buffer.ReadString
            frmAdmin.picEditPlayer.ZOrder 0
            frmAdmin.picEditPlayer.Visible = True
            frmAdmin.fraCharList.Visible = True
            frmAdmin.fraEditPlayer.Visible = False
            frmAdmin.lstChars.Clear
            For X = 1 To MAX_PLAYER_CHARS
                char = Trim$(buffer.ReadString)
                If char = "" Then
                    frmAdmin.lstChars.AddItem X & ". No Character"
                Else
                    frmAdmin.lstChars.AddItem X & ". " & char
                End If
            Next
            Set buffer = Nothing
            Exit Sub
        Case 2
            'Ready to edit player data
            frmAdmin.EditingAccount = buffer.ReadString
            frmAdmin.EditingPlayer = buffer.ReadString
            frmAdmin.picEditPlayer.ZOrder 0
            frmAdmin.picEditPlayer.Visible = True
            frmAdmin.fraCharList.Visible = False
            frmAdmin.fraEditPlayer.Visible = True
        Case 3
            'Ready to edit player data
            frmAdmin.EditingAccount = buffer.ReadString
            frmAdmin.EditingPlayer = buffer.ReadString
            frmAdmin.picEditPlayer.ZOrder 0
            frmAdmin.picEditPlayer.Visible = True
            frmAdmin.fraCharList.Visible = False
            frmAdmin.fraEditPlayer.Visible = True
    End Select
    
    'If still here then we are loading up the editor :D
    'Collect Player data first
    i = 0
    Call SetPlayerName(i, buffer.ReadString)
    Call SetPlayerLevel(i, buffer.ReadLong)
    Call SetPlayerPOINTS(i, buffer.ReadLong)
    For X = 1 To FaceEnum.Face_Count - 1
        Player(i).Face(X) = buffer.ReadLong
    Next
    For X = 1 To SpriteEnum.Sprite_Count - 1
        Player(i).Sprite(X) = buffer.ReadLong
    Next
    Player(i).Sex = buffer.ReadLong
    Call SetPlayerMap(i, buffer.ReadLong)
    Call SetPlayerX(i, buffer.ReadLong)
    Call SetPlayerY(i, buffer.ReadLong)
    Call SetPlayerDir(i, buffer.ReadLong)
    Call SetPlayerAccess(i, buffer.ReadLong)
    Call SetPlayerPK(i, buffer.ReadLong)
    Call SetPlayerClass(i, buffer.ReadLong)
    Player(i).InHouse = buffer.ReadLong
    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, buffer.ReadLong
    Next
    SetPlayerExp i, buffer.ReadLong
    Player(i).Vital(Vitals.HP) = buffer.ReadLong
    Player(i).Vital(Vitals.MP) = buffer.ReadLong
    For X = 1 To Equipment.Equipment_Count - 1
        SetPlayerEquipment i, buffer.ReadLong, X
    Next
    For X = 1 To MAX_PLAYER_SPELLS
        TempPlayerSpells(X) = buffer.ReadLong
    Next
    For X = 1 To MAX_INV
        TempPlayerInv(X).Num = buffer.ReadLong
        TempPlayerInv(X).Value = buffer.ReadLong
    Next
    
    InitPlayerEditor

    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEditPlayer", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePicture(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    If buffer.ReadLong = 0 Then
        i = buffer.ReadLong
        With Pictures(i)
            .pic = buffer.ReadLong
            .type = buffer.ReadLong
            .XOffset = buffer.ReadLong
            .YOffset = buffer.ReadLong
        End With
    Else
        i = buffer.ReadLong
        Pictures(i).pic = 0
    End If
    
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePicture", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Private Sub HandleHoldPlayer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    If buffer.ReadLong = 0 Then
        HoldPlayer = True
    Else
        HoldPlayer = False
    End If
    
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleHoldPlayer", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleGameOpts(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, X As Long, Name As String

   On Error GoTo errorhandler
   
   If frmAdmin.Visible = False Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    frmAdmin.picGAdmin.Visible = True
    frmAdmin.picGAdmin.ZOrder 0
    If buffer.ReadLong = 1 Then
        frmAdmin.chkNewBatForumlas.Value = 1
    Else
        frmAdmin.chkNewBatForumlas.Value = 0
    End If
    frmAdmin.scrlMaxLevel.Value = buffer.ReadLong
    frmAdmin.txtMainMenuMusic = buffer.ReadString
    frmAdmin.chkDisableItemLoss.Value = buffer.ReadLong
    frmAdmin.chkDisableExpLoss.Value = buffer.ReadLong
    Set buffer = Nothing

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleGameOpts", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePetEditor()
Dim i As Long

    

   On Error GoTo errorhandler

    With frmEditor_Pet
        Editor = EDITOR_PET
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_PETS
            .lstIndex.AddItem i & ": " & Trim$(Pet(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        PetEditorInit
    End With
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePetEditor", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Private Sub HandleUpdatePet(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long, i As Long
Dim buffer As clsBuffer
Dim PetSize As Long
Dim PetData() As Byte
    
    

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    n = buffer.ReadLong
    
    With Pet(n)
        .Num = buffer.ReadLong
        .Name = buffer.ReadString
        .Sprite = buffer.ReadLong
        .Range = buffer.ReadLong
        .Level = buffer.ReadLong
        .MaxLevel = buffer.ReadLong
        .ExpGain = buffer.ReadLong
        .LevelPnts = buffer.ReadLong
        .StatType = buffer.ReadByte
        .LevelingType = buffer.ReadByte
        For i = 1 To Stats.Stat_Count - 1
            .stat(i) = buffer.ReadByte
        Next
        For i = 1 To 4
            .spell(i) = buffer.ReadLong
        Next
    End With

    
    Set buffer = Nothing
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUpdatePet", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePetMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim dir As Long
Dim Movement As Long
Dim buffer As clsBuffer

    

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    dir = buffer.ReadLong
    Movement = buffer.ReadLong

    With Player(i).Pet
        .X = X
        .Y = Y
        .dir = dir
        .XOffset = 0
        .YOffset = 0
        .Moving = Movement

        Select Case .dir
            Case DIR_UP
                .YOffset = PIC_Y
            Case DIR_DOWN
                .YOffset = PIC_Y * -1
            Case DIR_LEFT
                .XOffset = PIC_X
            Case DIR_RIGHT
                .XOffset = PIC_X * -1
        End Select

    End With


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePetMove", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandlePetDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim dir As Long
Dim buffer As clsBuffer

    

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    dir = buffer.ReadLong

    Player(i).Pet.dir = dir



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePetDir", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub


Private Sub HandlePetVital(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim dir As Long
Dim buffer As clsBuffer

    

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    
    If buffer.ReadLong = 1 Then
        Player(i).Pet.MaxHp = buffer.ReadLong
        Player(i).Pet.Health = buffer.ReadLong
    Else
        Player(i).Pet.MaxMP = buffer.ReadLong
        Player(i).Pet.Mana = buffer.ReadLong
    End If


    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePetVital", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleClearPetSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    

   On Error GoTo errorhandler

    PetSpellBuffer = 0
    PetSpellBufferTimer = 0
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleClearPetSpellBuffer", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleRandomDungeonEditor()
    Dim i As Long


   On Error GoTo errorhandler

   'Removed Sorry


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRandomDungeonEditor", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleUpdateRandomDungeon(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim DungeonNum As Long
    Dim buffer As clsBuffer
    Dim DungeonSize As Long
    Dim DungeonData() As Byte
    

   On Error GoTo errorhandler

    'Removed


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUpdateRandomDungeon", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleRandomDungeonMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim MaxRooms As Long, i As Long
Dim StairsX As Long, StairsY As Long
Dim DungeonNum As Long
Dim buffer As clsBuffer
Dim RoomX As Long
Dim RoomY As Long
Dim RoomWidth As Long
Dim RoomHeight As Long
Dim MapNum As Long
Dim X As Long
Dim Y As Long


   On Error GoTo errorhandler
    'Remove


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRandomDungeonMap", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleProjectileEditor()
    Dim i As Long


   On Error GoTo errorhandler

    With frmEditor_Projectile
        Editor = EDITOR_PROJECTILE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_PROJECTILES
            .lstIndex.AddItem i & ": " & Trim$(Projectiles(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ProjectileEditorInit
    End With


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleProjectileEditor", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub HandleUpdateProjectile(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ProjectileNum As Long
    Dim buffer As clsBuffer
    Dim ProjectileSize As Long
    Dim ProjectileData() As Byte
    

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ProjectileNum = buffer.ReadLong
    
    ProjectileSize = LenB(Projectiles(ProjectileNum))
    ReDim ProjectileData(ProjectileSize - 1)
    ProjectileData = buffer.ReadBytes(ProjectileSize)
    CopyMemory ByVal VarPtr(Projectiles(ProjectileNum)), ByVal VarPtr(ProjectileData(0)), ProjectileSize
    
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUpdateProjectile", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleMapProjectile(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    
    With MapProjectiles(i)
        .ProjectileNum = buffer.ReadLong
        .Owner = buffer.ReadLong
        .OwnerType = buffer.ReadByte
        .dir = buffer.ReadByte
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Range = 0
        .Timer = GetTickCount + 60000
    End With

    Set buffer = Nothing
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapProjectile", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
