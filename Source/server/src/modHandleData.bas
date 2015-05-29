Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long

   On Error GoTo errorhandler

    GetAddress = FunAddr


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Sub InitMessages()

   On Error GoTo errorhandler

    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleCharSlot)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    'HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CEventChatReply) = GetAddress(AddressOf HandleEventChatReply)
    HandleDataSub(CEvent) = GetAddress(AddressOf HandleEvent)
    HandleDataSub(CRequestSwitchesAndVariables) = GetAddress(AddressOf HandleRequestSwitchesAndVariables)
    HandleDataSub(CSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(CBuyHouse) = GetAddress(AddressOf HandleBuyHouse)
    HandleDataSub(CVisit) = GetAddress(AddressOf HandleInviteToHouse)
    HandleDataSub(CAcceptVisit) = GetAddress(AddressOf HandleAcceptInvite)
    HandleDataSub(CPlaceFurniture) = GetAddress(AddressOf HandlePlaceFurniture)
    HandleDataSub(CSendMail) = GetAddress(AddressOf HandleSendMail)
    HandleDataSub(CDeleteMail) = GetAddress(AddressOf HandleDeleteMail)
    HandleDataSub(CReadMail) = GetAddress(AddressOf HandleReadMail)
    HandleDataSub(CTakeMailItem) = GetAddress(AddressOf HandleTakeMailItem)
    HandleDataSub(CRestartServer) = GetAddress(AddressOf HandleRestartServer)
    HandleDataSub(CNewMap) = GetAddress(AddressOf HandleNewMap)
    HandleDataSub(CEditFriend) = GetAddress(AddressOf HandleEditFriend)
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CRequestQuests) = GetAddress(AddressOf HandleRequestQuests)
    HandleDataSub(CPlayerHandleQuest) = GetAddress(AddressOf HandlePlayerHandleQuest)
    HandleDataSub(CQuestLogUpdate) = GetAddress(AddressOf HandleQuestLogUpdate)
    HandleDataSub(CRequestEditZone) = GetAddress(AddressOf HandleRequestEditZone)
    HandleDataSub(CSaveZones) = GetAddress(AddressOf HandleSaveZones)
    HandleDataSub(CAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(CEditAccountLogin) = GetAddress(AddressOf HandleEditAccountLogin)
    HandleDataSub(CAdmin) = GetAddress(AddressOf HandleAdmin)
    HandleDataSub(CServerOpts) = GetAddress(AddressOf HandleServerOpts)
    HandleDataSub(CSaveServerOpt) = GetAddress(AddressOf HandleSaveServerOpt)
    HandleDataSub(CRequestEditHouse) = GetAddress(AddressOf HandleRequestEditHouse)
    HandleDataSub(CSaveHouses) = GetAddress(AddressOf HandleSaveHouses)
    HandleDataSub(CEditPlayer) = GetAddress(AddressOf HandleEditPlayer)
    HandleDataSub(CSavePlayer) = GetAddress(AddressOf HandleSavePlayer)
    HandleDataSub(CEventTouch) = GetAddress(AddressOf HandleEventTouch)
    HandleDataSub(CGameOpts) = GetAddress(AddressOf HandleGameOpts)
    HandleDataSub(CSaveGameOpt) = GetAddress(AddressOf HandleSaveGameOpt)
    HandleDataSub(CMitigation) = GetAddress(AddressOf HandleMitigation)
    HandleDataSub(CRequestEditPet) = GetAddress(AddressOf HandleRequestEditPet)
    HandleDataSub(CSavePet) = GetAddress(AddressOf HandleSavePet)
    HandleDataSub(CRequestPets) = GetAddress(AddressOf HandleRequestPets)
    HandleDataSub(CPetMove) = GetAddress(AddressOf HandlePetMove)
    HandleDataSub(CSetBehaviour) = GetAddress(AddressOf HandleSetPetBehaviour)
    HandleDataSub(CReleasePet) = GetAddress(AddressOf HandleReleasePet)
    HandleDataSub(CPetSpell) = GetAddress(AddressOf HandlePetSpell)
    HandleDataSub(CPetUseStatPoint) = GetAddress(AddressOf HandleUsePetStatPoint)
    HandleDataSub(CRequestEditRandomDungeon) = GetAddress(AddressOf HandleRequestEditRandomDungeon)
    HandleDataSub(CSaveRandomDungeon) = GetAddress(AddressOf HandleSaveRandomDungeon)
    HandleDataSub(CRequestRandomDungeon) = GetAddress(AddressOf HandleRequestRandomDungeon)
    HandleDataSub(CRequestEditProjectiles) = GetAddress(AddressOf HandleRequestEditProjectiles)
    HandleDataSub(CSaveProjectile) = GetAddress(AddressOf HandleSaveProjectile)
    HandleDataSub(CRequestProjectiles) = GetAddress(AddressOf HandleRequestProjectiles)
    HandleDataSub(CClearProjectile) = GetAddress(AddressOf HandleClearProjectile)
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleData(ByVal Index As Long, ByRef data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long, x As Long


   On Error GoTo errorhandler

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(n) Then
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(Index, Name)
                
                Set Buffer = Nothing
                Set Buffer = New clsBuffer
                Buffer.WriteLong SSelChar
                Buffer.WriteLong MAX_PLAYER_CHARS
                
                For i = 1 To MAX_PLAYER_CHARS
                    If Player(Index).characters(i).Class <= 0 Then
                        Buffer.WriteString Trim$(Player(Index).characters(i).Name)
                        Buffer.WriteLong Player(Index).characters(i).Level
                        Buffer.WriteString ""
                        Buffer.WriteLong 0
                    Else
                        Buffer.WriteString Trim$(Player(Index).characters(i).Name)
                        Buffer.WriteLong Player(Index).characters(i).Level
                        Buffer.WriteString Trim$(Class(Player(Index).characters(i).Class).Name)
                        Buffer.WriteLong Player(Index).characters(i).Sex
                    End If
                    For x = 1 To FaceEnum.Face_Count - 1
                        Buffer.WriteLong Player(Index).characters(i).Face(x)
                    Next
                Next
                
                SendDataTo Index, Buffer.ToArray
                
                '' Check if character data has been created
                'If LenB(Trim$(Player(index).Characters(TempPlayer(index).CurChar).Name)) > 0 Then
                    ' we have a char!
                    'HandleUseChar index
                'Else
                    ' send new char shit
                    'If Not IsPlaying(index) Then
                        'Call SendNewCharClasses(index)
                    'End If
                'End If
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
            
            Set Buffer = Nothing
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleNewAccount", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long, x As Long, z As Long


   On Error GoTo errorhandler
   
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If
   
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ' Get the data
    Name = Buffer.ReadString
    
    If GetPlayerLogin(Index) = Trim$(Name) Then
        PlayerMsg Index, "You cannot delete your own account while online!", BrightRed
        Exit Sub
    End If
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If Trim$(Player(i).login) = Trim$(Name) Then
                AlertMsg i, "Your account has been removed by an admin!"
                ClearPlayer i
            End If
        End If
    Next
    
    If AccountCount > 0 Then
        For i = 1 To AccountCount
            If Trim$(account(i).login) = Trim$(Name) Then
                'Delete Stuff!
                For x = 1 To MAX_PLAYER_CHARS
                    If LenB(Trim$(account(i).characters(x))) > 0 Then
                        Call DeleteName(Trim$(account(i).characters(x)))
                    End If
                Next
                For x = i + 1 To AccountCount
                    account(x - 1).access = account(x).access
                    account(x - 1).ip = account(x).ip
                    account(x - 1).login = account(x).login
                    account(x - 1).pass = account(x).pass
                    For z = 1 To MAX_PLAYER_CHARS
                        account(x - 1).characters(z) = account(x).characters(z)
                    Next
                Next
                AccountCount = AccountCount - 1
                ReDim Preserve account(AccountCount)
            End If
        Next
    End If
    
    ' Everything went ok
    Call Kill(App.path & "\data\Accounts\" & Trim$(Name) & "\*.*")
    RmDir App.path & "\data\Accounts\" & Trim$(Name) & "\"
    Call AddLog("Account " & Trim$(Name) & " has been deleted.", ADMIN_LOG)
    
    For i = 1 To MAX_PLAYERS
        If GetPlayerAccess(i) >= ADMIN_CREATOR Then
            SendAccounts i
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleDelAccount", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long, x As Long, cMajor As Long, cMinor As Long, cRevision As Long


   On Error GoTo errorhandler

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
            cMajor = Buffer.ReadLong
            cMinor = Buffer.ReadLong
            cRevision = Buffer.ReadLong
            ' Check versions
            If cMajor <> CLng(App.Major) Or cMinor <> CLng(App.Minor) Or cRevision <> CLng(App.Revision) Then
                Call AlertMsg(Index, "Version outdated, please run the Eclipse Origins launcher instead of eo.exe!")
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(Index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(Index, "Multiple account logins is not authorized.")
                Exit Sub
            End If
            
            If IsBanned(Name, True) Then
                Call AlertMsg(Index, "You are currently banned and cannot login to " & Trim$(Options.Game_Name) & ".")
                Exit Sub
            End If

            ' Load the player
            Call LoadPlayer(Index, Name)
            
            Set Buffer = Nothing
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSelChar
            Buffer.WriteLong MAX_PLAYER_CHARS
            
            For i = 1 To MAX_PLAYER_CHARS
                If Player(Index).characters(i).Class <= 0 Then
                    Buffer.WriteString Trim$(Player(Index).characters(i).Name)
                    Buffer.WriteLong Player(Index).characters(i).Level
                    Buffer.WriteString ""
                    Buffer.WriteLong 0
                Else
                    Buffer.WriteString Trim$(Player(Index).characters(i).Name)
                    Buffer.WriteLong Player(Index).characters(i).Level
                    Buffer.WriteString Trim$(Class(Player(Index).characters(i).Class).Name)
                    Buffer.WriteLong Player(Index).characters(i).Sex
                End If
                For x = 1 To FaceEnum.Face_Count - 1
                    Buffer.WriteLong Player(Index).characters(i).Face(x)
                Next
            Next
            
            SendDataTo Index, Buffer.ToArray
            
            
            ' Check if character data has been created
            'If LenB(Trim$(Player(index).Characters(TempPlayer(index).CurChar).Name)) > 0 Then
                ' we have a char!
                'HandleUseChar index
            'Else
                ' send new char shit
                'If Not IsPlaying(index) Then
                    'Call SendNewCharClasses(index)
                'End If
            'End If
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            
            Set Buffer = Nothing
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleLogin", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub HandleCharSlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, char As CharacterRec
    Dim slot As Long
    Dim i As Long
    Dim n As Long, x As Long


   On Error GoTo errorhandler

    If Not IsPlaying(Index) Then
        If IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes data()
            slot = Buffer.ReadLong
            If Buffer.ReadLong = 1 Then
                'Del Char
                'Clear Mailbox and such...
                If FileExist(App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & "_char" & CStr(slot) & "_mail.ini", True) = True Then Kill App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & "_char" & CStr(slot) & "_mail.ini"
                If FileExist(App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & "_char" & CStr(slot) & "_bank.bin", True) = True Then Kill App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & "_char" & CStr(slot) & "_bank.bin"
                DeleteName Trim$(Player(Index).characters(slot).Name)
                Player(Index).characters(slot) = char
                Player(Index).characters(slot).Name = ""
                SavePlayer Index
                Set Buffer = Nothing
                Set Buffer = New clsBuffer
                Buffer.WriteLong SSelChar
                Buffer.WriteLong MAX_PLAYER_CHARS
                
                For i = 1 To MAX_PLAYER_CHARS
                    If Player(Index).characters(i).Class <= 0 Then
                        Buffer.WriteString Trim$(Player(Index).characters(i).Name)
                        Buffer.WriteLong Player(Index).characters(i).Level
                        Buffer.WriteString ""
                        Buffer.WriteLong 0
                    Else
                        Buffer.WriteString Trim$(Player(Index).characters(i).Name)
                        Buffer.WriteLong Player(Index).characters(i).Level
                        Buffer.WriteString Trim$(Class(Player(Index).characters(i).Class).Name)
                        Buffer.WriteLong Player(Index).characters(i).Sex
                    End If
                    For x = 1 To FaceEnum.Face_Count - 1
                        Buffer.WriteLong Player(Index).characters(i).Face(x)
                    Next
                Next
                
                SendDataTo Index, Buffer.ToArray
                Set Buffer = Nothing
                account(FindAccount(Player(Index).login)).characters(slot) = ""
                Exit Sub
                
            Else
                ' Check if character data has been created
                If LenB(Trim$(Player(Index).characters(slot).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar Index, slot
                    ClearBank Index
                    LoadBank Index, Trim$(Player(Index).login)
                Else
                    ' send new char shit
                    If Not IsPlaying(Index) Then
                        Call SendNewCharClasses(Index)
                        TempPlayer(Index).CurChar = slot
                    End If
                End If
            End If
            
            Set Buffer = Nothing
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleCharSlot", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim slot As Long
    Dim i As Long
    Dim n As Long, charcount As Long


   On Error GoTo errorhandler

    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteBytes data()
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        slot = TempPlayer(Index).CurChar
        
        Player(Index).characters(TempPlayer(Index).CurChar).Face(FaceEnum.Hair) = Buffer.ReadLong
        Player(Index).characters(TempPlayer(Index).CurChar).Face(FaceEnum.Head) = Buffer.ReadLong
        Player(Index).characters(TempPlayer(Index).CurChar).Face(FaceEnum.Eyes) = Buffer.ReadLong
        Player(Index).characters(TempPlayer(Index).CurChar).Face(FaceEnum.Eyebrows) = Buffer.ReadLong
        Player(Index).characters(TempPlayer(Index).CurChar).Face(FaceEnum.Ears) = Buffer.ReadLong
        Player(Index).characters(TempPlayer(Index).CurChar).Face(FaceEnum.Mouth) = Buffer.ReadLong
        Player(Index).characters(TempPlayer(Index).CurChar).Face(FaceEnum.Nose) = Buffer.ReadLong
        Player(Index).characters(TempPlayer(Index).CurChar).Face(FaceEnum.Shirt) = Buffer.ReadLong
        Player(Index).characters(TempPlayer(Index).CurChar).Face(FaceEnum.Etc) = Buffer.ReadLong
        
                
        For i = 1 To SpriteEnum.Sprite_Count - 1
            Player(Index).characters(TempPlayer(Index).CurChar).Sprite(i) = Buffer.ReadLong
        Next

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(Index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(Index, slot) Then
            Call AlertMsg(Index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name, charcount) Then
            Call AlertMsg(Index, "Sorry, but that name is in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(Index, Name, Sex, Class, Sprite, slot)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        ' log them in!!
        HandleUseChar Index, slot
        
        If charcount = 0 Then
            SetPlayerAccess Index, 4
            PlayerMsg Index, "You have been granted admin access by the server. Reason: First Character Created (Assuming User is Owner)", BrightRed
        End If
        
        Set Buffer = Nothing
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAddChar", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(Index), Index, Msg, QBColor(White))
    Call SendChatBubble(GetPlayerMap(Index), Index, TARGET_TYPE_PLAYER, Msg, White)
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEmoteMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    s = "[Global]" & GetPlayerName(Index) & ": " & Msg
    Call SayMsg_Global(Index, Msg, QBColor(White))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> Index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Msg & "'", TellColor)
            Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Cannot message yourself.", BrightRed)
    End If
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    movement = Buffer.ReadLong 'CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(Index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(Index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(Index).StunDuration > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(Index).InShop > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(Index) <> tmpX Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    If GetPlayerY(Index) <> tmpY Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, movement)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim invNum As Long
Dim Buffer As clsBuffer
    
    ' get inventory slot number

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    invNum = Buffer.ReadLong
    Set Buffer = Nothing

    UseItem Index, invNum


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUseItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long
    Dim Buffer As New clsBuffer
    
    ' can't attack whilst casting

   On Error GoTo errorhandler

    If TempPlayer(Index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAttack
    Buffer.WriteLong Index
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing

    ' Projectile check
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        If Item(GetPlayerEquipment(Index, Weapon)).Data1 > 0 Then 'Item has a projectile
            Call PlayerFireProjectile(Index)
            Exit Sub
        End If
    End If
            
    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> Index Then
            If IsPlaying(TempIndex) Then
                TryPlayerAttackPlayer Index, i
                If Player(TempIndex).characters(TempPlayer(TempIndex).CurChar).Pet.Alive = True Then
                    TryPlayerAttackPet Index, i
                End If
            End If
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc Index, i
    Next
    
    For i = 1 To MAX_ZONES
        For x = 1 To MAX_MAP_NPCS * 2
            If ZoneNpc(i).Npc(x).Num > 0 Then
                If ZoneNpc(i).Npc(x).Vital(HP) > 0 Then
                    If ZoneNpc(i).Npc(x).Map = GetPlayerMap(Index) Then
                        TryPlayerAttackZoneNpc Index, i, x
                    End If
                End If
            End If
        Next
    Next

    ' Check tradeskills
    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index) - 1
            y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
            x = GetPlayerX(Index) + 1
            y = GetPlayerY(Index)
    End Select
    
    CheckResource Index, x, y


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim Buffer As clsBuffer
Dim sMes As String
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    PointType = Buffer.ReadByte 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(Index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerRawStat(Index, PointType) >= 255 Then
            PlayerMsg Index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerStat(Index, Stats.Strength, GetPlayerRawStat(Index, Stats.Strength) + 1)
                sMes = "Strength"
            Case Stats.Endurance
                Call SetPlayerStat(Index, Stats.Endurance, GetPlayerRawStat(Index, Stats.Endurance) + 1)
                sMes = "Endurance"
            Case Stats.Intelligence
                Call SetPlayerStat(Index, Stats.Intelligence, GetPlayerRawStat(Index, Stats.Intelligence) + 1)
                sMes = "Intelligence"
            Case Stats.Agility
                Call SetPlayerStat(Index, Stats.Agility, GetPlayerRawStat(Index, Stats.Agility) + 1)
                sMes = "Agility"
            Case Stats.Willpower
                Call SetPlayerStat(Index, Stats.Willpower, GetPlayerRawStat(Index, Stats.Willpower) + 1)
                sMes = "Willpower"
        End Select
        SendPlayerData Index
        SendActionMsg GetPlayerMap(Index), "+1 " & sMes, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)

    Else
        Exit Sub
    End If

    ' Send the update
    'Call SendStats(Index)
    SendPlayerData Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUseStatPoint", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Name = Buffer.ReadString 'Parse(1)
    Set Buffer = Nothing
    i = FindPlayer(Name)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerInfoRequest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MONITOR Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(Index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(Index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp to yourself!", White)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleWarpMeTo", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MONITOR Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If Player(n).characters(TempPlayer(n).CurChar).InHouse > 0 Then
                Player(n).characters(TempPlayer(n).CurChar).InHouse = 0
            Else
                Player(n).characters(TempPlayer(n).CurChar).LastX = GetPlayerX(n)
                Player(n).characters(TempPlayer(n).CurChar).LastY = GetPlayerY(n)
                Player(n).characters(TempPlayer(n).CurChar).LastMap = GetPlayerMap(n)
            End If
            Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            If Player(Index).characters(TempPlayer(Index).CurChar).InHouse > 0 Then
                Player(n).characters(TempPlayer(n).CurChar).InHouse = Player(Index).characters(TempPlayer(Index).CurChar).InHouse
            End If
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(Index) & ".", BrightBlue)
            Call PlayerMsg(Index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp yourself to yourself!", White)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleWarpToMe", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
    Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleWarpTo", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Exit Sub
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    

    ' The sprite
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    Call SetPlayerSprite(Index, n)
    Call SendPlayerData(Index)
    Exit Sub


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSetSprite", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, 1)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestNewMap", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim MapNum As Long
    Dim x As Long
    Dim y As Long, z As Long, w As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Index)
    i = Map(MapNum).Revision + 1
    Call ClearMap(MapNum)
    
    Map(MapNum).Name = Buffer.ReadString
    Map(MapNum).Music = Buffer.ReadString
    Map(MapNum).BGS = Buffer.ReadString
    Map(MapNum).Revision = i
    Map(MapNum).Moral = Buffer.ReadByte
    Map(MapNum).Up = Buffer.ReadLong
    Map(MapNum).Down = Buffer.ReadLong
    Map(MapNum).Left = Buffer.ReadLong
    Map(MapNum).Right = Buffer.ReadLong
    Map(MapNum).BootMap = Buffer.ReadLong
    Map(MapNum).BootX = Buffer.ReadByte
    Map(MapNum).BootY = Buffer.ReadByte
    
    Map(MapNum).Weather = Buffer.ReadLong
    Map(MapNum).WeatherIntensity = Buffer.ReadLong
    
    Map(MapNum).Fog = Buffer.ReadLong
    Map(MapNum).FogSpeed = Buffer.ReadLong
    Map(MapNum).FogOpacity = Buffer.ReadLong
    
    Map(MapNum).Red = Buffer.ReadLong
    Map(MapNum).Green = Buffer.ReadLong
    Map(MapNum).Blue = Buffer.ReadLong
    Map(MapNum).Alpha = Buffer.ReadLong
    
    Map(MapNum).MaxX = Buffer.ReadByte
    Map(MapNum).MaxY = Buffer.ReadByte
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    ReDim Map(MapNum).ExTile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).Tile(x, y).Layer(i).x = Buffer.ReadLong
                Map(MapNum).Tile(x, y).Layer(i).y = Buffer.ReadLong
                Map(MapNum).Tile(x, y).Layer(i).Tileset = Buffer.ReadLong
            Next
            For z = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).Tile(x, y).Autotile(z) = Buffer.ReadLong
            Next
            Map(MapNum).Tile(x, y).type = Buffer.ReadByte
            Map(MapNum).Tile(x, y).Data1 = Buffer.ReadLong
            Map(MapNum).Tile(x, y).Data2 = Buffer.ReadLong
            Map(MapNum).Tile(x, y).Data3 = Buffer.ReadLong
            Map(MapNum).Tile(x, y).Data4 = Buffer.ReadString
            Map(MapNum).Tile(x, y).DirBlock = Buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(x) = Buffer.ReadLong
        Map(MapNum).NpcSpawnType(x) = Buffer.ReadLong
        Call ClearMapNpc(x, MapNum)
    Next
    
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            For i = 1 To ExMapLayer.Layer_Count - 1
                Map(MapNum).ExTile(x, y).Layer(i).x = Buffer.ReadLong
                Map(MapNum).ExTile(x, y).Layer(i).y = Buffer.ReadLong
                Map(MapNum).ExTile(x, y).Layer(i).Tileset = Buffer.ReadLong
            Next
            For z = 1 To ExMapLayer.Layer_Count - 1
                Map(MapNum).ExTile(x, y).Autotile(z) = Buffer.ReadLong
            Next
        Next
    Next
    
    'Event Data!
    Map(MapNum).EventCount = Buffer.ReadLong
        
    If Map(MapNum).EventCount > 0 Then
        ReDim Map(MapNum).Events(0 To Map(MapNum).EventCount)
        For i = 1 To Map(MapNum).EventCount
            With Map(MapNum).Events(i)
                .Name = Buffer.ReadString
                .Global = Buffer.ReadLong
                .x = Buffer.ReadLong
                .y = Buffer.ReadLong
                .PageCount = Buffer.ReadLong
            End With
            If Map(MapNum).Events(i).PageCount > 0 Then
                ReDim Map(MapNum).Events(i).Pages(0 To Map(MapNum).Events(i).PageCount)
                For x = 1 To Map(MapNum).Events(i).PageCount
                    With Map(MapNum).Events(i).Pages(x)
                        .chkVariable = Buffer.ReadLong
                        .VariableIndex = Buffer.ReadLong
                        .VariableCondition = Buffer.ReadLong
                        .VariableCompare = Buffer.ReadLong
                            
                        .chkSwitch = Buffer.ReadLong
                        .SwitchIndex = Buffer.ReadLong
                        .SwitchCompare = Buffer.ReadLong
                            
                        .chkHasItem = Buffer.ReadLong
                        .HasItemIndex = Buffer.ReadLong
                        .HasItemAmount = Buffer.ReadLong
                            
                        .chkSelfSwitch = Buffer.ReadLong
                        .SelfSwitchIndex = Buffer.ReadLong
                        .SelfSwitchCompare = Buffer.ReadLong
                            
                        .GraphicType = Buffer.ReadLong
                        .Graphic = Buffer.ReadLong
                        .GraphicX = Buffer.ReadLong
                        .GraphicY = Buffer.ReadLong
                        .GraphicX2 = Buffer.ReadLong
                        .GraphicY2 = Buffer.ReadLong
                            
                        .MoveType = Buffer.ReadLong
                        .MoveSpeed = Buffer.ReadLong
                        .MoveFreq = Buffer.ReadLong
                            
                        .MoveRouteCount = Buffer.ReadLong
                        
                        .IgnoreMoveRoute = Buffer.ReadLong
                        .RepeatMoveRoute = Buffer.ReadLong
                            
                        If .MoveRouteCount > 0 Then
                            ReDim Map(MapNum).Events(i).Pages(x).MoveRoute(0 To .MoveRouteCount)
                            For y = 1 To .MoveRouteCount
                                .MoveRoute(y).Index = Buffer.ReadLong
                                .MoveRoute(y).Data1 = Buffer.ReadLong
                                .MoveRoute(y).Data2 = Buffer.ReadLong
                                .MoveRoute(y).Data3 = Buffer.ReadLong
                                .MoveRoute(y).Data4 = Buffer.ReadLong
                                .MoveRoute(y).data5 = Buffer.ReadLong
                                .MoveRoute(y).data6 = Buffer.ReadLong
                            Next
                        End If
                            
                        .WalkAnim = Buffer.ReadLong
                        .DirFix = Buffer.ReadLong
                        .WalkThrough = Buffer.ReadLong
                        .ShowName = Buffer.ReadLong
                        .Trigger = Buffer.ReadLong
                        .CommandListCount = Buffer.ReadLong
                            
                        .Position = Buffer.ReadLong
                        .questnum = Buffer.ReadLong
                    End With
                        
                    If Map(MapNum).Events(i).Pages(x).CommandListCount > 0 Then
                        ReDim Map(MapNum).Events(i).Pages(x).CommandList(0 To Map(MapNum).Events(i).Pages(x).CommandListCount)
                        For y = 1 To Map(MapNum).Events(i).Pages(x).CommandListCount
                            Map(MapNum).Events(i).Pages(x).CommandList(y).CommandCount = Buffer.ReadLong
                            Map(MapNum).Events(i).Pages(x).CommandList(y).ParentList = Buffer.ReadLong
                            If Map(MapNum).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                ReDim Map(MapNum).Events(i).Pages(x).CommandList(y).Commands(1 To Map(MapNum).Events(i).Pages(x).CommandList(y).CommandCount)
                                For z = 1 To Map(MapNum).Events(i).Pages(x).CommandList(y).CommandCount
                                    With Map(MapNum).Events(i).Pages(x).CommandList(y).Commands(z)
                                        .Index = Buffer.ReadLong
                                        .Text1 = Buffer.ReadString
                                        .Text2 = Buffer.ReadString
                                        .Text3 = Buffer.ReadString
                                        .Text4 = Buffer.ReadString
                                        .Text5 = Buffer.ReadString
                                        .Data1 = Buffer.ReadLong
                                        .Data2 = Buffer.ReadLong
                                        .Data3 = Buffer.ReadLong
                                        .Data4 = Buffer.ReadLong
                                        .data5 = Buffer.ReadLong
                                        .data6 = Buffer.ReadLong
                                        .ConditionalBranch.CommandList = Buffer.ReadLong
                                        .ConditionalBranch.Condition = Buffer.ReadLong
                                        .ConditionalBranch.Data1 = Buffer.ReadLong
                                        .ConditionalBranch.Data2 = Buffer.ReadLong
                                        .ConditionalBranch.Data3 = Buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = Buffer.ReadLong
                                        .MoveRouteCount = Buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For w = 1 To .MoveRouteCount
                                                .MoveRoute(w).Index = Buffer.ReadLong
                                                .MoveRoute(w).Data1 = Buffer.ReadLong
                                                .MoveRoute(w).Data2 = Buffer.ReadLong
                                                .MoveRoute(w).Data3 = Buffer.ReadLong
                                                .MoveRoute(w).Data4 = Buffer.ReadLong
                                                .MoveRoute(w).data5 = Buffer.ReadLong
                                                .MoveRoute(w).data6 = Buffer.ReadLong
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

    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)
    Call SpawnGlobalEvents(MapNum)
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).characters(TempPlayer(i).CurChar).Map = MapNum Then
                SpawnMapEventsFor i, MapNum
            End If
        End If
    Next

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    ' Save the map
    Call SaveMap(MapNum)
    Call MapCache_Create(MapNum)
    Call ClearTempTile(MapNum)
    Call CacheResources(MapNum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i
    
    Call CacheMapBlocks(MapNum)
    UpdateMapReport

    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim Buffer As clsBuffer
    Dim i As Long, x As Long

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ' Get yes/no value
    s = Buffer.ReadLong 'Parse(1)
    Set Buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If

    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SpawnMapEventsFor(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)

    'send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next

    TempPlayer(Index).GettingMap = NO
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapDone
    SendDataTo Index, Buffer.ToArray()
    
    For i = 1 To MAX_ZONES
        If Len(Trim$(MapZones(i).Name)) > 0 Then
            If MapZones(i).MapCount > 0 Then
                For x = 1 To MapZones(i).MapCount
                    If MapZones(i).Maps(x) = GetPlayerMap(Index) Then
                        If Map(GetPlayerMap(Index)).Weather = 0 Then
                            SendSpecialEffect Index, EFFECT_TYPE_WEATHER, CByte(MapZones(i).CurrentWeather), CLng(MapZones(i).WeatherIntensity)
                        End If
                    End If
                Next
            End If
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleNeedMap", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    Call PlayerMapGetItem(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapGetItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim invNum As Long
    Dim Amount As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    invNum = Buffer.ReadLong 'CLng(Parse(1))
    Amount = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing
    
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(Index, invNum) < 1 Or GetPlayerInvItemNum(Index, invNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(Index, invNum)).Stackable = 1 Then
        If GetPlayerInvItemValue(Index, invNum) = 0 Then Call SetPlayerInvItemValue(Index, invNum, 1)
        If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(Index, invNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next

    CacheResources GetPlayerMap(Index)
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapRespawn", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim i As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1

    For i = 1 To MAX_MAPS

        If LenB(Trim$(Map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = i + 1
            tMapEnd = i + 1
        End If

    Next

    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    Call PlayerMsg(Index, s, Brown)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMapReport", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(Index) & "!", White)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot kick yourself!", White)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleKickPlayer", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim F As Long
    Dim s As String
    Dim Name As String

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MONITOR Then
        Exit Sub
    End If

    n = 1
    F = FreeFile
    Open App.path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        Call PlayerMsg(Index, n & ": Banned IP " & s & " by " & Name, White)
        n = n + 1
    Loop

    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleBanList", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String, Buffer As clsBuffer
    Dim File As Long
    Dim F As Long, i As Long, x As Long

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    x = Buffer.ReadLong
    Set Buffer = Nothing

    filename = App.path & "\data\banlist.bin"

    If Not FileExist("data\banlist.bin") Then
        F = FreeFile
        Open filename For Binary As #F
        Close #F
    End If
    
    If BanCount > 0 Then
        For i = x To BanCount - 1
            Bans(x).BanChar = Bans(x + 1).BanChar
            Bans(x).BanName = Bans(x + 1).BanName
            Bans(x).BanReason = Bans(x + 1).BanReason
            Bans(x).IPAddress = Bans(x + 1).IPAddress
        Next
        BanCount = BanCount - 1
        ReDim Preserve Bans(BanCount)
        PlayerMsg Index, "Ban removed!", BrightRed
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If GetPlayerAccess(i) >= ADMIN_CREATOR Then
                    SendAccounts i
                    SendBans i
                End If
            End If
        Next
        F = FreeFile
        Open filename For Binary As #F
        Put #F, , BanCount
        For i = 1 To BanCount
            Put #F, , Bans(i).BanName
            Put #F, , Bans(i).BanChar
            Put #F, , Bans(i).IPAddress
            Put #F, , Bans(i).BanReason
        Next
        Close #F
    Else
        PlayerMsg Index, "Error! No bans to delete!", BrightRed
    End If
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleBanDestroy", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long, s As String, acc As Boolean, i As Long, reason As String
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MONITOR Then
        Exit Sub
    End If
    
    s = Trim$(Buffer.ReadString)
    
    If Buffer.ReadLong = 1 Then
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Exit Sub
        End If
        acc = True
    Else
        n = FindPlayer(s)
    End If
    
    reason = Buffer.ReadString
    
    Set Buffer = Nothing
    If acc = True Then
        If Trim$(s) <> Trim$(Player(Index).login) Then
            n = FindAccount(Trim$(s))
            If n > 0 Then
                If account(n).access > GetPlayerAccess(Index) Then
                    Call PlayerMsg(Index, "A character on that account is a higher or same access admin then you!", White)
                Else
                    If Ban(0, Trim$(s), False, reason) = False Then
                        PlayerMsg Index, "Account is already banned!", BrightRed
                    End If
                End If
            Else
            
            End If
        Else
            Call PlayerMsg(Index, "You cannot ban yourself!", White)
        End If
    Else
        If n <> Index Then
            If n > 0 Then
                For i = 1 To MAX_PLAYER_CHARS
                    If Player(n).characters(i).access > GetPlayerAccess(Index) Then
                        Call PlayerMsg(Index, "A character on that account is a higher or same access admin then you!", White)
                    Else
                        If i = MAX_PLAYER_CHARS Then
                            If Ban(n, "", True, reason) = False Then
                                PlayerMsg Index, "Account is already banned!", BrightRed
                            End If
                            Exit Sub
                        End If
                    End If
                Next
            Else
                Call PlayerMsg(Index, "Player is no longer online.", White)
            End If
    
        Else
            Call PlayerMsg(Index, "You cannot ban yourself!", White)
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleBanPlayer", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    SendMapEventData (Index)

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEditMap
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditMap", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SItemEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimationEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditAnimation", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(Index) & " saved Animation #" & n & ".", ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveAnimation", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditNpc", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim npcnum As Long
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    npcnum = Buffer.ReadLong

    ' Prevent hacking
    If npcnum < 0 Or npcnum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(Npc(npcnum))
    ReDim NPCData(NPCSize - 1)
    NPCData = Buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(Npc(npcnum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(npcnum)
    Call SaveNpc(npcnum)
    Call AddLog(GetPlayerName(Index) & " saved Npc #" & npcnum & ".", ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveNpc", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditResource", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ResourceNum = Buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(Index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveResource", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SShopEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditShop", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    shopNum = Buffer.ReadLong

    ' Prevent hacking
    If shopNum < 0 Or shopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set Buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(shopNum)
    Call SaveShop(shopNum)
    Call AddLog(GetPlayerName(Index) & " saving shop #" & shopNum & ".", ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveShop", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditspell", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Spellnum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Spellnum = Buffer.ReadLong

    ' Prevent hacking
    If Spellnum < 0 Or Spellnum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(Spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(Spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(Spellnum)
    Call SaveSpell(Spellnum)
    Call AddLog(GetPlayerName(Index) & " saved Spell #" & Spellnum & ".", ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveSpell", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The access
    i = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
                Call PlayerMsg(Index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Invalid access level.", Red)
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSetAccess", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    Call SendWhosOnline(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleWhosOnline", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(Buffer.ReadString) 'Parse(1))
    SaveOptions
    Set Buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSetMotd", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim x As Long
    Dim y As Long
    Dim i As Long, z As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    x = Buffer.ReadLong 'CLng(Parse(1))
    y = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Prevent subscript out of range
    If x < 0 Or x > Map(GetPlayerMap(Index)).MaxX Or y < 0 Or y > Map(GetPlayerMap(Index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If Player(i).characters(TempPlayer(i).CurChar).InHouse = Player(Index).characters(TempPlayer(Index).CurChar).InHouse Then
                    If GetPlayerX(i) = x Then
                        If GetPlayerY(i) = y Then
                            ' Change target
                            If TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER And TempPlayer(Index).Target = i Then
                                TempPlayer(Index).Target = 0
                                TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                                TempPlayer(Index).TargetZone = 0
                                ' send target to player
                                SendTarget Index
                            Else
                                TempPlayer(Index).Target = i
                                TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER
                                TempPlayer(Index).TargetZone = 0
                                ' send target to player
                                SendTarget Index
                            End If
                            Exit Sub
                        End If
                    End If
                End If
                If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive Then
                    If Player(i).characters(TempPlayer(i).CurChar).Pet.x = x And Player(i).characters(TempPlayer(i).CurChar).Pet.y = y Then
                            ' Change target
                            If TempPlayer(Index).TargetType = TARGET_TYPE_PET And TempPlayer(Index).Target = i Then
                                TempPlayer(Index).Target = 0
                                TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                                ' send target to player
                                SendTarget Index
                            Else
                                TempPlayer(Index).Target = i
                                TempPlayer(Index).TargetType = TARGET_TYPE_PET
                                ' send target to player
                                SendTarget Index
                            End If
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next

    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index)).Npc(i).Num > 0 Then
            If MapNpc(GetPlayerMap(Index)).Npc(i).x = x Then
                If MapNpc(GetPlayerMap(Index)).Npc(i).y = y Then
                    If TempPlayer(Index).Target = i And TempPlayer(Index).TargetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(Index).Target = 0
                        TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                        TempPlayer(Index).TargetZone = 0
                        ' send target to player
                        SendTarget Index
                    Else
                        ' Change target
                        TempPlayer(Index).Target = i
                        TempPlayer(Index).TargetType = TARGET_TYPE_NPC
                        TempPlayer(Index).TargetZone = 0
                        ' send target to player
                        SendTarget Index
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next
    
    For i = 1 To MAX_ZONES
        For z = 1 To MAX_MAP_NPCS * 2
            If ZoneNpc(i).Npc(z).Num > 0 Then
                If ZoneNpc(i).Npc(z).Vital(HP) > 0 Then
                    If ZoneNpc(i).Npc(z).Map = GetPlayerMap(Index) Then
                        If ZoneNpc(i).Npc(z).x = x Then
                            If ZoneNpc(i).Npc(z).y = y Then
                                If TempPlayer(Index).Target = z And TempPlayer(Index).TargetType = TARGET_TYPE_ZONENPC And TempPlayer(Index).TargetZone = i Then
                                    ' Change target
                                    TempPlayer(Index).Target = 0
                                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                                    TempPlayer(Index).TargetZone = 0
                                    ' send target to player
                                    SendTarget Index
                                Else
                                    ' Change target
                                    TempPlayer(Index).Target = z
                                    TempPlayer(Index).TargetType = TARGET_TYPE_ZONENPC
                                    TempPlayer(Index).TargetZone = i
                                    ' send target to player
                                    SendTarget Index
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next
    Next
    
    If Player(Index).characters(TempPlayer(Index).CurChar).InHouse > 0 Then
        If Player(Index).characters(TempPlayer(Index).CurChar).InHouse = Index Then
            If Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex > 0 Then
                If Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount > 0 Then
                    For i = 1 To Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount
                        If x >= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).x And x <= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).x + Item(Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).ItemNum).FurnitureWidth - 1 Then
                            If y <= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).y And y >= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).y - Item(Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).ItemNum).FurnitureHeight + 1 Then
                                'Found an Item, get the index and lets pick it up!
                                x = FindOpenInvSlot(Index, Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).ItemNum)
                                If x > 0 Then
                                    GiveInvItem Index, Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).ItemNum, 0, True
                                    Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount = Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount - 1
                                    For x = i + 1 To Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount + 1
                                        Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(x - 1) = Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(x)
                                    Next
                                    ReDim Preserve Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount)
                                    SendFurnitureToHouse Index
                                    Exit Sub
                                Else
                                    PlayerMsg Index, "No inventory space available!", BrightRed
                                End If
                                Exit Sub
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSearch", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    Call SendPlayerSpells(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(Index, n)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleCast", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    Call CloseSocket(Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleQuit", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    

   On Error GoTo errorhandler

    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchInvSlots Index, oldSlot, newSlot


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSwapInvSlots", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleSwapSpellSlots(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    

   On Error GoTo errorhandler

    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        PlayerMsg Index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(Index).SpellCD(n) > GetTickCount Then
            PlayerMsg Index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchSpellSlots Index, oldSlot, newSlot


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSwapSpellSlots", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendPing
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleCheckPing", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleUnequip(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    PlayerUnequipItem Index, Buffer.ReadLong
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUnequip", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestPlayerData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendPlayerData Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestPlayerData", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestItems(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendItems Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestItems", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestAnimations(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendAnimations Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestAnimations", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestNPCS(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendNpcs Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestNPCS", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestResources(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendResources Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestResources", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestSpells(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendSpells Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestSpells", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestShops(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendShops Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestShops", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleSpawnItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    ' item
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong
        
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index)
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestLevelUp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < 4 Then Exit Sub
    SetPlayerExp Index, GetPlayerNextLevel(Index)
    CheckPlayerLevelUp Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestLevelUp", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleForgetSpell(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim spellslot As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    spellslot = Buffer.ReadLong
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(Index).spellBuffer.Spell = spellslot Then
        PlayerMsg Index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Player(Index).characters(TempPlayer(Index).CurChar).Spell(spellslot) = 0
    SendPlayerSpells Index
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleForgetSpell", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleCloseShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    TempPlayer(Index).InShop = 0


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleCloseShop", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleBuyItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim itemAmount As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    shopslot = Buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(Index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
            
        ' check has the cost item
        itemAmount = HasItem(Index, .costitem)
        If itemAmount = 0 Or itemAmount < .costvalue Then
            PlayerMsg Index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction Index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        TakeInvItem Index, .costitem, .costvalue
        GiveInvItem Index, .Item, .itemvalue
    End With
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Trade successful.", BrightGreen
    ResetShopAction Index
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleBuyItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleSellItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invslot As Long
    Dim ItemNum As Long
    Dim price As Long
    Dim multiplier As Double
    Dim Amount As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    invslot = Buffer.ReadLong
    
    ' if invalid, exit out
    If invslot < 1 Or invslot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(Index, invslot) < 1 Or GetPlayerInvItemNum(Index, invslot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    ItemNum = GetPlayerInvItemNum(Index, invslot)
    
    ' work out price
    multiplier = Shop(TempPlayer(Index).InShop).BuyRate / 100
    price = Item(ItemNum).price * multiplier
    
    ' item has cost?
    If price <= 0 Then
        PlayerMsg Index, "The shop doesn't want that item.", BrightRed
        ResetShopAction Index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem Index, ItemNum, 1
    GiveInvItem Index, 1, price
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Trade successful.", BrightGreen
    ResetShopAction Index
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSellItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleChangeBankSlots(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    
    PlayerSwitchBankSlots Index, oldSlot, newSlot
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleChangeBankSlots", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleWithdrawItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim BankSlot As Long
    Dim Amount As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    BankSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    TakeBankItem Index, BankSlot, Amount
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleWithdrawItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleDepositItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invslot As Long
    Dim Amount As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    invslot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    GiveBankItem Index, invslot, Amount
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleDepositItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleCloseBank(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    SaveBank Index
    SavePlayer Index
    
    TempPlayer(Index).InBank = False
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleCloseBank", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleAdminWarp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long
    Dim y As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    
    If GetPlayerAccess(Index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX Index, x
        SetPlayerY Index, y
        SendPlayerXYToMap Index
    End If
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAdminWarp", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleTradeRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs

   On Error GoTo errorhandler

    If TempPlayer(Index).TargetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(Index).Target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = Index Then
        PlayerMsg Index, "You can't trade with yourself.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).characters(TempPlayer(tradeTarget).CurChar).Map = Player(Index).characters(TempPlayer(Index).CurChar).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).characters(TempPlayer(tradeTarget).CurChar).x
    tY = Player(tradeTarget).characters(TempPlayer(tradeTarget).CurChar).y
    sX = Player(Index).characters(TempPlayer(Index).CurChar).x
    sY = Player(Index).characters(TempPlayer(Index).CurChar).y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg Index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = Index
    SendTradeRequest tradeTarget, Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleTradeRequest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleAcceptTradeRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim i As Long


   On Error GoTo errorhandler

    If TempPlayer(Index).InTrade > 0 Then
        TempPlayer(Index).TradeRequest = 0
    Else
        tradeTarget = TempPlayer(Index).TradeRequest
        If tradeTarget > 0 Then
            If IsPlaying(tradeTarget) Then
                ' let them know they're trading
                PlayerMsg Index, "You have accepted " & Trim$(GetPlayerName(tradeTarget)) & "'s trade request.", BrightGreen
                PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has accepted your trade request.", BrightGreen
                ' clear the tradeRequest server-side
                TempPlayer(Index).TradeRequest = 0
                TempPlayer(tradeTarget).TradeRequest = 0
                ' set that they're trading with each other
                TempPlayer(Index).InTrade = tradeTarget
                TempPlayer(tradeTarget).InTrade = Index
                ' clear out their trade offers
                For i = 1 To MAX_INV
                    TempPlayer(Index).TradeOffer(i).Num = 0
                    TempPlayer(Index).TradeOffer(i).Value = 0
                    TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                    TempPlayer(tradeTarget).TradeOffer(i).Value = 0
                Next
                ' Used to init the trade window clientside
                SendTrade Index, tradeTarget
                SendTrade tradeTarget, Index
                ' Send the offer data - Used to clear their client
                SendTradeUpdate Index, 0
                SendTradeUpdate Index, 1
                SendTradeUpdate tradeTarget, 0
                SendTradeUpdate tradeTarget, 1
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAcceptTradeRequest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleDeclineTradeRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, i As Long
   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    If i = 0 Then
        PlayerMsg TempPlayer(Index).TradeRequest, GetPlayerName(Index) & " has declined your trade request.", BrightRed
        PlayerMsg Index, "You decline the trade request.", BrightRed
        ' clear the tradeRequest server-side
        TempPlayer(Index).TradeRequest = 0
    Else
        PlayerMsg TempPlayer(Index).TradeRequest, GetPlayerName(Index) & " is busy!", BrightRed
        ' clear the tradeRequest server-side
        TempPlayer(Index).TradeRequest = 0
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleDeclineTradeRequest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleAcceptTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim ItemNum As Long
    

   On Error GoTo errorhandler

    TempPlayer(Index).AcceptTrade = True
    
    tradeTarget = TempPlayer(Index).InTrade
        
    If tradeTarget > 0 Then
    
        ' if not both of them accept, then exit
        If Not TempPlayer(tradeTarget).AcceptTrade Then
            SendTradeStatus Index, 2
            SendTradeStatus tradeTarget, 1
            Exit Sub
        End If
    
        ' take their items
        For i = 1 To MAX_INV
            ' player
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                ItemNum = Player(Index).characters(TempPlayer(Index).CurChar).Inv(TempPlayer(Index).TradeOffer(i).Num).Num
                If ItemNum > 0 Then
                    ' store temp
                    tmpTradeItem(i).Num = ItemNum
                    tmpTradeItem(i).Value = TempPlayer(Index).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot Index, TempPlayer(Index).TradeOffer(i).Num, tmpTradeItem(i).Value
                End If
            End If
            ' target
            If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
                ItemNum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                If ItemNum > 0 Then
                    ' store temp
                    tmpTradeItem2(i).Num = ItemNum
                    tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, tmpTradeItem2(i).Value
                End If
            End If
        Next
    
        ' taken all items. now they can't not get items because of no inventory space.
        For i = 1 To MAX_INV
            ' player
            If tmpTradeItem2(i).Num > 0 Then
                ' give away!
                GiveInvItem Index, tmpTradeItem2(i).Num, tmpTradeItem2(i).Value, False
            End If
            ' target
            If tmpTradeItem(i).Num > 0 Then
                ' give away!
                GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).Value, False
            End If
        Next
    
        SendInventory Index
        SendInventory tradeTarget
    
        ' they now have all the items. Clear out values + let them out of the trade.
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).Num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(Index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg Index, "Trade completed.", BrightGreen
        PlayerMsg tradeTarget, "Trade completed.", BrightGreen
    
        SendCloseTrade Index
        SendCloseTrade tradeTarget
            
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAcceptTrade", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleDeclineTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tradeTarget As Long


   On Error GoTo errorhandler

    tradeTarget = TempPlayer(Index).InTrade
    
    If tradeTarget > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).Num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(Index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg Index, "You declined the trade.", BrightRed
        PlayerMsg tradeTarget, GetPlayerName(Index) & " has declined the trade.", BrightRed
    
        SendCloseTrade Index
        SendCloseTrade tradeTarget
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleDeclineTrade", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleTradeItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invslot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim ItemNum As Long
    Dim i As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    invslot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If invslot <= 0 Or invslot > MAX_INV Then Exit Sub
    
    ItemNum = GetPlayerInvItemNum(Index, invslot)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invslot) Then
        Exit Sub
    End If

    If Item(ItemNum).Stackable = 1 Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invslot Then
                ' add amount
                TempPlayer(Index).TradeOffer(i).Value = TempPlayer(Index).TradeOffer(i).Value + Amount
                ' clamp to limits
                If TempPlayer(Index).TradeOffer(i).Value > GetPlayerInvItemValue(Index, invslot) Then
                    TempPlayer(Index).TradeOffer(i).Value = GetPlayerInvItemValue(Index, invslot)
                End If
                ' cancel any trade agreement
                TempPlayer(Index).AcceptTrade = False
                TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
                
                SendTradeStatus Index, 0
                SendTradeStatus TempPlayer(Index).InTrade, 0
                
                SendTradeUpdate Index, 0
                SendTradeUpdate TempPlayer(Index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invslot Then
                PlayerMsg Index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(Index).TradeOffer(EmptySlot).Num = invslot
    TempPlayer(Index).TradeOffer(EmptySlot).Value = Amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(Index).AcceptTrade = False
    TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleTradeItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleUntradeItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    tradeSlot = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(Index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(Index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(Index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(Index).AcceptTrade Then TempPlayer(Index).AcceptTrade = False
    If TempPlayer(TempPlayer(Index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUntradeItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleHotbarChange(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim sType As Long
    Dim slot As Long
    Dim hotbarNum As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    sType = Buffer.ReadLong
    slot = Buffer.ReadLong
    hotbarNum = Buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(hotbarNum).slot = 0
            Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If slot > 0 And slot <= MAX_INV Then
                If Player(Index).characters(TempPlayer(Index).CurChar).Inv(slot).Num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(Index, slot)).Name)) > 0 Then
                        Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(hotbarNum).slot = Player(Index).characters(TempPlayer(Index).CurChar).Inv(slot).Num
                        Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If slot > 0 And slot <= MAX_PLAYER_SPELLS Then
                If Player(Index).characters(TempPlayer(Index).CurChar).Spell(slot) > 0 Then
                    If Len(Trim$(Spell(Player(Index).characters(TempPlayer(Index).CurChar).Spell(slot)).Name)) > 0 Then
                        Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(hotbarNum).slot = Player(Index).characters(TempPlayer(Index).CurChar).Spell(slot)
                        Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar Index
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleHotbarChange", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleHotbarUse(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim slot As Long
    Dim i As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    slot = Buffer.ReadLong
    
    Select Case Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(Index).characters(TempPlayer(Index).CurChar).Inv(i).Num > 0 Then
                    If Player(Index).characters(TempPlayer(Index).CurChar).Inv(i).Num = Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(slot).slot Then
                        UseItem Index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(Index).characters(TempPlayer(Index).CurChar).Spell(i) > 0 Then
                    If Player(Index).characters(TempPlayer(Index).CurChar).Spell(i) = Player(Index).characters(TempPlayer(Index).CurChar).Hotbar(slot).slot Then
                        BufferSpell Index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleHotbarUse", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandlePartyRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' make sure it's a valid target

   On Error GoTo errorhandler

    If TempPlayer(Index).TargetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(Index).Target = Index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(Index).Target) Or Not IsPlaying(TempPlayer(Index).Target) Then Exit Sub
    If GetPlayerMap(TempPlayer(Index).Target) <> GetPlayerMap(Index) Then Exit Sub
    
    ' init the request
    Party_Invite Index, TempPlayer(Index).Target


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePartyRequest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleAcceptParty(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    Party_InviteAccept TempPlayer(Index).partyInvite, Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAcceptParty", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleDeclineParty(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    Party_InviteDecline TempPlayer(Index).partyInvite, Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleDeclineParty", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandlePartyLeave(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    Party_PlayerLeave Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePartyLeave", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleEventChatReply(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim eventID As Long, pageID As Long, reply As Long, i As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    eventID = Buffer.ReadLong
    pageID = Buffer.ReadLong
    reply = Buffer.ReadLong
    
    If TempPlayer(Index).EventProcessingCount > 0 Then
        For i = 1 To TempPlayer(Index).EventProcessingCount
            If TempPlayer(Index).EventProcessing(i).eventID = eventID And TempPlayer(Index).EventProcessing(i).pageID = pageID Then
                If TempPlayer(Index).EventProcessing(i).WaitingForResponse = 1 Then
                    If reply = 0 Then
                        If Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Index = EventType.evShowText Then
                            TempPlayer(Index).EventProcessing(i).WaitingForResponse = 0
                        End If
                    ElseIf reply > 0 Then
                        If Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Index = EventType.evShowChoices Then
                            Select Case reply
                                Case 1
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot - 1
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data1
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                                Case 2
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot - 1
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data2
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                                Case 3
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot - 1
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data3
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                                Case 4
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot - 1
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data4
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                            End Select
                        End If
                        TempPlayer(Index).EventProcessing(i).WaitingForResponse = 0
                    End If
                End If
            End If
        Next
    End If
    
    
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEventChatReply", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleEvent(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long, begineventprocessing As Boolean, z As Long, Buffer As clsBuffer

    ' Check tradeskills

   On Error GoTo errorhandler

    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index) - 1
            y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
            x = GetPlayerX(Index) + 1
            y = GetPlayerY(Index)
    End Select
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data
    i = Buffer.ReadLong
    Set Buffer = Nothing
    
    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
            If TempPlayer(Index).EventMap.EventPages(z).eventID = i Then
                i = z
                begineventprocessing = True
                Exit For
            End If
        Next
    End If
    
    If begineventprocessing = True Then
        If Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).CommandListCount > 0 Then
            'Process this event, it is action button and everything checks out.
            TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).Active = 1
            With TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID)
                .ActionTimer = GetTickCount
                .CurList = 1
                .CurSlot = 1
                .eventID = TempPlayer(Index).EventMap.EventPages(i).eventID
                .pageID = TempPlayer(Index).EventMap.EventPages(i).pageID
                .WaitingForResponse = 0
                ReDim .ListLeftOff(0 To Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).CommandListCount)
            End With
        End If
        begineventprocessing = False
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEvent", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestSwitchesAndVariables(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendSwitchesAndVariables (Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = Buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = Buffer.ReadString
    Next
    
    SaveSwitches
    SaveVariables
    
    Set Buffer = Nothing
    
    SendSwitchesAndVariables 0, True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleBuyHouse(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, price As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    i = Buffer.ReadLong
    
    If i = 1 Then
        If TempPlayer(Index).BuyHouseIndex > 0 Then
            price = HouseConfig(TempPlayer(Index).BuyHouseIndex).price
            If HasItem(Index, 1) >= price Then
                TakeInvItem Index, 1, price
                Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex = TempPlayer(Index).BuyHouseIndex
                PlayerMsg Index, "You just bought the " & Trim$(HouseConfig(TempPlayer(Index).BuyHouseIndex).ConfigName) & " house!", White
                Player(Index).characters(TempPlayer(Index).CurChar).LastMap = GetPlayerMap(Index)
                Player(Index).characters(TempPlayer(Index).CurChar).LastX = GetPlayerX(Index)
                Player(Index).characters(TempPlayer(Index).CurChar).LastY = GetPlayerY(Index)
                Player(Index).characters(TempPlayer(Index).CurChar).InHouse = Index
                Call PlayerWarp(Index, HouseConfig(Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex).BaseMap, HouseConfig(Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex).x, HouseConfig(Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex).y, True)
            Else
                PlayerMsg Index, "You cannot afford this house!", BrightRed
            End If
        End If
    End If
    
    TempPlayer(Index).BuyHouseIndex = 0
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleBuyHouse", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleInviteToHouse(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, invitee As Long, Name As String
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Name = Trim$(Buffer.ReadString)
    invitee = FindPlayer(Name)
    Set Buffer = Nothing
    
    If invitee = 0 Then PlayerMsg Index, "Player not found.", BrightRed: Exit Sub
    
    If Index = invitee Then PlayerMsg Index, "You cannot invite yourself to you own house!", BrightRed: Exit Sub
    
    If TempPlayer(invitee).InvitationIndex > 0 Then
        If TempPlayer(invitee).InvitationTimer > GetTickCount() Then
            PlayerMsg Index, Trim$(GetPlayerName(invitee)) & " is currently busy!", BrightRed
            Exit Sub
        End If
    End If
    
    If Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex > 0 Then
        If Player(Index).characters(TempPlayer(Index).CurChar).InHouse > 0 Then
            If Player(Index).characters(TempPlayer(Index).CurChar).InHouse = Index Then
                If Player(invitee).characters(TempPlayer(invitee).CurChar).InHouse > 0 Then
                    If Player(invitee).characters(TempPlayer(invitee).CurChar).InHouse = Index Then
                        PlayerMsg Index, Trim$(GetPlayerName(invitee)) & " is already in your house!", BrightRed
                    Else
                        PlayerMsg Index, Trim$(GetPlayerName(invitee)) & " is already visiting someone elses house!", BrightRed
                    End If
                Else
                    'Send invite
                    Set Buffer = New clsBuffer
                    Buffer.WriteLong SVisit
                    Buffer.WriteLong Index
                    SendDataTo invitee, Buffer.ToArray
                    TempPlayer(invitee).InvitationIndex = Index
                    TempPlayer(invitee).InvitationTimer = GetTickCount() + 15000
                    Set Buffer = Nothing
                End If
            Else
                PlayerMsg Index, "Only the house owner can invite other players into their house.", BrightRed
            End If
        Else
            PlayerMsg Index, "You must be inside your house before you can invite someone to visit!", BrightRed
        End If
    Else
        PlayerMsg Index, "You do not have a house to invite anyone to!", BrightRed
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleInviteToHouse", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleAcceptInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, response As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    response = Buffer.ReadLong
    Set Buffer = Nothing
    
    If response = 1 Then
        If TempPlayer(Index).InvitationIndex > 0 Then
            If TempPlayer(Index).InvitationTimer > GetTickCount() Then
                'Accept this invite
                If IsPlaying(TempPlayer(Index).InvitationIndex) Then
                    Player(Index).characters(TempPlayer(Index).CurChar).InHouse = TempPlayer(Index).InvitationIndex
                    Player(Index).characters(TempPlayer(Index).CurChar).LastX = GetPlayerX(Index)
                    Player(Index).characters(TempPlayer(Index).CurChar).LastY = GetPlayerY(Index)
                    Player(Index).characters(TempPlayer(Index).CurChar).LastMap = GetPlayerMap(Index)
                    TempPlayer(Index).InvitationTimer = 0
                    PlayerWarp Index, Player(TempPlayer(Index).InvitationIndex).characters(TempPlayer(TempPlayer(Index).InvitationIndex).CurChar).Map, HouseConfig(Player(TempPlayer(Index).InvitationIndex).characters(TempPlayer(TempPlayer(Index).InvitationIndex).CurChar).House.HouseIndex).x, HouseConfig(Player(TempPlayer(Index).InvitationIndex).characters(TempPlayer(TempPlayer(Index).InvitationIndex).CurChar).House.HouseIndex).y, True
                Else
                    TempPlayer(Index).InvitationTimer = 0
                    PlayerMsg Index, "Cannot find player!", BrightRed
                End If
            Else
                PlayerMsg Index, "Your invitation has expired, have your friend re-invite you.", BrightRed
            End If
        Else
            
        End If
    Else
        If IsPlaying(TempPlayer(Index).InvitationIndex) Then
            TempPlayer(Index).InvitationTimer = 0
            PlayerMsg TempPlayer(Index).InvitationIndex, Trim$(GetPlayerName(Index)) & " rejected your invitation", BrightRed
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAcceptInvite", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

Sub HandlePlaceFurniture(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, x As Long, y As Long, invslot As Long, ItemNum As Long, x1 As Long, y1 As Long, widthoffset As Long, heightoffset As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    invslot = Buffer.ReadLong
    Set Buffer = Nothing
    
    ItemNum = Player(Index).characters(TempPlayer(Index).CurChar).Inv(invslot).Num
    
    ' Prevent hacking
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If Player(Index).characters(TempPlayer(Index).CurChar).InHouse = Index Then
        If Item(ItemNum).type = ITEM_TYPE_FURNITURE Then
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
                
                'Ok, now we got to see what can be done about this furniture :/
                If Player(Index).characters(TempPlayer(Index).CurChar).InHouse <> Index Then
                    PlayerMsg Index, "You must be inside your house to place furniture!", BrightRed
                    Exit Sub
                End If
                
                If Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount >= HouseConfig(Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex).MaxFurniture Then
                    If HouseConfig(Player(Index).characters(TempPlayer(Index).CurChar).House.HouseIndex).MaxFurniture > 0 Then
                        PlayerMsg Index, "Your house cannot hold any more furniture!", BrightRed
                        Exit Sub
                    End If
                End If
                
                If x < 0 Or x > Map(GetPlayerMap(Index)).MaxX Then Exit Sub
                If y < 0 Or y > Map(GetPlayerMap(Index)).MaxY Then Exit Sub
                
                If Item(ItemNum).FurnitureWidth > 2 Then
                    x1 = x + (Item(ItemNum).FurnitureWidth / 2)
                    widthoffset = x1 - x
                    x1 = x1 - (Item(ItemNum).FurnitureWidth - widthoffset)
                Else
                    x1 = x
                End If
                
                x1 = x
                widthoffset = 0
                
                y1 = y
                
                If widthoffset > 0 Then
                
                    For x = x1 To x1 + widthoffset
                        For y = y1 To y1 - Item(ItemNum).FurnitureHeight + 1 Step -1
                            If Map(GetPlayerMap(Index)).Tile(x, y).type = TILE_TYPE_BLOCKED Then Exit Sub
                    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If i <> Index Then
                                        If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                            If Player(i).characters(TempPlayer(i).CurChar).InHouse = Player(Index).characters(TempPlayer(Index).CurChar).InHouse Then
                                                If Player(i).characters(TempPlayer(i).CurChar).x = x And Player(i).characters(TempPlayer(i).CurChar).y = y Then
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            
                            If Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount > 0 Then
                                For i = 1 To Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount
                                    If x >= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).x And x <= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).x + Item(Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).ItemNum).FurnitureWidth - 1 Then
                                        If y <= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).y And y >= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).y - Item(Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).ItemNum).FurnitureHeight + 1 Then
                                            'Blocked!
                                            Exit Sub
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    Next
                    
                    For x = x1 To x1 - (Item(ItemNum).FurnitureWidth - widthoffset) Step -1
                        For y = y1 To y1 - Item(ItemNum).FurnitureHeight + 1 Step -1
                            If Map(GetPlayerMap(Index)).Tile(x, y).type = TILE_TYPE_BLOCKED Then Exit Sub
                    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If i <> Index Then
                                        If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                            If Player(i).characters(TempPlayer(i).CurChar).InHouse = Player(Index).characters(TempPlayer(Index).CurChar).InHouse Then
                                                If Player(i).characters(TempPlayer(i).CurChar).x = x And Player(i).characters(TempPlayer(i).CurChar).y = y Then
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            
                            If Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount > 0 Then
                                For i = 1 To Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount
                                    If x >= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).x And x <= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).x + Item(Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).ItemNum).FurnitureWidth - 1 Then
                                        If y <= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).y And y >= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).y - Item(Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).ItemNum).FurnitureHeight + 1 Then
                                            'Blocked!
                                            Exit Sub
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    Next
                Else
                    For x = x1 To x1 + Item(ItemNum).FurnitureWidth - 1
                        For y = y1 To y1 - Item(ItemNum).FurnitureHeight + 1 Step -1
                            If Map(GetPlayerMap(Index)).Tile(x, y).type = TILE_TYPE_BLOCKED Then Exit Sub
                    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If i <> Index Then
                                        If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                            If Player(i).characters(TempPlayer(i).CurChar).InHouse = Player(Index).characters(TempPlayer(Index).CurChar).InHouse Then
                                                If Player(i).characters(TempPlayer(i).CurChar).x = x And Player(i).characters(TempPlayer(i).CurChar).y = y Then
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            
                            If Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount > 0 Then
                                For i = 1 To Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount
                                    If x >= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).x And x <= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).x + Item(Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).ItemNum).FurnitureWidth - 1 Then
                                        If y <= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).y And y >= Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).y - Item(Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(i).ItemNum).FurnitureHeight + 1 Then
                                            'Blocked!
                                            Exit Sub
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    Next
                End If
                
                x = x1
                y = y1

                'If all checks out, place furniture and send the update to everyone in the player's house.
                Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount = Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount + 1
                ReDim Preserve Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount)
                Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount).ItemNum = ItemNum
                Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount).x = x
                Player(Index).characters(TempPlayer(Index).CurChar).House.Furniture(Player(Index).characters(TempPlayer(Index).CurChar).House.FurnitureCount).y = y
                
                
                Call TakeInvItem(Index, ItemNum, 0)
                
                SendFurnitureToHouse Player(Index).characters(TempPlayer(Index).CurChar).InHouse
        End If
    Else
        PlayerMsg Index, "You cannot place furniture unless you are in your own house!", BrightRed
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlaceFurniture", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleSendMail(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, receiver As String, body As String, login As String, filename As String, i As Long, itemslot As Long, itemvalue As Long, ItemNum As Long, slot As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    receiver = Buffer.ReadString
    body = Buffer.ReadString
    itemslot = Buffer.ReadLong
    itemvalue = Buffer.ReadLong
    
    login = GetAccountFromCharacterName(receiver, slot)
    
    If Trim$(login) <> "" Then
        If itemslot > 0 Then
            ItemNum = GetPlayerInvItemNum(Index, itemslot)
            If TakeInvItem(Index, GetPlayerInvItemNum(Index, itemslot), itemvalue) = False Then
                Exit Sub
            End If
        End If
        filename = App.path & "\data\accounts\" & Trim$(login) & "\" & Trim$(login) & "_char" & CStr(slot) & "_mail.ini"
        i = Val(GetVar(filename, "Mail", "MessageCount"))
        i = i + 1
        PutVar filename, "Mail", "MessageCount", CStr(i)
        PutVar filename, "Message" & CStr(i), "Deleted", CStr(0)
        PutVar filename, "Message" & CStr(i), "Unread", CStr(1)
        PutVar filename, "Message" & CStr(i), "From", Trim$(Player(Index).characters(TempPlayer(Index).CurChar).Name)
        PutVar filename, "Message" & CStr(i), "Body", Trim$(Replace(body, vbNewLine, Chr$(237)))
        PutVar filename, "Message" & CStr(i), "ItemNum", CStr(ItemNum)
        PutVar filename, "Message" & CStr(i), "ItemVal", CStr(itemvalue)
        PutVar filename, "Message" & CStr(i), "Date", Format$(Now, "mm/dd/yy") & " at " & Format$(Now, "hh:mm AM/PM")
        If FindPlayer(Trim$(receiver)) > 0 Then
            SendUnreadMail FindPlayer(Trim$(receiver))
            SendMailBox FindPlayer(Trim$(receiver)), 0
        End If
    Else
        PlayerMsg Index, "Could not find player, " & Trim$(receiver) & ", letter not sent!", BrightRed
    End If
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSendMail", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleDeleteMail(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, filename As String
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    filename = App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & "_char" & CStr(TempPlayer(Index).CurChar) & "_mail.ini"
    PutVar filename, "Message" & CStr(i), "Deleted", CStr(1)
    Set Buffer = Nothing
    
    SendMailBox Index, 0
    SendUnreadMail Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleDeleteMail", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleReadMail(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, filename As String
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    filename = App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & "_char" & CStr(TempPlayer(Index).CurChar) & "_mail.ini"
    PutVar filename, "Message" & CStr(i), "Unread", CStr(0)
    Set Buffer = Nothing
    
    SendMailBox Index, 0
    SendUnreadMail Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleReadMail", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleTakeMailItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, filename As String, ItemNum As Long, itemvalue As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    filename = App.path & "\data\accounts\" & Trim$(Player(Index).login) & "\" & Trim$(Player(Index).login) & "_char" & CStr(TempPlayer(Index).CurChar) & "_mail.ini"
    ItemNum = Val(GetVar(filename, "Message" & CStr(i), "Itemnum"))
    itemvalue = Val(GetVar(filename, "Message" & CStr(i), "Itemval"))
    If ItemNum > 0 And itemvalue > 0 Then
        If GiveInvItem(Index, ItemNum, itemvalue, True) = True Then
            Call PutVar(filename, "Message" & CStr(i), "Itemnum", CStr(0))
            Call PutVar(filename, "Message" & CStr(i), "Itemval", CStr(0))
        End If
    End If
    Set Buffer = Nothing
    SendMailBox Index, 0


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleTakeMailItem", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRestartServer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, filename As String, ItemNum As Long, itemvalue As Long
    

   On Error GoTo errorhandler
    
    If Index <> 0 Then
        If Player(Index).characters(TempPlayer(Index).CurChar).access < ADMIN_CREATOR Then
            PlayerMsg Index, "You do not have the access needed to do that!", BrightRed
        End If
    End If
    
    If Options.DisableRemoteRestart = 1 Then Exit Sub
    
    If isShuttingDown = True Then
        shutDownType = 0
        isShuttingDown = False
        frmServer.cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        isShuttingDown = True
        frmServer.cmdShutDown.Caption = "Cancel"
        shutDownType = 1
        GlobalMsg "Server restart initiated for an update.", BrightRed
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRestartServer", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleNewMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, curmap As Long, Dir As Long, openmap As Long
        ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Dir = Buffer.ReadLong
    curmap = Player(Index).characters(TempPlayer(Index).CurChar).Map
    For i = curmap To MAX_MAPS
        If LenB(Trim$(Map(i).Name)) = 0 Or Trim$(Map(i).Name) = "New Map" Then
            openmap = i
        End If
        If openmap > 0 Then i = MAX_MAPS
    Next
    If openmap > 0 Then
        Select Case Dir
            Case DIR_UP
                Map(Player(Index).characters(TempPlayer(Index).CurChar).Map).Up = openmap
                Map(openmap).Down = Player(Index).characters(TempPlayer(Index).CurChar).Map
            Case DIR_DOWN
                Map(Player(Index).characters(TempPlayer(Index).CurChar).Map).Down = openmap
                Map(openmap).Up = Player(Index).characters(TempPlayer(Index).CurChar).Map
            Case DIR_LEFT
                Map(Player(Index).characters(TempPlayer(Index).CurChar).Map).Left = openmap
                Map(openmap).Right = Player(Index).characters(TempPlayer(Index).CurChar).Map
            Case DIR_RIGHT
                Map(Player(Index).characters(TempPlayer(Index).CurChar).Map).Right = openmap
                Map(openmap).Left = Player(Index).characters(TempPlayer(Index).CurChar).Map
        End Select
        Map(openmap).Revision = Map(openmap).Revision + 1
        Map(openmap).Name = "Linked Map"
        Map(curmap).Revision = Map(curmap).Revision + 1
        MapCache_Create openmap
        MapCache_Create curmap
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Player(i).characters(TempPlayer(i).CurChar).Map = openmap Or Player(i).characters(TempPlayer(i).CurChar).Map = curmap Then
                    PlayerWarp i, Player(i).characters(TempPlayer(i).CurChar).Map, Player(i).characters(TempPlayer(i).CurChar).x, Player(i).characters(TempPlayer(i).CurChar).y
                End If
            End If
        Next
        PlayerMsg Index, "Map created!", White
    Else
        PlayerMsg Index, "No open map found!", BrightRed
    End If
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleNewMap", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub HandleEditFriend(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, Name As String, x As Long, blankname As String * ACCOUNT_LENGTH


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Name = Buffer.ReadString
    If Buffer.ReadLong = 0 Then 'Add Friend
        For i = 1 To 25
            If Trim$(LCase(Player(Index).characters(TempPlayer(Index).CurChar).Friends(i))) = Trim$(LCase(Name)) Then
                'Already have that friend!
                PlayerMsg Index, "You have already added " & Trim$(Name) & " as a friend.", BrightRed
                x = 30
                Exit For
            ElseIf Trim$(Player(Index).characters(TempPlayer(Index).CurChar).Friends(i)) = "" Or Trim$(Player(Index).characters(TempPlayer(Index).CurChar).Friends(i)) = Trim$(blankname) Then
                If x = 0 Then x = i
            End If
        Next
        If x > 0 Then
            If x > 25 Then
            
            Else
                If GetAccountFromCharacterName(Trim$(Name), 0) <> "" Then
                    PlayerMsg Index, "You have added " & Trim$(Name) & " as a friend!", BrightBlue
                    Player(Index).characters(TempPlayer(Index).CurChar).Friends(x) = Trim$(Name)
                Else
                    PlayerMsg Index, "Could not find player " & Trim$(Name) & ". Please check for spelling mistakes and that the friend you are trying to add exists.", BrightRed
                End If
            End If
        Else
            PlayerMsg Index, "There is no more room on your friends list!", BrightRed
        End If
    Else 'Delete Friend
        For i = 1 To 25
            If Trim$(Player(Index).characters(TempPlayer(Index).CurChar).Friends(i)) = Trim$(Name) Then
                'Remove the Friend!
                Player(Index).characters(TempPlayer(Index).CurChar).Friends(i) = ""
                Exit For
            Else
                If i = 25 Then
                    PlayerMsg Index, "Could not find " & Trim$(Name) & " on your friends list!", BrightRed
                End If
            End If
        Next
    End If
    SavePlayer Index
    SendPlayerFriends Index
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEditFriend", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub


Sub HandleRequestEditQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    
    SendSwitchesAndVariables (Index)
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditQuest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleSaveQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte

   On Error GoTo errorhandler

Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    
    n = Buffer.ReadLong 'CLng(Parse(1))
    
    If n < 0 Or n > MAX_QUESTS Then
        Exit Sub
    End If
    
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateQuestToAll(n)
    Call SaveQuest(n)
    Call AddLog(GetPlayerName(Index) & " saved Quest #" & n & ".", ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveQuest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestQuests(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendQuests Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestQuests", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandlePlayerHandleQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim questnum As Long, Order As Long, i As Long, n As Long, testeditems(1 To MAX_INV) As Long
    Dim RemoveStartItems As Boolean
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    questnum = Buffer.ReadLong
    Order = Buffer.ReadLong '1 = accept quest, 2 = cancel quest
    
    If Order = 1 Then
        'Accept Quest
        If CanStartQuest(Index, questnum, True) Then
            Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_STARTED '1
            Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).CurrentTask = 1
            For i = 1 To 5
                Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i) = 0
            Next
            PlayerMsg Index, "New quest accepted: " & Trim$(Quest(questnum).Name) & "!", BrightGreen
            For i = 1 To MAX_INV
                For n = 1 To MAX_INV
                    If testeditems(n) = Player(Index).characters(TempPlayer(Index).CurChar).Inv(i).Num Then
                        n = MAX_INV
                    Else
                        If n = MAX_INV Then
                            CheckTasks Index, TASK_AQUIREITEMS, Player(Index).characters(TempPlayer(Index).CurChar).Inv(i).Num
                            CheckTasks Index, TASK_FETCHRETURN, Player(Index).characters(TempPlayer(Index).CurChar).Inv(i).Num
                            testeditems(n) = Player(Index).characters(TempPlayer(Index).CurChar).Inv(i).Num
                        End If
                    End If
                Next
            Next
        End If
    ElseIf Order = 2 Then
        If Quest(questnum).Abandonable = 1 Then
            Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_NOT_STARTED '2
            Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).CurrentTask = 1
            For i = 1 To 5
                Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i) = 0
            Next
            RemoveStartItems = True
            PlayerMsg Index, Trim$(Quest(questnum).Name) & " has been canceled!", BrightGreen
        Else
            Exit Sub
        End If
    End If
    
    If RemoveStartItems Then
        For i = 0 To 3
            If Quest(questnum).GiveItemBefore(i).Item > 0 Then
                TakeInvItem Index, Quest(questnum).GiveItemBefore(i).Item, Quest(questnum).GiveItemBefore(i).Value
            End If
        Next
    End If
    
    SavePlayer Index
    SendPlayerData Index
    SendPlayerQuests Index
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePlayerHandleQuest", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleQuestLogUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendPlayerQuests Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleQuestLogUpdate", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestEditZone(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, x As Long

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SZoneEdit
    For i = 1 To MAX_ZONES
        Buffer.WriteString Trim$(MapZones(i).Name)
        Buffer.WriteLong MapZones(i).MapCount
        If MapZones(i).MapCount > 0 Then
            For x = 1 To MapZones(i).MapCount
                Buffer.WriteLong MapZones(i).Maps(x)
            Next
        End If
        For x = 1 To MAX_MAP_NPCS * 2
            Buffer.WriteLong MapZones(i).NPCs(x)
        Next
        For x = 1 To 5
            Buffer.WriteByte MapZones(i).Weather(x)
        Next
        Buffer.WriteByte MapZones(i).WeatherIntensity
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditZone", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleSaveZones(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, x As Long, Count As Long, z As Long

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data
    Count = Buffer.ReadLong
    If Count > 0 Then
        For z = 1 To Count
            i = Buffer.ReadLong
            MapZones(i).Name = Trim$(Buffer.ReadString)
            MapZones(i).MapCount = Buffer.ReadLong
            ReDim MapZones(i).Maps(MapZones(i).MapCount)
            If MapZones(i).MapCount > 0 Then
                For x = 1 To MapZones(i).MapCount
                    MapZones(i).Maps(x) = Buffer.ReadLong
                Next
            End If
            For x = 1 To MAX_MAP_NPCS * 2
                MapZones(i).NPCs(x) = Buffer.ReadLong
            Next
            For x = 1 To 5
                MapZones(i).Weather(x) = Buffer.ReadByte
            Next
            MapZones(i).WeatherIntensity = Buffer.ReadByte
            Call SaveZone(i)
            Call SpawnZoneNpcs(i)
            For x = 1 To Player_HighIndex
                If IsPlaying(x) Then
                    If GetPlayerMap(x) = i Then
                        PlayerWarp x, GetPlayerMap(x), GetPlayerX(x), GetPlayerY(x)
                    End If
                End If
            Next
        Next
    End If
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveZones", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim Buffer As clsBuffer
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ServerOnline = False
    Msg = Buffer.ReadString 'Parse(1)
    
    Set Buffer = Nothing
    Call MsgBox(Msg, vbOKOnly, Options.Game_Name)
    SaveSetting "Eclipse Origins", "Server" & Options.Key, "Auto", "0"
    End


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub









Sub HandleEditAccountLogin(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, i As Long, x As Long
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    frmServer.lblEmail.Caption = "Email: " & Trim$(Buffer.ReadString)
    frmServer.lblAccountFName.Caption = "First Name: " & Trim$(Buffer.ReadString)
    frmServer.lblAccountLName.Caption = "Last Name: " & Trim$(Buffer.ReadString)
    i = Buffer.ReadLong
    frmServer.txtActivateCode.Text = ""
    frmServer.lblLicenseCount.Caption = "License Count: " & i
    frmServer.lstLicenses.Clear
    If i > 0 Then
        For x = 1 To i
            If x < 10 Then
                frmServer.lstLicenses.AddItem ("License 0" & x & " -  " & Trim$(Buffer.ReadString))
            ElseIf x < 100 Then
                frmServer.lstLicenses.AddItem ("License " & x & " -  " & Trim$(Buffer.ReadString))
            End If
        Next
    End If
    frmServer.fraLogin.Visible = False
    frmServer.picEditAccount.Visible = True

    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEditAccountLogin", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub HandleAdmin(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
   On Error GoTo errorhandler

    If GetPlayerAccess(Index) > 0 Then
        SendAdmin (Index)
    End If
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleAdmin", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleServerOpts(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If
    SendServerOpts Index
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleServerOpts", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleSaveServerOpt(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
   Dim Buffer As clsBuffer, i As Long, x As Long, s As String, s1 As String
   Dim iFileNumber As Integer
    

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    'i is the case number, I do not feel like making 20 different packets for editing simple server options...
    i = Buffer.ReadLong
    Select Case i
        Case 1 'News
            s = Buffer.ReadString
            News = s
            frmServer.txtNews.Text = News
            iFileNumber = FreeFile
            Open App.path & "\data\news.txt" For Output As #iFileNumber
            Print #iFileNumber, News
            Close #iFileNumber
        Case 2 'Credits
            s = Buffer.ReadString
            Credits = s
            frmServer.txtCredits.Text = Credits
            iFileNumber = FreeFile
            Open App.path & "\data\credits.txt" For Output As #iFileNumber
            Print #iFileNumber, Credits
            Close #iFileNumber
        Case 3 'MOTD
            s = Buffer.ReadString
            Options.MOTD = s
            SaveOptions
        Case 4 'Server Options
            s = Buffer.ReadString
            s1 = Buffer.ReadString
            Options.Game_Name = s
            Options.Website = s1
            SaveOptions
        Case 5 'Update Info
            s = Buffer.ReadString
            s1 = Buffer.ReadString
            Options.DataFolder = s
            Options.UpdateURL = s1
            frmServer.txtUpdateUrl.Text = s1
            frmServer.txtDataFolder.Text = s
            SaveOptions
        Case 6
            frmServer.cmdReloadClasses_Click
        Case 7
            frmServer.cmdReloadSpells_Click
        Case 8
            frmServer.cmdReloadNPCs_Click
        Case 9
            frmServer.cmdReloadResources_Click
        Case 10
            frmServer.cmdReloadMaps_Click
        Case 11
            frmServer.cmdReloadShops_Click
        Case 12
            frmServer.cmdReloadItems_Click
        Case 13
            frmServer.cmdReloadAnimations_Click
        Case 14
            If Buffer.ReadLong = 1 Then
                Options.StaffOnly = 1
                frmServer.chkStaffOnly.Value = 1
                SaveOptions
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If GetPlayerAccess(i) = 0 Then
                            AlertMsg i, "The server has been switched to staff only! Please come back later."
                        End If
                    End If
                Next
            Else
                Options.StaffOnly = 0
                frmServer.chkStaffOnly.Value = 0
                SaveOptions
            End If
    End Select
    Set Buffer = Nothing
    
    SendServerOpts Index
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveServerOpt", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestEditHouse(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, x As Long

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHouseEdit
    For i = 1 To MAX_HOUSES
        Buffer.WriteString Trim$(HouseConfig(i).ConfigName)
        Buffer.WriteLong HouseConfig(i).BaseMap
        Buffer.WriteLong HouseConfig(i).x
        Buffer.WriteLong HouseConfig(i).y
        Buffer.WriteLong HouseConfig(i).price
        Buffer.WriteLong HouseConfig(i).MaxFurniture
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditHouse", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub HandleSaveHouses(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, x As Long, Count As Long, z As Long

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data
    Count = Buffer.ReadLong
    If Count > 0 Then
        For z = 1 To Count
            i = Buffer.ReadLong
            HouseConfig(i).ConfigName = Trim$(Buffer.ReadString)
            HouseConfig(i).BaseMap = Buffer.ReadLong
            HouseConfig(i).x = Buffer.ReadLong
            HouseConfig(i).y = Buffer.ReadLong
            HouseConfig(i).price = Buffer.ReadLong
            HouseConfig(i).MaxFurniture = Buffer.ReadLong
            Call SaveHouse(i)
            For x = 1 To Player_HighIndex
                If IsPlaying(x) Then
                    If Player(x).characters(TempPlayer(x).CurChar).InHouse = i Then
                        
                    End If
                End If
            Next
        Next
    End If
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveHouses", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleEditPlayer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, x As Long, z As Long, y As Long, char As String, acc As String

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data
    i = Buffer.ReadLong
    Select Case i
        Case 0 'Online Player
            char = Trim$(Buffer.ReadString)
            x = FindPlayer(Trim$(char))
            Set Buffer = Nothing
            If x > 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEditPlayer
                Buffer.WriteLong 0
                Buffer.WriteString Trim$(Player(x).login)
                Buffer.WriteString char
                Buffer.WriteString GetPlayerName(x)
                Buffer.WriteLong GetPlayerLevel(x)
                Buffer.WriteLong GetPlayerPOINTS(x)
                For i = 1 To FaceEnum.Face_Count - 1
                    Buffer.WriteLong Player(x).characters(TempPlayer(x).CurChar).Face(i)
                Next
                For i = 1 To SpriteEnum.Sprite_Count - 1
                    Buffer.WriteLong Player(x).characters(TempPlayer(x).CurChar).Sprite(i)
                Next
                Buffer.WriteLong Player(x).characters(TempPlayer(x).CurChar).Sex
                Buffer.WriteLong GetPlayerMap(x)
                Buffer.WriteLong GetPlayerX(x)
                Buffer.WriteLong GetPlayerY(x)
                Buffer.WriteLong GetPlayerDir(x)
                Buffer.WriteLong GetPlayerAccess(x)
                Buffer.WriteLong GetPlayerPK(x)
                Buffer.WriteLong GetPlayerClass(x)
                Buffer.WriteLong Player(x).characters(TempPlayer(x).CurChar).InHouse
                For i = 1 To Stats.Stat_Count - 1
                    Buffer.WriteLong GetPlayerStat(x, i)
                Next
                Buffer.WriteLong GetPlayerExp(x)
                Buffer.WriteLong GetPlayerVital(x, HP)
                Buffer.WriteLong GetPlayerVital(x, MP)
                For i = 1 To Equipment.Equipment_Count - 1
                    Buffer.WriteLong GetPlayerEquipment(x, i)
                Next
                For i = 1 To MAX_PLAYER_SPELLS
                    Buffer.WriteLong GetPlayerSpell(x, i)
                Next
                For i = 1 To MAX_INV
                    Buffer.WriteLong GetPlayerInvItemNum(x, i)
                    Buffer.WriteLong GetPlayerInvItemValue(x, i)
                Next
                SendDataTo Index, Buffer.ToArray
                Set Buffer = Nothing
            End If
        Case 1 'Account (maybe online but no character selected)
            If GetPlayerAccess(Index) < ADMIN_CREATOR Then
                Exit Sub
            End If
            acc = Trim$(Buffer.ReadString)
            x = FindAccount(Trim$(acc))
            Set Buffer = Nothing
            If x > 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEditPlayer
                Buffer.WriteLong 1
                Buffer.WriteString acc
                For z = 1 To MAX_PLAYER_CHARS
                    Buffer.WriteString Trim$(account(x).characters(z))
                Next
                SendDataTo Index, Buffer.ToArray
                Set Buffer = Nothing
            End If
        Case 2 'Account and Character (maybe online, maybe not)
            If GetPlayerAccess(Index) < ADMIN_CREATOR Then
                Exit Sub
            End If
            acc = Trim$(Buffer.ReadString)
            z = Buffer.ReadLong
            Set Buffer = Nothing
            x = FindAccount(acc)
            If x > 0 Then
                If Trim$(account(x).characters(z)) = "" Then
                    PlayerMsg Index, "No character found in slot " & z & ".", BrightRed
                    Exit Sub
                Else
                    char = Trim$(account(x).characters(z))
                    y = FindPlayer(char)
                    If y > 0 Then
                        Set Buffer = New clsBuffer
                        Buffer.WriteLong SEditPlayer
                        Buffer.WriteLong 2
                        Buffer.WriteString acc
                        Buffer.WriteString char
                        Buffer.WriteString GetPlayerName(y)
                        Buffer.WriteLong GetPlayerLevel(y)
                        Buffer.WriteLong GetPlayerPOINTS(y)
                        For i = 1 To FaceEnum.Face_Count - 1
                            Buffer.WriteLong Player(y).characters(TempPlayer(y).CurChar).Face(i)
                        Next
                        For i = 1 To SpriteEnum.Sprite_Count - 1
                            Buffer.WriteLong Player(y).characters(TempPlayer(y).CurChar).Sprite(i)
                        Next
                        Buffer.WriteLong Player(y).characters(TempPlayer(y).CurChar).Sex
                        Buffer.WriteLong GetPlayerMap(y)
                        Buffer.WriteLong GetPlayerX(y)
                        Buffer.WriteLong GetPlayerY(y)
                        Buffer.WriteLong GetPlayerDir(y)
                        Buffer.WriteLong GetPlayerAccess(y)
                        Buffer.WriteLong GetPlayerPK(y)
                        Buffer.WriteLong GetPlayerClass(y)
                        Buffer.WriteLong Player(y).characters(TempPlayer(y).CurChar).InHouse
                        For i = 1 To Stats.Stat_Count - 1
                            Buffer.WriteLong GetPlayerStat(y, i)
                        Next
                        Buffer.WriteLong GetPlayerExp(y)
                        Buffer.WriteLong GetPlayerVital(y, HP)
                        Buffer.WriteLong GetPlayerVital(y, MP)
                        For i = 1 To Equipment.Equipment_Count - 1
                            Buffer.WriteLong GetPlayerEquipment(y, i)
                        Next
                        For i = 1 To MAX_PLAYER_SPELLS
                            Buffer.WriteLong GetPlayerSpell(y, i)
                        Next
                        For i = 1 To MAX_INV
                            Buffer.WriteLong GetPlayerInvItemNum(y, i)
                            Buffer.WriteLong GetPlayerInvItemValue(y, i)
                        Next
                            
                        SendDataTo Index, Buffer.ToArray
                        Set Buffer = Nothing
                        Exit Sub
                    Else
                        LoadPlayer 0, acc
                        For i = 1 To MAX_PLAYER_CHARS
                            If Trim$(Player(0).characters(i).Name) = char Then
                                TempPlayer(0).CurChar = i
                                Set Buffer = New clsBuffer
                                Buffer.WriteLong SEditPlayer
                                Buffer.WriteLong 3
                                Buffer.WriteString acc
                                Buffer.WriteString char
                                Buffer.WriteString GetPlayerName(0)
                                Buffer.WriteLong GetPlayerLevel(0)
                                Buffer.WriteLong GetPlayerPOINTS(0)
                                For y = 1 To FaceEnum.Face_Count - 1
                                    Buffer.WriteLong Player(0).characters(i).Face(y)
                                Next
                                For y = 1 To SpriteEnum.Sprite_Count - 1
                                    Buffer.WriteLong Player(0).characters(i).Sprite(y)
                                Next
                                Buffer.WriteLong Player(0).characters(TempPlayer(0).CurChar).Sex
                                Buffer.WriteLong GetPlayerMap(0)
                                Buffer.WriteLong GetPlayerX(0)
                                Buffer.WriteLong GetPlayerY(0)
                                Buffer.WriteLong GetPlayerDir(0)
                                Buffer.WriteLong GetPlayerAccess(0)
                                Buffer.WriteLong GetPlayerPK(0)
                                Buffer.WriteLong GetPlayerClass(0)
                                Buffer.WriteLong Player(0).characters(TempPlayer(0).CurChar).InHouse
                                For y = 1 To Stats.Stat_Count - 1
                                    Buffer.WriteLong GetPlayerStat(0, y)
                                Next
                                Buffer.WriteLong GetPlayerExp(0)
                                Buffer.WriteLong GetPlayerVital(0, HP)
                                Buffer.WriteLong GetPlayerVital(0, MP)
                                For y = 1 To Equipment.Equipment_Count - 1
                                    Buffer.WriteLong GetPlayerEquipment(0, y)
                                Next
                                For y = 1 To MAX_PLAYER_SPELLS
                                    Buffer.WriteLong GetPlayerSpell(0, y)
                                Next
                                For y = 1 To MAX_INV
                                    Buffer.WriteLong GetPlayerInvItemNum(0, y)
                                    Buffer.WriteLong GetPlayerInvItemValue(0, y)
                                Next
                                SendDataTo Index, Buffer.ToArray
                                Set Buffer = Nothing
                                Exit Sub
                            Else
                                If i = MAX_PLAYER_CHARS Then
                                    PlayerMsg Index, "Error, could not find account or character!", BrightRed
                                End If
                            End If
                        Next
                    End If
                End If
            Else
                PlayerMsg Index, "Error, could not find account or character!", BrightRed
            End If
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEditPlayer", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleSavePlayer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, x As Long, z As Long, y As Long, char As String, acc As String
    Dim a As Long
    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data
    acc = Trim$(Buffer.ReadString)
    char = Trim$(Buffer.ReadString)
    If FindPlayer(Trim$(char)) > 0 Then
        i = FindPlayer(Trim$(char))
        z = i
        x = TempPlayer(i).CurChar
    Else
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            PlayerMsg Index, "Account not found or player is offline!", BrightRed
            Exit Sub
        End If
        i = FindAccount(acc)
        If i > 0 Then
            LoadPlayer 0, acc
            For i = 1 To MAX_PLAYER_CHARS
                If Trim$(Player(0).characters(i).Name) = char Then
                    z = 0
                    x = i
                    Exit For
                Else
                    If i = MAX_PLAYER_CHARS Then
                        PlayerMsg Index, "An error occured while editing player. Server could not save changes.", BrightRed
                        Exit Sub
                    End If
                End If
            Next
        Else
            PlayerMsg Index, "Account not found. Edit player has been canceled.", BrightRed
        End If
    End If
    
    'If still here then we are saving player info!
    With Player(z).characters(x)
        .Sex = Buffer.ReadLong
        .Class = Buffer.ReadLong
        .Level = Buffer.ReadLong
        .Exp = Buffer.ReadLong
        a = Buffer.ReadLong
        If a > .access Then
            If a <= GetPlayerAccess(Index) Then
                .access = a
            Else
                PlayerMsg Index, "You cannot set anothers access higher than yours! [Access Not Saved]", BrightRed
            End If
        Else
            If GetPlayerAccess(Index) > .access Then
                .access = a
            End If
        End If
        .PK = Buffer.ReadLong
        .Vital(Vitals.HP) = Buffer.ReadLong
        .Vital(Vitals.MP) = Buffer.ReadLong
        .stat(Stats.Strength) = Buffer.ReadLong
        .stat(Stats.Endurance) = Buffer.ReadLong
        .stat(Stats.Intelligence) = Buffer.ReadLong
        .stat(Stats.Agility) = Buffer.ReadLong
        .stat(Stats.Willpower) = Buffer.ReadLong
        .Points = Buffer.ReadLong
        .Equipment(Equipment.Weapon) = Buffer.ReadLong
        .Equipment(Equipment.armor) = Buffer.ReadLong
        .Equipment(Equipment.Helmet) = Buffer.ReadLong
        .Equipment(Equipment.Shield) = Buffer.ReadLong
        If z > 0 Then
            PlayerWarp z, Buffer.ReadLong, Buffer.ReadLong, Buffer.ReadLong
        End If
        .Dir = Buffer.ReadLong
        
        For i = 1 To MAX_PLAYER_SPELLS
            SetPlayerSpell z, i, Buffer.ReadLong
        Next
        
        For i = 1 To MAX_INV
            SetPlayerInvItemNum z, i, Buffer.ReadLong
            SetPlayerInvItemValue z, i, Buffer.ReadLong
        Next
        
        If z > 0 Then
            SendDataToMap GetPlayerMap(z), PlayerData(z)
            SendInventory z
            SendSpells z
            SendStats z
            SendVital z, HP
            SendVital z, MP
            SendEXP z
        End If
        
        SavePlayer z
    
    End With
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSavePlayer", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleEventTouch(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, a As Long

   On Error GoTo errorhandler
   
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data
    i = Buffer.ReadLong
    If Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).Trigger = 1 Then
        'Process this event, it is on-touch and everything checks out.
        If Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(TempPlayer(Index).EventMap.EventPages(i).eventID).pageID).CommandListCount > 0 Then
            TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).Active = 1
            TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).ActionTimer = GetTickCount
            TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).CurList = 1
            TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).CurSlot = 1
            TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).eventID = TempPlayer(Index).EventMap.EventPages(i).eventID
            TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).pageID = TempPlayer(Index).EventMap.EventPages(i).pageID
            TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).WaitingForResponse = 0
            ReDim TempPlayer(Index).EventProcessing(TempPlayer(Index).EventMap.EventPages(i).eventID).ListLeftOff(0 To Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).CommandListCount)
        End If
    End If
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleEventTouch", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleGameOpts(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If
    SendGameOpts Index
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleGameOpts", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleSaveGameOpt(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
   Dim Buffer As clsBuffer, i As Long, x As Long, s As String, s1 As String
   Dim iFileNumber As Integer
    

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    NewOptions.CombatMode = Buffer.ReadLong
    NewOptions.MaxLevel = Buffer.ReadLong
    NewOptions.MainMenuMusic = Buffer.ReadString
    NewOptions.ItemLoss = Buffer.ReadLong
    NewOptions.ExpLoss = Buffer.ReadLong
    SaveOptions
    MAX_LEVELS = NewOptions.MaxLevel
    SendMaxes 0
    Set Buffer = Nothing
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveGameOpt", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleMitigation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
   On Error GoTo errorhandler
    Dim mitigationstr As String, rate As Double
    
    If NewOptions.CombatMode = 1 Then
        mitigationstr = "Crit %: "
        rate = Fix(GetPlayerStat(Index, Agility) / 3)
        mitigationstr = mitigationstr & rate & "%"
        mitigationstr = mitigationstr & " Block %: "
        If GetPlayerEquipment(Index, Shield) > 0 Then
            rate = Fix(GetPlayerStat(Index, Strength) / 3)
            mitigationstr = mitigationstr & rate & "%"
        Else
            mitigationstr = mitigationstr & "0%"
        End If
        mitigationstr = mitigationstr & " Dodge %: "
        rate = Fix(GetPlayerStat(Index, Agility) / 4)
        mitigationstr = mitigationstr & rate & "%"
        mitigationstr = mitigationstr & " Parry %: "
        rate = Fix(GetPlayerStat(Index, Agility) / 6)
        mitigationstr = mitigationstr & rate & "%"
        PlayerMsg Index, "Mitigation: " & mitigationstr, White
    End If

    
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMitigation", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
' ::::::::::::::::::::::::::::::
' :: Request edit Pet  packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditPet(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPetEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditPet", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
' :::::::::::::::::::::
' :: Save pet packet ::
' :::::::::::::::::::::
Private Sub HandleSavePet(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim petNum As Long
    Dim Buffer As clsBuffer
    Dim PetSize As Long
    Dim PetData() As Byte, i As Long

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    petNum = Buffer.ReadLong

    ' Prevent hacking
    If petNum < 0 Or petNum > MAX_PETS Then
        Exit Sub
    End If

    With Pet(petNum)
        .Num = Buffer.ReadLong
        .Name = Buffer.ReadString
        .Sprite = Buffer.ReadLong
        .Range = Buffer.ReadLong
        .Level = Buffer.ReadLong
        .MaxLevel = Buffer.ReadLong
        .ExpGain = Buffer.ReadLong
        .LevelPnts = Buffer.ReadLong
        .StatType = Buffer.ReadByte
        .LevelingType = Buffer.ReadByte
        For i = 1 To Stats.Stat_Count - 1
            .stat(i) = Buffer.ReadByte
        Next
        For i = 1 To 4
            .Spell(i) = Buffer.ReadLong
        Next
    End With
    ' Save it
    Call SendUpdatePetToAll(petNum)
    Call Savepet(petNum)
    Call AddLog(GetPlayerName(Index) & " saved Pet #" & petNum & ".", ADMIN_LOG)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSavePet", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub HandleRequestPets(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendPets Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestPets", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Private Sub HandlePetMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, i As Long, z As Long

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    
        ' Prevent subscript out of range
    If x < 0 Or x > Map(GetPlayerMap(Index)).MaxX Or y < 0 Or y > Map(GetPlayerMap(Index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        If i = Index Then
                            ' Change target
                            If TempPlayer(Index).PetTargetType = TARGET_TYPE_PLAYER And TempPlayer(Index).PetTarget = i Then
                                TempPlayer(Index).PetTarget = 0
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                                TempPlayer(Index).PetTargetZone = 0
                                TempPlayer(Index).PetBehavior = PET_BEHAVIOUR_GOTO
                                TempPlayer(Index).GoToX = x
                                TempPlayer(Index).GoToY = y
                                ' send target to player
                                Call PlayerMsg(Index, "Your pet is no longer following you.", BrightRed)
                            Else
                                TempPlayer(Index).PetTarget = i
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_PLAYER
                                TempPlayer(Index).PetTargetZone = 0
                                ' send target to player
                                TempPlayer(Index).PetBehavior = PET_BEHAVIOUR_FOLLOW
                               Call PlayerMsg(Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & " is now following you.", Blue)
                            End If
                        Else
                            ' Change target
                            If TempPlayer(Index).PetTargetType = TARGET_TYPE_PLAYER And TempPlayer(Index).PetTarget = i Then
                                TempPlayer(Index).PetTarget = 0
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                                TempPlayer(Index).PetTargetZone = 0
                                ' send target to player
                                Call PlayerMsg(Index, "Your pet is no longer targetting " & Trim$(GetPlayerName(i)) & ".", BrightRed)
                            Else
                                TempPlayer(Index).PetTarget = i
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_PLAYER
                                TempPlayer(Index).PetTargetZone = 0
                                ' send target to player
                                Call PlayerMsg(Index, "Your pet is now targetting " & Trim$(GetPlayerName(i)) & ".", BrightRed)
                            End If
                        End If
                        Exit Sub
                    End If
                End If
                If Player(i).characters(TempPlayer(i).CurChar).Pet.Alive = True And i <> Index Then
                    If Player(i).characters(TempPlayer(i).CurChar).Pet.x = x Then
                        If Player(i).characters(TempPlayer(i).CurChar).Pet.y = y Then
                            ' Change target
                            If TempPlayer(Index).PetTargetType = TARGET_TYPE_PET And TempPlayer(Index).PetTarget = i Then
                                TempPlayer(Index).PetTarget = 0
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                                TempPlayer(Index).PetTargetZone = 0
                                ' send target to player
                                Call PlayerMsg(Index, "Your pet is no longer targetting " & Trim$(GetPlayerName(i)) & "'s " & Trim$(Pet(Player(i).characters(TempPlayer(i).CurChar).Pet.Num).Name) & ".", BrightRed)
                            Else
                                TempPlayer(Index).PetTarget = i
                                TempPlayer(Index).PetTargetType = TARGET_TYPE_PET
                                TempPlayer(Index).PetTargetZone = 0
                                ' send target to player
                                Call PlayerMsg(Index, "Your pet is now targetting " & Trim$(GetPlayerName(i)) & "'s " & Trim$(Pet(Player(i).characters(TempPlayer(i).CurChar).Pet.Num).Name) & ".", BrightRed)
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    'Search For Target First
    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index)).Npc(i).Num > 0 Then
            If MapNpc(GetPlayerMap(Index)).Npc(i).x = x Then
                If MapNpc(GetPlayerMap(Index)).Npc(i).y = y Then
                    If TempPlayer(Index).PetTarget = i And TempPlayer(Index).PetTargetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(Index).PetTarget = 0
                        TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                        TempPlayer(Index).PetTargetZone = 0
                        ' send target to player
                        Call PlayerMsg(Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & "'s target is no longer a " & Trim$(Npc(MapNpc(GetPlayerMap(Index)).Npc(i).Num).Name) & "!", BrightRed)
                        Exit Sub
                    Else
                        ' Change target
                        TempPlayer(Index).PetTarget = i
                        TempPlayer(Index).PetTargetType = TARGET_TYPE_NPC
                        TempPlayer(Index).PetTargetZone = 0
                        ' send target to player
                        Call PlayerMsg(Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & "'s target is now a " & Trim$(Npc(MapNpc(GetPlayerMap(Index)).Npc(i).Num).Name) & "!", BrightRed)
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    For i = 1 To MAX_ZONES
        For z = 1 To MAX_MAP_NPCS * 2
            If ZoneNpc(i).Npc(z).Num > 0 Then
                If ZoneNpc(i).Npc(z).Vital(HP) > 0 Then
                    If ZoneNpc(i).Npc(z).Map = GetPlayerMap(Index) Then
                        If ZoneNpc(i).Npc(z).x = x Then
                            If ZoneNpc(i).Npc(z).y = y Then
                                If TempPlayer(Index).PetTarget = z And TempPlayer(Index).PetTargetType = TARGET_TYPE_ZONENPC And TempPlayer(Index).PetTargetZone = i Then
                                    ' Change target
                                    TempPlayer(Index).PetTarget = 0
                                    TempPlayer(Index).PetTargetType = TARGET_TYPE_NONE
                                    TempPlayer(Index).PetTargetZone = 0
                                    Call PlayerMsg(Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & "'s target is no longer a " & Trim$(Npc(ZoneNpc(i).Npc(z).Num).Name) & "!", BrightRed)
                                Else
                                    ' Change target
                                    TempPlayer(Index).PetTarget = z
                                    TempPlayer(Index).PetTargetType = TARGET_TYPE_ZONENPC
                                    TempPlayer(Index).PetTargetZone = i
                                    Call PlayerMsg(Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & "'s target is now a " & Trim$(Npc(ZoneNpc(i).Npc(z).Num).Name) & "!", BrightRed)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next
    Next
    
    
    TempPlayer(Index).PetBehavior = PET_BEHAVIOUR_GOTO
    TempPlayer(Index).PetTargetType = 0
    TempPlayer(Index).PetTargetZone = 0
    TempPlayer(Index).PetTarget = 0
    TempPlayer(Index).GoToX = x
    TempPlayer(Index).GoToY = y
    If TempPlayer(Index).GoToX = Player(Index).characters(TempPlayer(Index).CurChar).Pet.x And TempPlayer(Index).GoToY = Player(Index).characters(TempPlayer(Index).CurChar).Pet.y Then
        Select Case Player(Index).characters(TempPlayer(Index).CurChar).Pet.AttackBehaviour
            Case PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_GUARD
                SendActionMsg GetPlayerMap(Index), "Defensive Mode!", White, 0, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32, Index
            Case PET_ATTACK_BEHAVIOUR_GUARD
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT
                SendActionMsg GetPlayerMap(Index), "Agressive Mode!", White, 0, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32, Index
            Case Else
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.AttackBehaviour = PET_ATTACK_BEHAVIOUR_ATTACKONSIGHT
                SendActionMsg GetPlayerMap(Index), "Agressive Mode!", White, 0, Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32, Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32, Index
        End Select
        TempPlayer(Index).GoToX = -1
        TempPlayer(Index).GoToY = -1
    Else
        Call PlayerMsg(Index, "Your " & Trim$(Pet(Player(Index).characters(TempPlayer(Index).CurChar).Pet.Num).Name) & " is moving to " & TempPlayer(Index).GoToX & "," & TempPlayer(Index).GoToY & ".", Blue)
    End If
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePetMove", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Private Sub HandleSetPetBehaviour(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = True Then Player(Index).characters(TempPlayer(Index).CurChar).Pet.AttackBehaviour = Buffer.ReadLong
    
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSetPetBehaviour", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Private Sub HandleReleasePet(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = True Then ReleasePet (Index)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleReleasePet", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub HandlePetSpell(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferPetSpell(Index, n)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandlePetSpell", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleUsePetStatPoint(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim Buffer As clsBuffer
Dim sMes As String
    

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    PointType = Buffer.ReadByte 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = False Then Exit Sub
    ' Make sure they have points
    If Player(Index).characters(TempPlayer(Index).CurChar).Pet.Points > 0 Then
        ' make sure they're not maxed#
        If Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) >= 255 Then
            PlayerMsg Index, "You cannot spend any more points on that stat for your pet.", BrightRed
            Exit Sub
        End If
        

        Player(Index).characters(TempPlayer(Index).CurChar).Pet.Points = Player(Index).characters(TempPlayer(Index).CurChar).Pet.Points - 1

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) + 1
                sMes = "Strength"
            Case Stats.Endurance
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) + 1
                sMes = "Endurance"
            Case Stats.Intelligence
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) + 1
                sMes = "Intelligence"
            Case Stats.Agility
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) + 1
                sMes = "Agility"
            Case Stats.Willpower
                Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) = Player(Index).characters(TempPlayer(Index).CurChar).Pet.stat(PointType) + 1
                sMes = "Willpower"
        End Select
        SendActionMsg GetPlayerMap(Index), "+1 " & sMes, White, 1, (Player(Index).characters(TempPlayer(Index).CurChar).Pet.x * 32), (Player(Index).characters(TempPlayer(Index).CurChar).Pet.y * 32)

    Else
        Exit Sub
    End If

    ' Send the update
    'Call SendStats(Index)
    SendPlayerData Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleUseStatPoint", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' ::::::::::::::::::::::::::::::::::::::::
' :: Request edit Random Dungeon packet ::
' ::::::::::::::::::::::::::::::::::::::::
Sub HandleRequestEditRandomDungeon(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking

   On Error GoTo errorhandler

    'removed


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditRandomDungeon", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' :::::::::::::::::::::::::
' :: Save Random Dungeon ::
' :::::::::::::::::::::::::
Private Sub HandleSaveRandomDungeon(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim DungeonNum As Long
    Dim Buffer As clsBuffer
    Dim DungeonSize As Long
    Dim DungeonData() As Byte

    ' Prevent hacking

   On Error GoTo errorhandler

    'Removed


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveRandomDungeon", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestRandomDungeon(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    'removed


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestRandomDungeon", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestEditProjectiles(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SProjectileEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestEditProjectiles", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub HandleSaveProjectile(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ProjectileNum As Long
    Dim Buffer As clsBuffer
    Dim ProjectileSize As Long
    Dim ProjectileData() As Byte

    ' Prevent hacking

   On Error GoTo errorhandler

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ProjectileNum = Buffer.ReadLong

    ' Prevent hacking
    If ProjectileNum < 0 Or ProjectileNum > MAX_PROJECTILES Then
        Exit Sub
    End If

    ProjectileSize = LenB(Projectiles(ProjectileNum))
    ReDim ProjectileData(ProjectileSize - 1)
    ProjectileData = Buffer.ReadBytes(ProjectileSize)
    CopyMemory ByVal VarPtr(Projectiles(ProjectileNum)), ByVal VarPtr(ProjectileData(0)), ProjectileSize
    ' Save it
    Call SendUpdateProjectileToAll(ProjectileNum)
    Call SaveProjectile(ProjectileNum)
    Call AddLog(GetPlayerName(Index) & " saved Projectile #" & ProjectileNum & ".", ADMIN_LOG)
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleSaveProjectile", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleRequestProjectiles(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

   On Error GoTo errorhandler

    SendProjectiles Index

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleRequestProjectiles", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleClearProjectile(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim ProjectileNum As Long
    Dim TargetIndex As Long
    Dim TargetType As Byte
    Dim TargetZone As Long
    Dim MapNum As Long
    Dim Damage As Long
    Dim armor As Long
    Dim i As Long
    Dim npcnum As Long
    
   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ProjectileNum = Buffer.ReadLong
    TargetIndex = Buffer.ReadLong
    TargetType = Buffer.ReadByte
    TargetZone = Buffer.ReadLong
    Set Buffer = Nothing
    
    MapNum = GetPlayerMap(Index)
    
    Select Case MapProjectiles(MapNum, ProjectileNum).OwnerType
        Case TARGET_TYPE_PLAYER
            If MapProjectiles(MapNum, ProjectileNum).Owner = Index Then
                Select Case TargetType
                    Case TARGET_TYPE_PLAYER
                    
                        If IsPlaying(TargetIndex) Then
                            If TargetIndex <> Index Then
                                If CanPlayerAttackPlayer(Index, TargetIndex, True) = True Then
                            
                                    ' Get the damage we can do
                                    Damage = GetPlayerDamage(Index) + Projectiles(MapProjectiles(MapNum, ProjectileNum).ProjectileNum).Damage
                                    If NewOptions.CombatMode = 1 Then
                                        For i = 1 To Equipment.Equipment_Count - 1
                                            If GetPlayerEquipment(TargetIndex, i) > 0 Then
                                                armor = armor + Item(GetPlayerEquipment(TargetIndex, i)).Data2
                                            End If
                                        Next
                                        ' take away armour
                                        Damage = Damage - ((GetPlayerStat(TargetIndex, Willpower) * 2) + (GetPlayerLevel(TargetIndex) * 3) + armor)
                                    Else
                                        ' if the npc blocks, take away the block amount
                                        armor = CanPlayerBlock(TargetIndex)
                                        Damage = Damage - armor
            
                                        ' take away armour
                                        Damage = Damage - rand(1, (GetPlayerStat(TargetIndex, Agility) * 2))
            
                                        ' randomise for up to 10% lower than max hit
                                        Damage = rand(1, Damage)
                                    End If
                            
                                    If Damage < 1 Then Damage = 1
                            
                                    PlayerAttackPlayer Index, TargetIndex, Damage
                                End If
                            End If
                        End If
                        
                    Case TARGET_TYPE_NPC
                        npcnum = MapNpc(MapNum).Npc(TargetIndex).Num
                        If CanPlayerAttackNpc(Index, TargetIndex, True) = True Then
                            ' Get the damage we can do
                            Damage = GetPlayerDamage(Index) + Projectiles(MapProjectiles(MapNum, ProjectileNum).ProjectileNum).Damage
                            If NewOptions.CombatMode = 1 Then
                                Damage = Damage - ((Npc(MapNpc(MapNum).Npc(TargetIndex).Num).stat(Stats.Willpower) * 2) + (Npc(MapNpc(MapNum).Npc(TargetIndex).Num).Level * 3))
                            Else
                                ' if the npc blocks, take away the block amount
                                armor = CanNpcBlock(npcnum)
                                Damage = Damage - armor
            
                                ' take away armour
                                Damage = Damage - rand(1, (Npc(MapNpc(MapNum).Npc(TargetIndex).Num).stat(Stats.Agility) * 2))
                                ' randomise from 1 to max hit
                                Damage = rand(1, Damage)
                            End If
                            
                            If Damage < 1 Then Damage = 1
                            
                            PlayerAttackNpc Index, TargetIndex, Damage
                        End If
                        
                    Case TARGET_TYPE_PET
                        If IsPlaying(TargetIndex) Then
                            If Player(TargetIndex).characters(TempPlayer(TargetIndex).CurChar).Pet.Alive = True Then
                                If TargetIndex <> Index Then
                                    If CanPlayerAttackPet(Index, TargetIndex, True) = True Then
                                        ' Get the damage we can do
                                        Damage = GetPlayerDamage(Index) + Projectiles(MapProjectiles(MapNum, ProjectileNum).ProjectileNum).Damage
        
                                        ' if the npc blocks, take away the block amount
                                        armor = 0
                                        Damage = Damage - armor
        
                                        ' take away armour
                                        Damage = Damage - rand(1, (GetPlayerStat(TargetIndex, Agility) * 2))
        
                                        ' randomise for up to 10% lower than max hit
                                        Damage = rand(1, Damage)
                                
                                        If Damage < 1 Then Damage = 1
                                
                                        PlayerAttackPet Index, TargetIndex, Damage
                                    End If
                                End If
                            End If
                        End If
                    Case TARGET_TYPE_ZONENPC
                        npcnum = ZoneNpc(TargetZone).Npc(TargetIndex).Num
                        If CanPlayerAttackZoneNpc(Index, TargetZone, TargetIndex, True) = True Then
                            ' Get the damage we can do
                            Damage = GetPlayerDamage(Index) + Projectiles(MapProjectiles(MapNum, ProjectileNum).ProjectileNum).Damage
                            If NewOptions.CombatMode = 1 Then
                                Damage = Damage - ((Npc(ZoneNpc(TargetZone).Npc(TargetIndex).Num).stat(Stats.Willpower) * 2) + (Npc(ZoneNpc(TargetZone).Npc(TargetIndex).Num).Level * 3))
                            Else
                                ' if the npc blocks, take away the block amount
                                armor = CanNpcBlock(npcnum)
                                Damage = Damage - armor
            
                                ' take away armour
                                Damage = Damage - rand(1, (Npc(ZoneNpc(TargetZone).Npc(TargetIndex).Num).stat(Stats.Agility) * 2))
                                ' randomise from 1 to max hit
                                Damage = rand(1, Damage)
                            End If
                            
                            If Damage < 1 Then Damage = 1
                            
                            PlayerAttackZoneNpc Index, TargetZone, TargetIndex, Damage
                        End If
                End Select
            End If

        Case TARGET_TYPE_PET
            If MapProjectiles(MapNum, ProjectileNum).Owner = Index And Player(Index).characters(TempPlayer(Index).CurChar).Pet.Alive = True Then
                Select Case TargetType
                    Case TARGET_TYPE_PLAYER
                    
                        If IsPlaying(TargetIndex) Then
                            If TargetIndex <> Index Then
                                If CanPetAttackPlayer(Index, TargetIndex, True) = True Then
                            
                                    ' Get the damage we can do
                                    Damage = GetPetDamage(Index)
                                    If NewOptions.CombatMode = 1 Then
                                        For i = 1 To Equipment.Equipment_Count - 1
                                            If GetPlayerEquipment(TargetIndex, i) > 0 Then
                                                armor = armor + Item(GetPlayerEquipment(TargetIndex, i)).Data2
                                            End If
                                        Next
                                        ' take away armour
                                        Damage = Damage - ((GetPlayerStat(TargetIndex, Willpower) * 2) + (GetPlayerLevel(TargetIndex) * 3) + armor)
                                    Else
                                        ' if the npc blocks, take away the block amount
                                        armor = CanPlayerBlock(TargetIndex)
                                        Damage = Damage - armor
            
                                        ' take away armour
                                        Damage = Damage - rand(1, (GetPlayerStat(TargetIndex, Agility) * 2))
            
                                        ' randomise for up to 10% lower than max hit
                                        Damage = rand(1, Damage)
                                    End If
                            
                                    If Damage < 1 Then Damage = 1
                            
                                    PetAttackPlayer Index, TargetIndex, Damage
                                End If
                            End If
                        End If
                        
                    Case TARGET_TYPE_NPC
                        npcnum = MapNpc(MapNum).Npc(TargetIndex).Num
                        If CanPetAttackNpc(Index, TargetIndex, True) = True Then
                            ' Get the damage we can do
                            Damage = GetPetDamage(Index)
                            If NewOptions.CombatMode = 1 Then
                                Damage = Damage - ((Npc(MapNpc(MapNum).Npc(TargetIndex).Num).stat(Stats.Willpower) * 2) + (Npc(MapNpc(MapNum).Npc(TargetIndex).Num).Level * 3))
                            Else
                                ' if the npc blocks, take away the block amount
                                armor = CanNpcBlock(npcnum)
                                Damage = Damage - armor
            
                                ' take away armour
                                Damage = Damage - rand(1, (Npc(MapNpc(MapNum).Npc(TargetIndex).Num).stat(Stats.Agility) * 2))
                                ' randomise from 1 to max hit
                                Damage = rand(1, Damage)
                            End If
                            
                            If Damage < 1 Then Damage = 1
                            
                            PetAttackNpc Index, TargetIndex, Damage
                        End If
                        
                    Case TARGET_TYPE_PET
                        If IsPlaying(TargetIndex) Then
                            If Player(TargetIndex).characters(TempPlayer(TargetIndex).CurChar).Pet.Alive = True Then
                                If TargetIndex <> Index Then
                                    If CanPetAttackPet(Index, TargetIndex, True) = True Then
                                        ' Get the damage we can do
                                        Damage = GetPetDamage(Index)
        
                                        ' if the npc blocks, take away the block amount
                                        armor = 0
                                        Damage = Damage - armor
        
                                        ' take away armour
                                        Damage = Damage - rand(1, (GetPlayerStat(TargetIndex, Agility) * 2))
        
                                        ' randomise for up to 10% lower than max hit
                                        Damage = rand(1, Damage)
                                
                                        If Damage < 1 Then Damage = 1
                                
                                        PetAttackPet Index, TargetIndex, Damage
                                    End If
                                End If
                            End If
                        End If
                    Case TARGET_TYPE_ZONENPC
                        npcnum = ZoneNpc(TargetZone).Npc(TargetIndex).Num
                        If CanPetAttackZoneNpc(Index, TargetZone, TargetIndex, True) = True Then
                            ' Get the damage we can do
                            Damage = GetPetDamage(Index) + Projectiles(MapProjectiles(MapNum, ProjectileNum).ProjectileNum).Damage
                            If NewOptions.CombatMode = 1 Then
                                Damage = Damage - ((Npc(ZoneNpc(TargetZone).Npc(TargetIndex).Num).stat(Stats.Willpower) * 2) + (Npc(ZoneNpc(TargetZone).Npc(TargetIndex).Num).Level * 3))
                            Else
                                ' if the npc blocks, take away the block amount
                                armor = CanNpcBlock(npcnum)
                                Damage = Damage - armor
            
                                ' take away armour
                                Damage = Damage - rand(1, (Npc(ZoneNpc(TargetZone).Npc(TargetIndex).Num).stat(Stats.Agility) * 2))
                                ' randomise from 1 to max hit
                                Damage = rand(1, Damage)
                            End If
                            
                            If Damage < 1 Then Damage = 1
                            
                            PetAttackZoneNpc Index, TargetZone, TargetIndex, Damage
                        End If
                End Select
            End If
            
        Case TARGET_TYPE_NPC
         If MapProjectiles(MapNum, ProjectileNum).Owner > 0 Then
                Select Case TargetType
                    Case TARGET_TYPE_PLAYER
                    
                        If IsPlaying(TargetIndex) Then
                            'If TargetIndex <> Index Then
                                If CanNpcAttackPlayer(Index, TargetIndex, True) = True Then
                            
                                    ' Get the damage we can do
                                    Damage = GetPlayerDamage(Index) + Projectiles(MapProjectiles(MapNum, ProjectileNum).ProjectileNum).Damage
                                    If NewOptions.CombatMode = 1 Then
                                        For i = 1 To Equipment.Equipment_Count - 1
                                            If GetPlayerEquipment(TargetIndex, i) > 0 Then
                                                armor = armor + Item(GetPlayerEquipment(TargetIndex, i)).Data2
                                            End If
                                        Next
                                        ' take away armour
                                        Damage = Damage - ((GetPlayerStat(TargetIndex, Willpower) * 2) + (GetPlayerLevel(TargetIndex) * 3) + armor)
                                    Else
                                        ' if the npc blocks, take away the block amount
                                        armor = CanPlayerBlock(TargetIndex)
                                        Damage = Damage - armor
            
                                        ' take away armour
                                        Damage = Damage - rand(1, (GetPlayerStat(TargetIndex, Agility) * 2))
            
                                        ' randomise for up to 10% lower than max hit
                                        Damage = rand(1, Damage)
                                    End If
                            
                                    If Damage < 1 Then Damage = 1
                            
                                    NpcAttackPlayer Index, TargetIndex, Damage
                                End If
                            End If
                        'End If

                End Select
            End If
    End Select

    ClearMapProjectile MapNum, ProjectileNum

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleClearProjectile", "modHandleData", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
