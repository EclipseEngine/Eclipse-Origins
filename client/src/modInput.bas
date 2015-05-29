Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

Public Sub CheckKeys()

   On Error GoTo errorhandler

    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_SPACE) >= 0 Then SpaceDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    If GetAsyncKeyState(VK_HOME) >= 0 Then HomeDown = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CheckInputKeys()

   On Error GoTo errorhandler

    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
    If GetKeyState(vbKeyHome) < 0 Then
        HomeDown = True
    Else
        HomeDown = False
        HomeUp = True
    End If
    
    SpaceDown = False
    If chatOn = False Then
        If GetKeyState(vbKeyE) < 0 Or (GetKeyState(vbKeyRButton) < 0 And Options.ClicktoWalk = 1) Then
            SpaceDown = True
            CheckMapGetItem
        End If
    End If

    'Move Up
    If GetKeyState(vbKeyUp) < 0 Or (GetKeyState(vbKeyW) < 0 And chatOn = False) < 0 Then
        DirUp = True
        DirDown = False
        DirLeft = False
        DirRight = False
        Exit Sub
    Else
        DirUp = False
    End If

    'Move Right
    If GetKeyState(vbKeyRight) < 0 Or (GetKeyState(vbKeyD) < 0 And chatOn = False) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = True
        Exit Sub
    Else
        DirRight = False
    End If

    'Move down
    If GetKeyState(vbKeyDown) < 0 Or (GetKeyState(vbKeyS) < 0 And chatOn = False) < 0 Then
        DirUp = False
        DirDown = True
        DirLeft = False
        DirRight = False
        Exit Sub
    Else
        DirDown = False
    End If

    'Move left
    If GetKeyState(vbKeyLeft) < 0 Or (GetKeyState(vbKeyA) < 0 And chatOn = False) Then
        DirUp = False
        DirDown = False
        DirLeft = True
        DirRight = False
        Exit Sub
    Else
        DirLeft = False
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
Dim i As Long
Dim n As Long
Dim Command() As String
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    If InGame = False Then
        HandleMenuKeypress (KeyAscii)
        Exit Sub
    End If
    If CurrencyMenu > 0 Then
        If KeyAscii = vbKeyReturn Then
            CurrencyOk
            Exit Sub
        Else
            If KeyAscii = vbKeyBack Then
                If Len(CurrencyText) > 0 Then
                    CurrencyText = Left(CurrencyText, Len(CurrencyText) - 1)
                End If
            Else
                CurrencyText = CurrencyText & Chr(KeyAscii)
            End If
            Exit Sub
        End If
    End If
    If InMailbox Then
        If MailBoxMenu = 3 Then
            If KeyAscii = vbKeyTab Then
                If SelTextbox = 1 Then
                    SelTextbox = 2
                Else
                    SelTextbox = 1
                End If
                Exit Sub
            ElseIf KeyAscii = vbKeyReturn Then
                If SelTextbox = 2 Then
                    MailContent = MailContent & vbNewLine
                    Exit Sub
                End If
            Else
                If SelTextbox = 1 Then
                    If KeyAscii = vbKeyBack Then
                        If Len(MailToFrom) > 0 Then
                            MailToFrom = Left(MailToFrom, Len(MailToFrom) - 1)
                        End If
                    Else
                        MailToFrom = MailToFrom & Chr(KeyAscii)
                    End If
                    Exit Sub
                ElseIf SelTextbox = 2 Then
                    If KeyAscii = vbKeyBack Then
                        If Len(MailContent) > 0 Then
                            MailContent = Left(MailContent, Len(MailContent) - 1)
                        End If
                    Else
                        MailContent = MailContent & Chr(KeyAscii)
                    End If
                    Exit Sub
                End If
            End If
        End If
    End If
    'ChatText = Trim$(MyText)
    'MyText = LCase$(ChatText)
    If Not chatOn Then
        If Chr(KeyAscii) = "c" Then
            If Player(MyIndex).X = 0 Then
                If Player(MyIndex).dir = DIR_LEFT Then
                    If MsgBox("Do you wanna create a linked map here?", vbYesNo) = vbYes Then
                        Set buffer = New clsBuffer
                        buffer.WriteLong CNewMap
                        buffer.WriteLong DIR_LEFT
                        SendData buffer.ToArray
                        Set buffer = Nothing
                    End If
                    Exit Sub
                End If
            End If
            If Player(MyIndex).X = Map.MaxX Then
                If Player(MyIndex).dir = DIR_RIGHT Then
                    If MsgBox("Do you wanna create a linked map here?", vbYesNo) = vbYes Then
                        Set buffer = New clsBuffer
                        buffer.WriteLong CNewMap
                        buffer.WriteLong DIR_RIGHT
                        SendData buffer.ToArray
                        Set buffer = Nothing
                    End If
                    Exit Sub
                End If
            End If
            If Player(MyIndex).Y = 0 Then
                If Player(MyIndex).dir = DIR_UP Then
                    If MsgBox("Do you wanna create a linked map here?", vbYesNo) = vbYes Then
                        Set buffer = New clsBuffer
                        buffer.WriteLong CNewMap
                        buffer.WriteLong DIR_UP
                        SendData buffer.ToArray
                        Set buffer = Nothing
                    End If
                    Exit Sub
                End If
            End If
            If Player(MyIndex).Y = Map.MaxY Then
                If Player(MyIndex).dir = DIR_DOWN Then
                    If MsgBox("Do you wanna create a linked map here?", vbYesNo) = vbYes Then
                        Set buffer = New clsBuffer
                        buffer.WriteLong CNewMap
                        buffer.WriteLong DIR_DOWN
                        SendData buffer.ToArray
                        Set buffer = Nothing
                    End If
                    Exit Sub
                End If
            End If
                    End If
    End If
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
        ChatText = Trim$(MyText)
        If EventChat = True Then
            If EventChatType = 0 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CEventChatReply
                buffer.WriteLong EventReplyID
                buffer.WriteLong EventReplyPage
                buffer.WriteLong 0
                SendData buffer.ToArray
                Set buffer = Nothing
                ClearEventChat
                InEvent = False
                Exit Sub
            End If
        End If
        If chatOn Then
            ' Broadcast message
            If Left$(ChatText, 1) = "'" Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                If Len(ChatText) > 0 Then
                    Call BroadcastMsg(ChatText)
                End If
                MyText = vbNullString
                UpdateShowChatText
                Exit Sub
            End If
            ' Emote message
            If Left$(ChatText, 1) = "-" Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                If Len(ChatText) > 0 Then
                    Call EmoteMsg(ChatText)
                End If
                MyText = vbNullString
                UpdateShowChatText
                Exit Sub
            End If
            ' Player message
            If Left$(ChatText, 1) = "!" Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                Name = vbNullString
                ' Get the desired player from the user text
                For i = 1 To Len(ChatText)
                    If Mid$(ChatText, i, 1) <> Space(1) Then
                        Name = Name & Mid$(ChatText, i, 1)
                    Else
                        Exit For
                    End If
                Next
                ' Make sure they are actually sending something
                If Len(ChatText) - i > 0 Then
                    ChatText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                    ' Send the message to the player
                    Call PlayerMsg(ChatText, Name)
                Else
                    Call AddText("Usage: !playername (message)", AlertColor)
                End If
                MyText = vbNullString
                UpdateShowChatText
                Exit Sub
            End If
            If Left$(MyText, 1) = "/" Then
                Command = Split(MyText, Space(1))
                Select Case Command(0)
                    Case "/help"
                        Call AddText("Social Commands:", HelpColor)
                        Call AddText("'msghere = Broadcast Message", HelpColor)
                        Call AddText("-msghere = Emote Message", HelpColor)
                        Call AddText("!namehere msghere = Player Message", HelpColor)
                        Call AddText("Available Commands: /info, /who, /fps, /fpslock", HelpColor)
                    Case "/invite"
                        If UBound(Command) > 0 Then
                            Set buffer = New clsBuffer
                            buffer.WriteLong CVisit
                            buffer.WriteString Command(1)
                            SendData buffer.ToArray
                            Set buffer = Nothing
                        End If
                    Case "/info"
                        ' Checks to make sure we have more than one string in the array
                        If UBound(Command) < 1 Then
                            AddText "Usage: /info (name)", AlertColor
                            GoTo Continue
                        End If
                        If IsNumeric(Command(1)) Then
                            AddText "Usage: /info (name)", AlertColor
                            GoTo Continue
                        End If
                        Set buffer = New clsBuffer
                        buffer.WriteLong CPlayerInfoRequest
                        buffer.WriteString Command(1)
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        ' Whos Online
                    Case "/who"
                        SendWhosOnline
                        ' Checking fps
                    Case "/fps"
                        BFPS = Not BFPS
                        If BFPS = True Then
                            frmAdmin.chkShowFPS.Value = 1
                        Else
                            frmAdmin.chkShowFPS.Value = 0
                        End If
                        ' toggle fps lock
                    Case "/fpslock"
                        FPS_Lock = Not FPS_Lock
                        ' Request stats
                    Case "/stats"
                        Set buffer = New clsBuffer
                        buffer.WriteLong CGetStats
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        ' // Monitor Admin Commands //
                        ' Admin Help
                    Case "/admin"
                        If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue
                        If frmAdmin.Visible = False Then
                            Set buffer = New clsBuffer
                            buffer.WriteLong CAdmin
                            SendData buffer.ToArray
                            Set buffer = Nothing
                        Else
                            frmAdmin.Visible = False
                        End If
                    Case "/mitigation"
                        Set buffer = New clsBuffer
                        buffer.WriteLong CMitigation
                        SendData buffer.ToArray
                        Set buffer = Nothing
                        ' Kicking a player
                    Case "/kick"
                        If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue
                        If UBound(Command) < 1 Then
                            AddText "Usage: /kick (name)", AlertColor
                            GoTo Continue
                        End If
                        If IsNumeric(Command(1)) Then
                            AddText "Usage: /kick (name)", AlertColor
                            GoTo Continue
                        End If
                        SendKick Command(1)
                        ' // Mapper Admin Commands //
                        ' Location
                    Case "/loc"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                        BLoc = Not BLoc
                        If BLoc = True Then
                            frmAdmin.chkShowLoc.Value = 1
                        Else
                            frmAdmin.chkShowLoc.Value = 0
                        End If
                        ' Map Editor
                    Case "/editmap"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                                            SendRequestEditMap
                    Case "/editzones"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                                            SendRequestEditZone
                        ' Warping to a player
                    Case "/warpmeto"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                        If UBound(Command) < 1 Then
                            AddText "Usage: /warpmeto (name)", AlertColor
                            GoTo Continue
                        End If
                        If IsNumeric(Command(1)) Then
                            AddText "Usage: /warpmeto (name)", AlertColor
                            GoTo Continue
                        End If
                        WarpMeTo Command(1)
                        ' Warping a player to you
                    Case "/warptome"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                        If UBound(Command) < 1 Then
                            AddText "Usage: /warptome (name)", AlertColor
                            GoTo Continue
                        End If
                        If IsNumeric(Command(1)) Then
                            AddText "Usage: /warptome (name)", AlertColor
                            GoTo Continue
                        End If
                        WarpToMe Command(1)
                        ' Warping to a map
                    Case "/warpto"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                        If UBound(Command) < 1 Then
                            AddText "Usage: /warpto (map #)", AlertColor
                            GoTo Continue
                        End If
                        If Not IsNumeric(Command(1)) Then
                            AddText "Usage: /warpto (map #)", AlertColor
                            GoTo Continue
                        End If
                        n = CLng(Command(1))
                        ' Check to make sure its a valid map #
                        If n > 0 And n <= MAX_MAPS Then
                            Call WarpTo(n)
                        Else
                            Call AddText("Invalid map number.", Red)
                        End If
                        ' Setting sprite
                    Case "/setsprite"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                        If UBound(Command) < 1 Then
                            AddText "Usage: /setsprite (sprite #)", AlertColor
                            GoTo Continue
                        End If
                        If Not IsNumeric(Command(1)) Then
                            AddText "Usage: /setsprite (sprite #)", AlertColor
                            GoTo Continue
                        End If
                        SendSetSprite CLng(Command(1))
                        ' Map report
                    Case "/mapreport"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                        SendMapReport
                        ' Respawn request
                    Case "/respawn"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                        SendMapRespawn
                        ' MOTD change
                    Case "/motd"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                        If UBound(Command) < 1 Then
                            AddText "Usage: /motd (new motd)", AlertColor
                            GoTo Continue
                        End If
                        SendMOTDChange Right$(ChatText, Len(ChatText) - 5)
                        ' Check the ban list
                    Case "/banlist"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                        SendBanList
                        ' Banning a player
                    Case "/ban"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                        If UBound(Command) < 1 Then
                            AddText "Usage: /ban (name)", AlertColor
                            GoTo Continue
                        End If
                        SendBan Command(1), InputBox("Input a reason for banning this player.", "Ban " & Command(1)), False
                        ' // Developer Admin Commands //
                        ' Editing item request
                    Case "/edititem"
                        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                        SendRequestEditItem
                    ' Editing animation request
                    Case "/editanimation"
                        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                        SendRequestEditAnimation
                        ' Editing npc request
                    Case "/editnpc"
                        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                        SendRequestEditNpc
                    Case "/editresource"
                        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                        SendRequestEditResource
                        ' Editing shop request
                    Case "/editshop"
                        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                        SendRequestEditShop
                        ' Editing spell request
                    Case "/editspell"
                        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                        SendRequestEditSpell
                        ' Editing spell request
                    Case "/editpet"
                        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                        SendRequestEditPet
                        'Edit projectiles
                    Case "/editprojectiles"
                        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                    
                        SendRequestEditProjectiles
                    Case "/editquest"
                        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                        SendRequestEditQuest
                        ' // Creator Admin Commands //
                        ' Giving another player access
                    Case "/setaccess"
                        If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                        If UBound(Command) < 2 Then
                            AddText "Usage: /setaccess (name) (access)", AlertColor
                            GoTo Continue
                        End If
                        If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                            AddText "Usage: /setaccess (name) (access)", AlertColor
                            GoTo Continue
                        End If
                        SendSetAccess Command(1), CLng(Command(2))
                        ' Packet debug mode
                    Case "/debug"
                        If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                        DEBUG_MODE = (Not DEBUG_MODE)
                    Case Else
                        AddText "Not a valid command!", HelpColor
                End Select
                'continue label where we go instead of exiting the sub
Continue:
                MyText = vbNullString
                UpdateShowChatText
                chatOn = False
                Exit Sub
            End If
            ' Say message
            If Len(ChatText) > 0 Then
                Call SayMsg(ChatText)
            End If
            MyText = vbNullString
            UpdateShowChatText
            chatOn = False
            Exit Sub
        Else
            chatOn = True
        End If

    ElseIf KeyAscii = vbKeyEscape Then
        logoutGame
    End If
    If chatOn Then
        ' Handle when the user presses the backspace key
        If (KeyAscii = vbKeyBack) Then
            If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1): UpdateShowChatText
        End If
        ' And if neither, then add the character to the user's text buffer
        If (KeyAscii <> vbKeyReturn) Then
            If (KeyAscii <> vbKeyBack) Then
                MyText = MyText & ChrW$(KeyAscii)
                UpdateShowChatText
            End If
        End If
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function HandleGame_MouseMove() As Boolean

   On Error GoTo errorhandler

    If HideHotbar = False Then
        'Check for hotbar..
        If GlobalX >= HotbarPnlBounds.Left And GlobalX <= HotbarPnlBounds.Left + HotbarPnlBounds.Right Then
            If GlobalY >= HotbarPnlBounds.Top And GlobalY <= HotbarPnlBounds.Top + HotbarPnlBounds.Bottom Then
                Hotbar_MouseMove
                HandleGame_MouseMove = True
                Exit Function
            End If
        End If
    End If
    
    'Check for inventory
    If CurrentGameMenu = 1 Then
        If GlobalX >= InvItemsBounds.Left And GlobalX <= InvItemsBounds.Left + InvItemsBounds.Right Then
            If GlobalY >= InvItemsBounds.Top And GlobalY <= InvItemsBounds.Top + InvItemsBounds.Bottom Then
                Inventory_MouseMove
                HandleGame_MouseMove = True
                Exit Function
            End If
        End If
    End If
    'Check for Spells
    If CurrentGameMenu = 2 Then
        If GlobalX >= SpellIconsBounds.Left And GlobalX <= SpellIconsBounds.Left + SpellIconsBounds.Right Then
            If GlobalY >= SpellIconsBounds.Top And GlobalY <= SpellIconsBounds.Top + SpellIconsBounds.Bottom Then
                Spells_MouseMove GlobalX, GlobalY
                HandleGame_MouseMove = True
                Exit Function
            End If
        End If
    End If
    If CurrentGameMenu = 3 Then
        If GlobalX >= CharacterPnlBounds.Left And GlobalX <= CharacterPnlBounds.Left + CharacterPnlBounds.Right Then
            If GlobalY >= CharacterPnlBounds.Top And GlobalY <= CharacterPnlBounds.Top + CharacterPnlBounds.Bottom Then
                Character_MouseMove GlobalX, GlobalY
                HandleGame_MouseMove = True
                Exit Function
            End If
        End If
    End If
    If InBank Then
        If GlobalX >= BankPnlBounds.Left And GlobalX <= BankPnlBounds.Left + BankPnlBounds.Right Then
            If GlobalY >= BankPnlBounds.Top And GlobalY <= BankPnlBounds.Top + BankPnlBounds.Bottom Then
                Bank_MouseMove GlobalX, GlobalY
                HandleGame_MouseMove = True
                Exit Function
            End If
        End If
    End If
    If InShop > 0 Then
        If GlobalX >= ShopPnlBounds.Left And GlobalX <= ShopPnlBounds.Left + ShopPnlBounds.Right Then
            If GlobalY >= ShopPnlBounds.Top And GlobalY <= ShopPnlBounds.Top + ShopPnlBounds.Bottom Then
                Shop_MouseMove GlobalX, GlobalY
                HandleGame_MouseMove = True
                Exit Function
            End If
        End If
    End If
    If InTrade > 0 Then
        If GlobalX >= TradePnlBounds.Left And GlobalX <= TradePnlBounds.Left + TradePnlBounds.Right Then
            If GlobalY >= TradePnlBounds.Top And GlobalY <= TradePnlBounds.Top + TradePnlBounds.Bottom Then
                Trade_MouseMove GlobalX, GlobalY
                HandleGame_MouseMove = True
                Exit Function
            End If
        End If
    End If
    If InMailbox Then
        If MailBoxMenu = 2 Then
            If GlobalX >= ReadLetterPnlBounds.Left And GlobalX <= ReadLetterPnlBounds.Left + ReadLetterPnlBounds.Right Then
                If GlobalY >= ReadLetterPnlBounds.Top And GlobalY <= ReadLetterPnlBounds.Top + ReadLetterPnlBounds.Bottom Then
                    ReadLetterPnl_MouseMove GlobalX, GlobalY
                    HandleGame_MouseMove = True
                    Exit Function
                End If
            End If
        ElseIf MailBoxMenu = 3 Then
            If GlobalX >= SendMailPnlBounds.Left And GlobalX <= SendMailPnlBounds.Left + SendMailPnlBounds.Right Then
                If GlobalY >= SendMailPnlBounds.Top And GlobalY <= SendMailPnlBounds.Top + SendMailPnlBounds.Bottom Then
                    SendMail_MouseMove GlobalX, GlobalY
                    HandleGame_MouseMove = True
                    Exit Function
                End If
            End If
        End If
    End If
    ' hide the descriptions
    ItemDescVisible = False
    SpellDescVisible = False


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleGame_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub Hotbar_MouseDown(Button As Long)
Dim SlotNum As Long

   On Error GoTo errorhandler

    SlotNum = IsHotbarSlot(GlobalX, GlobalY)

    If Button = 1 Then
        If SlotNum <> 0 Then
            SendHotbarUse SlotNum
        End If
    ElseIf Button = 2 Then
        If SlotNum <> 0 Then
            SendHotbarChange 0, 0, SlotNum
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Hotbar_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub Hotbar_MouseUp(X As Long, Y As Long)
Dim SlotNum As Long

   On Error GoTo errorhandler

    SlotNum = IsHotbarSlot(X, Y)
    If DragInvSlotNum > 0 Then
        ' hotbar
        If SlotNum > 0 Then
            SendHotbarChange 1, DragInvSlotNum, SlotNum
            DragInvSlotNum = 0
            DragSpell = 0
            Exit Sub
        Else
            DragInvSlotNum = 0
            DragSpell = 0
            Exit Sub
        End If
    End If
    If DragSpell > 0 Then
        If SlotNum > 0 Then
            SendHotbarChange 2, DragSpell, SlotNum
            DragInvSlotNum = 0
            DragSpell = 0
            Exit Sub
        Else
            DragInvSlotNum = 0
            DragSpell = 0
            Exit Sub
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Hotbar_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub Hotbar_MouseMove()
    Dim SlotNum As Long, X As Long, Y As Long

   On Error GoTo errorhandler

    X = GlobalX
    Y = GlobalY
    SlotNum = IsHotbarSlot(X, Y)

    If SlotNum <> 0 Then
        If Hotbar(SlotNum).sType = 1 Then ' item
            UpdateDescWindow Hotbar(SlotNum).slot, GlobalX + 2, GlobalY + 2
            LastItemDesc = Hotbar(SlotNum).slot ' set it so you don't re-set values
            Exit Sub
        ElseIf Hotbar(SlotNum).sType = 2 Then ' spell
            UpdateSpellWindow Hotbar(SlotNum).slot, GlobalX + 2, GlobalY + 2
            LastSpellDesc = Hotbar(SlotNum).slot  ' set it so you don't re-set values
            Exit Sub
        End If
    End If
    ItemDescVisible = False
    LastItemDesc = 0 ' no item was last loaded
    SpellDescVisible = False
    LastSpellDesc = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Hotbar_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Inventory_MouseMove()
    Dim InvNum As Long, X As Long, Y As Long, i As Long

   On Error GoTo errorhandler

    X = GlobalX
    Y = GlobalY
    If DragInvSlotNum > 0 Then
        Exit Sub
    Else
        InvNum = IsInvItem(X, Y)

        If InvNum <> 0 Then
            ' exit out if we're offering that item
            If InTrade Then
                For i = 1 To MAX_INV
                    If TradeYourOffer(i).Num = InvNum Then
                        ' is currency?
                        If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)).Stackable = 1 Then
                            ' only exit out if we're offering all of it
                            If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).Num) Then
                                Exit Sub
                            End If
                        Else
                            Exit Sub
                        End If
                    End If
                Next
            End If
            If InMailbox Then
                If MailBoxMenu = 3 Then
                    If MailItem = InvNum Then Exit Sub
                End If
            End If
            UpdateDescWindow GetPlayerInvItemNum(MyIndex, InvNum), X, Y
            LastItemDesc = GetPlayerInvItemNum(MyIndex, InvNum) ' set it so you don't re-set values
            Exit Sub
        End If
    End If
    ItemDescVisible = False
    LastItemDesc = 0 ' no item was last loaded




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Inventory_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Function HandleGame_DblClick() As Boolean
Dim X As Long, Y As Long, isHandled As Boolean

   On Error GoTo errorhandler

    X = GlobalX
    Y = GlobalY
        If CurrentGameMenu = 1 Then
            If X >= InventoryPnlBounds.Left And X <= InventoryPnlBounds.Left + InventoryPnlBounds.Right Then
                If Y >= InventoryPnlBounds.Top And Y <= InventoryPnlBounds.Top + InventoryPnlBounds.Bottom Then
                    Inventory_DblClick X, Y
                    HandleGame_DblClick = True
                    Exit Function
                End If
            End If
        End If
        If CurrentGameMenu = 2 Then
            If X >= SpellsPnlBounds.Left And X <= SpellsPnlBounds.Left + SpellsPnlBounds.Right Then
                If Y >= SpellsPnlBounds.Top And Y <= SpellsPnlBounds.Top + SpellsPnlBounds.Bottom Then
                    Spells_DblClick X, Y
                    HandleGame_DblClick = True
                    Exit Function
                End If
            End If
        End If
        If InBank Then
            If X >= BankPnlBounds.Left And X <= BankPnlBounds.Left + BankPnlBounds.Right Then
                If Y >= BankPnlBounds.Top And Y <= BankPnlBounds.Top + BankPnlBounds.Bottom Then
                    Bank_DblClick X, Y
                    HandleGame_DblClick = True
                    Exit Function
                End If
            End If
        End If
        If InShop > 0 Then
            If X >= ShopPnlBounds.Left And X <= ShopPnlBounds.Left + ShopPnlBounds.Right Then
                If Y >= ShopPnlBounds.Top And Y <= ShopPnlBounds.Top + ShopPnlBounds.Bottom Then
                    Shop_DblClick X, Y
                    HandleGame_DblClick = True
                    Exit Function
                End If
            End If
        End If
        If InTrade > 0 Then
            If X >= TradePnlBounds.Left And X <= TradePnlBounds.Left + TradePnlBounds.Right Then
                If Y >= TradePnlBounds.Top And Y <= TradePnlBounds.Top + TradePnlBounds.Bottom Then
                    Trade_DblClick X, Y
                    HandleGame_DblClick = True
                    Exit Function
                End If
            End If
        End If
        If InMailbox Then
            If MailBoxMenu = 2 Then
                If X >= ReadLetterPnlBounds.Left And X <= ReadLetterPnlBounds.Left + ReadLetterPnlBounds.Right Then
                    If Y >= ReadLetterPnlBounds.Top And Y <= ReadLetterPnlBounds.Top + ReadLetterPnlBounds.Bottom Then
                        ReadLetterPnl_DblClick X, Y
                        HandleGame_DblClick = True
                        Exit Function
                    End If
                End If
            ElseIf MailBoxMenu = 3 Then
                If X >= SendMailPnlBounds.Left And X <= SendMailPnlBounds.Left + SendMailPnlBounds.Right Then
                    If Y >= SendMailPnlBounds.Top And Y <= SendMailPnlBounds.Top + SendMailPnlBounds.Bottom Then
                        SendMail_DblClick X, Y
                        HandleGame_DblClick = True
                        Exit Function
                    End If
                End If
            End If
        End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleGame_DblClick", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function HandleGame_MouseDown() As Boolean
Dim X As Long, Y As Long, isHandled As Boolean

   On Error GoTo errorhandler

    X = GlobalX
    Y = GlobalY
        If CurrencyMenu > 0 Then
            If X >= CurrencyPanelBounds.Left And X <= CurrencyPanelBounds.Left + CurrencyPanelBounds.Right Then
                If Y >= CurrencyPanelBounds.Top And Y <= CurrencyPanelBounds.Top + CurrencyPanelBounds.Bottom Then
                    Currency_MouseDown GlobalX, GlobalY
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
            If dialogueIndex > 0 Then
            If X >= DialoguePanelBounds.Left And X <= DialoguePanelBounds.Left + DialoguePanelBounds.Right Then
                If Y >= DialoguePanelBounds.Top And Y <= DialoguePanelBounds.Top + DialoguePanelBounds.Bottom Then
                    Dialogue_MouseDown GlobalX, GlobalY
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
            If EventChat = True Then
            If X >= EventPnlBounds.Left And X <= EventPnlBounds.Left + EventPnlBounds.Right Then
                If Y >= EventPnlBounds.Top And Y <= EventPnlBounds.Top + EventPnlBounds.Bottom Then
                    Event_MouseDown GlobalX, GlobalY
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If HideChat = False Then
            If X >= ChatboxPnlBounds.Left And X <= ChatboxPnlBounds.Left + ChatboxPnlBounds.Right Then
                If Y >= ChatboxPnlBounds.Top And Y <= ChatboxPnlBounds.Top + ChatboxPnlBounds.Bottom Then
                    Chatbox_MouseDown X, Y
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If HideMenu = False Then
            If X >= GameMenuPanelBounds.Left And X <= GameMenuPanelBounds.Left + GameMenuPanelBounds.Right Then
                If Y >= GameMenuPanelBounds.Top And Y <= GameMenuPanelBounds.Top + GameMenuPanelBounds.Bottom Then
                    GameMenu_MouseDown X, Y
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 1 Then
            If X >= InventoryPnlBounds.Left And X <= InventoryPnlBounds.Left + InventoryPnlBounds.Right Then
                If Y >= InventoryPnlBounds.Top And Y <= InventoryPnlBounds.Top + InventoryPnlBounds.Bottom Then
                    Inventory_MouseDown X, Y
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 2 Then
            If X >= SpellsPnlBounds.Left And X <= SpellsPnlBounds.Left + SpellsPnlBounds.Right Then
                If Y >= SpellsPnlBounds.Top And Y <= SpellsPnlBounds.Top + SpellsPnlBounds.Bottom Then
                    Spells_MouseDown X, Y
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 3 Then
            If X >= CharacterPnlBounds.Left And X <= CharacterPnlBounds.Left + CharacterPnlBounds.Right Then
                If Y >= CharacterPnlBounds.Top And Y <= CharacterPnlBounds.Top + CharacterPnlBounds.Bottom Then
                    Character_MouseDown X, Y
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 4 Then
            If X >= OptionsPnlBounds.Left And X <= OptionsPnlBounds.Left + OptionsPnlBounds.Right Then
                If Y >= OptionsPnlBounds.Top And Y <= OptionsPnlBounds.Top + OptionsPnlBounds.Bottom Then
                    Options_MouseDown X, Y
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 5 Then
            If X >= PartyPnlBounds.Left And X <= PartyPnlBounds.Left + PartyPnlBounds.Right Then
                If Y >= PartyPnlBounds.Top And Y <= PartyPnlBounds.Top + PartyPnlBounds.Bottom Then
                    Party_MouseDown X, Y
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 6 Then
            If X >= FriendsPnlBounds.Left And X <= FriendsPnlBounds.Left + FriendsPnlBounds.Right Then
                If Y >= FriendsPnlBounds.Top And Y <= FriendsPnlBounds.Top + FriendsPnlBounds.Bottom Then
                    Friends_MouseDown X, Y
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 7 Then
            If X >= QuestsPnlBounds.Left And X <= QuestsPnlBounds.Left + QuestsPnlBounds.Right Then
                If Y >= QuestsPnlBounds.Top And Y <= QuestsPnlBounds.Top + QuestsPnlBounds.Bottom Then
                    Quests_MouseDown X, Y
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 8 Then
            If X >= PetPanelBounds.Left And X <= PetPanelBounds.Left + PetPanelBounds.Right Then
                If Y >= PetPanelBounds.Top And Y <= PetPanelBounds.Top + PetPanelBounds.Bottom Then
                    Pets_MouseDown X, Y
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If HideHotbar = False Then
            If X >= HotbarPnlBounds.Left And X <= HotbarPnlBounds.Left + HotbarPnlBounds.Right Then
                If Y >= HotbarPnlBounds.Top And Y <= HotbarPnlBounds.Top + HotbarPnlBounds.Bottom Then
                    Hotbar_MouseDown MouseBtn
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If InBank Then
            If X >= BankPnlBounds.Left And X <= BankPnlBounds.Left + BankPnlBounds.Right Then
                If Y >= BankPnlBounds.Top And Y <= BankPnlBounds.Top + BankPnlBounds.Bottom Then
                    Bank_MouseDown GlobalX, GlobalY
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If InShop > 0 Then
            If X >= ShopPnlBounds.Left And X <= ShopPnlBounds.Left + ShopPnlBounds.Right Then
                If Y >= ShopPnlBounds.Top And Y <= ShopPnlBounds.Top + ShopPnlBounds.Bottom Then
                    Shop_MouseDown GlobalX, GlobalY
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If InTrade > 0 Then
            If X >= TradePnlBounds.Left And X <= TradePnlBounds.Left + TradePnlBounds.Right Then
                If Y >= TradePnlBounds.Top And Y <= TradePnlBounds.Top + TradePnlBounds.Bottom Then
                    Trade_MouseDown GlobalX, GlobalY
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If
        
        If InMailbox Then
            Select Case MailBoxMenu
                Case 0
                    If X >= MailboxPnlBounds.Left And X <= MailboxPnlBounds.Left + MailboxPnlBounds.Right Then
                        If Y >= MailboxPnlBounds.Top And Y <= MailboxPnlBounds.Top + MailboxPnlBounds.Bottom Then
                            MailboxPnl_MouseDown GlobalX, GlobalY
                            HandleGame_MouseDown = True
                            Exit Function
                        End If
                    End If
                Case 1
                    If X >= InboxPnlBounds.Left And X <= InboxPnlBounds.Left + InboxPnlBounds.Right Then
                        If Y >= InboxPnlBounds.Top And Y <= InboxPnlBounds.Top + InboxPnlBounds.Bottom Then
                            InboxPnl_MouseDown GlobalX, GlobalY
                            HandleGame_MouseDown = True
                            Exit Function
                        End If
                    End If
                Case 2
                    If X >= ReadLetterPnlBounds.Left And X <= ReadLetterPnlBounds.Left + ReadLetterPnlBounds.Right Then
                        If Y >= ReadLetterPnlBounds.Top And Y <= ReadLetterPnlBounds.Top + ReadLetterPnlBounds.Bottom Then
                            ReadLetterPnl_MouseDown GlobalX, GlobalY
                            HandleGame_MouseDown = True
                            Exit Function
                        End If
                    End If
                Case 3
                    If X >= SendMailPnlBounds.Left And X <= SendMailPnlBounds.Left + SendMailPnlBounds.Right Then
                        If Y >= SendMailPnlBounds.Top And Y <= SendMailPnlBounds.Top + SendMailPnlBounds.Bottom Then
                            SendMail_MouseDown GlobalX, GlobalY
                            HandleGame_MouseDown = True
                            Exit Function
                        End If
                    End If
            End Select
        End If
        
        If InQuestLog Then
            If X >= QuestLogPanelBounds.Left And X <= QuestLogPanelBounds.Left + QuestLogPanelBounds.Right Then
                If Y >= QuestLogPanelBounds.Top And Y <= QuestLogPanelBounds.Top + QuestLogPanelBounds.Bottom Then
                    QuestLog_MouseDown GlobalX, GlobalY
                    HandleGame_MouseDown = True
                    Exit Function
                End If
            End If
        End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleGame_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function HandleGame_MouseUp() As Boolean
Dim X As Long, Y As Long, isHandled As Boolean, buffer As clsBuffer, i As Long

   On Error GoTo errorhandler

    X = GlobalX
    Y = GlobalY
        If CurrencyMenu > 0 Then
            If X >= CurrencyPanelBounds.Left And X <= CurrencyPanelBounds.Left + CurrencyPanelBounds.Right Then
                If Y >= CurrencyPanelBounds.Top And Y <= CurrencyPanelBounds.Top + CurrencyPanelBounds.Bottom Then
                    Currency_MouseUp GlobalX, GlobalY
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        If dialogueIndex > 0 Then
            If X >= DialoguePanelBounds.Left And X <= DialoguePanelBounds.Left + DialoguePanelBounds.Right Then
                If Y >= DialoguePanelBounds.Top And Y <= DialoguePanelBounds.Top + DialoguePanelBounds.Bottom Then
                    Dialogue_MouseUp GlobalX, GlobalY
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If HideChat = False Then
            If X >= ChatboxPnlBounds.Left And X <= ChatboxPnlBounds.Left + ChatboxPnlBounds.Right Then
                If Y >= ChatboxPnlBounds.Top And Y <= ChatboxPnlBounds.Top + ChatboxPnlBounds.Bottom Then
                    Chatbox_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If HideMenu = False Then
            If X >= GameMenuPanelBounds.Left And X <= GameMenuPanelBounds.Left + GameMenuPanelBounds.Right Then
                If Y >= GameMenuPanelBounds.Top And Y <= GameMenuPanelBounds.Top + GameMenuPanelBounds.Bottom Then
                    GameMenu_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 1 Then
            If X >= InventoryPnlBounds.Left And X <= InventoryPnlBounds.Left + InventoryPnlBounds.Right Then
                If Y >= InventoryPnlBounds.Top And Y <= InventoryPnlBounds.Top + InventoryPnlBounds.Bottom Then
                    Inventory_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 2 Then
            If X >= SpellsPnlBounds.Left And X <= SpellsPnlBounds.Left + SpellsPnlBounds.Right Then
                If Y >= SpellsPnlBounds.Top And Y <= SpellsPnlBounds.Top + SpellsPnlBounds.Bottom Then
                    Spells_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 3 Then
            If X >= CharacterPnlBounds.Left And X <= CharacterPnlBounds.Left + CharacterPnlBounds.Right Then
                If Y >= CharacterPnlBounds.Top And Y <= CharacterPnlBounds.Top + CharacterPnlBounds.Bottom Then
                    Character_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 4 Then
            If X >= OptionsPnlBounds.Left And X <= OptionsPnlBounds.Left + OptionsPnlBounds.Right Then
                If Y >= OptionsPnlBounds.Top And Y <= OptionsPnlBounds.Top + OptionsPnlBounds.Bottom Then
                    Options_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 5 Then
            If X >= PartyPnlBounds.Left And X <= PartyPnlBounds.Left + PartyPnlBounds.Right Then
                If Y >= PartyPnlBounds.Top And Y <= PartyPnlBounds.Top + PartyPnlBounds.Bottom Then
                    Party_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 6 Then
            If X >= FriendsPnlBounds.Left And X <= FriendsPnlBounds.Left + FriendsPnlBounds.Right Then
                If Y >= FriendsPnlBounds.Top And Y <= FriendsPnlBounds.Top + FriendsPnlBounds.Bottom Then
                    Friends_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 7 Then
            If X >= QuestsPnlBounds.Left And X <= QuestsPnlBounds.Left + QuestsPnlBounds.Right Then
                If Y >= QuestsPnlBounds.Top And Y <= QuestsPnlBounds.Top + QuestsPnlBounds.Bottom Then
                    Quests_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If CurrentGameMenu = 8 Then
            If X >= PetPanelBounds.Left And X <= PetPanelBounds.Left + PetPanelBounds.Right Then
                If Y >= PetPanelBounds.Top And Y <= PetPanelBounds.Top + PetPanelBounds.Bottom Then
                    Pets_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If HideHotbar = False Then
            If X >= HotbarPnlBounds.Left And X <= HotbarPnlBounds.Left + HotbarPnlBounds.Right Then
                If Y >= HotbarPnlBounds.Top And Y <= HotbarPnlBounds.Top + HotbarPnlBounds.Bottom Then
                    Hotbar_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If InBank Then
            If X >= BankPnlBounds.Left And X <= BankPnlBounds.Left + BankPnlBounds.Right Then
                If Y >= BankPnlBounds.Top And Y <= BankPnlBounds.Top + BankPnlBounds.Bottom Then
                    Bank_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If InShop > 0 Then
            If X >= ShopPnlBounds.Left And X <= ShopPnlBounds.Left + ShopPnlBounds.Right Then
                If Y >= ShopPnlBounds.Top And Y <= ShopPnlBounds.Top + ShopPnlBounds.Bottom Then
                    Shop_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If InTrade > 0 Then
            If X >= TradePnlBounds.Left And X <= TradePnlBounds.Left + TradePnlBounds.Right Then
                If Y >= TradePnlBounds.Top And Y <= TradePnlBounds.Top + TradePnlBounds.Bottom Then
                    Trade_MouseUp X, Y
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If InMailbox Then
            Select Case MailBoxMenu
                Case 0
                    If X >= MailboxPnlBounds.Left And X <= MailboxPnlBounds.Left + MailboxPnlBounds.Right Then
                        If Y >= MailboxPnlBounds.Top And Y <= MailboxPnlBounds.Top + MailboxPnlBounds.Bottom Then
                            MailboxPnl_MouseUp GlobalX, GlobalY
                            HandleGame_MouseUp = True
                            Exit Function
                        End If
                    End If
                Case 1
                    If X >= InboxPnlBounds.Left And X <= InboxPnlBounds.Left + InboxPnlBounds.Right Then
                        If Y >= InboxPnlBounds.Top And Y <= InboxPnlBounds.Top + InboxPnlBounds.Bottom Then
                            InboxPnl_MouseUp GlobalX, GlobalY
                            HandleGame_MouseUp = True
                            Exit Function
                        End If
                    End If
                Case 2
                    If X >= ReadLetterPnlBounds.Left And X <= ReadLetterPnlBounds.Left + ReadLetterPnlBounds.Right Then
                        If Y >= ReadLetterPnlBounds.Top And Y <= ReadLetterPnlBounds.Top + ReadLetterPnlBounds.Bottom Then
                            ReadLetterPnl_MouseUp GlobalX, GlobalY
                            HandleGame_MouseUp = True
                            Exit Function
                        End If
                    End If
                Case 3
                    If X >= SendMailPnlBounds.Left And X <= SendMailPnlBounds.Left + SendMailPnlBounds.Right Then
                        If Y >= SendMailPnlBounds.Top And Y <= SendMailPnlBounds.Top + SendMailPnlBounds.Bottom Then
                            SendMail_MouseUp GlobalX, GlobalY
                            HandleGame_MouseUp = True
                            Exit Function
                        End If
                    End If
            End Select
        End If
        
        If InQuestLog Then
            If X >= QuestLogPanelBounds.Left And X <= QuestLogPanelBounds.Left + QuestLogPanelBounds.Right Then
                If Y >= QuestLogPanelBounds.Top And Y <= QuestLogPanelBounds.Top + QuestLogPanelBounds.Bottom Then
                    QuestLog_MouseUp GlobalX, GlobalY
                    HandleGame_MouseUp = True
                    Exit Function
                End If
            End If
        End If
        
        If DragInvSlotNum > 0 Then
            If Player(MyIndex).InHouse = MyIndex Then
                If Item(PlayerInv(DragInvSlotNum).Num).type = ITEM_TYPE_FURNITURE Then
                    Set buffer = New clsBuffer
                    buffer.WriteLong CPlaceFurniture
                    i = GlobalX
                    i = TileView.Left + (((i) + Camera.Left) \ PIC_X)
                    buffer.WriteLong i
                    i = GlobalY
                    i = TileView.Top + (((i) + Camera.Top) \ PIC_Y) + Item(PlayerInv(DragInvSlotNum).Num).FurnitureHeight
                    buffer.WriteLong i
                    buffer.WriteLong DragInvSlotNum
                    SendData buffer.ToArray
                    Set buffer = Nothing
                End If
            End If
        End If
    DragInvSlotNum = 0
    DragBankSlotNum = 0
    DragSpell = 0
    DragTradeSlotNum = 0


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleGame_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function HandleLoginPanel_MouseDown(X As Long, Y As Long)
Dim isHandled As Boolean
    'Check Textbox 1

   On Error GoTo errorhandler

    If X >= LUsernameBounds.Left And X <= LUsernameBounds.Left + LUsernameBounds.Right Then
        If Y >= LUsernameBounds.Top And Y <= LUsernameBounds.Top + LUsernameBounds.Bottom Then
            SelTextbox = 1
            isHandled = True
        End If
    End If
    'Check Textbox 2
    If isHandled = False Then
        If X >= LPasswordBounds.Left And X <= LPasswordBounds.Left + LPasswordBounds.Right Then
            If Y >= LPasswordBounds.Top And Y <= LPasswordBounds.Top + LPasswordBounds.Bottom Then
                SelTextbox = 2
                isHandled = True
            End If
        End If
    End If
    'Login Button
    If isHandled = False Then
        If X >= LoginButtonBounds.Left And X <= LoginButtonBounds.Left + LoginButtonBounds.Right Then
            If Y >= LoginButtonBounds.Top And Y <= LoginButtonBounds.Top + LoginButtonBounds.Bottom Then
                LoginButtonState = 2
            End If
        End If
    End If
    'Login Button
    If isHandled = False Then
        If X >= SaveInfoCheckBounds.Left And X <= SaveInfoCheckBounds.Left + SaveInfoCheckBounds.Right Then
            If Y >= SaveInfoCheckBounds.Top And Y <= SaveInfoCheckBounds.Top + SaveInfoCheckBounds.Bottom Then
                If Servers(ServerIndex).SavePass = 1 Then
                    Servers(ServerIndex).SavePass = 0
                    SaveServers
                Else
                    Servers(ServerIndex).SavePass = 1
                    SaveServers
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleLoginPanel_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function HandleLoginPanel_MouseUp(X As Long, Y As Long)
    'Login Button

   On Error GoTo errorhandler

    If X >= LoginButtonBounds.Left And X <= LoginButtonBounds.Left + LoginButtonBounds.Right Then
        If Y >= LoginButtonBounds.Top And Y <= LoginButtonBounds.Top + LoginButtonBounds.Bottom Then
            If LoginButtonState = 2 Then
                MenuState MENU_STATE_LOGIN
                LoginButtonState = 0
            Else
                LoginButtonState = 0
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleLoginPanel_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function HandleNewCharPanel_MouseDown(X As Long, Y As Long)
Dim isHandled As Boolean
    'Check Textbox 1

   On Error GoTo errorhandler

    If X >= NCTextboxBounds.Left And X <= NCTextboxBounds.Left + NCTextboxBounds.Right Then
        If Y >= NCTextboxBounds.Top And Y <= NCTextboxBounds.Top + NCTextboxBounds.Bottom Then
            SelTextbox = 1
            isHandled = True
        End If
    End If
    If isHandled = False Then
        If X >= PrevClassBounds.Left And X <= PrevClassBounds.Left + PrevClassBounds.Right Then
            If Y >= PrevClassBounds.Top And Y <= PrevClassBounds.Top + PrevClassBounds.Bottom Then
                PrevClassState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= NextClassBounds.Left And X <= NextClassBounds.Left + NextClassBounds.Right Then
            If Y >= NextClassBounds.Top And Y <= NextClassBounds.Top + NextClassBounds.Bottom Then
                NextClassState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= MaleButtonBounds.Left And X <= MaleButtonBounds.Left + MaleButtonBounds.Right Then
            If Y >= MaleButtonBounds.Top And Y <= MaleButtonBounds.Top + MaleButtonBounds.Bottom Then
                MaleButtonState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= FemaleButtonBounds.Left And X <= FemaleButtonBounds.Left + FemaleButtonBounds.Right Then
            If Y >= FemaleButtonBounds.Top And Y <= FemaleButtonBounds.Top + FemaleButtonBounds.Bottom Then
                FemaleButtonState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= NextHairBounds.Left And X <= NextHairBounds.Left + NextHairBounds.Right Then
            If Y >= NextHairBounds.Top And Y <= NextHairBounds.Top + NextHairBounds.Bottom Then
                NextHairState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= NextHeadBounds.Left And X <= NextHeadBounds.Left + NextHeadBounds.Right Then
            If Y >= NextHeadBounds.Top And Y <= NextHeadBounds.Top + NextHeadBounds.Bottom Then
                NextHeadState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= NextEyeBounds.Left And X <= NextEyeBounds.Left + NextEyeBounds.Right Then
            If Y >= NextEyeBounds.Top And Y <= NextEyeBounds.Top + NextEyeBounds.Bottom Then
                NextEyeState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= NextEyebrowBounds.Left And X <= NextEyebrowBounds.Left + NextEyebrowBounds.Right Then
            If Y >= NextEyebrowBounds.Top And Y <= NextEyebrowBounds.Top + NextEyebrowBounds.Bottom Then
                NextEyebrowState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= NextEarBounds.Left And X <= NextEarBounds.Left + NextEarBounds.Right Then
            If Y >= NextEarBounds.Top And Y <= NextEarBounds.Top + NextEarBounds.Bottom Then
                NextEarState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= NextMouthBounds.Left And X <= NextMouthBounds.Left + NextMouthBounds.Right Then
            If Y >= NextMouthBounds.Top And Y <= NextMouthBounds.Top + NextMouthBounds.Bottom Then
                NextMouthState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= NextNoseBounds.Left And X <= NextNoseBounds.Left + NextNoseBounds.Right Then
            If Y >= NextNoseBounds.Top And Y <= NextNoseBounds.Top + NextNoseBounds.Bottom Then
                NextNoseState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= NextShirtBounds.Left And X <= NextShirtBounds.Left + NextShirtBounds.Right Then
            If Y >= NextShirtBounds.Top And Y <= NextShirtBounds.Top + NextShirtBounds.Bottom Then
                NextShirtState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= NextExtraBounds.Left And X <= NextExtraBounds.Left + NextExtraBounds.Right Then
            If Y >= NextExtraBounds.Top And Y <= NextExtraBounds.Top + NextExtraBounds.Bottom Then
                NextExtraState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= PrevHairBounds.Left And X <= PrevHairBounds.Left + PrevHairBounds.Right Then
            If Y >= PrevHairBounds.Top And Y <= PrevHairBounds.Top + PrevHairBounds.Bottom Then
                PrevHairState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= PrevHeadBounds.Left And X <= PrevHeadBounds.Left + PrevHeadBounds.Right Then
            If Y >= PrevHeadBounds.Top And Y <= PrevHeadBounds.Top + PrevHeadBounds.Bottom Then
                PrevHeadState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= PrevEyeBounds.Left And X <= PrevEyeBounds.Left + PrevEyeBounds.Right Then
            If Y >= PrevEyeBounds.Top And Y <= PrevEyeBounds.Top + PrevEyeBounds.Bottom Then
                PrevEyeState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= PrevEyebrowBounds.Left And X <= PrevEyebrowBounds.Left + PrevEyebrowBounds.Right Then
            If Y >= PrevEyebrowBounds.Top And Y <= PrevEyebrowBounds.Top + PrevEyebrowBounds.Bottom Then
                PrevEyebrowState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= PrevEarBounds.Left And X <= PrevEarBounds.Left + PrevEarBounds.Right Then
            If Y >= PrevEarBounds.Top And Y <= PrevEarBounds.Top + PrevEarBounds.Bottom Then
                PrevEarState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= PrevMouthBounds.Left And X <= PrevMouthBounds.Left + PrevMouthBounds.Right Then
            If Y >= PrevMouthBounds.Top And Y <= PrevMouthBounds.Top + PrevMouthBounds.Bottom Then
                PrevMouthState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= PrevNoseBounds.Left And X <= PrevNoseBounds.Left + PrevNoseBounds.Right Then
            If Y >= PrevNoseBounds.Top And Y <= PrevNoseBounds.Top + PrevNoseBounds.Bottom Then
                PrevNoseState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= PrevShirtBounds.Left And X <= PrevShirtBounds.Left + PrevShirtBounds.Right Then
            If Y >= PrevShirtBounds.Top And Y <= PrevShirtBounds.Top + PrevShirtBounds.Bottom Then
                PrevShirtState = 2
                isHandled = True
            End If
        End If
    End If
    If isHandled = False Then
        If X >= PrevExtraBounds.Left And X <= PrevExtraBounds.Left + PrevExtraBounds.Right Then
            If Y >= PrevExtraBounds.Top And Y <= PrevExtraBounds.Top + PrevExtraBounds.Bottom Then
                PrevExtraState = 2
                isHandled = True
            End If
        End If
    End If
    'Accept Char Button
    If isHandled = False Then
        If X >= NCAcceptBounds.Left And X <= NCAcceptBounds.Left + NCAcceptBounds.Right Then
            If Y >= NCAcceptBounds.Top And Y <= NCAcceptBounds.Top + NCAcceptBounds.Bottom Then
                NCAcceptState = 2
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleNewCharPanel_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function HandleNewCharPanel_MouseUp(X As Long, Y As Long)
    'PrevClass

   On Error GoTo errorhandler
    If X >= PrevClassBounds.Left And X <= PrevClassBounds.Left + PrevClassBounds.Right Then
        If Y >= PrevClassBounds.Top And Y <= PrevClassBounds.Top + PrevClassBounds.Bottom Then
            If PrevClassState = 2 Then
                newCharClass = newCharClass - 1
                If newCharClass <= 0 Then newCharClass = Max_Classes
                ResetNewChar
                PrevClassState = 0
            Else
                PrevClassState = 0
            End If
        End If
    End If
    'NextClass
    If X >= NextClassBounds.Left And X <= NextClassBounds.Left + NextClassBounds.Right Then
        If Y >= NextClassBounds.Top And Y <= NextClassBounds.Top + NextClassBounds.Bottom Then
            If NextClassState = 2 Then
                newCharClass = newCharClass + 1
                If newCharClass > Max_Classes Then newCharClass = 1
                ResetNewChar
                NextClassState = 0
            Else
                NextClassState = 0
            End If
        End If
    End If
    'MaleButton
    If X >= MaleButtonBounds.Left And X <= MaleButtonBounds.Left + MaleButtonBounds.Right Then
        If Y >= MaleButtonBounds.Top And Y <= MaleButtonBounds.Top + MaleButtonBounds.Bottom Then
            If MaleButtonState = 2 Then
                NewCharSex = SEX_MALE
                ResetNewChar
                MaleButtonState = 0
            Else
                MaleButtonState = 0
            End If
        End If
    End If
    'FemaleButton
    If X >= FemaleButtonBounds.Left And X <= FemaleButtonBounds.Left + FemaleButtonBounds.Right Then
        If Y >= FemaleButtonBounds.Top And Y <= FemaleButtonBounds.Top + FemaleButtonBounds.Bottom Then
            If FemaleButtonState = 2 Then
                NewCharSex = SEX_FEMALE
                ResetNewChar
                FemaleButtonState = 0
            Else
                FemaleButtonState = 0
            End If
        End If
    End If
    If CharMode = 1 Then
        If X >= NextHairBounds.Left And X <= NextHairBounds.Left + NextHairBounds.Right Then
            If Y >= NextHairBounds.Top And Y <= NextHairBounds.Top + NextHairBounds.Bottom Then
                If NextHairState = 2 Then
                    NewCharChange 0, 1
                    NextHairState = 0
                Else
                    NextHairState = 0
                End If
            End If
        End If
    End If
    If X >= NextHeadBounds.Left And X <= NextHeadBounds.Left + NextHeadBounds.Right Then
        If Y >= NextHeadBounds.Top And Y <= NextHeadBounds.Top + NextHeadBounds.Bottom Then
            If NextHeadState = 2 Then
                If CharMode = 1 Then
                    NewCharChange 1, 1
                Else
                    NewCharChange 9, 1
                End If
                NextHeadState = 0
            Else
                NextHeadState = 0
            End If
        End If
    End If
    If CharMode = 1 Then
        If X >= NextEyebrowBounds.Left And X <= NextEyebrowBounds.Left + NextEyebrowBounds.Right Then
            If Y >= NextEyebrowBounds.Top And Y <= NextEyebrowBounds.Top + NextEyebrowBounds.Bottom Then
                If NextEyebrowState = 2 Then
                    NewCharChange 2, 1
                    NextEyebrowState = 0
                Else
                    NextEyebrowState = 0
                End If
            End If
        End If
        If X >= NextEyeBounds.Left And X <= NextEyeBounds.Left + NextEyeBounds.Right Then
            If Y >= NextEyeBounds.Top And Y <= NextEyeBounds.Top + NextEyeBounds.Bottom Then
                If NextEyeState = 2 Then
                    NewCharChange 3, 1
                    NextEyeState = 0
                Else
                    NextEyeState = 0
                End If
            End If
        End If
        If X >= NextEarBounds.Left And X <= NextEarBounds.Left + NextEarBounds.Right Then
            If Y >= NextEarBounds.Top And Y <= NextEarBounds.Top + NextEarBounds.Bottom Then
                If NextEarState = 2 Then
                    NewCharChange 4, 1
                    NextEarState = 0
                Else
                    NextEarState = 0
                End If
            End If
        End If
        If X >= NextMouthBounds.Left And X <= NextMouthBounds.Left + NextMouthBounds.Right Then
            If Y >= NextMouthBounds.Top And Y <= NextMouthBounds.Top + NextMouthBounds.Bottom Then
                If NextMouthState = 2 Then
                    NewCharChange 5, 1
                    NextMouthState = 0
                Else
                    NextMouthState = 0
                End If
            End If
        End If
        If X >= NextNoseBounds.Left And X <= NextNoseBounds.Left + NextNoseBounds.Right Then
            If Y >= NextNoseBounds.Top And Y <= NextNoseBounds.Top + NextNoseBounds.Bottom Then
                If NextNoseState = 2 Then
                    NewCharChange 6, 1
                    NextNoseState = 0
                Else
                    NextNoseState = 0
                End If
            End If
        End If
        If X >= NextShirtBounds.Left And X <= NextShirtBounds.Left + NextShirtBounds.Right Then
            If Y >= NextShirtBounds.Top And Y <= NextShirtBounds.Top + NextShirtBounds.Bottom Then
                If NextShirtState = 2 Then
                    NewCharChange 7, 1
                    NextShirtState = 0
                Else
                    NextShirtState = 0
                End If
            End If
        End If
        If X >= NextExtraBounds.Left And X <= NextExtraBounds.Left + NextExtraBounds.Right Then
            If Y >= NextExtraBounds.Top And Y <= NextExtraBounds.Top + NextExtraBounds.Bottom Then
                If NextExtraState = 2 Then
                    NewCharChange 8, 1
                    NextExtraState = 0
                Else
                    NextExtraState = 0
                End If
            End If
        End If
        If X >= PrevHairBounds.Left And X <= PrevHairBounds.Left + PrevHairBounds.Right Then
            If Y >= PrevHairBounds.Top And Y <= PrevHairBounds.Top + PrevHairBounds.Bottom Then
                If PrevHairState = 2 Then
                    NewCharChange 0, 2
                    PrevHairState = 0
                Else
                    PrevHairState = 0
                End If
            End If
        End If
    End If
    If X >= PrevHeadBounds.Left And X <= PrevHeadBounds.Left + PrevHeadBounds.Right Then
        If Y >= PrevHeadBounds.Top And Y <= PrevHeadBounds.Top + PrevHeadBounds.Bottom Then
            If PrevHeadState = 2 Then
                If CharMode = 1 Then
                    NewCharChange 1, 2
                Else
                    NewCharChange 9, 2
                End If
                PrevHeadState = 0
            Else
                PrevHeadState = 0
            End If
        End If
    End If
    If CharMode = 1 Then
        If X >= PrevEyebrowBounds.Left And X <= PrevEyebrowBounds.Left + PrevEyebrowBounds.Right Then
            If Y >= PrevEyebrowBounds.Top And Y <= PrevEyebrowBounds.Top + PrevEyebrowBounds.Bottom Then
                If PrevEyebrowState = 2 Then
                    NewCharChange 2, 2
                    PrevEyebrowState = 0
                Else
                    PrevEyebrowState = 0
                End If
            End If
        End If
        If X >= PrevEyeBounds.Left And X <= PrevEyeBounds.Left + PrevEyeBounds.Right Then
            If Y >= PrevEyeBounds.Top And Y <= PrevEyeBounds.Top + PrevEyeBounds.Bottom Then
                If PrevEyeState = 2 Then
                    NewCharChange 3, 2
                    PrevEyeState = 0
                Else
                    PrevEyeState = 0
                End If
            End If
        End If
        If X >= PrevEarBounds.Left And X <= PrevEarBounds.Left + PrevEarBounds.Right Then
            If Y >= PrevEarBounds.Top And Y <= PrevEarBounds.Top + PrevEarBounds.Bottom Then
                If PrevEarState = 2 Then
                    NewCharChange 4, 2
                    PrevEarState = 0
                Else
                    PrevEarState = 0
                End If
            End If
        End If
        If X >= PrevMouthBounds.Left And X <= PrevMouthBounds.Left + PrevMouthBounds.Right Then
            If Y >= PrevMouthBounds.Top And Y <= PrevMouthBounds.Top + PrevMouthBounds.Bottom Then
                If PrevMouthState = 2 Then
                    NewCharChange 5, 2
                    PrevMouthState = 0
                Else
                    PrevMouthState = 0
                End If
            End If
        End If
        If X >= PrevNoseBounds.Left And X <= PrevNoseBounds.Left + PrevNoseBounds.Right Then
            If Y >= PrevNoseBounds.Top And Y <= PrevNoseBounds.Top + PrevNoseBounds.Bottom Then
                If PrevNoseState = 2 Then
                    NewCharChange 6, 2
                    PrevNoseState = 0
                Else
                    PrevNoseState = 0
                End If
            End If
        End If
        If X >= PrevShirtBounds.Left And X <= PrevShirtBounds.Left + PrevShirtBounds.Right Then
            If Y >= PrevShirtBounds.Top And Y <= PrevShirtBounds.Top + PrevShirtBounds.Bottom Then
                If PrevShirtState = 2 Then
                    NewCharChange 7, 2
                    PrevShirtState = 0
                Else
                    PrevShirtState = 0
                End If
            End If
        End If
        If X >= PrevExtraBounds.Left And X <= PrevExtraBounds.Left + PrevExtraBounds.Right Then
            If Y >= PrevExtraBounds.Top And Y <= PrevExtraBounds.Top + PrevExtraBounds.Bottom Then
                If PrevExtraState = 2 Then
                    NewCharChange 8, 2
                    PrevExtraState = 0
                Else
                    PrevExtraState = 0
                End If
            End If
        End If
    End If
    'FACE
    'NCAccept
    If X >= NCAcceptBounds.Left And X <= NCAcceptBounds.Left + NCAcceptBounds.Right Then
        If Y >= NCAcceptBounds.Top And Y <= NCAcceptBounds.Top + NCAcceptBounds.Bottom Then
            If NCAcceptState = 2 Then
                Call MenuState(MENU_STATE_ADDCHAR)
                NCAcceptState = 0
            Else
                NCAcceptState = 0
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleNewCharPanel_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function HandleSelCharPanel_MouseDown(X As Long, Y As Long)
Dim isHandled As Boolean
    'UseChar Button

   On Error GoTo errorhandler

    If isHandled = False Then
        If X >= UseCharButtonBounds.Left And X <= UseCharButtonBounds.Left + UseCharButtonBounds.Right Then
            If Y >= UseCharButtonBounds.Top And Y <= UseCharButtonBounds.Top + UseCharButtonBounds.Bottom Then
                UseCharButtonState = 2
            End If
        End If
    End If
    'DelChar Button
    If isHandled = False Then
        If X >= DelCharButtonBounds.Left And X <= DelCharButtonBounds.Left + DelCharButtonBounds.Right Then
            If Y >= DelCharButtonBounds.Top And Y <= DelCharButtonBounds.Top + DelCharButtonBounds.Bottom Then
                DelCharButtonState = 2
            End If
        End If
    End If
    'NewChar Button
    If isHandled = False Then
        If X >= NewCharButtonBounds.Left And X <= NewCharButtonBounds.Left + NewCharButtonBounds.Right Then
            If Y >= NewCharButtonBounds.Top And Y <= NewCharButtonBounds.Top + NewCharButtonBounds.Bottom Then
                NewCharButtonState = 2
            End If
        End If
    End If
    'PrevChar Button
    If isHandled = False Then
        If X >= PrevCharBounds.Left And X <= PrevCharBounds.Left + PrevCharBounds.Right Then
            If Y >= PrevCharBounds.Top And Y <= PrevCharBounds.Top + PrevCharBounds.Bottom Then
                prevCharState = 2
            End If
        End If
    End If
    'NextChar Button
    If isHandled = False Then
        If X >= NextCharBounds.Left And X <= NextCharBounds.Left + NextCharBounds.Right Then
            If Y >= NextCharBounds.Top And Y <= NextCharBounds.Top + NextCharBounds.Bottom Then
                NextCharState = 2
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleSelCharPanel_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function
Public Function HandleSelCharPanel_MouseUp(X As Long, Y As Long)
Dim buffer As clsBuffer
    'UseChar Button

   On Error GoTo errorhandler

    If X >= UseCharButtonBounds.Left And X <= UseCharButtonBounds.Left + UseCharButtonBounds.Right Then
        If Y >= UseCharButtonBounds.Top And Y <= UseCharButtonBounds.Top + UseCharButtonBounds.Bottom Then
            If UseCharButtonState = 2 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CUseChar
                buffer.WriteLong SelectedChar
                buffer.WriteLong 0
                SendData buffer.ToArray
                MenuStage = 0
                Set buffer = Nothing
                UseCharButtonState = 0
            Else
                UseCharButtonState = 0
            End If
        End If
    End If
    'DelChar Button
    If X >= DelCharButtonBounds.Left And X <= DelCharButtonBounds.Left + DelCharButtonBounds.Right Then
        If Y >= DelCharButtonBounds.Top And Y <= DelCharButtonBounds.Top + DelCharButtonBounds.Bottom Then
            If DelCharButtonState = 2 Then
                If MsgBox("Are you sure you want to delete the character?", vbYesNo, "Character Deletion") = vbYes Then
                    Set buffer = New clsBuffer
                    buffer.WriteLong CUseChar
                    buffer.WriteLong SelectedChar
                    buffer.WriteLong 1
                    SendData buffer.ToArray
                    Set buffer = Nothing
                End If
                DelCharButtonState = 0
            Else
                DelCharButtonState = 0
            End If
        End If
    End If
    'NewChar Button
    If X >= NewCharButtonBounds.Left And X <= NewCharButtonBounds.Left + NewCharButtonBounds.Right Then
        If Y >= NewCharButtonBounds.Top And Y <= NewCharButtonBounds.Top + NewCharButtonBounds.Bottom Then
            If NewCharButtonState = 2 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CUseChar
                buffer.WriteLong SelectedChar
                buffer.WriteLong 0
                SendData buffer.ToArray
                MenuStage = 0
                Set buffer = Nothing
                NewCharButtonState = 0
            Else
                NewCharButtonState = 0
            End If
        End If
    End If
    'PrevChar Button
    If X >= PrevCharBounds.Left And X <= PrevCharBounds.Left + PrevCharBounds.Right Then
        If Y >= PrevCharBounds.Top And Y <= PrevCharBounds.Top + PrevCharBounds.Bottom Then
            If prevCharState = 2 Then
                SelectedChar = SelectedChar - 1
                If SelectedChar <= 0 Then
                    SelectedChar = UBound(CharSelection)
                End If
                prevCharState = 0
            Else
                prevCharState = 0
            End If
        End If
    End If
    'NextChar Button
    If X >= NextCharBounds.Left And X <= NextCharBounds.Left + NextCharBounds.Right Then
        If Y >= NextCharBounds.Top And Y <= NextCharBounds.Top + NextCharBounds.Bottom Then
            If NextCharState = 2 Then
                SelectedChar = SelectedChar + 1
                If SelectedChar > UBound(CharSelection) Then
                    SelectedChar = 1
                End If
                NextCharState = 0
            Else
                NextCharState = 0
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleSelCharPanel_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function HandleRegisterPanel_MouseDown(X As Long, Y As Long)
Dim isHandled As Boolean
    'Check Textbox 1

   On Error GoTo errorhandler

    If X >= RUsernameBounds.Left And X <= RUsernameBounds.Left + RUsernameBounds.Right Then
        If Y >= RUsernameBounds.Top And Y <= RUsernameBounds.Top + RUsernameBounds.Bottom Then
            SelTextbox = 1
            isHandled = True
        End If
    End If
    'Check Textbox 2
    If isHandled = False Then
        If X >= RPasswordBounds.Left And X <= RPasswordBounds.Left + RPasswordBounds.Right Then
            If Y >= RPasswordBounds.Top And Y <= RPasswordBounds.Top + RPasswordBounds.Bottom Then
                SelTextbox = 2
                isHandled = True
            End If
        End If
    End If
    'Check Textbox 3
    If isHandled = False Then
        If X >= RPassword2Bounds.Left And X <= RPassword2Bounds.Left + RPassword2Bounds.Right Then
            If Y >= RPassword2Bounds.Top And Y <= RPassword2Bounds.Top + RPassword2Bounds.Bottom Then
                SelTextbox = 3
                isHandled = True
            End If
        End If
    End If
    'Register Button
    If isHandled = False Then
        If X >= RegisterButtonBounds.Left And X <= RegisterButtonBounds.Left + RegisterButtonBounds.Right Then
            If Y >= RegisterButtonBounds.Top And Y <= RegisterButtonBounds.Top + RegisterButtonBounds.Bottom Then
                RegisterButtonState = 2
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleRegisterPanel_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function HandleRegisterPanel_MouseUp(X As Long, Y As Long)
    'Register Button

   On Error GoTo errorhandler

    If X >= RegisterButtonBounds.Left And X <= RegisterButtonBounds.Left + RegisterButtonBounds.Right Then
        If Y >= RegisterButtonBounds.Top And Y <= RegisterButtonBounds.Top + RegisterButtonBounds.Bottom Then
            If RegisterButtonState = 2 Then
                If txtPassword <> TxtPassword2 Then
                    MsgBox "Passwords do not match!", vbOKOnly, Servers(ServerIndex).Game_Name
                    Exit Function
                End If
                MenuState MENU_STATE_NEWACCOUNT
                RegisterButtonState = 0
            Else
                RegisterButtonState = 0
            End If
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "HandleRegisterPanel_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Sub HandleMenu_MouseDown()
    Dim X As Long, Y As Long, isHandled As Boolean

   On Error GoTo errorhandler

    X = GlobalX
    Y = GlobalY
    If InIntro = 1 Then
        If IntroSkip = 1 Then
            InIntro = 0
            FadeType = 0
            FadeAmount = 0
        End If
        Exit Sub
    End If
    If X >= GUIContainerX And X <= GUIContainerX + GUIContainerWidth Then
        If Y >= GUIContainerY And Y <= GUIContainerY + GUIContainerHeight Then
            'In Container... Now we have to narrow down the click
            'First, check the menus...
            Select Case MenuStage
                Case 1
                    'Login
                    If X >= LoginPanelBounds.Left And X <= LoginPanelBounds.Left + LoginPanelBounds.Right Then
                        If Y >= LoginPanelBounds.Top And Y <= LoginPanelBounds.Top + LoginPanelBounds.Bottom Then
                            isHandled = HandleLoginPanel_MouseDown(GlobalX, GlobalY)
                        End If
                    End If
                Case 2
                    'Register
                    If X >= RegisterPanelBounds.Left And X <= RegisterPanelBounds.Left + RegisterPanelBounds.Right Then
                        If Y >= RegisterPanelBounds.Top And Y <= RegisterPanelBounds.Top + RegisterPanelBounds.Bottom Then
                            isHandled = HandleRegisterPanel_MouseDown(GlobalX, GlobalY)
                        End If
                    End If
                Case 3
                    'Credits
                Case 4
                    'Char
                    If X >= CharPanelBounds.Left And X <= CharPanelBounds.Left + CharPanelBounds.Right Then
                        If Y >= CharPanelBounds.Top And Y <= CharPanelBounds.Top + CharPanelBounds.Bottom Then
                            isHandled = HandleSelCharPanel_MouseDown(GlobalX, GlobalY)
                        End If
                    End If
                Case 5
                    'Char
                    If X >= NewCharPanelBounds.Left And X <= NewCharPanelBounds.Left + NewCharPanelBounds.Right Then
                        If Y >= NewCharPanelBounds.Top And Y <= NewCharPanelBounds.Top + NewCharPanelBounds.Bottom Then
                            isHandled = HandleNewCharPanel_MouseDown(GlobalX, GlobalY)
                        End If
                    End If
            End Select
            'If we made it this far then they clicked on nothing or one of the buttons, lets check them before leaving.
            'Login Button
            If isHandled = False Then
                If X >= OpenLoginBounds.Left And X <= OpenLoginBounds.Left + OpenLoginBounds.Right Then
                    If Y >= OpenLoginBounds.Top And Y <= OpenLoginBounds.Top + OpenLoginBounds.Bottom Then
                        OpenLoginState = 2
                    Else
                        OpenLoginState = 0
                    End If
                Else
                        OpenLoginState = 0
                End If
                If X >= OpenRegisterBounds.Left And X <= OpenRegisterBounds.Left + OpenRegisterBounds.Right Then
                    If Y >= OpenRegisterBounds.Top And Y <= OpenRegisterBounds.Top + OpenRegisterBounds.Bottom Then
                        OpenRegisterState = 2
                    Else
                        OpenRegisterState = 0
                    End If
                Else
                        OpenRegisterState = 0
                End If
                If X >= OpenCreditsBounds.Left And X <= OpenCreditsBounds.Left + OpenCreditsBounds.Right Then
                    If Y >= OpenCreditsBounds.Top And Y <= OpenCreditsBounds.Top + OpenCreditsBounds.Bottom Then
                        OpenCreditsState = 2
                    Else
                        OpenCreditsState = 0
                    End If
                Else
                        OpenCreditsState = 0
                End If
                If X >= ExitGameBounds.Left And X <= ExitGameBounds.Left + ExitGameBounds.Right Then
                    If Y >= ExitGameBounds.Top And Y <= ExitGameBounds.Top + ExitGameBounds.Bottom Then
                        ExitGameState = 2
                    Else
                        ExitGameState = 0
                    End If
                Else
                        ExitGameState = 0
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMenu_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub HandleMenu_MouseUp()
Dim X As Long, Y As Long, isHandled As Boolean

   On Error GoTo errorhandler

    X = GlobalX
    Y = GlobalY
    If X >= GUIContainerX And X <= GUIContainerX + GUIContainerWidth Then
        If Y >= GUIContainerY And Y <= GUIContainerY + GUIContainerHeight Then
            'In Container... Now we have to narrow down the click
            'First, check the menus...
            Select Case MenuStage
                Case 1
                    'Login
                    If X >= LoginPanelBounds.Left And X <= LoginPanelBounds.Left + LoginPanelBounds.Right Then
                        If Y >= LoginPanelBounds.Top And Y <= LoginPanelBounds.Top + LoginPanelBounds.Bottom Then
                            isHandled = HandleLoginPanel_MouseUp(GlobalX, GlobalY)
                        End If
                    End If
                Case 2
                    'Register
                    If X >= RegisterPanelBounds.Left And X <= RegisterPanelBounds.Left + RegisterPanelBounds.Right Then
                        If Y >= RegisterPanelBounds.Top And Y <= RegisterPanelBounds.Top + RegisterPanelBounds.Bottom Then
                            isHandled = HandleRegisterPanel_MouseUp(GlobalX, GlobalY)
                        End If
                    End If
                Case 4
                    'Char
                    If X >= CharPanelBounds.Left And X <= CharPanelBounds.Left + CharPanelBounds.Right Then
                        If Y >= CharPanelBounds.Top And Y <= CharPanelBounds.Top + CharPanelBounds.Bottom Then
                            isHandled = HandleSelCharPanel_MouseUp(GlobalX, GlobalY)
                        End If
                    End If
                Case 5
                    'Char
                    If X >= NewCharPanelBounds.Left And X <= NewCharPanelBounds.Left + NewCharPanelBounds.Right Then
                        If Y >= NewCharPanelBounds.Top And Y <= NewCharPanelBounds.Top + NewCharPanelBounds.Bottom Then
                            isHandled = HandleNewCharPanel_MouseUp(GlobalX, GlobalY)
                        End If
                    End If
            End Select
                    'If we made it this far then they clicked on nothing or one of the buttons, lets check them before leaving.
            'Login Button
            If isHandled = False Then
                If OpenLoginBounds.Right > 0 And OpenLoginBounds.Bottom > 0 Then
                    If X >= OpenLoginBounds.Left And X <= OpenLoginBounds.Left + OpenLoginBounds.Right Then
                        If Y >= OpenLoginBounds.Top And Y <= OpenLoginBounds.Top + OpenLoginBounds.Bottom Then
                            If OpenLoginState = 2 Then
                                'Open Login Panel
                                If MenuStage = 1 Then
                                    MenuStage = 0
                                    SelTextbox = 0
                                Else
                                    MenuStage = 1
                                    SelTextbox = 1
                                    If Servers(ServerIndex).SavePass = 1 Then
                                        TxtUsername = Trim$(Servers(ServerIndex).Username)
                                        txtPassword = Trim$(Servers(ServerIndex).Password)
                                        TxtPassword2 = ""
                                    Else
                                        TxtUsername = ""
                                        txtPassword = ""
                                        TxtPassword2 = ""
                                    End If
                                End If
                                isHandled = True
                            End If
                        End If
                    End If
                End If
            End If
            If isHandled = False Then
                If OpenRegisterBounds.Right > 0 And OpenRegisterBounds.Bottom > 0 Then
                    If X >= OpenRegisterBounds.Left And X <= OpenRegisterBounds.Left + OpenRegisterBounds.Right Then
                        If Y >= OpenRegisterBounds.Top And Y <= OpenRegisterBounds.Top + OpenRegisterBounds.Bottom Then
                            If OpenRegisterState = 2 Then
                                'Open Register Panel
                                If MenuStage = 2 Then
                                    MenuStage = 0
                                    SelTextbox = 0
                                    TxtUsername = ""
                                    txtPassword = ""
                                    TxtPassword2 = ""
                                Else
                                    MenuStage = 2
                                    SelTextbox = 1
                                End If
                                isHandled = True
                            End If
                        End If
                    End If
                End If
            End If
            If isHandled = False Then
                If OpenCreditsBounds.Right > 0 And OpenCreditsBounds.Bottom > 0 Then
                    If X >= OpenCreditsBounds.Left And X <= OpenCreditsBounds.Left + OpenCreditsBounds.Right Then
                        If Y >= OpenCreditsBounds.Top And Y <= OpenCreditsBounds.Top + OpenCreditsBounds.Bottom Then
                            If OpenCreditsState = 2 Then
                                'Open Credits Panel
                                If MenuStage = 3 Then
                                    MenuStage = 0
                                Else
                                    MenuStage = 3
                                End If
                                isHandled = True
                            End If
                        End If
                    End If
                End If
            End If
            If isHandled = False Then
                If ExitGameBounds.Right > 0 And ExitGameBounds.Bottom > 0 Then
                    If X >= ExitGameBounds.Left And X <= ExitGameBounds.Left + ExitGameBounds.Right Then
                        If Y >= ExitGameBounds.Top And Y <= ExitGameBounds.Top + ExitGameBounds.Bottom Then
                            If ExitGameState = 2 Then
                                'Exit Game Button
                                DestroyGame
                                isHandled = True
                            End If
                        End If
                    End If
                End If
            End If
            End If
    End If
    OpenLoginState = 0
    OpenRegisterState = 0
    OpenCreditsState = 0
    ExitGameState = 0
    LoginButtonState = 0
    RegisterButtonState = 0


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMenu_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Public Sub HandleMenuKeypress(KeyAscii As Integer)
    Dim buffer As clsBuffer

   On Error GoTo errorhandler

    If InIntro = 1 Then
        If IntroSkip = 1 Then
            InIntro = 0
            FadeType = 0
            FadeAmount = 0
        End If
        Exit Sub
    End If
    If KeyAscii = vbKeyEscape Then DestroyGame
    If KeyAscii = vbKeyTab And TabTick > GetTickCount Then Exit Sub
    If MenuStage = 1 Then
        If KeyAscii = vbKeyTab Then
            If SelTextbox = 1 Then
                SelTextbox = 2
            Else
                SelTextbox = 1
            End If
            TabTick = GetTickCount + 200
        ElseIf KeyAscii = vbKeyReturn Then
            If SelTextbox = 1 Then
                SelTextbox = 2
            Else
                MenuState MENU_STATE_LOGIN
            End If
        Else
            If SelTextbox = 1 Then
                If KeyAscii = vbKeyBack Then
                    If Len(TxtUsername) > 0 Then
                        TxtUsername = Left(TxtUsername, Len(TxtUsername) - 1)
                    End If
                Else
                    TxtUsername = TxtUsername & Chr(KeyAscii)
                End If
            ElseIf SelTextbox = 2 Then
                If KeyAscii = vbKeyBack Then
                    If Len(txtPassword) > 0 Then
                        txtPassword = Left(txtPassword, Len(txtPassword) - 1)
                    End If
                Else
                    txtPassword = txtPassword & Chr(KeyAscii)
                End If
            End If
        End If
    ElseIf MenuStage = 2 Then
        If KeyAscii = vbKeyTab Then
            If SelTextbox = 1 Then
                SelTextbox = 2
            ElseIf SelTextbox = 2 Then
                SelTextbox = 3
            ElseIf SelTextbox = 3 Then
                SelTextbox = 1
            End If
            TabTick = GetTickCount + 200
        ElseIf KeyAscii = vbKeyReturn Then
            If SelTextbox = 1 Then
                SelTextbox = 2
            ElseIf SelTextbox = 2 Then
                SelTextbox = 3
            ElseIf SelTextbox = 3 Then
                MenuState MENU_STATE_NEWACCOUNT
            End If
        Else
            If SelTextbox = 1 Then
                If KeyAscii = vbKeyBack Then
                    If Len(TxtUsername) > 0 Then
                        TxtUsername = Left(TxtUsername, Len(TxtUsername) - 1)
                    End If
                Else
                    TxtUsername = TxtUsername & Chr(KeyAscii)
                End If
            ElseIf SelTextbox = 2 Then
                If KeyAscii = vbKeyBack Then
                    If Len(txtPassword) > 0 Then
                        txtPassword = Left(txtPassword, Len(txtPassword) - 1)
                    End If
                Else
                    txtPassword = txtPassword & Chr(KeyAscii)
                End If
            ElseIf SelTextbox = 3 Then
                If KeyAscii = vbKeyBack Then
                    If Len(TxtPassword2) > 0 Then
                        TxtPassword2 = Left(TxtPassword2, Len(TxtPassword2) - 1)
                    End If
                Else
                    TxtPassword2 = TxtPassword2 & Chr(KeyAscii)
                End If
            End If
        End If
    ElseIf MenuStage = 4 Then
        If KeyAscii = Asc("a") Or KeyAscii = Asc("A") Or KeyAscii = vbKeyLeft Then
            SelectedChar = SelectedChar - 1
            If SelectedChar <= 0 Then SelectedChar = UBound(CharSelection)
        ElseIf KeyAscii = Asc("d") Or KeyAscii = Asc("D") Or KeyAscii = vbKeyRight Then
            SelectedChar = SelectedChar + 1
            If SelectedChar > UBound(CharSelection) Then SelectedChar = 1
        ElseIf KeyAscii = vbKeyReturn Then
            Set buffer = New clsBuffer
            buffer.WriteLong CUseChar
            buffer.WriteLong SelectedChar
            buffer.WriteLong 0
            SendData buffer.ToArray
            MenuStage = 0
            frmLoad.Visible = True
            SetStatus "Sending character status to server."
            Set buffer = Nothing
        ElseIf KeyAscii = vbKeyDelete Then
            If MsgBox("Are you sure you want to delete the character?", vbYesNo, "Character Deletion") = vbYes Then
                Set buffer = New clsBuffer
                buffer.WriteLong CUseChar
                buffer.WriteLong SelectedChar
                buffer.WriteLong 1
                SendData buffer.ToArray
                Set buffer = Nothing
            End If
        End If
    ElseIf MenuStage = 5 Then
        If KeyAscii = vbKeyTab Then
            TabTick = GetTickCount + 200
        ElseIf KeyAscii = vbKeyReturn Then
            MenuState MENU_STATE_ADDCHAR
        Else
            If SelTextbox = 1 Then
                If KeyAscii = vbKeyBack Then
                    If Len(TxtUsername) > 0 Then
                        TxtUsername = Left(TxtUsername, Len(TxtUsername) - 1)
                    End If
                Else
                    TxtUsername = TxtUsername & Chr(KeyAscii)
                End If
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "HandleMenuKeypress", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub


Sub Chatbox_MouseDown(X As Long, Y As Long)
    'My Chat Line

   On Error GoTo errorhandler

    If X >= MyTextBounds.Left And X <= MyTextBounds.Left + MyTextBounds.Right Then
        If Y >= MyTextBounds.Top And Y <= MyTextBounds.Top + MyTextBounds.Bottom Then
            chatOn = True
            Exit Sub
        End If
    End If
    'ChatUp Button
    If X >= ChatUpBtnBounds.Left And X <= ChatUpBtnBounds.Left + ChatUpBtnBounds.Right Then
        If Y >= ChatUpBtnBounds.Top And Y <= ChatUpBtnBounds.Top + ChatUpBtnBounds.Bottom Then
            ChatUpBtnState = 2
            Exit Sub
        End If
    End If
    'ChatDown Button
    If X >= ChatDownBtnBounds.Left And X <= ChatDownBtnBounds.Left + ChatDownBtnBounds.Right Then
        If Y >= ChatDownBtnBounds.Top And Y <= ChatDownBtnBounds.Top + ChatDownBtnBounds.Bottom Then
            ChatDownBtnState = 2
            Exit Sub
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Chatbox_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub Chatbox_MouseUp(X As Long, Y As Long)
    'ChatUpBtn

   On Error GoTo errorhandler

    If X >= ChatUpBtnBounds.Left And X <= ChatUpBtnBounds.Left + ChatUpBtnBounds.Right Then
        If Y >= ChatUpBtnBounds.Top And Y <= ChatUpBtnBounds.Top + ChatUpBtnBounds.Bottom Then
            If ChatUpBtnState = 2 Then
                dialogueHandler 2
                ChatUpBtnState = 0
            Else
                ChatUpBtnState = 0
            End If
        End If
    End If
    'ChatDownBtn
    If X >= ChatDownBtnBounds.Left And X <= ChatDownBtnBounds.Left + ChatDownBtnBounds.Right Then
        If Y >= ChatDownBtnBounds.Top And Y <= ChatDownBtnBounds.Top + ChatDownBtnBounds.Bottom Then
            If ChatDownBtnState = 2 Then
                            ChatDownBtnState = 0
            Else
                ChatDownBtnState = 0
            End If
        End If
    End If
    DragInvSlotNum = 0
    DragSpell = 0


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Chatbox_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub Inventory_MouseDown(X As Long, Y As Long)
    Dim InvNum As Long, Button As Long, Value As Long, multiplier As Double


   On Error GoTo errorhandler

    InvNum = IsInvItem(X, Y)
    Button = MouseBtn

    If Button = 1 Then
        'Inventory Button
        If X >= CloseInvBtnBounds.Left And X <= CloseInvBtnBounds.Left + CloseInvBtnBounds.Right Then
            If Y >= CloseInvBtnBounds.Top And Y <= CloseInvBtnBounds.Top + CloseInvBtnBounds.Bottom Then
                CloseInvBtnState = 2
                Exit Sub
            End If
        End If
        
        If InvNum <> 0 Then
            If InMailbox Then Exit Sub
            DragInvSlotNum = InvNum
            IsReallyShop = False
            AddText Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Desc), BrightBlue
            ' are we in a shop?
            If InShop > 0 Then
                multiplier = Shop(InShop).BuyRate / 100
                Value = Item(GetPlayerInvItemNum(MyIndex, InvNum)).Price * multiplier
                If Value > 0 Then
                    AddText "You can sell this item for " & Value & " gold.", White
                Else
                    AddText "The shop does not want this item.", BrightRed
                End If
            End If
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 And Not InMailbox Then
            If InvNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
                    If GetPlayerInvItemValue(MyIndex, InvNum) > 1 Then
                        CurrencyMenu = 1 ' drop
                        CurrencyCaption = "How many do you want to drop?"
                        CurrencyItem = InvNum
                        CurrencyText = ""
                    Else
                        Call SendDropItem(InvNum, 0)
                    End If
                Else
                    Call SendDropItem(InvNum, 0)
                End If
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Inventory_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Sub Inventory_MouseUp(X As Long, Y As Long)
    Dim i As Long
    Dim rec_pos As rect, buffer As clsBuffer
    'CloseInvBtn

   On Error GoTo errorhandler

    If X >= CloseInvBtnBounds.Left And X <= CloseInvBtnBounds.Left + CloseInvBtnBounds.Right Then
        If Y >= CloseInvBtnBounds.Top And Y <= CloseInvBtnBounds.Top + CloseInvBtnBounds.Bottom Then
            If CloseInvBtnState = 2 Then
                If CurrentGameMenu = 1 Then
                    CurrentGameMenu = 0
                End If
                Exit Sub
                CloseInvBtnState = 0
            Else
                CloseInvBtnState = 0
            End If
        End If
    End If
    If DragTradeSlotNum > 0 Then
        UntradeItem DragTradeSlotNum
        DragTradeSlotNum = 0
        Exit Sub
    End If
    If DragMailboxItem = 1 Then
        If MailBoxMenu = 2 Then
            Set buffer = New clsBuffer
            buffer.WriteLong CTakeMailItem
            buffer.WriteLong Mail(SelectedMail).Index
            SendData buffer.ToArray
            Set buffer = Nothing
        ElseIf MailBoxMenu = 3 Then
            MailItem = 0
            MailItemValue = 0
        End If
        Exit Sub
    End If

    If DragInvSlotNum > 0 Then
        If IsReallyShop = False Then
            i = IsInvItem(X, Y, True)
            If i <> DragInvSlotNum And i > 0 Then
                SendChangeInvSlots DragInvSlotNum, i
            End If
        Else
            If X >= InvItemsBounds.Left And X <= InvItemsBounds.Left + InvItemsBounds.Right Then
                If Y >= InvItemsBounds.Top And Y <= InvItemsBounds.Top + InvItemsBounds.Bottom Then
                    'Buy item
                    BuyItem DragInvSlotNum
                    DragInvSlotNum = 0
                    IsReallyShop = False
                End If
            End If
        End If
    End If
    If DragBankSlotNum > 0 Then
         If GetBankItemNum(DragBankSlotNum) = ITEM_TYPE_NONE Then Exit Sub
                  If Item(GetBankItemNum(DragBankSlotNum)).Stackable = 1 Then
                CurrencyMenu = 3 ' withdraw
                CurrencyCaption = "How many do you want to withdraw?"
                CurrencyItem = DragBankSlotNum
                CurrencyText = ""
                Exit Sub
            End If
                 WithdrawItem DragBankSlotNum, 0
         Exit Sub
    End If
    DragBankSlotNum = 0
    DragInvSlotNum = 0
    DragSpell = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Inventory_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Inventory_DblClick(X As Long, Y As Long)
    Dim InvNum As Long
    Dim Value As Long
    Dim multiplier As Double
    Dim i As Long


   On Error GoTo errorhandler

    DragInvSlotNum = 0
    DragSpell = 0
    InvNum = IsInvItem(X, Y)

    If InvNum <> 0 Then
        ' are we in a shop?
        If InShop > 0 Then
            SellItem InvNum
            Exit Sub
        End If
            ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
                CurrencyMenu = 2 ' deposit
                CurrencyCaption = "How many do you want to deposit?"
                CurrencyItem = InvNum
                CurrencyText = vbNullString
                Exit Sub
            End If
                        Call DepositItem(InvNum, 0)
            Exit Sub
        End If
            ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For i = 1 To MAX_INV
                If TradeYourOffer(i).Num = InvNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)).Stackable = 1 Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).Num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
                    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
                CurrencyMenu = 4 ' offer in trade
                CurrencyCaption = "How many do you want to trade?"
                CurrencyItem = InvNum
                CurrencyText = ""
                Exit Sub
            End If
                    Call TradeItem(InvNum, 0)
            Exit Sub
        End If
            If InMailbox Then
            If MailBoxMenu = 3 Then
                If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
                    CurrencyMenu = 5 ' offer in message
                    CurrencyCaption = "How many do you want to give?"
                    CurrencyItem = InvNum
                    CurrencyText = ""
                    Exit Sub
                Else
                    MailItem = InvNum
                    MailItemValue = 1
                    Exit Sub
                End If
            End If
        End If
            ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(InvNum)
        Exit Sub
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Inventory_DblClick", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
        
End Sub
Sub GameMenu_MouseDown(X As Long, Y As Long)
    'Inventory Button

   On Error GoTo errorhandler

    If X >= InventoryBtnBounds.Left And X <= InventoryBtnBounds.Left + InventoryBtnBounds.Right Then
        If Y >= InventoryBtnBounds.Top And Y <= InventoryBtnBounds.Top + InventoryBtnBounds.Bottom Then
            InventoryBtnState = 2
            Exit Sub
        End If
    End If
    'Skills Button
    If X >= SkillsBtnBounds.Left And X <= SkillsBtnBounds.Left + SkillsBtnBounds.Right Then
        If Y >= SkillsBtnBounds.Top And Y <= SkillsBtnBounds.Top + SkillsBtnBounds.Bottom Then
            SkillsBtnState = 2
            Exit Sub
        End If
    End If
    'Character Button
    If X >= CharacterBtnBounds.Left And X <= CharacterBtnBounds.Left + CharacterBtnBounds.Right Then
        If Y >= CharacterBtnBounds.Top And Y <= CharacterBtnBounds.Top + CharacterBtnBounds.Bottom Then
            CharacterBtnState = 2
            Exit Sub
        End If
    End If
    'Options Button
    If X >= OptionsBtnBounds.Left And X <= OptionsBtnBounds.Left + OptionsBtnBounds.Right Then
        If Y >= OptionsBtnBounds.Top And Y <= OptionsBtnBounds.Top + OptionsBtnBounds.Bottom Then
            OptionsBtnState = 2
            Exit Sub
        End If
    End If
    'Trade Button
    If X >= TradeBtnBounds.Left And X <= TradeBtnBounds.Left + TradeBtnBounds.Right Then
        If Y >= TradeBtnBounds.Top And Y <= TradeBtnBounds.Top + TradeBtnBounds.Bottom Then
            TradeBtnState = 2
            Exit Sub
        End If
    End If
    'Party Button
    If X >= PartyBtnBounds.Left And X <= PartyBtnBounds.Left + PartyBtnBounds.Right Then
        If Y >= PartyBtnBounds.Top And Y <= PartyBtnBounds.Top + PartyBtnBounds.Bottom Then
            PartyBtnState = 2
            Exit Sub
        End If
    End If
    'Friends Button
    If X >= FriendsBtnBounds.Left And X <= FriendsBtnBounds.Left + FriendsBtnBounds.Right Then
        If Y >= FriendsBtnBounds.Top And Y <= FriendsBtnBounds.Top + FriendsBtnBounds.Bottom Then
            FriendsBtnState = 2
            Exit Sub
        End If
    End If
    'Quests Button
    If X >= QuestsBtnBounds.Left And X <= QuestsBtnBounds.Left + QuestsBtnBounds.Right Then
        If Y >= QuestsBtnBounds.Top And Y <= QuestsBtnBounds.Top + QuestsBtnBounds.Bottom Then
            QuestsBtnState = 2
            Exit Sub
        End If
    End If
    'Pets Button
    If X >= PetsBtnBounds.Left And X <= PetsBtnBounds.Left + PetsBtnBounds.Right Then
        If Y >= PetsBtnBounds.Top And Y <= PetsBtnBounds.Top + PetsBtnBounds.Bottom Then
            PetsBtnState = 2
            Exit Sub
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "GameMenu_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub GameMenu_MouseUp(X As Long, Y As Long)
    'InventoryBtn

   On Error GoTo errorhandler

    If X >= InventoryBtnBounds.Left And X <= InventoryBtnBounds.Left + InventoryBtnBounds.Right Then
        If Y >= InventoryBtnBounds.Top And Y <= InventoryBtnBounds.Top + InventoryBtnBounds.Bottom Then
            If InventoryBtnState = 2 Then
                If CurrentGameMenu = 1 Then
                    CurrentGameMenu = 0
                Else
                    CurrentGameMenu = 1
                End If
                Exit Sub
                InventoryBtnState = 0
            Else
                InventoryBtnState = 0
            End If
        End If
    End If
    'SkillsBtn
    If X >= SkillsBtnBounds.Left And X <= SkillsBtnBounds.Left + SkillsBtnBounds.Right Then
        If Y >= SkillsBtnBounds.Top And Y <= SkillsBtnBounds.Top + SkillsBtnBounds.Bottom Then
            If SkillsBtnState = 2 Then
                If CurrentGameMenu = 2 Then
                    CurrentGameMenu = 0
                Else
                    CurrentGameMenu = 2
                End If
                Exit Sub
                SkillsBtnState = 0
            Else
                SkillsBtnState = 0
            End If
        End If
    End If
    'CharacterBtn
    If X >= CharacterBtnBounds.Left And X <= CharacterBtnBounds.Left + CharacterBtnBounds.Right Then
        If Y >= CharacterBtnBounds.Top And Y <= CharacterBtnBounds.Top + CharacterBtnBounds.Bottom Then
            If CharacterBtnState = 2 Then
                If CurrentGameMenu = 3 Then
                    CurrentGameMenu = 0
                Else
                    CurrentGameMenu = 3
                End If
                Exit Sub
                CharacterBtnState = 0
            Else
                CharacterBtnState = 0
            End If
        End If
    End If
    'OptionsBtn
    If X >= OptionsBtnBounds.Left And X <= OptionsBtnBounds.Left + OptionsBtnBounds.Right Then
        If Y >= OptionsBtnBounds.Top And Y <= OptionsBtnBounds.Top + OptionsBtnBounds.Bottom Then
            If OptionsBtnState = 2 Then
                If CurrentGameMenu = 4 Then
                    CurrentGameMenu = 0
                Else
                    CurrentGameMenu = 4
                End If
                Exit Sub
                OptionsBtnState = 0
            Else
                OptionsBtnState = 0
            End If
        End If
    End If
    'TradeBtn
    If X >= TradeBtnBounds.Left And X <= TradeBtnBounds.Left + TradeBtnBounds.Right Then
        If Y >= TradeBtnBounds.Top And Y <= TradeBtnBounds.Top + TradeBtnBounds.Bottom Then
            If TradeBtnState = 2 Then
                If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                    SendTradeRequest
                    ' play sound
                    PlaySound Sound_ButtonClick, -1, -1
                Else
                    AddText "Invalid trade target.", BrightRed
                End If
                Exit Sub
                TradeBtnState = 0
            Else
                TradeBtnState = 0
            End If
        End If
    End If
    'PartyBtn
    If X >= PartyBtnBounds.Left And X <= PartyBtnBounds.Left + PartyBtnBounds.Right Then
        If Y >= PartyBtnBounds.Top And Y <= PartyBtnBounds.Top + PartyBtnBounds.Bottom Then
            If PartyBtnState = 2 Then
                If CurrentGameMenu = 5 Then
                    CurrentGameMenu = 0
                Else
                    CurrentGameMenu = 5
                End If
                Exit Sub
                PartyBtnState = 0
            Else
                PartyBtnState = 0
            End If
        End If
    End If
    'FriendsBtn
    If X >= FriendsBtnBounds.Left And X <= FriendsBtnBounds.Left + FriendsBtnBounds.Right Then
        If Y >= FriendsBtnBounds.Top And Y <= FriendsBtnBounds.Top + FriendsBtnBounds.Bottom Then
            If FriendsBtnState = 2 Then
                If CurrentGameMenu = 6 Then
                    CurrentGameMenu = 0
                Else
                    CurrentGameMenu = 6
                End If
                Exit Sub
                FriendsBtnState = 0
            Else
                FriendsBtnState = 0
            End If
        End If
    End If
    'QuestsBtn
    If X >= QuestsBtnBounds.Left And X <= QuestsBtnBounds.Left + QuestsBtnBounds.Right Then
        If Y >= QuestsBtnBounds.Top And Y <= QuestsBtnBounds.Top + QuestsBtnBounds.Bottom Then
            If QuestsBtnState = 2 Then
                If CurrentGameMenu = 7 Then
                    CurrentGameMenu = 0
                Else
                    CurrentGameMenu = 7
                    QuestSelection = 0
                End If
                Exit Sub
                QuestsBtnState = 0
            Else
                QuestsBtnState = 0
            End If
        End If
    End If
    'PetsBtn
    If X >= PetsBtnBounds.Left And X <= PetsBtnBounds.Left + PetsBtnBounds.Right Then
        If Y >= PetsBtnBounds.Top And Y <= PetsBtnBounds.Top + PetsBtnBounds.Bottom Then
            If PetsBtnState = 2 Then
                If CurrentGameMenu = 8 Then
                    CurrentGameMenu = 0
                Else
                    If Player(MyIndex).Pet.Alive Then
                        CurrentGameMenu = 8
                    Else
                        CurrentGameMenu = 0
                        AddText "You do not have a pet!", BrightRed, 255
                    End If
                End If
                Exit Sub
                PetsBtnState = 0
            Else
                PetsBtnState = 0
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "GameMenu_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub Spells_MouseDown(X As Long, Y As Long)
Dim Spellnum As Long, Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn

    Spellnum = IsPlayerSpell(X, Y)
    If Button = 1 Then ' left click
        If X >= CloseSpellsBtnBounds.Left And X <= CloseSpellsBtnBounds.Left + CloseSpellsBtnBounds.Right Then
            If Y >= CloseSpellsBtnBounds.Top And Y <= CloseSpellsBtnBounds.Top + CloseSpellsBtnBounds.Bottom Then
                CloseSpellsBtnState = 2
                Exit Sub
            End If
        End If
            If Spellnum <> 0 Then
                AddText Trim$(spell(PlayerSpells(Spellnum)).Desc), BrightBlue
                DragSpell = Spellnum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If Spellnum <> 0 Then
            dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(spell(PlayerSpells(Spellnum)).Name) & "?", DIALOGUE_TYPE_FORGET, True, Spellnum
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Spells_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Spells_DblClick(X As Long, Y As Long)
Dim Spellnum As Long, Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn

    Spellnum = IsPlayerSpell(GlobalX, GlobalY)

    If Spellnum <> 0 Then
        Call CastSpell(Spellnum)
        Exit Sub
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Spells_DblClick", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As rect
    Dim i As Long


   On Error GoTo errorhandler

    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS

        If PlayerSpells(i) > 0 And PlayerSpells(i) <= MAX_SPELLS Then

            With tempRec
                .Top = SpellIconsBounds.Top + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellIconsBounds.Left + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With
                    If X >= SpellIconsBounds.Left And X <= SpellIconsBounds.Left + SpellIconsBounds.Right Then
                If Y >= SpellIconsBounds.Top And Y <= SpellIconsBounds.Top + SpellIconsBounds.Bottom Then
                    If X >= tempRec.Left And X <= tempRec.Right Then
                        If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                            IsPlayerSpell = i
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If

    Next




   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub Spells_MouseMove(X As Long, Y As Long)
Dim SpellSlot As Long, Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn

    SpellSlot = IsPlayerSpell(X, Y)
    If DragSpell > 0 Then
        Else
        If SpellSlot <> 0 Then
            UpdateSpellWindow PlayerSpells(SpellSlot), 0, 0
            LastSpellDesc = PlayerSpells(SpellSlot)
            Exit Sub
        End If
    End If
    SpellDescVisible = False
    LastSpellDesc = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Spells_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Sub Spells_MouseUp(X As Long, Y As Long)
Dim SpellSlot As Long, Button As Long, i As Long, rec As rect, rec_pos As rect

   On Error GoTo errorhandler

    Button = MouseBtn
    'CloseSpellsBtn
    If X >= CloseSpellsBtnBounds.Left And X <= CloseSpellsBtnBounds.Left + CloseSpellsBtnBounds.Right Then
        If Y >= CloseSpellsBtnBounds.Top And Y <= CloseSpellsBtnBounds.Top + CloseSpellsBtnBounds.Bottom Then
            If CloseSpellsBtnState = 2 Then
                If CurrentGameMenu = 2 Then
                    CurrentGameMenu = 0
                End If
                Exit Sub
                CloseSpellsBtnState = 0
            Else
                CloseSpellsBtnState = 0
            End If
        End If
    End If
    If DragSpell > 0 Then
        ' drag + drop
        For i = 1 To MAX_PLAYER_SPELLS
            With rec_pos
                .Top = SpellIconsBounds.Top + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellIconsBounds.Left + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With


            If X >= SpellIconsBounds.Left And X <= SpellIconsBounds.Left + SpellIconsBounds.Right Then
                If Y >= SpellIconsBounds.Top And Y <= SpellIconsBounds.Top + SpellIconsBounds.Bottom Then
                    If X >= rec_pos.Left And X <= rec_pos.Right Then
                        If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                            If DragSpell <> i Then
                                SendChangeSpellSlots DragSpell, i
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If

    DragSpell = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Spells_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Character_MouseUp(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect
    'AddStrength Btn

   On Error GoTo errorhandler

    If X >= AddStrengthBtnBounds.Left And X <= AddStrengthBtnBounds.Left + AddStrengthBtnBounds.Right Then
        If Y >= AddStrengthBtnBounds.Top And Y <= AddStrengthBtnBounds.Top + AddStrengthBtnBounds.Bottom Then
            If AddStrengthBtnState = 2 Then
                If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
                SendTrainStat Stats.Strength
                AddStrengthBtnState = 0
            Else
                AddStrengthBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'AddEndurance Btn
    If X >= AddEnduranceBtnBounds.Left And X <= AddEnduranceBtnBounds.Left + AddEnduranceBtnBounds.Right Then
        If Y >= AddEnduranceBtnBounds.Top And Y <= AddEnduranceBtnBounds.Top + AddEnduranceBtnBounds.Bottom Then
            If AddEnduranceBtnState = 2 Then
                If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
                SendTrainStat Stats.Endurance
                AddEnduranceBtnState = 0
            Else
                AddEnduranceBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'AddIntelligence Btn
    If X >= AddIntelligenceBtnBounds.Left And X <= AddIntelligenceBtnBounds.Left + AddIntelligenceBtnBounds.Right Then
        If Y >= AddIntelligenceBtnBounds.Top And Y <= AddIntelligenceBtnBounds.Top + AddIntelligenceBtnBounds.Bottom Then
            If AddIntelligenceBtnState = 2 Then
                If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
                SendTrainStat Stats.Intelligence
                AddIntelligenceBtnState = 0
            Else
                AddIntelligenceBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'AddAgility Btn
    If X >= AddAgilityBtnBounds.Left And X <= AddAgilityBtnBounds.Left + AddAgilityBtnBounds.Right Then
        If Y >= AddAgilityBtnBounds.Top And Y <= AddAgilityBtnBounds.Top + AddAgilityBtnBounds.Bottom Then
            If AddAgilityBtnState = 2 Then
                If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
                SendTrainStat Stats.Agility
                AddAgilityBtnState = 0
            Else
                AddAgilityBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'AddWillpower Btn
    If X >= AddWillpowerBtnBounds.Left And X <= AddWillpowerBtnBounds.Left + AddWillpowerBtnBounds.Right Then
        If Y >= AddWillpowerBtnBounds.Top And Y <= AddWillpowerBtnBounds.Top + AddWillpowerBtnBounds.Bottom Then
            If AddWillpowerBtnState = 2 Then
                If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
                SendTrainStat Stats.Willpower
                AddWillpowerBtnState = 0
            Else
                AddWillpowerBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'CloseCharacter Btn
    If X >= CloseCharacterBtnBounds.Left And X <= CloseCharacterBtnBounds.Left + CloseCharacterBtnBounds.Right Then
        If Y >= CloseCharacterBtnBounds.Top And Y <= CloseCharacterBtnBounds.Top + CloseCharacterBtnBounds.Bottom Then
            If CloseCharacterBtnState = 2 Then
                If CurrentGameMenu = 3 Then
                    CurrentGameMenu = 0
                End If
                Exit Sub
                CloseCharacterBtnState = 0
            Else
                CloseCharacterBtnState = 0
            End If
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Character_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Character_MouseDown(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    'AddStrength Btn
    If X >= AddStrengthBtnBounds.Left And X <= AddStrengthBtnBounds.Left + AddStrengthBtnBounds.Right Then
        If Y >= AddStrengthBtnBounds.Top And Y <= AddStrengthBtnBounds.Top + AddStrengthBtnBounds.Bottom Then
            AddStrengthBtnState = 2
            Exit Sub
        End If
    End If
    'AddEndurance Btn
    If X >= AddEnduranceBtnBounds.Left And X <= AddEnduranceBtnBounds.Left + AddEnduranceBtnBounds.Right Then
        If Y >= AddEnduranceBtnBounds.Top And Y <= AddEnduranceBtnBounds.Top + AddEnduranceBtnBounds.Bottom Then
            AddEnduranceBtnState = 2
            Exit Sub
        End If
    End If
    'AddIntelligence Btn
    If X >= AddIntelligenceBtnBounds.Left And X <= AddIntelligenceBtnBounds.Left + AddIntelligenceBtnBounds.Right Then
        If Y >= AddIntelligenceBtnBounds.Top And Y <= AddIntelligenceBtnBounds.Top + AddIntelligenceBtnBounds.Bottom Then
            AddIntelligenceBtnState = 2
            Exit Sub
        End If
    End If
    'AddAgility Btn
    If X >= AddAgilityBtnBounds.Left And X <= AddAgilityBtnBounds.Left + AddAgilityBtnBounds.Right Then
        If Y >= AddAgilityBtnBounds.Top And Y <= AddAgilityBtnBounds.Top + AddAgilityBtnBounds.Bottom Then
            AddAgilityBtnState = 2
            Exit Sub
        End If
    End If
    'AddWillpower Btn
    If X >= AddWillpowerBtnBounds.Left And X <= AddWillpowerBtnBounds.Left + AddWillpowerBtnBounds.Right Then
        If Y >= AddWillpowerBtnBounds.Top And Y <= AddWillpowerBtnBounds.Top + AddWillpowerBtnBounds.Bottom Then
            AddWillpowerBtnState = 2
            Exit Sub
        End If
    End If
    If IsEqItem(X, Y) > 0 Then
        SendUnequip IsEqItem(X, Y)
    End If
    'CloseCharacter Btn
    If X >= CloseCharacterBtnBounds.Left And X <= CloseCharacterBtnBounds.Left + CloseCharacterBtnBounds.Right Then
        If Y >= CloseCharacterBtnBounds.Top And Y <= CloseCharacterBtnBounds.Top + CloseCharacterBtnBounds.Bottom Then
            CloseCharacterBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Character_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Character_MouseMove(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect

   On Error GoTo errorhandler

    If IsEqItem(X, Y) <> 0 Then
        UpdateDescWindow GetPlayerEquipment(MyIndex, IsEqItem(X, Y)), 0, 0
        LastItemDesc = GetPlayerEquipment(MyIndex, IsEqItem(X, Y)) ' set it so you don't re-set values
        Exit Sub
    End If

    ItemDescVisible = False
    LastItemDesc = 0 ' no item was last loaded




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Character_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function IsBankItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As rect
Dim i As Long



   On Error GoTo errorhandler

    IsBankItem = 0
    For i = 1 To MAX_BANK
        If GetBankItemNum(i) > 0 And GetBankItemNum(i) <= MAX_ITEMS Then
                With tempRec
                .Top = BankItemsBounds.Top + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankItemsBounds.Left + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
                    If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    If X >= BankItemsBounds.Left And X <= BankItemsBounds.Left + BankItemsBounds.Right Then
                        If Y >= BankItemsBounds.Top And Y <= BankItemsBounds.Top + BankItemsBounds.Bottom Then
                            IsBankItem = i
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsBankItem", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As rect
Dim i As Long


   On Error GoTo errorhandler

    IsShopItem = 0

    For i = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(i).Item > 0 And Shop(InShop).TradeItem(i).Item <= MAX_ITEMS Then
            With tempRec
                .Top = ShopItemsBounds.Top + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                .Bottom = .Top + PIC_Y
                .Left = ShopItemsBounds.Left + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    If X >= ShopItemsBounds.Left And X <= ShopItemsBounds.Left + ShopItemsBounds.Right Then
                        If Y >= ShopItemsBounds.Top And Y <= ShopItemsBounds.Top + ShopItemsBounds.Bottom Then
                            IsShopItem = i
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsShopItem", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Function IsInvItem(ByVal X As Long, ByVal Y As Long, Optional ByVal slot As Boolean = False) As Long
    Dim tempRec As rect
    Dim i As Long


   On Error GoTo errorhandler

    IsInvItem = 0

    For i = 1 To MAX_INV

        If (GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS) Or slot = True Then

            With tempRec
                .Top = ((InvOffsetY + 32) * ((i - 1) \ InvColumns)) + InvItemsBounds.Top
                .Bottom = .Top + PIC_Y
                .Left = ((InvOffsetX + 32) * (((i - 1) Mod InvColumns))) + InvItemsBounds.Left
                .Right = .Left + PIC_X
            End With
                    If X >= InventoryPnlBounds.Left And X <= InventoryPnlBounds.Left + InventoryPnlBounds.Right Then
                If Y >= InventoryPnlBounds.Top And Y <= InventoryPnlBounds.Top + InventoryPnlBounds.Bottom Then
                    If X >= tempRec.Left And X <= tempRec.Right Then
                        If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                            IsInvItem = i
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If

    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsInvItem", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function IsEqItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As rect
    Dim i As Long


   On Error GoTo errorhandler

    IsEqItem = 0
    If X >= PlayerWeaponSlotBounds.Left And X <= PlayerWeaponSlotBounds.Left + PlayerWeaponSlotBounds.Right Then
        If Y >= PlayerWeaponSlotBounds.Top And Y <= PlayerWeaponSlotBounds.Top + PlayerWeaponSlotBounds.Bottom Then
            If GetPlayerEquipment(MyIndex, Equipment.Weapon) > 0 And GetPlayerEquipment(MyIndex, Equipment.Weapon) <= MAX_ITEMS Then
                IsEqItem = Equipment.Weapon
            End If
        End If
    End If
    If X >= PlayerArmorSlotBounds.Left And X <= PlayerArmorSlotBounds.Left + PlayerArmorSlotBounds.Right Then
        If Y >= PlayerArmorSlotBounds.Top And Y <= PlayerArmorSlotBounds.Top + PlayerArmorSlotBounds.Bottom Then
            If GetPlayerEquipment(MyIndex, Equipment.Armor) > 0 And GetPlayerEquipment(MyIndex, Equipment.Armor) <= MAX_ITEMS Then
                IsEqItem = Equipment.Armor
            End If
        End If
    End If
    If X >= PlayerHelmetSlotBounds.Left And X <= PlayerHelmetSlotBounds.Left + PlayerHelmetSlotBounds.Right Then
        If Y >= PlayerHelmetSlotBounds.Top And Y <= PlayerHelmetSlotBounds.Top + PlayerHelmetSlotBounds.Bottom Then
            If GetPlayerEquipment(MyIndex, Equipment.Helmet) > 0 And GetPlayerEquipment(MyIndex, Equipment.Helmet) <= MAX_ITEMS Then
                IsEqItem = Equipment.Helmet
            End If
        End If
    End If
    If X >= PlayerShieldSlotBounds.Left And X <= PlayerShieldSlotBounds.Left + PlayerShieldSlotBounds.Right Then
        If Y >= PlayerShieldSlotBounds.Top And Y <= PlayerShieldSlotBounds.Top + PlayerShieldSlotBounds.Bottom Then
            If GetPlayerEquipment(MyIndex, Equipment.Shield) > 0 And GetPlayerEquipment(MyIndex, Equipment.Shield) <= MAX_ITEMS Then
                IsEqItem = Equipment.Shield
            End If
        End If
    End If



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsEqItem", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub Options_MouseDown(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    'MusicOn Btn
    If X >= MusicOnBtnBounds.Left And X <= MusicOnBtnBounds.Left + MusicOnBtnBounds.Right Then
        If Y >= MusicOnBtnBounds.Top And Y <= MusicOnBtnBounds.Top + MusicOnBtnBounds.Bottom Then
            MusicOnBtnState = 2
            Exit Sub
        End If
    End If
    'MusicOff Btn
    If X >= MusicOffBtnBounds.Left And X <= MusicOffBtnBounds.Left + MusicOffBtnBounds.Right Then
        If Y >= MusicOffBtnBounds.Top And Y <= MusicOffBtnBounds.Top + MusicOffBtnBounds.Bottom Then
            MusicOffBtnState = 2
            Exit Sub
        End If
    End If
    'SoundOn Btn
    If X >= SoundOnBtnBounds.Left And X <= SoundOnBtnBounds.Left + SoundOnBtnBounds.Right Then
        If Y >= SoundOnBtnBounds.Top And Y <= SoundOnBtnBounds.Top + SoundOnBtnBounds.Bottom Then
            SoundOnBtnState = 2
            Exit Sub
        End If
    End If
    'SoundOff Btn
    If X >= SoundOffBtnBounds.Left And X <= SoundOffBtnBounds.Left + SoundOffBtnBounds.Right Then
        If Y >= SoundOffBtnBounds.Top And Y <= SoundOffBtnBounds.Top + SoundOffBtnBounds.Bottom Then
            SoundOffBtnState = 2
            Exit Sub
        End If
    End If
    'FullscreenOn Btn
    If X >= FullScreenOnBtnBounds.Left And X <= FullScreenOnBtnBounds.Left + FullScreenOnBtnBounds.Right Then
        If Y >= FullScreenOnBtnBounds.Top And Y <= FullScreenOnBtnBounds.Top + FullScreenOnBtnBounds.Bottom Then
            FullScreenOnBtnState = 2
            Exit Sub
        End If
    End If
    'FullscreenOff Btn
    If X >= FullScreenOffBtnBounds.Left And X <= FullScreenOffBtnBounds.Left + FullScreenOffBtnBounds.Right Then
        If Y >= FullScreenOffBtnBounds.Top And Y <= FullScreenOffBtnBounds.Top + FullScreenOffBtnBounds.Bottom Then
            FullScreenOffBtnState = 2
            Exit Sub
        End If
    End If
    'CloseOptions Btn
    If X >= CloseOptionsBtnBounds.Left And X <= CloseOptionsBtnBounds.Left + CloseOptionsBtnBounds.Right Then
        If Y >= CloseOptionsBtnBounds.Top And Y <= CloseOptionsBtnBounds.Top + CloseOptionsBtnBounds.Bottom Then
            CloseOptionsBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Options_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Options_MouseUp(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, MusicFile As String
    'MusicOn Btn

   On Error GoTo errorhandler

    If X >= MusicOnBtnBounds.Left And X <= MusicOnBtnBounds.Left + MusicOnBtnBounds.Right Then
        If Y >= MusicOnBtnBounds.Top And Y <= MusicOnBtnBounds.Top + MusicOnBtnBounds.Bottom Then
            If MusicOnBtnState = 2 Then
                Options.Music = 1
                ' start music playing
                MusicFile = Trim$(Map.Music)
                If Not MusicFile = "None." Then
                    PlayMusic MusicFile
                Else
                    StopMusic
                End If
                ' save to config.ini
                SaveOptions
                Exit Sub
                MusicOnBtnState = 0
            Else
                MusicOnBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'MusicOff Btn
    If X >= MusicOffBtnBounds.Left And X <= MusicOffBtnBounds.Left + MusicOffBtnBounds.Right Then
        If Y >= MusicOffBtnBounds.Top And Y <= MusicOffBtnBounds.Top + MusicOffBtnBounds.Bottom Then
            If MusicOffBtnState = 2 Then
                Options.Music = 0
                ' stop music playing
                StopMusic
                ' save to config.ini
                SaveOptions
                Exit Sub
                MusicOffBtnState = 0
            Else
                MusicOffBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'SoundOn Btn
    If X >= SoundOnBtnBounds.Left And X <= SoundOnBtnBounds.Left + SoundOnBtnBounds.Right Then
        If Y >= SoundOnBtnBounds.Top And Y <= SoundOnBtnBounds.Top + SoundOnBtnBounds.Bottom Then
            If SoundOnBtnState = 2 Then
                Options.sound = 1
                ' save to config.ini
                SaveOptions
                Exit Sub
                SoundOnBtnState = 0
            Else
                SoundOnBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'SoundOff Btn
    If X >= SoundOffBtnBounds.Left And X <= SoundOffBtnBounds.Left + SoundOffBtnBounds.Right Then
        If Y >= SoundOffBtnBounds.Top And Y <= SoundOffBtnBounds.Top + SoundOffBtnBounds.Bottom Then
            If SoundOffBtnState = 2 Then
                StopAllSounds
                Options.sound = 0
                ' save to config.ini
                SaveOptions
                Exit Sub
                SoundOffBtnState = 0
            Else
                SoundOffBtnState = 0
            End If
            Exit Sub
        End If
    End If
   
   
    'CloseOptions Btn
    If X >= CloseOptionsBtnBounds.Left And X <= CloseOptionsBtnBounds.Left + CloseOptionsBtnBounds.Right Then
        If Y >= CloseOptionsBtnBounds.Top And Y <= CloseOptionsBtnBounds.Top + CloseOptionsBtnBounds.Bottom Then
            If CloseOptionsBtnState = 2 Then
                If CurrentGameMenu = 4 Then
                    CurrentGameMenu = 0
                End If
                Exit Sub
                CloseOptionsBtnState = 0
            Else
                CloseOptionsBtnState = 0
            End If
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Options_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Party_MouseUp(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, MusicFile As String
    'PartyInvite Btn

   On Error GoTo errorhandler

    If X >= PartyInviteBtnBounds.Left And X <= PartyInviteBtnBounds.Left + PartyInviteBtnBounds.Right Then
        If Y >= PartyInviteBtnBounds.Top And Y <= PartyInviteBtnBounds.Top + PartyInviteBtnBounds.Bottom Then
            If PartyInviteBtnState = 2 Then
                If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                    SendPartyRequest
                Else
                    AddText "Invalid invitation target.", BrightRed
                End If
                Exit Sub
                PartyInviteBtnState = 0
            Else
                PartyInviteBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'PartyLeave Btn
    If X >= PartyLeaveBtnBounds.Left And X <= PartyLeaveBtnBounds.Left + PartyLeaveBtnBounds.Right Then
        If Y >= PartyLeaveBtnBounds.Top And Y <= PartyLeaveBtnBounds.Top + PartyLeaveBtnBounds.Bottom Then
            If PartyLeaveBtnState = 2 Then
                If Party.Leader > 0 Then
                    SendPartyLeave
                Else
                    AddText "You are not in a party.", BrightRed
                End If
                Exit Sub
                PartyLeaveBtnState = 0
            Else
                PartyLeaveBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'CloseParty Btn
    If X >= ClosePartyBtnBounds.Left And X <= ClosePartyBtnBounds.Left + ClosePartyBtnBounds.Right Then
        If Y >= ClosePartyBtnBounds.Top And Y <= ClosePartyBtnBounds.Top + ClosePartyBtnBounds.Bottom Then
            If ClosePartyBtnState = 2 Then
                If CurrentGameMenu = 5 Then
                    CurrentGameMenu = 0
                End If
                Exit Sub
                ClosePartyBtnState = 0
            Else
                ClosePartyBtnState = 0
            End If
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Party_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Party_MouseDown(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    'PartyInvite Btn
    If X >= PartyInviteBtnBounds.Left And X <= PartyInviteBtnBounds.Left + PartyInviteBtnBounds.Right Then
        If Y >= PartyInviteBtnBounds.Top And Y <= PartyInviteBtnBounds.Top + PartyInviteBtnBounds.Bottom Then
            PartyInviteBtnState = 2
            Exit Sub
        End If
    End If
    'PartyLeave Btn
    If X >= PartyLeaveBtnBounds.Left And X <= PartyLeaveBtnBounds.Left + PartyLeaveBtnBounds.Right Then
        If Y >= PartyLeaveBtnBounds.Top And Y <= PartyLeaveBtnBounds.Top + PartyLeaveBtnBounds.Bottom Then
            PartyLeaveBtnState = 2
            Exit Sub
        End If
    End If
    'CloseParty Btn
    If X >= ClosePartyBtnBounds.Left And X <= ClosePartyBtnBounds.Left + ClosePartyBtnBounds.Right Then
        If Y >= ClosePartyBtnBounds.Top And Y <= ClosePartyBtnBounds.Top + ClosePartyBtnBounds.Bottom Then
            ClosePartyBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Party_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Friends_MouseDown(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Button As Long
Dim dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long, i As Long, z As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    'Check List First
    'FriendsList
    dX = FriendsListBounds.Left
    dY = FriendsListBounds.Top
    dw = FriendsListBounds.Right
    dH = FriendsListBounds.Bottom
    If X >= dX And X <= dX + dw And Y >= dY And Y <= dY + dH Then
        If dw > 0 And dH > 0 And (dw + dX) > 0 And (dY + dH) > 0 Then
            If dH / 14 >= 1 Then
                FriendSelection = 0
                lineCount = Fix(dH / 14)
                maxScroll = Int(FriendCount / lineCount)
                If FriendListScroll > maxScroll Then FriendListScroll = 1
                For i = (lineCount * FriendListScroll) + 1 To (lineCount * FriendListScroll) + lineCount + 1
                    z = i - (lineCount * FriendListScroll) - 1
                    If i > FriendCount Then
                        Exit For
                    End If
                    If GlobalX >= dX And GlobalX <= dX + dw And GlobalY >= dY + 2 + (z * 14) And GlobalY <= dY + 2 + ((z + 1) * 14) Then
                        If FriendSelection = i Then FriendSelection = 0 Else FriendSelection = i
                        'Do mailbox stuff here!
                        If InMailbox Then
                            If MailBoxMenu = 3 Then
                                If SelTextbox = 1 Then
                                    MailToFrom = Trim$(FriendsList(FriendSelection))
                                End If
                            End If
                        End If
                    Else
                        If i = (lineCount * FriendListScroll) + lineCount + 1 Then
                            Exit For
                        End If
                    End If
                Next
                If FriendSelection > 0 Then Exit Sub
            End If
        End If
    End If
    'FriendsUp Btn
    If X >= FriendsUpBtnBounds.Left And X <= FriendsUpBtnBounds.Left + FriendsUpBtnBounds.Right Then
        If Y >= FriendsUpBtnBounds.Top And Y <= FriendsUpBtnBounds.Top + FriendsUpBtnBounds.Bottom Then
            FriendsUpBtnState = 2
            Exit Sub
        End If
    End If
    'FriendsDown Btn
    If X >= FriendsDownBtnBounds.Left And X <= FriendsDownBtnBounds.Left + FriendsDownBtnBounds.Right Then
        If Y >= FriendsDownBtnBounds.Top And Y <= FriendsDownBtnBounds.Top + FriendsDownBtnBounds.Bottom Then
            FriendsDownBtnState = 2
            Exit Sub
        End If
    End If
    'AddFriend Btn
    If X >= AddFriendBtnBounds.Left And X <= AddFriendBtnBounds.Left + AddFriendBtnBounds.Right Then
        If Y >= AddFriendBtnBounds.Top And Y <= AddFriendBtnBounds.Top + AddFriendBtnBounds.Bottom Then
            AddFriendBtnState = 2
            Exit Sub
        End If
    End If
    'DelFriend Btn
    If X >= DelFriendBtnBounds.Left And X <= DelFriendBtnBounds.Left + DelFriendBtnBounds.Right Then
        If Y >= DelFriendBtnBounds.Top And Y <= DelFriendBtnBounds.Top + DelFriendBtnBounds.Bottom Then
            DelFriendBtnState = 2
            Exit Sub
        End If
    End If
    'CloseFriends Btn
    If X >= CloseFriendsBtnBounds.Left And X <= CloseFriendsBtnBounds.Left + CloseFriendsBtnBounds.Right Then
        If Y >= CloseFriendsBtnBounds.Top And Y <= CloseFriendsBtnBounds.Top + CloseFriendsBtnBounds.Bottom Then
            CloseFriendsBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Friends_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Friends_MouseUp(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, MusicFile As String, lineCount As Long, maxScroll As Long
    'FriendsUp Btn

   On Error GoTo errorhandler

    If X >= FriendsUpBtnBounds.Left And X <= FriendsUpBtnBounds.Left + FriendsUpBtnBounds.Right Then
        If Y >= FriendsUpBtnBounds.Top And Y <= FriendsUpBtnBounds.Top + FriendsUpBtnBounds.Bottom Then
            If FriendsUpBtnState = 2 Then
                If FriendListScroll > 0 Then
                    FriendListScroll = FriendListScroll - 1
                    FriendSelection = 0
                End If
                Exit Sub
                FriendsUpBtnState = 0
            Else
                FriendsUpBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'FriendsDown Btn
    If X >= FriendsDownBtnBounds.Left And X <= FriendsDownBtnBounds.Left + FriendsDownBtnBounds.Right Then
        If Y >= FriendsDownBtnBounds.Top And Y <= FriendsDownBtnBounds.Top + FriendsDownBtnBounds.Bottom Then
            If FriendsDownBtnState = 2 Then
                lineCount = FriendsListBounds.Bottom / 14
                maxScroll = FriendCount / lineCount
                If FriendListScroll < maxScroll Then
                    FriendListScroll = FriendListScroll + 1
                    FriendSelection = 0
                End If
                Exit Sub
                FriendsDownBtnState = 0
            Else
                FriendsDownBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'AddFriend Btn
    If X >= AddFriendBtnBounds.Left And X <= AddFriendBtnBounds.Left + AddFriendBtnBounds.Right Then
        If Y >= AddFriendBtnBounds.Top And Y <= AddFriendBtnBounds.Top + AddFriendBtnBounds.Bottom Then
            If AddFriendBtnState = 2 Then
                If FriendCount < 25 Then
                    CurrencyMenu = 6 ' add friend
                    CurrencyCaption = "Who do you want to add as a friend?"
                    CurrencyItem = 0
                    CurrencyText = ""
                Else
                    AddText "Your friends list is full!", BrightRed
                End If
                Exit Sub
                AddFriendBtnState = 0
            Else
                AddFriendBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'DelFriend Btn
    If X >= DelFriendBtnBounds.Left And X <= DelFriendBtnBounds.Left + DelFriendBtnBounds.Right Then
        If Y >= DelFriendBtnBounds.Top And Y <= DelFriendBtnBounds.Top + DelFriendBtnBounds.Bottom Then
            If DelFriendBtnState = 2 Then
                If FriendSelection > 0 Then
                    If FriendIndex(FriendSelection) > 0 Then
                        dialogue "Remove Friend?", "Do you want to remove " & Trim$(FriendsList(FriendSelection)) & " from your friends list?", DIALOGUE_TYPE_REMOVEFRIEND, True, , Trim$(FriendsList(FriendSelection))
                    End If
                Else
                    AddText "No friend selected to remove!", BrightRed
                End If
                Exit Sub
                DelFriendBtnState = 0
            Else
                DelFriendBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'CloseFriends Btn
    If X >= CloseFriendsBtnBounds.Left And X <= CloseFriendsBtnBounds.Left + CloseFriendsBtnBounds.Right Then
        If Y >= CloseFriendsBtnBounds.Top And Y <= CloseFriendsBtnBounds.Top + CloseFriendsBtnBounds.Bottom Then
            If CloseFriendsBtnState = 2 Then
                If CurrentGameMenu = 6 Then
                    CurrentGameMenu = 0
                End If
                Exit Sub
                CloseFriendsBtnState = 0
            Else
                CloseFriendsBtnState = 0
            End If
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Friends_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Pets_MouseDown(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Button As Long
Dim dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long, i As Long, z As Long, buffer As clsBuffer

   On Error GoTo errorhandler

    Button = MouseBtn
    If Player(MyIndex).Pet.Points > 0 Then
        'PetAddStrBtn
        If X >= PetAddStrBtnBounds.Left And X <= PetAddStrBtnBounds.Left + PetAddStrBtnBounds.Right Then
            If Y >= PetAddStrBtnBounds.Top And Y <= PetAddStrBtnBounds.Top + PetAddStrBtnBounds.Bottom Then
                PetAddStrBtnState = 2
                Exit Sub
            End If
        End If
        'PetAddEndBtn
        If X >= PetAddEndBtnBounds.Left And X <= PetAddEndBtnBounds.Left + PetAddEndBtnBounds.Right Then
            If Y >= PetAddEndBtnBounds.Top And Y <= PetAddEndBtnBounds.Top + PetAddEndBtnBounds.Bottom Then
                PetAddEndBtnState = 2
                Exit Sub
            End If
        End If
        'PetAddIntBtn
        If X >= PetAddIntBtnBounds.Left And X <= PetAddIntBtnBounds.Left + PetAddIntBtnBounds.Right Then
            If Y >= PetAddIntBtnBounds.Top And Y <= PetAddIntBtnBounds.Top + PetAddIntBtnBounds.Bottom Then
                PetAddIntBtnState = 2
                Exit Sub
            End If
        End If
        'PetAddAgiBtn
        If X >= PetAddAgiBtnBounds.Left And X <= PetAddAgiBtnBounds.Left + PetAddAgiBtnBounds.Right Then
            If Y >= PetAddAgiBtnBounds.Top And Y <= PetAddAgiBtnBounds.Top + PetAddAgiBtnBounds.Bottom Then
                PetAddAgiBtnState = 2
                Exit Sub
            End If
        End If
        'PetAddWillBtn
        If X >= PetAddWillBtnBounds.Left And X <= PetAddWillBtnBounds.Left + PetAddWillBtnBounds.Right Then
            If Y >= PetAddWillBtnBounds.Top And Y <= PetAddWillBtnBounds.Top + PetAddWillBtnBounds.Bottom Then
                PetAddWillBtnState = 2
                Exit Sub
            End If
        End If
    End If
    
    If Pet(Player(MyIndex).Pet.Num).spell(1) > 0 And Pet(Player(MyIndex).Pet.Num).spell(1) <= MAX_SPELLS Then
        'PetSpell1Panel
        If X >= PetSpell1PanelBounds.Left And X <= PetSpell1PanelBounds.Left + PetSpell1PanelBounds.Right Then
            If Y >= PetSpell1PanelBounds.Top And Y <= PetSpell1PanelBounds.Top + PetSpell1PanelBounds.Bottom Then
                Set buffer = New clsBuffer
                buffer.WriteLong CPetSpell
                buffer.WriteLong 1
                SendData buffer.ToArray
                Set buffer = Nothing
                PetSpellBuffer = 1
                PetSpellBufferTimer = GetTickCount
                Exit Sub
            End If
        End If
        'PetSpell1Label
        If X >= PetSpell1LabelBounds.Left And X <= PetSpell1LabelBounds.Left + PetSpell1LabelBounds.Right Then
            If Y >= PetSpell1LabelBounds.Top And Y <= PetSpell1LabelBounds.Top + PetSpell1LabelBounds.Bottom Then
                Set buffer = New clsBuffer
                buffer.WriteLong CPetSpell
                buffer.WriteLong 1
                SendData buffer.ToArray
                Set buffer = Nothing
                PetSpellBuffer = 1
                PetSpellBufferTimer = GetTickCount
                Exit Sub
            End If
        End If
    End If
    
    If Pet(Player(MyIndex).Pet.Num).spell(2) > 0 And Pet(Player(MyIndex).Pet.Num).spell(2) <= MAX_SPELLS Then
        'PetSpell2Panel
        If X >= PetSpell2PanelBounds.Left And X <= PetSpell2PanelBounds.Left + PetSpell2PanelBounds.Right Then
            If Y >= PetSpell2PanelBounds.Top And Y <= PetSpell2PanelBounds.Top + PetSpell2PanelBounds.Bottom Then
                Set buffer = New clsBuffer
                buffer.WriteLong CPetSpell
                buffer.WriteLong 2
                SendData buffer.ToArray
                Set buffer = Nothing
                PetSpellBuffer = 2
                PetSpellBufferTimer = GetTickCount
                Exit Sub
            End If
        End If
        'PetSpell2Label
        If X >= PetSpell2LabelBounds.Left And X <= PetSpell2LabelBounds.Left + PetSpell2LabelBounds.Right Then
            If Y >= PetSpell2LabelBounds.Top And Y <= PetSpell2LabelBounds.Top + PetSpell2LabelBounds.Bottom Then
                Set buffer = New clsBuffer
                buffer.WriteLong CPetSpell
                buffer.WriteLong 2
                SendData buffer.ToArray
                Set buffer = Nothing
                PetSpellBuffer = 2
                PetSpellBufferTimer = GetTickCount
                Exit Sub
            End If
        End If
    End If
    
    If Pet(Player(MyIndex).Pet.Num).spell(3) > 0 And Pet(Player(MyIndex).Pet.Num).spell(3) <= MAX_SPELLS Then
        'PetSpell3Panel
        If X >= PetSpell3PanelBounds.Left And X <= PetSpell3PanelBounds.Left + PetSpell3PanelBounds.Right Then
            If Y >= PetSpell3PanelBounds.Top And Y <= PetSpell3PanelBounds.Top + PetSpell3PanelBounds.Bottom Then
                Set buffer = New clsBuffer
                buffer.WriteLong CPetSpell
                buffer.WriteLong 3
                SendData buffer.ToArray
                Set buffer = Nothing
                PetSpellBuffer = 3
                PetSpellBufferTimer = GetTickCount
                Exit Sub
            End If
        End If
        'PetSpell3Label
        If X >= PetSpell3LabelBounds.Left And X <= PetSpell3LabelBounds.Left + PetSpell3LabelBounds.Right Then
            If Y >= PetSpell3LabelBounds.Top And Y <= PetSpell3LabelBounds.Top + PetSpell3LabelBounds.Bottom Then
                Set buffer = New clsBuffer
                buffer.WriteLong CPetSpell
                buffer.WriteLong 3
                SendData buffer.ToArray
                Set buffer = Nothing
                PetSpellBuffer = 3
                PetSpellBufferTimer = GetTickCount
                Exit Sub
            End If
        End If
    End If
    
    If Pet(Player(MyIndex).Pet.Num).spell(4) > 0 And Pet(Player(MyIndex).Pet.Num).spell(4) <= MAX_SPELLS Then
        'PetSpell4Panel
        If X >= PetSpell4PanelBounds.Left And X <= PetSpell4PanelBounds.Left + PetSpell4PanelBounds.Right Then
            If Y >= PetSpell4PanelBounds.Top And Y <= PetSpell4PanelBounds.Top + PetSpell4PanelBounds.Bottom Then
                Set buffer = New clsBuffer
                buffer.WriteLong CPetSpell
                buffer.WriteLong 4
                SendData buffer.ToArray
                Set buffer = Nothing
                PetSpellBuffer = 4
                PetSpellBufferTimer = GetTickCount
                Exit Sub
            End If
        End If
        'PetSpell4Label
        If X >= PetSpell4LabelBounds.Left And X <= PetSpell4LabelBounds.Left + PetSpell4LabelBounds.Right Then
            If Y >= PetSpell4LabelBounds.Top And Y <= PetSpell4LabelBounds.Top + PetSpell4LabelBounds.Bottom Then
                Set buffer = New clsBuffer
                buffer.WriteLong CPetSpell
                buffer.WriteLong 4
                SendData buffer.ToArray
                Set buffer = Nothing
                PetSpellBuffer = 4
                PetSpellBufferTimer = GetTickCount
                Exit Sub
            End If
        End If
    End If

    'PetReleaseLabel
    If X >= PetReleaseLabelBounds.Left And X <= PetReleaseLabelBounds.Left + PetReleaseLabelBounds.Right Then
        If Y >= PetReleaseLabelBounds.Top And Y <= PetReleaseLabelBounds.Top + PetReleaseLabelBounds.Bottom Then
            Set buffer = New clsBuffer
            buffer.WriteLong CReleasePet
            SendData buffer.ToArray
            Set buffer = Nothing
            Exit Sub
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Pets_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Pets_MouseUp(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, MusicFile As String, lineCount As Long, maxScroll As Long
    'QuestUp Btn

   On Error GoTo errorhandler
    If Player(MyIndex).Pet.Points > 0 Then
        'PetAddStrBtn
        If X >= PetAddStrBtnBounds.Left And X <= PetAddStrBtnBounds.Left + PetAddStrBtnBounds.Right Then
            If Y >= PetAddStrBtnBounds.Top And Y <= PetAddStrBtnBounds.Top + PetAddStrBtnBounds.Bottom Then
                If PetAddStrBtnState = 2 Then
                    SendTrainPetStat Stats.Strength
                    Exit Sub
                    PetAddStrBtnState = 0
                Else
                    PetAddStrBtnState = 0
                End If
                Exit Sub
            End If
        End If
        'PetAddEndBtn
        If X >= PetAddEndBtnBounds.Left And X <= PetAddEndBtnBounds.Left + PetAddEndBtnBounds.Right Then
            If Y >= PetAddEndBtnBounds.Top And Y <= PetAddEndBtnBounds.Top + PetAddEndBtnBounds.Bottom Then
                If PetAddEndBtnState = 2 Then
                    SendTrainPetStat Stats.Endurance
                    Exit Sub
                    PetAddEndBtnState = 0
                Else
                    PetAddEndBtnState = 0
                End If
                Exit Sub
            End If
        End If
        'PetAddIntBtn
        If X >= PetAddIntBtnBounds.Left And X <= PetAddIntBtnBounds.Left + PetAddIntBtnBounds.Right Then
            If Y >= PetAddIntBtnBounds.Top And Y <= PetAddIntBtnBounds.Top + PetAddIntBtnBounds.Bottom Then
                If PetAddIntBtnState = 2 Then
                    SendTrainPetStat Stats.Intelligence
                    Exit Sub
                    PetAddIntBtnState = 0
                Else
                    PetAddIntBtnState = 0
                End If
                Exit Sub
            End If
        End If
        'PetAddAgiBtn
        If X >= PetAddAgiBtnBounds.Left And X <= PetAddAgiBtnBounds.Left + PetAddAgiBtnBounds.Right Then
            If Y >= PetAddAgiBtnBounds.Top And Y <= PetAddAgiBtnBounds.Top + PetAddAgiBtnBounds.Bottom Then
                If PetAddAgiBtnState = 2 Then
                    SendTrainPetStat Stats.Agility
                    Exit Sub
                    PetAddAgiBtnState = 0
                Else
                    PetAddAgiBtnState = 0
                End If
                Exit Sub
            End If
        End If
        'PetAddWillBtn
        If X >= PetAddWillBtnBounds.Left And X <= PetAddWillBtnBounds.Left + PetAddWillBtnBounds.Right Then
            If Y >= PetAddWillBtnBounds.Top And Y <= PetAddWillBtnBounds.Top + PetAddWillBtnBounds.Bottom Then
                If PetAddWillBtnState = 2 Then
                    SendTrainPetStat Stats.Willpower
                    Exit Sub
                    PetAddWillBtnState = 0
                Else
                    PetAddWillBtnState = 0
                End If
                Exit Sub
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Pets_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Quests_MouseDown(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Button As Long
Dim dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long, i As Long, z As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    'Check List First
    'QuestsList
    dX = QuestListBounds.Left
    dY = QuestListBounds.Top
    dw = QuestListBounds.Right
    dH = QuestListBounds.Bottom
    If X >= dX And X <= dX + dw And Y >= dY And Y <= dY + dH Then
        If dw > 0 And dH > 0 And (dw + dX) > 0 And (dY + dH) > 0 Then
            If dH / 14 >= 1 Then
                QuestSelection = 0
                lineCount = Fix(dH / 14)
                maxScroll = Int(QuestCount / lineCount)
                If QuestListScroll > maxScroll Then QuestListScroll = 1
                For i = (lineCount * QuestListScroll) + 1 To (lineCount * QuestListScroll) + lineCount + 1
                    z = i - (lineCount * QuestListScroll) - 1
                    If i > QuestCount Then
                        Exit For
                    End If
                    If GlobalX >= dX And GlobalX <= dX + dw And GlobalY >= dY + 2 + (z * 14) And GlobalY <= dY + 2 + ((z + 1) * 14) Then
                        QuestSelection = i
                    Else
                        If i = (lineCount * QuestListScroll) + lineCount + 1 Then
                            Exit For
                        End If
                    End If
                Next
                If QuestSelection > 0 Then Exit Sub
            End If
        End If
    End If
    'QuestUp Btn
    If X >= QuestUpBtnBounds.Left And X <= QuestUpBtnBounds.Left + QuestUpBtnBounds.Right Then
        If Y >= QuestUpBtnBounds.Top And Y <= QuestUpBtnBounds.Top + QuestUpBtnBounds.Bottom Then
            QuestUpBtnState = 2
            Exit Sub
        End If
    End If
    'QuestDown Btn
    If X >= QuestDownBtnBounds.Left And X <= QuestDownBtnBounds.Left + QuestDownBtnBounds.Right Then
        If Y >= QuestDownBtnBounds.Top And Y <= QuestDownBtnBounds.Top + QuestDownBtnBounds.Bottom Then
            QuestDownBtnState = 2
            Exit Sub
        End If
    End If
    'QuestInfo Btn
    If X >= QuestInfoBtnBounds.Left And X <= QuestInfoBtnBounds.Left + QuestInfoBtnBounds.Right Then
        If Y >= QuestInfoBtnBounds.Top And Y <= QuestInfoBtnBounds.Top + QuestInfoBtnBounds.Bottom Then
            QuestInfoBtnState = 2
            Exit Sub
        End If
    End If
    'CloseQuests Btn
    If X >= CloseQuestsBtnBounds.Left And X <= CloseQuestsBtnBounds.Left + CloseQuestsBtnBounds.Right Then
        If Y >= CloseQuestsBtnBounds.Top And Y <= CloseQuestsBtnBounds.Top + CloseQuestsBtnBounds.Bottom Then
            CloseQuestsBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Quests_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Quests_MouseUp(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, MusicFile As String, lineCount As Long, maxScroll As Long
    'QuestUp Btn

   On Error GoTo errorhandler

    If X >= QuestUpBtnBounds.Left And X <= QuestUpBtnBounds.Left + QuestUpBtnBounds.Right Then
        If Y >= QuestUpBtnBounds.Top And Y <= QuestUpBtnBounds.Top + QuestUpBtnBounds.Bottom Then
            If QuestUpBtnState = 2 Then
                If QuestListScroll > 0 Then
                    QuestListScroll = QuestListScroll - 1
                    QuestSelection = 0
                End If
                Exit Sub
                QuestUpBtnState = 0
            Else
                QuestUpBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'QuestDown Btn
    If X >= QuestDownBtnBounds.Left And X <= QuestDownBtnBounds.Left + QuestDownBtnBounds.Right Then
        If Y >= QuestDownBtnBounds.Top And Y <= QuestDownBtnBounds.Top + QuestDownBtnBounds.Bottom Then
            If QuestDownBtnState = 2 Then
                lineCount = QuestListBounds.Bottom / 14
                maxScroll = QuestCount / lineCount
                If QuestListScroll < maxScroll Then
                    QuestListScroll = QuestListScroll + 1
                    QuestSelection = 0
                End If
                Exit Sub
                QuestDownBtnState = 0
            Else
                QuestDownBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'QuestInfo Btn
    If X >= QuestInfoBtnBounds.Left And X <= QuestInfoBtnBounds.Left + QuestInfoBtnBounds.Right Then
        If Y >= QuestInfoBtnBounds.Top And Y <= QuestInfoBtnBounds.Top + QuestInfoBtnBounds.Bottom Then
            If QuestInfoBtnState = 2 Then
                If QuestSelection > 0 Then
                    InQuestLog = True
                    QuestLogPage = 0
                    QuestLogFunction = 1
                    QuestLogQuest = QuestIndex(QuestSelection)
                Else
                    AddText "No quest selected!", BrightRed
                End If
                Exit Sub
                QuestInfoBtnState = 0
            Else
                QuestInfoBtnState = 0
            End If
            Exit Sub
        End If
    End If
    'CloseQuests Btn
    If X >= CloseQuestsBtnBounds.Left And X <= CloseQuestsBtnBounds.Left + CloseQuestsBtnBounds.Right Then
        If Y >= CloseQuestsBtnBounds.Top And Y <= CloseQuestsBtnBounds.Top + CloseQuestsBtnBounds.Bottom Then
            If CloseQuestsBtnState = 2 Then
                If CurrentGameMenu = 7 Then
                    CurrentGameMenu = 0
                End If
                Exit Sub
                CloseQuestsBtnState = 0
            Else
                CloseQuestsBtnState = 0
            End If
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Quests_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Bank_MouseMove(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, BankNum As Long

   On Error GoTo errorhandler

    If DragBankSlotNum > 0 Then

    Else
        BankNum = IsBankItem(X, Y)
            If BankNum <> 0 Then
            UpdateDescWindow Bank.Item(BankNum).Num, X, Y
            LastBankDesc = Bank.Item(BankNum).Num
            Exit Sub
        Else
            LastBankDesc = 0
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Bank_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Bank_MouseDown(X As Long, Y As Long)
    Dim BankNum As Long, Button As Long


   On Error GoTo errorhandler

    BankNum = IsBankItem(X, Y)
    Button = MouseBtn

    If Button = 1 Then
        'Bank Button
        If X >= CloseBankBtnBounds.Left And X <= CloseBankBtnBounds.Left + CloseBankBtnBounds.Right Then
            If Y >= CloseBankBtnBounds.Top And Y <= CloseBankBtnBounds.Top + CloseBankBtnBounds.Bottom Then
                CloseBankBtnState = 2
                Exit Sub
            End If
        End If
            BankNum = IsBankItem(X, Y)
        If BankNum <> 0 Then
            DragBankSlotNum = BankNum
        End If

    ElseIf Button = 2 Then

    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Bank_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Bank_MouseUp(X As Long, Y As Long)
    Dim i As Long
    Dim rec_pos As rect, buffer As clsBuffer
    'CloseBankBtn

   On Error GoTo errorhandler

    If X >= CloseBankBtnBounds.Left And X <= CloseBankBtnBounds.Left + CloseBankBtnBounds.Right Then
        If Y >= CloseBankBtnBounds.Top And Y <= CloseBankBtnBounds.Top + CloseBankBtnBounds.Bottom Then
            If CloseBankBtnState = 2 Then
                InBank = False
                Exit Sub
                CloseBankBtnState = 0
            Else
                CloseBankBtnState = 0
            End If
        End If
    End If
    ' TODO : Add sub to change bankslots client side first so there's no delay in switching
    If DragBankSlotNum > 0 Then
        For i = 1 To MAX_BANK
            With rec_pos
                .Top = BankItemsBounds.Top + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankItemsBounds.Left + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If X >= BankItemsBounds.Left And X <= BankItemsBounds.Left + BankItemsBounds.Right Then
                        If Y >= BankItemsBounds.Top And Y <= BankItemsBounds.Top + BankItemsBounds.Bottom Then
                            If DragBankSlotNum <> i Then
                                ChangeBankSlots DragBankSlotNum, i
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    If DragInvSlotNum > 0 Then
        If X >= BankItemsBounds.Left And X <= BankItemsBounds.Left + BankItemsBounds.Right Then
            If Y >= BankItemsBounds.Top And Y <= BankItemsBounds.Top + BankItemsBounds.Bottom Then
                'Lets Deposit the dragged item
                If Item(GetPlayerInvItemNum(MyIndex, DragInvSlotNum)).Stackable = 1 Then
                    CurrencyMenu = 2 ' deposit
                    CurrencyCaption = "How many do you want to deposit?"
                    CurrencyItem = DragInvSlotNum
                    CurrencyText = ""
                    Exit Sub
                End If
                                Call DepositItem(DragInvSlotNum, 0)
            End If
        End If
    End If

    DragBankSlotNum = 0
    DragInvSlotNum = 0





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Bank_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Bank_DblClick(X As Long, Y As Long)
    Dim BankNum As Long
    Dim Value As Long
    Dim multiplier As Double
    Dim i As Long


   On Error GoTo errorhandler

    DragBankSlotNum = 0

    BankNum = IsBankItem(GlobalX, GlobalY)
    If BankNum <> 0 Then
         If Item(GetBankItemNum(BankNum)).type = ITEM_TYPE_NONE Then Exit Sub
                  If Item(GetBankItemNum(BankNum)).Stackable = 1 Then
                CurrencyMenu = 3 ' withdraw
                CurrencyCaption = "How many do you want to withdraw?"
                CurrencyItem = BankNum
                CurrencyText = ""
                Exit Sub
            End If
                 WithdrawItem BankNum, 0
         Exit Sub
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Bank_DblClick", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub Shop_MouseMove(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Shopslot As Long

   On Error GoTo errorhandler

    Shopslot = IsShopItem(X, Y)
    If DragInvSlotNum > 0 Then
        Exit Sub
    ElseIf Shopslot <> 0 Then
        UpdateDescWindow Shop(InShop).TradeItem(Shopslot).Item, 0, 0
        LastItemDesc = Shop(InShop).TradeItem(Shopslot).Item
        Exit Sub
    End If
    ItemDescVisible = False
    LastItemDesc = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Shop_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Shop_MouseDown(X As Long, Y As Long)
    Dim Shopitem As Long, Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn

    Shopitem = IsShopItem(X, Y)
    If Button = 1 Then
        'Shop Button
        If X >= CloseShopBtnBounds.Left And X <= CloseShopBtnBounds.Left + CloseShopBtnBounds.Right Then
            If Y >= CloseShopBtnBounds.Top And Y <= CloseShopBtnBounds.Top + CloseShopBtnBounds.Bottom Then
                CloseShopBtnState = 2
                Exit Sub
            End If
        End If
        If Shopitem > 0 Then
            Select Case ShopAction
                Case 0 ' no action, give cost
                    With Shop(InShop).TradeItem(Shopitem)
                        AddText "You can buy this item for " & .CostValue & " " & Trim$(Item(.CostItem).Name) & ".", White
                            If Button = 1 Then
                                If Shopitem <> 0 Then
                                    DragInvSlotNum = Shopitem
                                    IsReallyShop = True
                                    ItemDescVisible = False
                                    LastItemDesc = 0
                                End If
                            End If
                    End With
                Case 1 ' buy item
                    ' buy item code
                    BuyItem Shopitem
            End Select
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Shop_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Sub Shop_MouseUp(X As Long, Y As Long)
    Dim i As Long
    Dim rec_pos As rect, buffer As clsBuffer
    'CloseShopBtn

   On Error GoTo errorhandler

    If X >= CloseShopBtnBounds.Left And X <= CloseShopBtnBounds.Left + CloseShopBtnBounds.Right Then
        If Y >= CloseShopBtnBounds.Top And Y <= CloseShopBtnBounds.Top + CloseShopBtnBounds.Bottom Then
            If CloseShopBtnState = 2 Then
                CloseShop
                Exit Sub
                CloseShopBtnState = 0
            Else
                CloseShopBtnState = 0
            End If
        End If
    End If
    If DragInvSlotNum > 0 Then
        If IsReallyShop = False Then
            If X >= ShopItemsBounds.Left And X <= ShopItemsBounds.Left + ShopItemsBounds.Right Then
                If Y >= ShopItemsBounds.Top And Y <= ShopItemsBounds.Top + ShopItemsBounds.Bottom Then
                    'If Item was dragged from inventory into shop
                    SellItem DragInvSlotNum
                    DragInvSlotNum = 0
                    Exit Sub
                End If
            End If
        End If
    End If

    DragInvSlotNum = 0
    IsReallyShop = False
    DragInvSlotNum = 0





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Shop_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Shop_DblClick(X As Long, Y As Long)
    Dim ShopNum As Long
    Dim Value As Long
    Dim multiplier As Double
    Dim i As Long


   On Error GoTo errorhandler

    DragInvSlotNum = 0
    IsReallyShop = False
    ShopNum = IsShopItem(GlobalX, GlobalY)
    If ShopNum > 0 Then
        BuyItem ShopNum
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Shop_DblClick", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean) As Long
    Dim tempRec As rect
    Dim i As Long
    Dim ItemNum As Long


   On Error GoTo errorhandler

    IsTradeItem = 0

    For i = 1 To MAX_INV
        If Yours Then
            ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)
        Else
            ItemNum = TradeTheirOffer(i).Num
        End If

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                    If Yours = True Then
                With tempRec
                    .Top = YourTradePnlBounds.Top - 2 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = YourTradePnlBounds.Left + 4 + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With
            Else
                With tempRec
                    .Top = TheirTradePnlBounds.Top - 2 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = TheirTradePnlBounds.Left + 4 + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With
            End If

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    If Yours = True And X >= YourTradePnlBounds.Left And X <= YourTradePnlBounds.Left + YourTradePnlBounds.Right And Y >= YourTradePnlBounds.Top And Y <= YourTradePnlBounds.Top + YourTradePnlBounds.Bottom Then
                        IsTradeItem = i
                        Exit Function
                    Else
                        If Yours = False And X >= TheirTradePnlBounds.Left And X <= TheirTradePnlBounds.Left + TheirTradePnlBounds.Right And Y >= TheirTradePnlBounds.Top And Y <= TheirTradePnlBounds.Top + TheirTradePnlBounds.Bottom Then
                            IsTradeItem = i
                            Exit Function
                        Else
                                                End If
                    End If
                End If
            End If
        End If

    Next




   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsTradeItem", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Sub Trade_MouseMove(X As Long, Y As Long)
Dim Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    'Check both windows.....
    If X >= YourTradePnlBounds.Left And X <= YourTradePnlBounds.Left + YourTradePnlBounds.Right Then
        If Y >= YourTradePnlBounds.Top And Y <= YourTradePnlBounds.Top + YourTradePnlBounds.Bottom Then
            Call YourTrade_MouseMove(X, Y)
        End If
    End If
    If X >= TheirTradePnlBounds.Left And X <= TheirTradePnlBounds.Left + TheirTradePnlBounds.Right Then
        If Y >= TheirTradePnlBounds.Top And Y <= TheirTradePnlBounds.Top + TheirTradePnlBounds.Bottom Then
            Call TheirTrade_MouseMove(X, Y)
        End If
    End If





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Trade_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Trade_MouseDown(X As Long, Y As Long)
Dim Button As Long, TradeSlot

   On Error GoTo errorhandler

    Button = MouseBtn
    'Check both windows.....
    If X >= YourTradePnlBounds.Left And X <= YourTradePnlBounds.Left + YourTradePnlBounds.Right Then
        If Y >= YourTradePnlBounds.Top And Y <= YourTradePnlBounds.Top + YourTradePnlBounds.Bottom Then
            TradeSlot = IsTradeItem(X, Y, True)
            If TradeSlot > 0 And TradeSlot <= 35 Then
                If GetPlayerInvItemNum(MyIndex, TradeSlot) > 0 Then
                    DragTradeSlotNum = TradeSlot
                    ItemDescVisible = False
                    LastItemDesc = 0
                    Exit Sub
                End If
            End If
        End If
    End If
    'AcceptTrade Button
    If X >= AcceptTradeBtnBounds.Left And X <= AcceptTradeBtnBounds.Left + AcceptTradeBtnBounds.Right Then
        If Y >= AcceptTradeBtnBounds.Top And Y <= AcceptTradeBtnBounds.Top + AcceptTradeBtnBounds.Bottom Then
            AcceptTradeBtnState = 2
            Exit Sub
        End If
    End If
    'DeclineTrade Button
    If X >= DeclineTradeBtnBounds.Left And X <= DeclineTradeBtnBounds.Left + DeclineTradeBtnBounds.Right Then
        If Y >= DeclineTradeBtnBounds.Top And Y <= DeclineTradeBtnBounds.Top + DeclineTradeBtnBounds.Bottom Then
            DeclineTradeBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Trade_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Sub Trade_MouseUp(X As Long, Y As Long)
Dim Button As Long, TradeSlot

   On Error GoTo errorhandler

    Button = MouseBtn
    'Check both windows.....
    If X >= YourTradePnlBounds.Left And X <= YourTradePnlBounds.Left + YourTradePnlBounds.Right Then
        If Y >= YourTradePnlBounds.Top And Y <= YourTradePnlBounds.Top + YourTradePnlBounds.Bottom Then
            If DragInvSlotNum > 0 Then
                If IsReallyShop = False Then
                    If Item(GetPlayerInvItemNum(MyIndex, DragInvSlotNum)).Stackable = 1 Then
                        CurrencyMenu = 4 ' offer in trade
                        CurrencyCaption = "How many do you want to trade?"
                        CurrencyItem = DragInvSlotNum
                        CurrencyText = ""
                        Exit Sub
                    End If
                                    Call TradeItem(DragInvSlotNum, 0)
                End If
            End If
        End If
    End If
    'AcceptTradeBtn
    If X >= AcceptTradeBtnBounds.Left And X <= AcceptTradeBtnBounds.Left + AcceptTradeBtnBounds.Right Then
        If Y >= AcceptTradeBtnBounds.Top And Y <= AcceptTradeBtnBounds.Top + AcceptTradeBtnBounds.Bottom Then
            If AcceptTradeBtnState = 2 Then
                AcceptTrade
                Exit Sub
                AcceptTradeBtnState = 0
            Else
                AcceptTradeBtnState = 0
            End If
        End If
    End If
    'DeclineTradeBtn
    If X >= DeclineTradeBtnBounds.Left And X <= DeclineTradeBtnBounds.Left + DeclineTradeBtnBounds.Right Then
        If Y >= DeclineTradeBtnBounds.Top And Y <= DeclineTradeBtnBounds.Top + DeclineTradeBtnBounds.Bottom Then
            If DeclineTradeBtnState = 2 Then
                DeclineTrade
                Exit Sub
                DeclineTradeBtnState = 0
            Else
                DeclineTradeBtnState = 0
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Trade_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub YourTrade_MouseMove(X As Long, Y As Long)
Dim ItemNum As Long, Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    ItemNum = IsTradeItem(X, Y, True)
    If DragTradeSlotNum > 0 Or DragInvSlotNum > 0 Then
        Exit Sub
    ElseIf ItemNum <> 0 And ItemNum < 35 Then
        If TradeYourOffer(ItemNum).Num > 0 Then
            UpdateDescWindow GetPlayerInvItemNum(MyIndex, TradeYourOffer(ItemNum).Num), 0, 0
            LastItemDesc = GetPlayerInvItemNum(MyIndex, TradeYourOffer(ItemNum).Num)
            Exit Sub
        End If
    End If
    ItemDescVisible = False
    LastItemDesc = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "YourTrade_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub TheirTrade_MouseMove(X As Long, Y As Long)
Dim ItemNum As Long, Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    ItemNum = IsTradeItem(X, Y, False)
    If DragTradeSlotNum > 0 Or DragInvSlotNum > 0 Then
        Exit Sub
    ElseIf ItemNum <> 0 And ItemNum < 35 Then
        If TradeTheirOffer(ItemNum).Num > 0 Then
            UpdateDescWindow TradeTheirOffer(ItemNum).Num, 0, 0
            LastItemDesc = TradeTheirOffer(ItemNum).Num
            Exit Sub
        End If
    End If
    ItemDescVisible = False
    LastItemDesc = 0




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "TheirTrade_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Sub Trade_DblClick(X As Long, Y As Long)
    Dim TradeNum As Long
    Dim i As Long


   On Error GoTo errorhandler

    DragTradeSlotNum = 0
    TradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If TradeNum > 0 Then
        UntradeItem TradeNum
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Trade_DblClick", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub MailboxPnl_MouseDown(X As Long, Y As Long)
Dim Button As Long, TradeSlot

   On Error GoTo errorhandler

    Button = MouseBtn
    'CheckMail Button
    If X >= CheckMailBtnBounds.Left And X <= CheckMailBtnBounds.Left + CheckMailBtnBounds.Right Then
        If Y >= CheckMailBtnBounds.Top And Y <= CheckMailBtnBounds.Top + CheckMailBtnBounds.Bottom Then
            CheckMailBtnState = 2
            Exit Sub
        End If
    End If
    'SendMail Button
    If X >= SendMailBtnBounds.Left And X <= SendMailBtnBounds.Left + SendMailBtnBounds.Right Then
        If Y >= SendMailBtnBounds.Top And Y <= SendMailBtnBounds.Top + SendMailBtnBounds.Bottom Then
            SendMailBtnState = 2
            Exit Sub
        End If
    End If
    'CloseMailBox Button
    If X >= CloseMailboxBtnBounds.Left And X <= CloseMailboxBtnBounds.Left + CloseMailboxBtnBounds.Right Then
        If Y >= CloseMailboxBtnBounds.Top And Y <= CloseMailboxBtnBounds.Top + CloseMailboxBtnBounds.Bottom Then
            CloseMailboxBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MailboxPnl_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub MailboxPnl_MouseUp(X As Long, Y As Long)
Dim Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    'CheckMailBtn
    If X >= CheckMailBtnBounds.Left And X <= CheckMailBtnBounds.Left + CheckMailBtnBounds.Right Then
        If Y >= CheckMailBtnBounds.Top And Y <= CheckMailBtnBounds.Top + CheckMailBtnBounds.Bottom Then
            If CheckMailBtnState = 2 Then
                MailBoxMenu = 1
                MailToFrom = ""
                MailContent = ""
                MailItem = 0
                MailItemValue = 0
                Exit Sub
                CheckMailBtnState = 0
            Else
                CheckMailBtnState = 0
            End If
        End If
    End If
    'SendMailBtn
    If X >= SendMailBtnBounds.Left And X <= SendMailBtnBounds.Left + SendMailBtnBounds.Right Then
        If Y >= SendMailBtnBounds.Top And Y <= SendMailBtnBounds.Top + SendMailBtnBounds.Bottom Then
            If SendMailBtnState = 2 Then
                MailBoxMenu = 3
                SelTextbox = 1
                MailToFrom = ""
                MailContent = ""
                MailItem = 0
                MailItemValue = 0
                Exit Sub
                SendMailBtnState = 0
            Else
                SendMailBtnState = 0
            End If
        End If
    End If
    'CloseMailBoxBtn
    If X >= CloseMailboxBtnBounds.Left And X <= CloseMailboxBtnBounds.Left + CloseMailboxBtnBounds.Right Then
        If Y >= CloseMailboxBtnBounds.Top And Y <= CloseMailboxBtnBounds.Top + CloseMailboxBtnBounds.Bottom Then
            If CloseMailboxBtnState = 2 Then
                MailItem = 0
                MailItemValue = 0
                MailBoxMenu = 0
                InMailbox = False
                Exit Sub
                CloseMailboxBtnState = 0
            Else
                CloseMailboxBtnState = 0
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "MailboxPnl_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub InboxPnl_MouseDown(X As Long, Y As Long)
Dim Button As Long, dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long
Dim i As Long, z As Long, buffer As clsBuffer

   On Error GoTo errorhandler

    Button = MouseBtn
     'Check List First
    'QuestsList
    dX = InboxListBounds.Left
    dY = InboxListBounds.Top
    dw = InboxListBounds.Right
    dH = InboxListBounds.Bottom
    If X >= dX And X <= dX + dw And Y >= dY And Y <= dY + dH Then
        If dw > 0 And dH > 0 And (dw + dX) > 0 And (dY + dH) > 0 Then
            If dH / 14 >= 1 Then
                lineCount = Fix(dH / 14)
                maxScroll = Int(MailCount / lineCount)
                If InboxListScroll > maxScroll Then InboxListScroll = 1
                For i = (lineCount * InboxListScroll) + 1 To (lineCount * InboxListScroll) + lineCount + 1
                    z = i - (lineCount * InboxListScroll) - 1
                    If i > MailCount Then
                        Exit For
                    End If
                    If GlobalX >= dX And GlobalX <= dX + dw And GlobalY >= dY + 2 + (z * 14) And GlobalY <= dY + 2 + ((z + 1) * 14) Then
                        'Found a message
                        MailBoxMenu = 2
                        SelectedMail = i - 1
                        Set buffer = New clsBuffer
                        buffer.WriteLong CReadMail
                        buffer.WriteLong Mail(SelectedMail).Index
                        SendData buffer.ToArray
                        Set buffer = Nothing
                        Exit Sub
                    Else
                        If i = (lineCount * InboxListScroll) + lineCount + 1 Then
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    End If
    'ScrlInboxUpBtn Button
    If X >= ScrlInboxUpBtnBounds.Left And X <= ScrlInboxUpBtnBounds.Left + ScrlInboxUpBtnBounds.Right Then
        If Y >= ScrlInboxUpBtnBounds.Top And Y <= ScrlInboxUpBtnBounds.Top + ScrlInboxUpBtnBounds.Bottom Then
            ScrlInboxUpBtnState = 2
            Exit Sub
        End If
    End If
    'ScrlInboxDownBtn Button
    If X >= ScrlInboxDownBtnBounds.Left And X <= ScrlInboxDownBtnBounds.Left + ScrlInboxDownBtnBounds.Right Then
        If Y >= ScrlInboxDownBtnBounds.Top And Y <= ScrlInboxDownBtnBounds.Top + ScrlInboxDownBtnBounds.Bottom Then
            ScrlInboxDownBtnState = 2
            Exit Sub
        End If
    End If
    'CloseInbox Button
    If X >= CloseInboxBtnBounds.Left And X <= CloseInboxBtnBounds.Left + CloseInboxBtnBounds.Right Then
        If Y >= CloseInboxBtnBounds.Top And Y <= CloseInboxBtnBounds.Top + CloseInboxBtnBounds.Bottom Then
            CloseInboxBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "InboxPnl_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub InboxPnl_MouseUp(X As Long, Y As Long)
Dim Button As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    'ScrlInboxUpBtn
    If X >= ScrlInboxUpBtnBounds.Left And X <= ScrlInboxUpBtnBounds.Left + ScrlInboxUpBtnBounds.Right Then
        If Y >= ScrlInboxUpBtnBounds.Top And Y <= ScrlInboxUpBtnBounds.Top + ScrlInboxUpBtnBounds.Bottom Then
            If ScrlInboxUpBtnState = 2 Then
                If InboxListScroll > 0 Then
                    InboxListScroll = InboxListScroll - 1
                End If
                Exit Sub
                ScrlInboxUpBtnState = 0
            Else
                ScrlInboxUpBtnState = 0
            End If
        End If
    End If
    Dim lineCount As Long, maxScroll As Long
    'ScrlInboxDownBtn
    If X >= ScrlInboxDownBtnBounds.Left And X <= ScrlInboxDownBtnBounds.Left + ScrlInboxDownBtnBounds.Right Then
        If Y >= ScrlInboxDownBtnBounds.Top And Y <= ScrlInboxDownBtnBounds.Top + ScrlInboxDownBtnBounds.Bottom Then
            If ScrlInboxDownBtnState = 2 Then
                lineCount = InboxListBounds.Bottom / 14
                maxScroll = MailCount / lineCount
                If InboxListScroll < maxScroll Then
                    InboxListScroll = InboxListScroll + 1
                End If
                Exit Sub
                ScrlInboxDownBtnState = 0
            Else
                ScrlInboxDownBtnState = 0
            End If
        End If
    End If
    'CloseInboxBtn
    If X >= CloseInboxBtnBounds.Left And X <= CloseInboxBtnBounds.Left + CloseInboxBtnBounds.Right Then
        If Y >= CloseInboxBtnBounds.Top And Y <= CloseInboxBtnBounds.Top + CloseInboxBtnBounds.Bottom Then
            If CloseInboxBtnState = 2 Then
                MailBoxMenu = 0
                Exit Sub
                CloseInboxBtnState = 0
            Else
                CloseInboxBtnState = 0
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "InboxPnl_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ReadLetterPnl_MouseDown(X As Long, Y As Long)
Dim Button As Long, dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long
Dim i As Long, z As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    If X >= ItemReceivedBounds.Left And X <= ItemReceivedBounds.Left + ItemReceivedBounds.Right Then
        If Y >= ItemReceivedBounds.Top And Y <= ItemReceivedBounds.Top + ItemReceivedBounds.Bottom Then
            If Mail(SelectedMail).ItemNum > 0 Then
                DragMailboxItem = 1
            End If
        End If
    End If

    'ReplyLetterBtn Button
    If X >= ReplyLetterBtnBounds.Left And X <= ReplyLetterBtnBounds.Left + ReplyLetterBtnBounds.Right Then
        If Y >= ReplyLetterBtnBounds.Top And Y <= ReplyLetterBtnBounds.Top + ReplyLetterBtnBounds.Bottom Then
            ReplyLetterBtnState = 2
            Exit Sub
        End If
    End If
    'TrashLetterBtn Button
    If X >= TrashLetterBtnBounds.Left And X <= TrashLetterBtnBounds.Left + TrashLetterBtnBounds.Right Then
        If Y >= TrashLetterBtnBounds.Top And Y <= TrashLetterBtnBounds.Top + TrashLetterBtnBounds.Bottom Then
            TrashLetterBtnState = 2
            Exit Sub
        End If
    End If
    'StopReading Button
    If X >= StopReadingBtnBounds.Left And X <= StopReadingBtnBounds.Left + StopReadingBtnBounds.Right Then
        If Y >= StopReadingBtnBounds.Top And Y <= StopReadingBtnBounds.Top + StopReadingBtnBounds.Bottom Then
            StopReadingBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ReadLetterPnl_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ReadLetterPnl_DblClick(X As Long, Y As Long)
Dim Button As Long, dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long
Dim i As Long, z As Long, buffer As clsBuffer

   On Error GoTo errorhandler

    Button = MouseBtn
    If X >= ItemReceivedBounds.Left And X <= ItemReceivedBounds.Left + ItemReceivedBounds.Right Then
        If Y >= ItemReceivedBounds.Top And Y <= ItemReceivedBounds.Top + ItemReceivedBounds.Bottom Then
            'If there is an item, take it.
            If Mail(SelectedMail).ItemNum > 0 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CTakeMailItem
                buffer.WriteLong Mail(SelectedMail).Index
                SendData buffer.ToArray
                Set buffer = Nothing
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ReadLetterPnl_DblClick", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ReadLetterPnl_MouseMove(X As Long, Y As Long)
Dim Button As Long, dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long
Dim i As Long, z As Long, buffer As clsBuffer

   On Error GoTo errorhandler

    Button = MouseBtn
    If X >= ItemReceivedBounds.Left And X <= ItemReceivedBounds.Left + ItemReceivedBounds.Right Then
        If Y >= ItemReceivedBounds.Top And Y <= ItemReceivedBounds.Top + ItemReceivedBounds.Bottom Then
            'If there is an item, take it.
            If Mail(SelectedMail).ItemNum > 0 Then
                UpdateDescWindow Mail(SelectedMail).ItemNum, 0, 0
                LastItemDesc = Mail(SelectedMail).ItemNum
                ItemDescVisible = True
                Exit Sub
            End If
        End If
    End If
    ItemDescVisible = False




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ReadLetterPnl_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub ReadLetterPnl_MouseUp(X As Long, Y As Long)
Dim Button As Long, buffer As clsBuffer

   On Error GoTo errorhandler

    Button = MouseBtn
    'ReplyLetterBtn
    If X >= ReplyLetterBtnBounds.Left And X <= ReplyLetterBtnBounds.Left + ReplyLetterBtnBounds.Right Then
        If Y >= ReplyLetterBtnBounds.Top And Y <= ReplyLetterBtnBounds.Top + ReplyLetterBtnBounds.Bottom Then
            If ReplyLetterBtnState = 2 Then
                MailToFrom = Trim$(Mail(SelectedMail).From)
                MailContent = ""
                MailItem = 0
                MailItemValue = 0
                MailBoxMenu = 3
                SelTextbox = 2
                Exit Sub
                ReplyLetterBtnState = 0
            Else
                ReplyLetterBtnState = 0
            End If
        End If
    End If
    'TrashLetterBtn
    If X >= TrashLetterBtnBounds.Left And X <= TrashLetterBtnBounds.Left + TrashLetterBtnBounds.Right Then
        If Y >= TrashLetterBtnBounds.Top And Y <= TrashLetterBtnBounds.Top + TrashLetterBtnBounds.Bottom Then
            If TrashLetterBtnState = 2 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CDeleteMail
                buffer.WriteLong Mail(SelectedMail).Index
                SendData buffer.ToArray
                Set buffer = Nothing
                MailBoxMenu = 1
                SelectedMail = 0
                MailItem = 0
                MailItemValue = 0
                Exit Sub
                TrashLetterBtnState = 0
            Else
                TrashLetterBtnState = 0
            End If
        End If
    End If
    'StopReadingBtn
    If X >= StopReadingBtnBounds.Left And X <= StopReadingBtnBounds.Left + StopReadingBtnBounds.Right Then
        If Y >= StopReadingBtnBounds.Top And Y <= StopReadingBtnBounds.Top + StopReadingBtnBounds.Bottom Then
            If StopReadingBtnState = 2 Then
                SelectedMail = 0
                MailBoxMenu = 1
                Exit Sub
                StopReadingBtnState = 0
            Else
                StopReadingBtnState = 0
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ReadLetterPnl_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendMail_MouseDown(X As Long, Y As Long)
Dim Button As Long, dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long
Dim i As Long, z As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    If MailItem > 0 Then
        If X >= SendMessageItemBounds.Left And X <= SendMessageItemBounds.Left + SendMessageItemBounds.Right Then
            If Y >= SendMessageItemBounds.Top And Y <= SendMessageItemBounds.Top + SendMessageItemBounds.Bottom Then
                DragMailboxItem = 1
                Exit Sub
            End If
        End If
    End If
    If X >= SendMessageToBounds.Left And X <= SendMessageToBounds.Left + SendMessageToBounds.Right Then
        If Y >= SendMessageToBounds.Top And Y <= SendMessageToBounds.Top + SendMessageToBounds.Bottom Then
            SelTextbox = 1
            Exit Sub
        End If
    End If
    If X >= SendMessageTextBounds.Left And X <= SendMessageTextBounds.Left + SendMessageTextBounds.Right Then
        If Y >= SendMessageTextBounds.Top And Y <= SendMessageTextBounds.Top + SendMessageTextBounds.Bottom Then
            SelTextbox = 2
            Exit Sub
        End If
    End If

    'SendMessageBtn Button
    If X >= SendMessageBtnBounds.Left And X <= SendMessageBtnBounds.Left + SendMessageBtnBounds.Right Then
        If Y >= SendMessageBtnBounds.Top And Y <= SendMessageBtnBounds.Top + SendMessageBtnBounds.Bottom Then
            SendMessageBtnState = 2
            Exit Sub
        End If
    End If
    'DiscardMessageBtn Button
    If X >= DiscardMessageBtnBounds.Left And X <= DiscardMessageBtnBounds.Left + DiscardMessageBtnBounds.Right Then
        If Y >= DiscardMessageBtnBounds.Top And Y <= DiscardMessageBtnBounds.Top + DiscardMessageBtnBounds.Bottom Then
            DiscardMessageBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMail_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendMail_MouseUp(X As Long, Y As Long)
Dim Button As Long, buffer As clsBuffer

   On Error GoTo errorhandler

    Button = MouseBtn
    If DragInvSlotNum > 0 Then
        If X >= SendMessageItemBounds.Left And X <= SendMessageItemBounds.Left + SendMessageItemBounds.Right Then
            If Y >= SendMessageItemBounds.Top And Y <= SendMessageItemBounds.Top + SendMessageItemBounds.Bottom Then
                If Item(GetPlayerInvItemNum(MyIndex, DragInvSlotNum)).Stackable = 1 Then
                    CurrencyMenu = 5 ' offer in message
                    CurrencyCaption = "How many do you want to give?"
                    CurrencyItem = DragInvSlotNum
                    CurrencyText = ""
                    Exit Sub
                Else
                    MailItem = DragInvSlotNum
                    MailItemValue = 1
                    Exit Sub
                End If
            End If
        End If
    End If
    'SendMessageBtn
    If X >= SendMessageBtnBounds.Left And X <= SendMessageBtnBounds.Left + SendMessageBtnBounds.Right Then
        If Y >= SendMessageBtnBounds.Top And Y <= SendMessageBtnBounds.Top + SendMessageBtnBounds.Bottom Then
            If SendMessageBtnState = 2 Then
                If Len(Trim$(MailToFrom)) <= 0 Or Len(Trim$(MailContent)) <= 0 Then
                    AddText "No text found in message or recepiant left blank. Message could not be sent!", BrightRed
                Else
                    'Send Message if Possible
                    Set buffer = New clsBuffer
                    buffer.WriteLong CSendMail
                    buffer.WriteString Trim$(MailToFrom)
                    buffer.WriteString Trim$(MailContent)
                    buffer.WriteLong MailItem
                    buffer.WriteLong MailItemValue
                    SendData buffer.ToArray
                    Set buffer = Nothing
                End If
                MailBoxMenu = 0
                MailItem = 0
                MailItemValue = 0
                Exit Sub
                SendMessageBtnState = 0
            Else
                SendMessageBtnState = 0
            End If
        End If
    End If
    'DiscardMessageBtn
    If X >= DiscardMessageBtnBounds.Left And X <= DiscardMessageBtnBounds.Left + DiscardMessageBtnBounds.Right Then
        If Y >= DiscardMessageBtnBounds.Top And Y <= DiscardMessageBtnBounds.Top + DiscardMessageBtnBounds.Bottom Then
            If DiscardMessageBtnState = 2 Then
                MailItem = 0
                MailItemValue = 0
                MailBoxMenu = 0
                Exit Sub
                DiscardMessageBtnState = 0
            Else
                DiscardMessageBtnState = 0
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMail_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendMail_MouseMove(X As Long, Y As Long)
Dim Button As Long, dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long
Dim i As Long, z As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    If X >= SendMessageItemBounds.Left And X <= SendMessageItemBounds.Left + SendMessageItemBounds.Right Then
        If Y >= SendMessageItemBounds.Top And Y <= SendMessageItemBounds.Top + SendMessageItemBounds.Bottom Then
            If MailItem > 0 Then
                UpdateDescWindow GetPlayerInvItemNum(MyIndex, MailItem), 0, 0
                LastItemDesc = GetPlayerInvItemNum(MyIndex, MailItem)
                Exit Sub
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMail_MouseMove", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub SendMail_DblClick(X As Long, Y As Long)
Dim Button As Long, dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long
Dim i As Long, z As Long

   On Error GoTo errorhandler

    Button = MouseBtn




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendMail_DblClick", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub QuestLog_MouseDown(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Button As Long
Dim dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long, i As Long, z As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    If QuestLogFunction = 1 Then
        If X >= QuestDeclineLblBounds.Left And X <= QuestDeclineLblBounds.Left + QuestDeclineLblBounds.Right Then
            If Y >= QuestDeclineLblBounds.Top And Y <= QuestDeclineLblBounds.Top + QuestDeclineLblBounds.Bottom Then
                QuestLogPage = QuestLogPage - 1
                If QuestLogPage = -1 Then QuestLogPage = 2
            End If
        End If
            If X >= QuestAcceptLblBounds.Left And X <= QuestAcceptLblBounds.Left + QuestAcceptLblBounds.Right Then
            If Y >= QuestAcceptLblBounds.Top And Y <= QuestAcceptLblBounds.Top + QuestAcceptLblBounds.Bottom Then
                QuestLogPage = QuestLogPage + 1
                If QuestLogPage = 3 Then QuestLogPage = 0
            End If
        End If
        If Player(MyIndex).PlayerQuest(QuestLogQuest).state = QUEST_STARTED Then
            If X >= QuitQuestLblBounds.Left And X <= QuitQuestLblBounds.Left + QuitQuestLblBounds.Right Then
                If Y >= QuitQuestLblBounds.Top And Y <= QuitQuestLblBounds.Top + QuitQuestLblBounds.Bottom Then
                    PlayerHandleQuest QuestLogQuest, 2
                    InQuestLog = False
                End If
            End If
        End If
            'CloseQuestLogBtn Button
        If X >= CloseQuestLogBtnBounds.Left And X <= CloseQuestLogBtnBounds.Left + CloseQuestLogBtnBounds.Right Then
            If Y >= CloseQuestLogBtnBounds.Top And Y <= CloseQuestLogBtnBounds.Top + CloseQuestLogBtnBounds.Bottom Then
                CloseQuestLogBtnState = 2
                Exit Sub
            End If
        End If
    ElseIf QuestLogFunction = 0 Then
        If X >= QuestDeclineLblBounds.Left And X <= QuestDeclineLblBounds.Left + QuestDeclineLblBounds.Right Then
            If Y >= QuestDeclineLblBounds.Top And Y <= QuestDeclineLblBounds.Top + QuestDeclineLblBounds.Bottom Then
                InQuestLog = False
            End If
        End If
            If X >= QuestAcceptLblBounds.Left And X <= QuestAcceptLblBounds.Left + QuestAcceptLblBounds.Right Then
            If Y >= QuestAcceptLblBounds.Top And Y <= QuestAcceptLblBounds.Top + QuestAcceptLblBounds.Bottom Then
                PlayerHandleQuest QuestLogQuest, 1
                InQuestLog = False
                RefreshQuestLog
            End If
        End If
    ElseIf QuestLogFunction = 2 Then
        'CloseQuestLogBtn Button
        If X >= CloseQuestLogBtnBounds.Left And X <= CloseQuestLogBtnBounds.Left + CloseQuestLogBtnBounds.Right Then
            If Y >= CloseQuestLogBtnBounds.Top And Y <= CloseQuestLogBtnBounds.Top + CloseQuestLogBtnBounds.Bottom Then
                CloseQuestLogBtnState = 2
                Exit Sub
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "QuestLog_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub QuestLog_MouseUp(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, MusicFile As String, lineCount As Long, maxScroll As Long

   On Error GoTo errorhandler

    If InQuestLog = True Then
        If QuestLogFunction = 1 Or QuestLogFunction = 2 Then
            'CloseQuestLogBtn
            If X >= CloseQuestLogBtnBounds.Left And X <= CloseQuestLogBtnBounds.Left + CloseQuestLogBtnBounds.Right Then
                If Y >= CloseQuestLogBtnBounds.Top And Y <= CloseQuestLogBtnBounds.Top + CloseQuestLogBtnBounds.Bottom Then
                    If CloseQuestLogBtnState = 2 Then
                        InQuestLog = False
                        Exit Sub
                        CloseQuestLogBtnState = 0
                    Else
                        CloseQuestLogBtnState = 0
                    End If
                End If
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "QuestLog_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Currency_MouseDown(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Button As Long
Dim dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long, i As Long, z As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    'CurrencyOkBtn Button
    If X >= CurrencyOkBtnBounds.Left And X <= CurrencyOkBtnBounds.Left + CurrencyOkBtnBounds.Right Then
        If Y >= CurrencyOkBtnBounds.Top And Y <= CurrencyOkBtnBounds.Top + CurrencyOkBtnBounds.Bottom Then
            CurrencyOkBtnState = 2
            Exit Sub
        End If
    End If
    'CurrencyCancelBtn Button
    If X >= CurrencyCancelBtnBounds.Left And X <= CurrencyCancelBtnBounds.Left + CurrencyCancelBtnBounds.Right Then
        If Y >= CurrencyCancelBtnBounds.Top And Y <= CurrencyCancelBtnBounds.Top + CurrencyCancelBtnBounds.Bottom Then
            CurrencyCancelBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Currency_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Currency_MouseUp(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, MusicFile As String, lineCount As Long, maxScroll As Long
    'CurrencyOkBtn

   On Error GoTo errorhandler

    If X >= CurrencyOkBtnBounds.Left And X <= CurrencyOkBtnBounds.Left + CurrencyOkBtnBounds.Right Then
        If Y >= CurrencyOkBtnBounds.Top And Y <= CurrencyOkBtnBounds.Top + CurrencyOkBtnBounds.Bottom Then
            If CurrencyOkBtnState = 2 Then
                CurrencyOk
                Exit Sub
                CurrencyOkBtnState = 0
            Else
                CurrencyOkBtnState = 0
            End If
        End If
    End If
    'CurrencyCancelBtn
    If X >= CurrencyCancelBtnBounds.Left And X <= CurrencyCancelBtnBounds.Left + CurrencyCancelBtnBounds.Right Then
        If Y >= CurrencyCancelBtnBounds.Top And Y <= CurrencyCancelBtnBounds.Top + CurrencyCancelBtnBounds.Bottom Then
            If CurrencyCancelBtnState = 2 Then
                CurrencyText = vbNullString
                CurrencyItem = 0
                CurrencyMenu = 0 ' clear
                Exit Sub
                CurrencyCancelBtnState = 0
            Else
                CurrencyCancelBtnState = 0
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Currency_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub CurrencyOk()
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    If IsNumeric(CurrencyText) And CurrencyMenu <> 6 Then
        Select Case CurrencyMenu
            Case 1 ' drop item
                If Val(CurrencyText) > GetPlayerInvItemValue(MyIndex, CurrencyItem) Then CurrencyText = GetPlayerInvItemValue(MyIndex, CurrencyItem)
                SendDropItem CurrencyItem, Val(CurrencyText)
            Case 2 ' deposit item
                If Val(CurrencyText) > GetPlayerInvItemValue(MyIndex, CurrencyItem) Then CurrencyText = GetPlayerInvItemValue(MyIndex, CurrencyItem)
                DepositItem CurrencyItem, Val(CurrencyText)
            Case 3 ' withdraw item
                WithdrawItem CurrencyItem, Val(CurrencyText)
            Case 4 ' offer trade item
                If Val(CurrencyText) > GetPlayerInvItemValue(MyIndex, CurrencyItem) Then CurrencyText = GetPlayerInvItemValue(MyIndex, CurrencyItem)
                TradeItem CurrencyItem, Val(CurrencyText)
            Case 5
                If Val(CurrencyText) > 0 Then
                    If Val(CurrencyText) > GetPlayerInvItemValue(MyIndex, CurrencyItem) Then CurrencyText = GetPlayerInvItemValue(MyIndex, CurrencyItem)
                    MailItem = CurrencyItem
                    MailItemValue = Val(CurrencyText)
                End If
        End Select
    Else
        If CurrencyMenu = 6 Then
            If Len(Trim$(CurrencyText)) > 0 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CEditFriend
                buffer.WriteString Trim$(CurrencyText)
                buffer.WriteLong 0
                SendData buffer.ToArray
                Set buffer = Nothing
            End If
        Else
            AddText "Please enter a valid amount.", BrightRed
            Exit Sub
        End If
    End If
    CurrencyItem = 0
    CurrencyText = vbNullString
    CurrencyMenu = 0 ' clear


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CurrencyOk", "modInput", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub Dialogue_MouseDown(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Button As Long
Dim dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long, i As Long, z As Long

   On Error GoTo errorhandler

    Button = MouseBtn
    'DialogueYesBtn Button
    If X >= DialogueYesBtnBounds.Left And X <= DialogueYesBtnBounds.Left + DialogueYesBtnBounds.Right Then
        If Y >= DialogueYesBtnBounds.Top And Y <= DialogueYesBtnBounds.Top + DialogueYesBtnBounds.Bottom Then
            DialogueYesBtnState = 2
            Exit Sub
        End If
    End If
    'DialogueNoBtn Button
    If X >= DialogueNoBtnBounds.Left And X <= DialogueNoBtnBounds.Left + DialogueNoBtnBounds.Right Then
        If Y >= DialogueNoBtnBounds.Top And Y <= DialogueNoBtnBounds.Top + DialogueNoBtnBounds.Bottom Then
            DialogueNoBtnState = 2
            Exit Sub
        End If
    End If
    'DialogueOkayBtn Button
    If X >= DialogueOkayBtnBounds.Left And X <= DialogueOkayBtnBounds.Left + DialogueOkayBtnBounds.Right Then
        If Y >= DialogueOkayBtnBounds.Top And Y <= DialogueOkayBtnBounds.Top + DialogueOkayBtnBounds.Bottom Then
            DialogueOkayBtnState = 2
            Exit Sub
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Dialogue_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Dialogue_MouseUp(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, MusicFile As String, lineCount As Long, maxScroll As Long
    'DialogueYesBtn

   On Error GoTo errorhandler

    If X >= DialogueYesBtnBounds.Left And X <= DialogueYesBtnBounds.Left + DialogueYesBtnBounds.Right Then
        If Y >= DialogueYesBtnBounds.Top And Y <= DialogueYesBtnBounds.Top + DialogueYesBtnBounds.Bottom Then
            If DialogueYesBtnState = 2 Then
                dialogueHandler 2
                dialogueIndex = 0
                Exit Sub
                DialogueYesBtnState = 0
            Else
                DialogueYesBtnState = 0
            End If
        End If
    End If
    'DialogueNoBtn
    If X >= DialogueNoBtnBounds.Left And X <= DialogueNoBtnBounds.Left + DialogueNoBtnBounds.Right Then
        If Y >= DialogueNoBtnBounds.Top And Y <= DialogueNoBtnBounds.Top + DialogueNoBtnBounds.Bottom Then
            If DialogueNoBtnState = 2 Then
                dialogueHandler 3
                dialogueIndex = 0
                Exit Sub
                DialogueNoBtnState = 0
            Else
                DialogueNoBtnState = 0
            End If
        End If
    End If
    'DialogueOkayBtn
    If X >= DialogueOkayBtnBounds.Left And X <= DialogueOkayBtnBounds.Left + DialogueOkayBtnBounds.Right Then
        If Y >= DialogueOkayBtnBounds.Top And Y <= DialogueOkayBtnBounds.Top + DialogueOkayBtnBounds.Bottom Then
            If DialogueOkayBtnState = 2 Then
                dialogueHandler 1
                dialogueIndex = 0
                Exit Sub
                DialogueOkayBtnState = 0
            Else
                DialogueOkayBtnState = 0
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Dialogue_MouseUp", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub Event_MouseDown(X As Long, Y As Long)
Dim rec As rect, rec_pos As rect, Button As Long
Dim dX As Long, dY As Long, dw As Long, dH As Long, lineCount As Long, maxScroll As Long, i As Long, z As Long
Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Button = MouseBtn
        If EventChatType = 1 Then
            If EventChoiceVisible(1) Then
                'Response1Lbl
                dX = Response1LblBounds.Left
                dY = Response1LblBounds.Top
                dw = Response1LblBounds.Right
                dH = Response1LblBounds.Bottom
                If dw > 0 And dH > 0 And (dw + dX) > 0 And (dY + dH) > 0 Then
                    If dH / 14 > 0 Then
                        If GlobalX >= dX And GlobalX <= dX + dw And GlobalY >= dY And GlobalY <= dY + dH Then
                            Set buffer = New clsBuffer
                            buffer.WriteLong CEventChatReply
                            buffer.WriteLong EventReplyID
                            buffer.WriteLong EventReplyPage
                            buffer.WriteLong 1
                            SendData buffer.ToArray
                            Set buffer = Nothing
                            ClearEventChat
                            InEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            End If
                    If EventChoiceVisible(2) Then
                'Response2Lbl
                dX = Response2LblBounds.Left
                dY = Response2LblBounds.Top
                dw = Response2LblBounds.Right
                dH = Response2LblBounds.Bottom
                If dw > 0 And dH > 0 And (dw + dX) > 0 And (dY + dH) > 0 Then
                    If dH / 14 > 0 Then
                        If GlobalX >= dX And GlobalX <= dX + dw And GlobalY >= dY And GlobalY <= dY + dH Then
                            Set buffer = New clsBuffer
                            buffer.WriteLong CEventChatReply
                            buffer.WriteLong EventReplyID
                            buffer.WriteLong EventReplyPage
                            buffer.WriteLong 2
                            SendData buffer.ToArray
                            Set buffer = Nothing
                            ClearEventChat
                            InEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            End If
                    If EventChoiceVisible(3) Then
                'Response3Lbl
                dX = Response3LblBounds.Left
                dY = Response3LblBounds.Top
                dw = Response3LblBounds.Right
                dH = Response3LblBounds.Bottom
                If dw > 0 And dH > 0 And (dw + dX) > 0 And (dY + dH) > 0 Then
                    If dH / 14 > 0 Then
                        If GlobalX >= dX And GlobalX <= dX + dw And GlobalY >= dY And GlobalY <= dY + dH Then
                            Set buffer = New clsBuffer
                            buffer.WriteLong CEventChatReply
                            buffer.WriteLong EventReplyID
                            buffer.WriteLong EventReplyPage
                            buffer.WriteLong 3
                            SendData buffer.ToArray
                            Set buffer = Nothing
                            ClearEventChat
                            InEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            End If
                    If EventChoiceVisible(4) Then
                'Response4Lbl
                dX = Response4LblBounds.Left
                dY = Response4LblBounds.Top
                dw = Response4LblBounds.Right
                dH = Response4LblBounds.Bottom
                If dw > 0 And dH > 0 And (dw + dX) > 0 And (dY + dH) > 0 Then
                    If dH / 14 > 0 Then
                        If GlobalX >= dX And GlobalX <= dX + dw And GlobalY >= dY And GlobalY <= dY + dH Then
                            Set buffer = New clsBuffer
                            buffer.WriteLong CEventChatReply
                            buffer.WriteLong EventReplyID
                            buffer.WriteLong EventReplyPage
                            buffer.WriteLong 4
                            SendData buffer.ToArray
                            Set buffer = Nothing
                            ClearEventChat
                            InEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Else
            'EventContinueLbl
            dX = EventContinueLblBounds.Left
            dY = EventContinueLblBounds.Top
            dw = EventContinueLblBounds.Right
            dH = EventContinueLblBounds.Bottom
            If dw > 0 And dH > 0 And (dw + dX) > 0 And (dY + dH) > 0 Then
                If dH / 44 > 0 Then
                    If GlobalX >= dX And GlobalX <= dX + dw And GlobalY >= dY And GlobalY <= dY + dH Then
                        Set buffer = New clsBuffer
                        buffer.WriteLong CEventChatReply
                        buffer.WriteLong EventReplyID
                        buffer.WriteLong EventReplyPage
                        buffer.WriteLong 0
                        SendData buffer.ToArray
                        Set buffer = Nothing
                        ClearEventChat
                        InEvent = False
                        Exit Sub
                    End If
                End If
            End If
        End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Event_MouseDown", "modInput", Err.Number, Err.Description, Erl
    Err.Clear

End Sub





