Attribute VB_Name = "modQuests"
Option Explicit
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS_ITEMS As Byte = 10

Public Const TASK_KILLNPCS As Byte = 1
Public Const TASK_TALKEVENT As Byte = 2
Public Const TASK_AQUIREITEMS As Byte = 3
Public Const TASK_FETCHRETURN As Byte = 4
Public Const TASK_KILLPLAYERS As Byte = 5
Public Const TASK_GOTOMAP As Byte = 6
Public Const TASK_GETRESOURCES As Byte = 7

Public Const QUEST_NOT_STARTED As Byte = 0
Public Const QUEST_STARTED As Byte = 1
Public Const QUEST_COMPLETED As Byte = 2
Public Const QUEST_COMPLETED_BUT As Byte = 3

'Types
Public Quest(1 To 250) As QuestRec

Public Type PlayerQuestRec
    State As Integer
    CurrentTask As Integer
    TaskCount(1 To 5) As Integer
End Type

Private Type QuestItemRec
    Item As Long
    Value As Long
End Type

Public Type TaskRec
    type As Long
    data(1 To 8) As Long
    Text(1 To 4) As String * 200
    TaskDesc As String * 300
    EndOnCompletion As Byte
End Type

Public Type QuestRec
    Name As String * 30
    QuestDesc As String * 300
    ClassReq As Byte
    LevelReq As Byte
    QuestCompleteReq As Byte
    ItemReq As Byte
    NumQuestCompleteReq As Byte
    SwitchReq As Byte
    VariableReq As Byte
    RequiredClass As Long
    RequiredLevel As Long
    RequiredQuest As Long
    RequiredItem As Long
    RequiredItemVal As Long
    RequiredQuestCount As Long
    RequiredSwitchNum As Long
    RequiredSwitchSet As Long
    RequiredVariableNum As Long
    RequiredVariableCompare As Long
    RequiredVariableCompareTo As Long
    GiveItemBefore(0 To 3) As QuestItemRec
    TakeItemBefore(0 To 3) As QuestItemRec
    GiveItemAfter(0 To 3) As QuestItemRec
    TakeItemAfter(0 To 3) As QuestItemRec
    TeleportBefore As Byte
    BeforeMap As Long
    BeforeX As Long
    BeforeY As Long
    Task(1 To MAX_TASKS) As TaskRec
    RewardItem(1 To 3) As QuestItemRec
    RewardSpell(1 To 2) As Long
    TeleportAfter As Byte
    AfterMap As Long
    AfterX As Long
    AfterY As Long
    GiveExp As Long
    RestoreHealth As Byte
    RestoreMana As Byte
    SetPlayerVar As Long
    SetPlayerVarMod As Byte
    SetPlayerVarValue As Long
    SetPlayerSwitch As Long
    SetPlayerSwitchValue As Long
    Repeatable As Byte
    Abandonable As Byte
    QuestLogBefore As Byte
    QuestLogAfter As Byte
    QuestLogBeforeDesc As String * 200
    QuestLogAfterDesc As String * 200
End Type

' //////////////
' // DATABASE //
' //////////////

Sub SaveQuests()
    Dim i As Long

   On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveQuests", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SaveQuest(ByVal questnum As Long)
    Dim filename As String
    Dim F As Long, i As Long

   On Error GoTo errorhandler

    filename = App.path & "\data\quests\quest" & questnum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        With Quest(questnum)
            Put #F, , Quest(questnum)
        End With
    Close #F


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SaveQuest", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub LoadQuests()
    Dim filename As String
    Dim i As Integer
    Dim F As Long, n As Long
    Dim sLen As Long
    

   On Error GoTo errorhandler

    Call CheckQuests

    For i = 1 To MAX_QUESTS
        SetLoadingProgress "Loading Quests.", 29, i / MAX_QUESTS
        DoEvents
        filename = App.path & "\data\quests\quest" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Quest(i)
        Close #F
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadQuests", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub CheckQuests()
    Dim i As Long

   On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        If Not FileExist("\Data\quests\quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckQuests", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearQuest(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = vbNullString


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearQuest", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearQuests()
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
        SetLoadingProgress "Clearing Quests.", 16, i / MAX_QUESTS
        DoEvents
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearQuests", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendQuests(ByVal Index As Long)
    Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(Index, i)
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendQuests", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateQuestToAll(ByVal questnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(questnum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(questnum)), QuestSize
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong questnum
    Buffer.WriteBytes QuestData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateQuestToAll", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal questnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(questnum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(questnum)), QuestSize
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong questnum
    Buffer.WriteBytes QuestData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendUpdateQuestTo", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendPlayerQuests(ByVal Index As Long)
    Dim i As Long, x As Long
    Dim Buffer As clsBuffer


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest
        For i = 1 To MAX_QUESTS
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(i).State
            Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(i).CurrentTask
            For x = 1 To 5
                Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(i).TaskCount(x)
            Next
        Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerQuests", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SendPlayerQuest(ByVal Index As Long, ByVal questnum As Long)
    Dim Buffer As clsBuffer, i As Long


   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State
    Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).CurrentTask
    For i = 1 To 5
        Buffer.WriteLong Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i)
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendPlayerQuest", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

'sends a message to the client that is shown on the screen
Public Sub QuestMessage(ByVal Index As Long, ByVal questnum As Long, ByVal message As String, ByVal QuestNumForStart As Long, Optional ByVal questover As Long = 0)
    Dim Buffer As clsBuffer

   On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SQuestMessage
    Buffer.WriteLong questnum
    Buffer.WriteLong questover
    Buffer.WriteString Trim$(message)
    Buffer.WriteLong QuestNumForStart
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "QuestMessage", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
    
End Sub

' ///////////////
' // Functions //
' ///////////////

Public Function CanStartQuest(ByVal Index As Long, ByVal questnum As Long, Optional ByVal ActuallyStarting As Boolean = False) As Boolean
    Dim i As Long, n As Long, tempinv(1 To MAX_INV) As PlayerInvRec

   On Error GoTo errorhandler

    CanStartQuest = False
    If questnum < 1 Or questnum > MAX_QUESTS Then Exit Function
    If QuestInProgress(Index, questnum) Then Exit Function

        
    'check if now a completed quest can be repeated
    If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_COMPLETED Then
        If Quest(questnum).Repeatable = 1 Then
            Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_COMPLETED_BUT
        Else
            CanStartQuest = False
            Exit Function
        End If
    End If
    
    'Check if player has the quest 0 (not started) or 3 (completed but it can be started again)
    If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_NOT_STARTED Or Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_COMPLETED_BUT Then
        'Check if player's level is right
        If Quest(questnum).LevelReq = 0 Or Quest(questnum).RequiredLevel <= Player(Index).characters(TempPlayer(Index).CurChar).Level Then
            

            If Quest(questnum).ItemReq = 1 Then
                If HasItem(Index, Quest(questnum).RequiredItem) < Quest(questnum).RequiredItemVal Then
                    PlayerMsg Index, "You need " & Trim$(Item(Quest(questnum).RequiredItem).Name) & " to take this quest!", BrightRed
                    Exit Function
                End If
            End If
            
            If Quest(questnum).ClassReq = 1 Then
                If GetPlayerClass(Index) <> Quest(questnum).RequiredClass Then
                    PlayerMsg Index, "You must be a " & Trim$(Class(Quest(questnum).RequiredClass).Name) & " to accept this quest!", BrightRed
                    Exit Function
                End If
            End If
            
            If Quest(questnum).LevelReq = 1 Then
                If GetPlayerLevel(Index) < Quest(questnum).RequiredLevel Then
                    PlayerMsg Index, "You must be at least level " & STR(Quest(questnum).RequiredLevel) & " to accept this quest!", BrightRed
                    Exit Function
                End If
            End If
            
            If Quest(questnum).QuestCompleteReq = 1 Then
                If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(Quest(questnum).RequiredQuest).State <> QUEST_COMPLETED And Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(Quest(questnum).RequiredQuest).State <> QUEST_COMPLETED_BUT Then
                    PlayerMsg Index, "You must have completed " & Trim$(Quest(Quest(questnum).RequiredQuest).Name) & " to accept this quest!", BrightRed
                    Exit Function
                End If
            End If
            
            n = 0
            For i = 1 To MAX_QUESTS
                If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(i).State = QUEST_COMPLETED Or Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(i).State = QUEST_COMPLETED_BUT Then
                    n = n + 1
                End If
            Next
            
            If Quest(questnum).NumQuestCompleteReq = 1 Then
                If n < Quest(questnum).RequiredQuestCount Then
                    PlayerMsg Index, "You must complete more quests before accepthing this one!", BrightRed
                    Exit Function
                End If
            End If
            
            If Quest(questnum).VariableReq = 1 Then
                If Quest(questnum).RequiredVariableNum >= 0 And Quest(questnum).RequiredVariableNum < MAX_VARIABLES Then
                    Select Case Quest(questnum).RequiredVariableCompare
                        Case 0
                            If Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).RequiredVariableNum + 1) <> Quest(questnum).RequiredVariableCompareTo Then
                                Exit Function
                            End If
                        Case 1
                            If Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).RequiredVariableNum + 1) < Quest(questnum).RequiredVariableCompareTo Then
                                Exit Function
                            End If
                        Case 2
                            If Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).RequiredVariableNum + 1) > Quest(questnum).RequiredVariableCompareTo Then
                                Exit Function
                            End If
                        Case 3
                            If Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).RequiredVariableNum + 1) <= Quest(questnum).RequiredVariableCompareTo Then
                                Exit Function
                            End If
                        Case 4
                            If Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).RequiredVariableNum + 1) >= Quest(questnum).RequiredVariableCompareTo Then
                                Exit Function
                            End If
                        Case 5
                            If Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).RequiredVariableNum + 1) = Quest(questnum).RequiredVariableCompareTo Then
                                Exit Function
                            End If
                    End Select
                End If
            End If
            
            If Quest(questnum).SwitchReq = 1 Then
                If Quest(questnum).RequiredSwitchNum >= 0 And Quest(questnum).RequiredSwitchNum < MAX_SWITCHES Then
                    If Quest(questnum).RequiredSwitchSet = 0 Then
                        'Want True
                        If Player(Index).characters(TempPlayer(Index).CurChar).Switches(Quest(questnum).RequiredSwitchNum + 1) = 0 Then
                            Exit Function
                        End If
                    Else
                        If Player(Index).characters(TempPlayer(Index).CurChar).Switches(Quest(questnum).RequiredSwitchNum + 1) = 1 Then
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            For i = 1 To MAX_INV
                tempinv(i) = Player(Index).characters(TempPlayer(Index).CurChar).Inv(i)
            Next
            
            For i = 0 To 3
                If Quest(questnum).TakeItemBefore(i).Item > 0 Then
                    If HasItem(Index, Quest(questnum).TakeItemBefore(i).Item) >= Quest(questnum).TakeItemBefore(i).Value Then
                        TakeInvItem Index, Quest(questnum).TakeItemBefore(i).Item, Quest(questnum).TakeItemBefore(i).Value
                    Else
                        For n = 1 To MAX_INV
                            Player(Index).characters(TempPlayer(Index).CurChar).Inv(n) = tempinv(n)
                        Next
                        PlayerMsg Index, "You need " & Quest(questnum).TakeItemBefore(i).Value & " " & Trim$(Item(Quest(questnum).TakeItemBefore(i).Item).Name) & "(s) to take this quest!", BrightRed
                        Exit Function
                    End If
                End If
            Next
            
            For i = 0 To 3
                If Quest(questnum).GiveItemBefore(i).Item > 0 Then
                    If GiveInvItem(Index, Quest(questnum).GiveItemBefore(i).Item, Quest(questnum).GiveItemBefore(i).Value, True, False) = False Then
                        For n = 1 To MAX_INV
                            Player(Index).characters(TempPlayer(Index).CurChar).Inv(n) = tempinv(n)
                        Next
                        PlayerMsg Index, "You need more inventory space to begin this quest!", BrightRed
                        Exit Function
                    End If
                End If
            Next
            
            If ActuallyStarting = False Then
                For n = 1 To MAX_INV
                    Player(Index).characters(TempPlayer(Index).CurChar).Inv(n) = tempinv(n)
                Next
                SendInventory Index
            Else
                'Do the "before" teleport
                If Quest(questnum).TeleportBefore = 1 Then
                    PlayerWarp Index, Quest(questnum).BeforeMap, Quest(questnum).BeforeX, Quest(questnum).BeforeY
                End If
            End If
            CanStartQuest = True
        Else
            PlayerMsg Index, "You need to be a higher level to take this quest!", BrightRed
            Exit Function
        End If
    Else
        PlayerMsg Index, "You can't start that quest again!", BrightRed
        Exit Function
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanStartQuest", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function CanEndQuest(ByVal Index As Long, questnum As Long, EndNow As Boolean) As Boolean
Dim i As Long, x As Long

   On Error GoTo errorhandler

    CanEndQuest = False
    If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_STARTED Then
        If Quest(questnum).Task(Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).CurrentTask).EndOnCompletion = True Or EndNow = True Then
            CanEndQuest = True
        End If
        If CanEndQuest = False And EndNow = False Then
            x = Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).CurrentTask
            For i = x + 1 To 10
                If Quest(questnum).Task(i).type > 0 Then
                    CanEndQuest = False
                    Exit Function
                Else
                    If i = 10 Then
                        CanEndQuest = True
                        Exit Function
                    End If
                End If
            Next
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "CanEndQuest", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function QuestInProgress(ByVal Index As Long, ByVal questnum As Long) As Boolean

   On Error GoTo errorhandler

    QuestInProgress = False
    If questnum < 1 Or questnum > MAX_QUESTS Then Exit Function
    
    If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_STARTED Then
        QuestInProgress = True
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "QuestInProgress", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function QuestCompleted(ByVal Index As Long, ByVal questnum As Long) As Boolean

   On Error GoTo errorhandler

    QuestCompleted = False
    If questnum < 1 Or questnum > MAX_QUESTS Then Exit Function
    
    If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_COMPLETED Or Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "QuestCompleted", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function GetQuestNum(ByVal QuestName As String) As Long
    Dim i As Long

   On Error GoTo errorhandler

    GetQuestNum = 0
    
    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) = Trim$(QuestName) Then
            GetQuestNum = i
            Exit For
        End If
    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetQuestNum", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function GetItemNum(ByVal ItemName As String) As Long
    Dim i As Long

   On Error GoTo errorhandler

    GetItemNum = 0
    
    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) = Trim$(ItemName) Then
            GetItemNum = i
            Exit For
        End If
    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetItemNum", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Sub CheckTasks(ByVal Index As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim i As Long
    

   On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        If QuestInProgress(Index, i) Then
            If TaskType = Quest(i).Task(Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(i).CurrentTask).type Then
                Call CheckTask(Index, i, TaskType, TargetIndex)
            End If
        End If
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckTasks", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Function AdvanceQuestTask(Index, questnum) As Boolean
    Dim i As Long, x As Long, z As Long

   On Error GoTo errorhandler

    x = Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).CurrentTask
    AdvanceQuestTask = False
    For i = x + 1 To 10
        If Quest(questnum).Task(i).type > 0 Then
            Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).CurrentTask = i
            For z = 1 To 5
                Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(z) = 0
            Next
            AdvanceQuestTask = True
            SendPlayerQuest Index, questnum
            Exit Function
        End If
    Next


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "AdvanceQuestTask", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Sub CheckTask(ByVal Index As Long, ByVal questnum As Long, ByVal TaskType As Long, ByVal TargetIndex As Long, Optional forcecomplete As Boolean = False)
    Dim CurrentTask As Long, i As Long, taskcomplete As Boolean, x As Long, z As Long, p As Long

   On Error GoTo errorhandler

    CurrentTask = Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).CurrentTask
    
    Select Case TaskType
        Case TASK_KILLNPCS
            taskcomplete = True
            With Quest(questnum).Task(CurrentTask)
                For i = 1 To 4
                    If .data(i) > 0 Then
                        If .data(i) = TargetIndex Then
                            If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i) < Quest(questnum).Task(CurrentTask).data(i + 4) Then
                                Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i) = Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i) + 1
                                PlayerMsg Index, "Quest: " + Trim$(Quest(questnum).Name) + " - " + Trim$(Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i)) + "/" + Trim$(Quest(questnum).Task(CurrentTask).data(4 + 1)) + " " + Trim$(Npc(.data(i)).Name) + " killed.", Yellow
                                If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i) < .data(4 + i) Then
                                    taskcomplete = False
                                End If
                            End If
                        Else
                            If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i) < Quest(questnum).Task(CurrentTask).data(i + 4) Then
                                If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i) < .data(4 + i) Then
                                    taskcomplete = False
                                End If
                            End If
                        End If
                    End If
                Next
            End With
            If taskcomplete = True Or forcecomplete = True Then
                PlayerMsg Index, "Quest: " & Trim$(Quest(questnum).Name) & " - Task completed.", Yellow
                If CanEndQuest(Index, questnum, False) Then
                    EndQuest Index, questnum
                Else
                    If AdvanceQuestTask(Index, questnum) = True Then
                        For i = 1 To 5
                            Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(i) = 0
                        Next
                    End If
                End If
            End If
        Case TASK_TALKEVENT
            PlayerMsg Index, "Quest: " & Trim$(Quest(questnum).Name) & " - Task completed.", Yellow
            If CanEndQuest(Index, questnum, False) Then
                EndQuest Index, questnum
            Else
                If AdvanceQuestTask(Index, questnum) = True Then
                End If
            End If
        Case TASK_AQUIREITEMS
            For x = 1 To 4
                If TargetIndex = Quest(questnum).Task(CurrentTask).data(x) Then
                    Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = 0
                    For i = 1 To MAX_INV
                        If GetPlayerInvItemNum(Index, i) = TargetIndex Then
                            If Item(GetPlayerInvItemNum(Index, i)).Stackable = 1 Then
                                Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = GetPlayerInvItemValue(Index, i)
                            Else
                                Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) + 1
                            End If
                        End If
                    Next
                    
                    If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) > Quest(questnum).Task(CurrentTask).data(x + 4) Then Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = Quest(questnum).Task(CurrentTask).data(x + 4)
                    PlayerMsg Index, "Quest: " + Trim$(Quest(questnum).Name) + " - You have " + Trim$(Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x)) + "/" + Trim$(Quest(questnum).Task(CurrentTask).data(x + 4)) + " " + Trim$(Item(TargetIndex).Name), Yellow
                End If
            Next
            p = 0
            For x = 1 To 4
                If Quest(questnum).Task(CurrentTask).data(x) > 0 Then
                    If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) >= Quest(questnum).Task(CurrentTask).data(x + 4) Or forcecomplete Then
                    Else
                        p = 1
                    End If
                End If
            Next
            If p = 0 Then
                PlayerMsg Index, "Quest: " & Trim$(Quest(questnum).Name) & " - Task completed.", Yellow
                If CanEndQuest(Index, questnum, False) Then
                    EndQuest Index, questnum
                    Exit Sub
                Else
                    If AdvanceQuestTask(Index, questnum) = True Then
                        Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = 0
                        Exit Sub
                    End If
                End If
            End If
            
        Case TASK_GOTOMAP
            If TargetIndex = Quest(questnum).Task(CurrentTask).data(1) Then
                PlayerMsg Index, "Quest: " & Trim$(Quest(questnum).Name) & " - Task completed.", Yellow
                If CanEndQuest(Index, questnum, False) Then
                    EndQuest Index, questnum
                Else
                    If AdvanceQuestTask(Index, questnum) = True Then
                    End If
                End If
            End If
        Case TASK_KILLPLAYERS
            Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(1) = Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(1) + 1
            PlayerMsg Index, "Quest: " + Trim$(Quest(questnum).Name) + " - " + Trim$(Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(1)) + "/" + Trim$(Quest(questnum).Task(CurrentTask).data(1)) + " players killed.", Yellow
            If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(1) >= Quest(questnum).Task(CurrentTask).data(1) Or forcecomplete Then
                PlayerMsg Index, "Quest: " & Trim$(Quest(questnum).Name) & " - Task completed.", Yellow
                If CanEndQuest(Index, questnum, False) Then
                    EndQuest Index, questnum
                Else
                    If AdvanceQuestTask(Index, questnum) = True Then
                    End If
                End If
            End If
        
        Case TASK_FETCHRETURN
            If TargetIndex > 0 Then
                For x = 1 To 4
                    If TargetIndex = Quest(questnum).Task(CurrentTask).data(x) Then
                        Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = 0
                        For i = 1 To MAX_INV
                            If GetPlayerInvItemNum(Index, i) = TargetIndex Then
                                If Item(GetPlayerInvItemNum(Index, i)).Stackable = 1 Then
                                    Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = GetPlayerInvItemValue(Index, i)
                                Else
                                    Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) + 1
                                End If
                            End If
                        Next
                        
                        If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) > Quest(questnum).Task(CurrentTask).data(x + 4) Then Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = Quest(questnum).Task(CurrentTask).data(x + 4)
                        PlayerMsg Index, "Quest: " & Trim$(Quest(questnum).Name) & " - You have " & Trim$(Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x)) & "/" + Trim$(Quest(questnum).Task(CurrentTask).data(x + 4)) + " " + Trim$(Item(TargetIndex).Name), Yellow
                    End If
                Next
                
            Else
                'Talking to the event. Check if we have the items, if so, lets give them up and move on.
                If TargetIndex = -1 Then
                    p = 0
                    For x = 1 To 4
                        If Quest(questnum).Task(CurrentTask).data(x) > 0 Then
                            If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) >= Quest(questnum).Task(CurrentTask).data(x + 4) Or forcecomplete Then
                            Else
                                p = 1
                            End If
                        End If
                    Next
                    If p = 0 Then
                        ' Take all 4 items and move on!
                        For z = 1 To 4
                            If Quest(questnum).Task(CurrentTask).data(z) > 0 Then
                                TakeInvItem Index, Quest(questnum).Task(CurrentTask).data(z), Quest(questnum).Task(CurrentTask).data(z + 4)
                            End If
                        Next
                        PlayerMsg Index, "Quest: " & Trim$(Quest(questnum).Name) & " - Task completed.", Yellow
                        If CanEndQuest(Index, questnum, False) Then
                            EndQuest Index, questnum
                            Exit Sub
                        Else
                            If AdvanceQuestTask(Index, questnum) = True Then
                                Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = 0
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
            
        Case TASK_GETRESOURCES
            For x = 1 To 4
                If TargetIndex = Quest(questnum).Task(CurrentTask).data(x) Then
                    Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) + 1
                    If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) > Quest(questnum).Task(CurrentTask).data(x + 4) Then Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) = Quest(questnum).Task(CurrentTask).data(x + 4)
                    PlayerMsg Index, "Quest: " + Trim$(Quest(questnum).Name) + " - " + Trim$(Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x)) + "/" + Trim$(Quest(questnum).Task(CurrentTask).data(x + 4)) + " " & Trim$(Resource(Quest(questnum).Task(CurrentTask).data(x)).Name) & " depleated.", Yellow
                End If
            Next
            For x = 1 To 4
                If Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(x) >= Quest(questnum).Task(CurrentTask).data(x + 1) Or forcecomplete Then
                    If x = 4 Then
                        PlayerMsg Index, "Quest: " & Trim$(Quest(questnum).Name) & " - Task completed.", Yellow
                        If CanEndQuest(Index, questnum, False) Then
                            EndQuest Index, questnum
                        Else
                            If AdvanceQuestTask(Index, questnum) = True Then
                            End If
                        End If
                    End If
                Else
                    Exit Sub
                End If
            Next
            
        Case Else
            If forcecomplete Then
                PlayerMsg Index, "Quest: " & Trim$(Quest(questnum).Name) & " - Task completed.", Yellow
                If CanEndQuest(Index, questnum, False) Then
                    EndQuest Index, questnum
                Else
                    If AdvanceQuestTask(Index, questnum) = True Then
                    End If
                End If
            End If
    End Select
    SavePlayer Index
    SendPlayerData Index
    SendPlayerQuests Index


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CheckTask", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub EndQuest(ByVal Index As Long, ByVal questnum As Long)
    Dim i As Long, n As Long, z As Long, tempinv(1 To MAX_INV) As PlayerInvRec, spellslot(1 To 2) As Long, p As Long, y As Long
    

   On Error GoTo errorhandler
    For i = 1 To MAX_INV
        tempinv(i).Num = Player(Index).characters(TempPlayer(Index).CurChar).Inv(i).Num
        tempinv(i).Value = Player(Index).characters(TempPlayer(Index).CurChar).Inv(i).Value
    Next
    
    'give rewards
    For i = 1 To 3
        If Quest(questnum).RewardItem(i).Item > 0 Then
            If GiveInvItem(Index, Quest(questnum).RewardItem(i).Item, Quest(questnum).RewardItem(i).Value, True, False, True) = True Then
            
            Else
                For z = 1 To MAX_INV
                    Player(Index).characters(TempPlayer(Index).CurChar).Inv(z).Num = tempinv(z).Num
                    Player(Index).characters(TempPlayer(Index).CurChar).Inv(z).Value = tempinv(z).Value
                    SendInventoryUpdate Index, z
                Next
                Exit Sub
            End If
        End If
    Next
    
    For i = 1 To 2
        y = 0
        If Quest(questnum).RewardSpell(i) > 0 Then
            For p = 1 To MAX_PLAYER_SPELLS
                If Player(Index).characters(TempPlayer(Index).CurChar).Spell(p) = Quest(questnum).RewardSpell(i) Then
                    'Already know the spell... no big deal.
                    y = 1
                Else
                    If p = MAX_PLAYER_SPELLS And y = 0 Then
                        If FindOpenSpellSlot(Index) > 0 Then
                            spellslot(i) = FindOpenSpellSlot(Index)
                            SetPlayerSpell Index, spellslot(i), Quest(questnum).RewardSpell(i)
                        Else
                            For z = 1 To 2
                                If spellslot(z) > 0 Then
                                    SetPlayerSpell Index, spellslot(z), 0
                                End If
                            Next
                            For z = 1 To MAX_INV
                                Player(Index).characters(TempPlayer(Index).CurChar).Inv(z).Num = tempinv(z).Num
                                Player(Index).characters(TempPlayer(Index).CurChar).Inv(z).Value = tempinv(z).Value
                                SendInventoryUpdate Index, z
                            Next
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End If
    Next
    
    SendPlayerSpells Index
    
    'Easy part done.
    GivePlayerEXP Index, Quest(questnum).GiveExp
    
    If Quest(questnum).RestoreHealth = 1 Then
        SetPlayerVital Index, HP, GetPlayerMaxVital(Index, HP)
    End If
    
    If Quest(questnum).RestoreMana = 1 Then
        SetPlayerVital Index, MP, GetPlayerMaxVital(Index, MP)
    End If
    
    If Quest(questnum).SetPlayerVar > 0 Then
        Select Case Quest(questnum).SetPlayerVarMod
            Case 0 'Set
                Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).SetPlayerVar) = Quest(questnum).SetPlayerVarValue
            Case 1 'Add
                Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).SetPlayerVar) = Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).SetPlayerVar) + Quest(questnum).SetPlayerVarValue
            Case 2 'Subtract
                Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).SetPlayerVar) = Player(Index).characters(TempPlayer(Index).CurChar).Variables(Quest(questnum).SetPlayerVar) - Quest(questnum).SetPlayerVarValue
        End Select
    End If
    
    If Quest(questnum).SetPlayerSwitch > 0 Then
        If Quest(questnum).SetPlayerSwitchValue = 0 Then 'True
            Player(Index).characters(TempPlayer(Index).CurChar).Switches(Quest(questnum).SetPlayerSwitch) = 1
        Else 'False
            Player(Index).characters(TempPlayer(Index).CurChar).Switches(Quest(questnum).SetPlayerSwitch) = 0
        End If
    End If
    
    'If all is successful then finish up :D
    Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).State = QUEST_COMPLETED
    Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).CurrentTask = 0
    For z = 1 To 5
        Player(Index).characters(TempPlayer(Index).CurChar).PlayerQuest(questnum).TaskCount(z) = 0
    Next
    
    QuestMessage Index, questnum, Trim$(Quest(questnum).QuestDesc), 0, 1

    'mark quest as completed in chat
    PlayerMsg Index, Trim$(Quest(questnum).Name) & " completed!", Green
    
    SavePlayer Index
    SendEXP Index
    Call SendStats(Index)
    SendPlayerData Index
    SendPlayerQuests Index
    
    For i = 1 To MAX_INV
        Call CheckTasks(Index, TASK_AQUIREITEMS, GetPlayerInvItemNum(Index, i))
        Call CheckTasks(Index, TASK_FETCHRETURN, GetPlayerInvItemNum(Index, i))
    Next
    
    If Quest(questnum).TeleportAfter = 1 Then
        PlayerWarp Index, Quest(questnum).AfterMap, Quest(questnum).AfterX, Quest(questnum).AfterY
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EndQuest", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
