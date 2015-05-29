Attribute VB_Name = "modQuests"

Option Explicit

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const EDITOR_TASKS As Byte = 7

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

Public Quest_Changed(1 To 250) As Boolean

Public QuestEditorPage As Long
Public QuestEditorTask As Long

'Types
Public quest(1 To 250) As QuestRec

Public Type PlayerQuestRec
    state As Integer
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
    classReq As Byte
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

' ////////////
' // Editor //
' ////////////

Public Sub QuestEditorInit()
Dim i As Long

   On Error GoTo errorhandler

    If frmEditor_Quest.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1
    QuestEditorPage = 1
    QuestEditorInitPage
    frmEditor_Quest.cmdPrevStep.Enabled = False
    frmEditor_Quest.cmdNextStep.Enabled = True
    frmEditor_Quest.fraStep1.Visible = True
    frmEditor_Quest.fraStep2.Visible = False
    frmEditor_Quest.fraStep3.Visible = False
    
    With frmEditor_Quest
        .txtQuestName.Text = Trim$(quest(EditorIndex).Name)
    
    End With

    Quest_Changed(EditorIndex) = True


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "QuestEditorInit", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub QuestEditorOk()
Dim i As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        If Quest_Changed(i) Then
            Call SendSaveQuest(i)
        End If
    Next
    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "QuestEditorOk", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub QuestEditorCancel()

   On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Quest
    ClearChanged_Quest
    ClearQuests
    SendRequestQuests


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "QuestEditorCancel", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub ClearChanged_Quest()

   On Error GoTo errorhandler

    ZeroMemory Quest_Changed(1), MAX_QUESTS * 2 ' 2 = boolean length


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearChanged_Quest", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' //////////////
' // DATABASE //
' //////////////

Sub ClearQuest(ByVal Index As Long)

   On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(quest(Index)), LenB(quest(Index)))
    quest(Index).Name = vbNullString


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
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearQuests", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendRequestEditQuest()
Dim buffer As clsBuffer


   On Error GoTo errorhandler
   
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditQuest
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestEditQuest", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub SendSaveQuest(ByVal questnum As Long)
Dim buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    QuestSize = LenB(quest(questnum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(quest(questnum)), QuestSize
    buffer.WriteLong CSaveQuest
    buffer.WriteLong questnum
    buffer.WriteBytes QuestData
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendSaveQuest", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub SendRequestQuests()
Dim buffer As clsBuffer


   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestQuests
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SendRequestQuests", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub UpdateQuestLog()
    Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CQuestLogUpdate
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateQuestLog", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub PlayerHandleQuest(ByVal questnum As Long, ByVal Order As Long)
    Dim buffer As clsBuffer

   On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerHandleQuest
    buffer.WriteLong questnum
    buffer.WriteLong Order '1=accept quest, 2=cancel quest
    SendData buffer.ToArray()
    Set buffer = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PlayerHandleQuest", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function QuestInProgress(ByVal questnum As Long) As Boolean

   On Error GoTo errorhandler

    QuestInProgress = False
    If questnum < 1 Or questnum > MAX_QUESTS Then Exit Function
    If Player(MyIndex).PlayerQuest(questnum).state = QUEST_STARTED Then 'Status=1 means started
        QuestInProgress = True
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "QuestInProgress", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function QuestCompleted(ByVal questnum As Long) As Boolean

   On Error GoTo errorhandler

    QuestCompleted = False
    If questnum < 1 Or questnum > MAX_QUESTS Then Exit Function
    If Player(MyIndex).PlayerQuest(questnum).state = QUEST_COMPLETED Then
        QuestCompleted = True
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "QuestCompleted", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function GetQuestNum(ByVal questname As String) As Long
    Dim i As Long

   On Error GoTo errorhandler

    GetQuestNum = 0
    For i = 1 To MAX_QUESTS
        If Trim$(quest(i).Name) = Trim$(questname) Then
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

Public Sub RefreshQuestLog()
    Dim i As Long, itemstr As String, CurrentTask As Long, questnum As Long, Index As Long, X As Long, z As Long
    Dim questslst() As String, actualquest() As Long
    

   On Error GoTo errorhandler
   
    ReDim questslst(1 To MAX_QUESTS)
    ReDim actualquest(1 To MAX_QUESTS)

    Index = MyIndex
    z = 1
    For i = 1 To MAX_QUESTS
        QuestList(i) = ""
        QuestIndex(i) = 0
    Next
    For i = 1 To MAX_QUESTS
        If QuestInProgress(i) Then
            itemstr = Trim$(quest(i).Name)
            questnum = i
            CurrentTask = Player(Index).PlayerQuest(i).CurrentTask
            Select Case quest(questnum).Task(Player(MyIndex).PlayerQuest(i).CurrentTask).type
                Case TASK_KILLNPCS
                    'itemstr = itemstr + " - " + Trim$(Player(Index).PlayerQuest(questnum).CurrentCount) + "/" + Trim$(quest(questnum).task(CurrentTask).Amount) + " " + Trim$(NPC(quest(questnum).task(CurrentTask).NPC).Name) + " killed."
                Case TASK_AQUIREITEMS
                    'itemstr = itemstr + " - " + Trim$(Player(Index).PlayerQuest(questnum).CurrentCount) + "/" + Trim$(quest(questnum).task(CurrentTask).Amount) + " " + Trim$(Item(quest(questnum).task(CurrentTask).Item).Name)
                Case TASK_KILLPLAYERS
                    'itemstr = itemstr + " - " + Trim$(Player(Index).PlayerQuest(questnum).CurrentCount) + "/" + Trim$(quest(questnum).task(CurrentTask).Amount) + " players killed."
                'Case task_trainresource
                    'itemstr = itemstr + " - " + Trim$(Player(Index).PlayerQuest(questnum).CurrentCount) + "/" + Trim$(quest(questnum).task(CurrentTask).Amount) + " hits."
                Case Else
                    'itemstr = itemstr + " - " + Trim$(quest(questnum).task(CurrentTask).TaskLog)
            End Select
            QuestList(z) = itemstr
            QuestIndex(z) = i
            z = z + 1
        ElseIf QuestCompleted(i) Then
            If quest(i).QuestLogAfter = 1 Then
                itemstr = Trim$(quest(i).Name)
                QuestList(z) = itemstr
                QuestIndex(z) = i
                z = z + 1
            End If
        Else
            If quest(i).QuestLogBefore = 1 Then
                itemstr = Trim$(quest(i).Name)
                QuestList(z) = itemstr
                QuestIndex(z) = i
                z = z + 1
            End If
        End If
    Next

    QuestCount = z - 1
    QuestListScroll = 0
    QuestSelection = 0


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "RefreshQuestLog", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub QuestEditorInitPage()
Dim i As Long

   On Error GoTo errorhandler
   
    Select Case QuestEditorPage
        Case 1
            frmEditor_Quest.txtQuestName.Text = Trim$(quest(EditorIndex).Name)
            frmEditor_Quest.txtQuestOffer.Text = Trim$(quest(EditorIndex).QuestDesc)
            
            If quest(EditorIndex).classReq = 1 Then
                frmEditor_Quest.classReq.Value = 1
                If Max_Classes > 0 Then
                    If quest(EditorIndex).RequiredClass > 0 And quest(EditorIndex).RequiredClass <= Max_Classes Then
                        frmEditor_Quest.cmbClassReq.ListIndex = quest(EditorIndex).RequiredClass - 1
                    Else
                        frmEditor_Quest.cmbClassReq.ListIndex = 0
                    End If
                End If
            Else
                frmEditor_Quest.classReq.Value = 0
                If Max_Classes > 0 Then
                    frmEditor_Quest.cmbClassReq.ListIndex = 0
                End If
            End If
            
            
            If quest(EditorIndex).LevelReq = 1 Then
                frmEditor_Quest.chkLevelReq.Value = 1
                If quest(EditorIndex).RequiredLevel > 0 And quest(EditorIndex).RequiredLevel <= MAX_LEVELS Then
                    frmEditor_Quest.scrlLevelReq.Value = quest(EditorIndex).RequiredLevel
                Else
                    frmEditor_Quest.scrlLevelReq.Value = 1
                End If
            Else
                frmEditor_Quest.chkLevelReq.Value = 0
                frmEditor_Quest.scrlLevelReq.Value = 1
            End If
            
            
            If quest(EditorIndex).QuestCompleteReq = 1 Then
                frmEditor_Quest.chkQuestReq.Value = 1
                If MAX_QUESTS > 0 Then
                    If quest(EditorIndex).RequiredQuest > 0 And quest(EditorIndex).RequiredQuest <= MAX_QUESTS Then
                        frmEditor_Quest.cmbQuestReq.ListIndex = quest(EditorIndex).RequiredQuest - 1
                    Else
                        frmEditor_Quest.cmbQuestReq.ListIndex = 0
                    End If
                End If
            Else
                frmEditor_Quest.chkQuestReq.Value = 0
                If MAX_QUESTS > 0 Then
                    frmEditor_Quest.cmbQuestReq.ListIndex = 0
                End If
            End If
            
            If quest(EditorIndex).ItemReq = 1 Then
                frmEditor_Quest.chkItemReq = 1
                If MAX_ITEMS > 0 Then
                    If quest(EditorIndex).RequiredItem > 0 And quest(EditorIndex).RequiredItem <= MAX_ITEMS Then
                        frmEditor_Quest.cmbItemReq.ListIndex = quest(EditorIndex).RequiredItem - 1
                        If quest(EditorIndex).RequiredItemVal > 0 Then
                            frmEditor_Quest.scrlItemReqVal.Value = quest(EditorIndex).RequiredItemVal
                        End If
                    Else
                        If MAX_ITEMS > 0 Then
                            frmEditor_Quest.cmbItemReq.ListIndex = 0
                        End If
                        frmEditor_Quest.cmbItemReq.ListIndex = 0
                        frmEditor_Quest.scrlItemReqVal.Value = 1
                    End If
                Else
                    frmEditor_Quest.scrlItemReqVal.Value = 1
                End If
            Else
                frmEditor_Quest.chkItemReq.Value = 0
                If MAX_ITEMS > 0 Then
                    frmEditor_Quest.cmbItemReq.ListIndex = 0
                End If
                frmEditor_Quest.scrlItemReqVal.Value = 1
            End If
            
            If quest(EditorIndex).NumQuestCompleteReq = 1 Then
                frmEditor_Quest.chkNumQuestReq.Value = 1
                If quest(EditorIndex).RequiredQuestCount > 0 And quest(EditorIndex).RequiredQuestCount <= MAX_QUESTS Then
                    frmEditor_Quest.scrlQuestCompleteCount.Value = quest(EditorIndex).RequiredQuestCount
                End If
            Else
                frmEditor_Quest.chkNumQuestReq.Value = 0
                frmEditor_Quest.scrlQuestCompleteCount.Value = 1
            End If
            
            If quest(EditorIndex).VariableReq = 1 Then
                frmEditor_Quest.chkVariableReq.Value = 1
                If MAX_VARIABLES > 0 Then
                    If quest(EditorIndex).RequiredVariableNum >= 0 And quest(EditorIndex).RequiredVariableNum < MAX_VARIABLES Then
                        frmEditor_Quest.cmbPlayerVarReq.ListIndex = quest(EditorIndex).RequiredVariableNum
                        frmEditor_Quest.cmbVariableReqCompare.ListIndex = quest(EditorIndex).RequiredVariableCompare
                        frmEditor_Quest.txtVariableReq.Text = str(quest(EditorIndex).RequiredVariableCompareTo)
                    Else
                        frmEditor_Quest.cmbPlayerVarReq.ListIndex = 0
                        frmEditor_Quest.cmbVariableReqCompare.ListIndex = 0
                        frmEditor_Quest.txtVariableReq.Text = "0"
                    End If
                Else
                    frmEditor_Quest.cmbVariableReqCompare.ListIndex = 0
                    frmEditor_Quest.txtVariableReq.Text = "0"
                End If
            Else
                frmEditor_Quest.chkVariableReq.Value = 0
                If MAX_VARIABLES > 0 Then
                    frmEditor_Quest.cmbPlayerVarReq.ListIndex = 0
                End If
                frmEditor_Quest.cmbVariableReqCompare.ListIndex = 0
                frmEditor_Quest.txtVariableReq.Text = "0"
            End If
            
            If quest(EditorIndex).SwitchReq = 1 Then
                frmEditor_Quest.chkSwitchReq.Value = 1
                If MAX_SWITCHES > 0 Then
                    If quest(EditorIndex).RequiredSwitchNum >= 0 And quest(EditorIndex).RequiredSwitchNum < MAX_SWITCHES Then
                        frmEditor_Quest.cmbPlayerSwitchReq.ListIndex = quest(EditorIndex).RequiredSwitchNum
                        frmEditor_Quest.cmbSwitchReqCompare.ListIndex = quest(EditorIndex).RequiredSwitchSet
                    Else
                        frmEditor_Quest.cmbPlayerSwitchReq.ListIndex = 0
                        frmEditor_Quest.cmbSwitchReqCompare.ListIndex = 0
                    End If
                Else
                    frmEditor_Quest.cmbSwitchReqCompare.ListIndex = 0
                End If
            Else
                frmEditor_Quest.chkSwitchReq.Value = 0
                If MAX_SWITCHES > 0 Then
                    frmEditor_Quest.cmbPlayerSwitchReq.ListIndex = 0
                End If
                frmEditor_Quest.cmbSwitchReqCompare.ListIndex = 0
            End If
        Case 2
            For i = 0 To 3
                If quest(EditorIndex).GiveItemBefore(i).Item > 0 And quest(EditorIndex).GiveItemBefore(i).Item <= MAX_ITEMS Then
                    frmEditor_Quest.cmbGiveItem(i).ListIndex = quest(EditorIndex).GiveItemBefore(i).Item
                    If quest(EditorIndex).GiveItemBefore(i).Value > 0 Then
                        frmEditor_Quest.scrlGiveItemVal(i).Value = quest(EditorIndex).GiveItemBefore(i).Value
                    Else
                        frmEditor_Quest.scrlGiveItemVal(i).Value = 1
                    End If
                Else
                    frmEditor_Quest.cmbGiveItem(i).ListIndex = 0
                    frmEditor_Quest.scrlGiveItemVal(i).Value = 1
                End If
            Next
            For i = 0 To 3
                If quest(EditorIndex).TakeItemBefore(i).Item > 0 And quest(EditorIndex).TakeItemBefore(i).Item <= MAX_ITEMS Then
                    frmEditor_Quest.cmbTakeItem(i).ListIndex = quest(EditorIndex).TakeItemBefore(i).Item
                    If quest(EditorIndex).TakeItemBefore(i).Value > 0 Then
                        frmEditor_Quest.scrlTakeItemVal(i).Value = quest(EditorIndex).TakeItemBefore(i).Value
                    Else
                        frmEditor_Quest.scrlTakeItemVal(i).Value = 1
                    End If
                Else
                    frmEditor_Quest.cmbTakeItem(i).ListIndex = 0
                    frmEditor_Quest.scrlTakeItemVal(i).Value = 1
                End If
            Next
            If quest(EditorIndex).TeleportBefore = 1 Then
                frmEditor_Quest.chkTeleportOnStart.Value = 1
                frmEditor_Quest.scrlTeleMap.Value = quest(EditorIndex).BeforeMap
                frmEditor_Quest.scrlTeleX.Value = quest(EditorIndex).BeforeX
                frmEditor_Quest.scrlTeleY.Value = quest(EditorIndex).BeforeY
            Else
                frmEditor_Quest.chkTeleportOnStart.Value = 0
                frmEditor_Quest.scrlTeleMap.Value = 1
                frmEditor_Quest.scrlTeleX.Value = 1
                frmEditor_Quest.scrlTeleY.Value = 1
            End If
        Case 3
        
        Case 4
            For i = 1 To 3
                If quest(EditorIndex).RewardItem(i).Item > 0 And quest(EditorIndex).RewardItem(i).Item <= MAX_ITEMS Then
                    frmEditor_Quest.cmbRewardItem(i).ListIndex = quest(EditorIndex).RewardItem(i).Item
                    If quest(EditorIndex).RewardItem(i).Value > 0 Then
                        frmEditor_Quest.scrlRewardItemVal(i).Value = quest(EditorIndex).RewardItem(i).Value
                    Else
                        frmEditor_Quest.scrlRewardItemVal(i).Value = 1
                    End If
                Else
                    frmEditor_Quest.cmbRewardItem(i).ListIndex = 0
                    frmEditor_Quest.scrlRewardItemVal(i).Value = 1
                End If
            Next
            For i = 1 To 2
                If quest(EditorIndex).RewardSpell(i) > 0 And quest(EditorIndex).RewardSpell(i) <= MAX_SPELLS Then
                    frmEditor_Quest.cmbSpellReward(i).ListIndex = quest(EditorIndex).RewardSpell(i)
                Else
                    frmEditor_Quest.cmbSpellReward(i).ListIndex = 0
                End If
            Next
            frmEditor_Quest.chkTeleAfter.Value = quest(EditorIndex).TeleportAfter
            If quest(EditorIndex).AfterMap = 0 Then quest(EditorIndex).AfterMap = 1
            frmEditor_Quest.scrlAfterMap.Value = quest(EditorIndex).AfterMap
            frmEditor_Quest.scrlAfterX.Value = quest(EditorIndex).AfterX
            frmEditor_Quest.scrlAfterY.Value = quest(EditorIndex).AfterY
            frmEditor_Quest.scrlGiveExp.Value = quest(EditorIndex).GiveExp
            frmEditor_Quest.chkRestoreHealth.Value = quest(EditorIndex).RestoreHealth
            frmEditor_Quest.chkRestoreMana.Value = quest(EditorIndex).RestoreMana
            frmEditor_Quest.cmbSetPlayerVar.ListIndex = quest(EditorIndex).SetPlayerVar
            frmEditor_Quest.cmbSetVariableCompare.ListIndex = quest(EditorIndex).SetPlayerVarMod
            frmEditor_Quest.txtSetPlayerVarValue.Text = str(quest(EditorIndex).SetPlayerVarValue)
            frmEditor_Quest.cmbSetPlayerSwitch.ListIndex = quest(EditorIndex).SetPlayerSwitch
            frmEditor_Quest.cmbSetSwitchCompare.ListIndex = quest(EditorIndex).SetPlayerSwitchValue
            frmEditor_Quest.chkAbandonable.Value = quest(EditorIndex).Abandonable
            frmEditor_Quest.chkRepeat.Value = quest(EditorIndex).Repeatable
            frmEditor_Quest.chkShowBeforeQuest.Value = quest(EditorIndex).QuestLogBefore
            frmEditor_Quest.chkShowAfterQuest.Value = quest(EditorIndex).QuestLogAfter
            frmEditor_Quest.txtBeforeQuest.Text = Trim$(quest(EditorIndex).QuestLogBeforeDesc)
            frmEditor_Quest.txtAfterQuest.Text = Trim$(quest(EditorIndex).QuestLogAfterDesc)
        
    End Select
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "QuestEditorInitPage", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub QuestEditorInitTask()
Dim i As Long

   On Error GoTo errorhandler
    frmEditor_Quest.fraGatherResources.Visible = False
    frmEditor_Quest.fraGotoMap.Visible = False
    frmEditor_Quest.fraKillPlayers.Visible = False
    frmEditor_Quest.fraAquireDeliverItem.Visible = False
    frmEditor_Quest.fraAquireItems.Visible = False
    frmEditor_Quest.fraTalkToEvent.Visible = False
    frmEditor_Quest.fraKillNpcs.Visible = False
    frmEditor_Quest.chkEndQuestOnCompletion.Visible = True
    frmEditor_Quest.lblCurTask.Caption = "Task: " & QuestEditorTask
    frmEditor_Quest.fraCurrentTask.Caption = "Task: " & QuestEditorTask
    
    frmEditor_Quest.cmbTaskType.ListIndex = quest(EditorIndex).Task(QuestEditorTask).type
    
    If quest(EditorIndex).Task(QuestEditorTask).type = 0 Then
        frmEditor_Quest.txtTaskDesc.Text = ""
        frmEditor_Quest.txtTaskDesc.Enabled = False
    Else
        frmEditor_Quest.txtTaskDesc.Text = Trim$(quest(EditorIndex).Task(QuestEditorTask).TaskDesc)
        frmEditor_Quest.txtTaskDesc.Enabled = True
    End If
    
    Select Case quest(EditorIndex).Task(QuestEditorTask).type
        Case 0
            frmEditor_Quest.chkEndQuestOnCompletion.Visible = False
        Case 1
            frmEditor_Quest.fraKillNpcs.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Value = quest(EditorIndex).Task(QuestEditorTask).EndOnCompletion
            For i = 1 To 4
                If quest(EditorIndex).Task(QuestEditorTask).data(i) > 0 Then
                    frmEditor_Quest.cmbKillNPC(i).ListIndex = quest(EditorIndex).Task(QuestEditorTask).data(i)
                    frmEditor_Quest.scrlKillNPCCount(i).Value = quest(EditorIndex).Task(QuestEditorTask).data(4 + i)
                Else
                    frmEditor_Quest.cmbKillNPC(i).ListIndex = 0
                    frmEditor_Quest.scrlKillNPCCount(i).Value = 1
                End If
            Next
        Case 2
            frmEditor_Quest.fraTalkToEvent.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Value = quest(EditorIndex).Task(QuestEditorTask).EndOnCompletion
            frmEditor_Quest.txtEventTask.Text = Trim$(quest(EditorIndex).Task(QuestEditorTask).Text(1))
        Case 3
            frmEditor_Quest.fraAquireItems.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Value = quest(EditorIndex).Task(QuestEditorTask).EndOnCompletion
            For i = 1 To 4
                If quest(EditorIndex).Task(QuestEditorTask).data(i) > 0 Then
                    frmEditor_Quest.cmbAquireItem(i).ListIndex = quest(EditorIndex).Task(QuestEditorTask).data(i)
                    frmEditor_Quest.scrlAquireItemVal(i).Value = quest(EditorIndex).Task(QuestEditorTask).data(4 + i)
                Else
                    frmEditor_Quest.cmbAquireItem(i).ListIndex = 0
                    frmEditor_Quest.scrlAquireItemVal(i).Value = 1
                End If
            Next
        Case 4
            frmEditor_Quest.fraAquireDeliverItem.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Value = quest(EditorIndex).Task(QuestEditorTask).EndOnCompletion
            For i = 1 To 4
                If quest(EditorIndex).Task(QuestEditorTask).data(i) > 0 Then
                    frmEditor_Quest.cmbAquireDeliverItem(i).ListIndex = quest(EditorIndex).Task(QuestEditorTask).data(i)
                    frmEditor_Quest.scrlAquireDeliverItemVal(i).Value = quest(EditorIndex).Task(QuestEditorTask).data(4 + i)
                Else
                    frmEditor_Quest.cmbAquireDeliverItem(i).ListIndex = 0
                    frmEditor_Quest.scrlAquireDeliverItemVal(i).Value = 1
                End If
            Next
            frmEditor_Quest.txtAquireDeliverEventName.Text = quest(EditorIndex).Task(QuestEditorTask).Text(1)
        Case 5
            frmEditor_Quest.fraKillPlayers.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Value = quest(EditorIndex).Task(QuestEditorTask).EndOnCompletion
            If quest(EditorIndex).Task(QuestEditorTask).data(1) > 0 Then
                frmEditor_Quest.scrlKillPlayer.Value = quest(EditorIndex).Task(QuestEditorTask).data(1)
            Else
                quest(EditorIndex).Task(QuestEditorTask).data(1) = 1
                frmEditor_Quest.scrlKillPlayer.Value = quest(EditorIndex).Task(QuestEditorTask).data(1)
            End If
        Case 6
            frmEditor_Quest.fraGotoMap.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Value = quest(EditorIndex).Task(QuestEditorTask).EndOnCompletion
            If quest(EditorIndex).Task(QuestEditorTask).data(1) > 0 And quest(EditorIndex).Task(QuestEditorTask).data(1) <= MAX_MAPS Then
                frmEditor_Quest.scrlGotoMap.Value = quest(EditorIndex).Task(QuestEditorTask).data(1)
            Else
                quest(EditorIndex).Task(QuestEditorTask).data(1) = 1
                frmEditor_Quest.scrlGotoMap.Value = quest(EditorIndex).Task(QuestEditorTask).data(1)
            End If
            frmEditor_Quest.txtGoToMap.Text = Trim$(quest(EditorIndex).Task(QuestEditorTask).Text(1))
        Case 7
            frmEditor_Quest.fraGatherResources.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Visible = True
            frmEditor_Quest.chkEndQuestOnCompletion.Value = quest(EditorIndex).Task(QuestEditorTask).EndOnCompletion
            For i = 1 To 4
                If quest(EditorIndex).Task(QuestEditorTask).data(i) > 0 Then
                    frmEditor_Quest.cmbGatherResource(i).ListIndex = quest(EditorIndex).Task(QuestEditorTask).data(i)
                    frmEditor_Quest.scrlGatherResourceAmount(i).Value = quest(EditorIndex).Task(QuestEditorTask).data(4 + i)
                Else
                    frmEditor_Quest.cmbGatherResource(i).ListIndex = 0
                    frmEditor_Quest.scrlGatherResourceAmount(i).Value = 1
                End If
            Next
    End Select
    
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "QuestEditorInitTask", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearQuestTask(questnum As Long, tasknum As Long)
Dim i As Long, X As Long


   On Error GoTo errorhandler
    With quest(questnum).Task(tasknum)
        .type = 0
        For i = 1 To 8
            .data(i) = 0
        Next
        For i = 1 To 4
            .Text(i) = ""
        Next
    End With
   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearQuestTask", "modQuests", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
