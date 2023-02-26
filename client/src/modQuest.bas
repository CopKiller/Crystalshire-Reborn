Attribute VB_Name = "modQuests"
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////
Option Explicit

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Byte = 70
Public Const MAX_QUESTS_ITEMS As Byte = 10    'Alatar v1.2

Public Const QUEST_TYPE_GOSLAY As Byte = 1
Public Const QUEST_TYPE_GOGATHER As Byte = 2
Public Const QUEST_TYPE_GOTALK As Byte = 3
Public Const QUEST_TYPE_GOREACH As Byte = 4
Public Const QUEST_TYPE_GOGIVE As Byte = 5
Public Const QUEST_TYPE_GOKILL As Byte = 6
Public Const QUEST_TYPE_GOTRAIN As Byte = 7
Public Const QUEST_TYPE_GOGET As Byte = 8

Public Const QUEST_NOT_STARTED As Byte = 0
Public Const QUEST_STARTED As Byte = 1
Public Const QUEST_COMPLETED As Byte = 2
Public Const QUEST_COMPLETED_BUT As Byte = 3
Public Const QUEST_COMPLETED_DIARY As Byte = 4
Public Const QUEST_COMPLETED_TIME As Byte = 5

Public Quest_Changed(1 To MAX_QUESTS) As Boolean

'Types
Public Quest(1 To MAX_QUESTS) As QuestRec

'Alatar v1.2
Private Type QuestRequiredItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestGiveItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestTakeItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestRewardItemRec
    Item As Long
    Value As Long
End Type
'/Alatar v1.2

Public Type TaskTimerRec
    Active As Byte            ' Is Active?
    TimerType As Byte         ' 0=Days; 1=Hours; 2=Minutes; 3=Seconds.
    Timer As Currency             ' Time with /\

    Teleport As Byte          ' Teleport cannot end task in time.
    MapNum As Integer         ' Map Number to teleport /\
    ResetType As Byte         ' 0=Resetar Task ; 1=Resetar Quest.
    x As Byte
    Y As Byte
    
    Msg As String * TASK_DEFEAT_LENGTH
End Type

Public Type TaskRec
    Order As Byte
    NPC As Integer
    Item As Integer
    Map As Integer
    Resource As Integer
    Amount As Long
    TaskLog As String * 150
    QuestEnd As Boolean
    
    ' Task Timer
    TaskTimer As TaskTimerRec
End Type

Public Type QuestRec
    'Alatar v1.2
    Name As String * NAME_LENGTH
    Repeat As Byte
    Time As Long
    QuestLog As String * 100
    Speech As String * 200
    GiveItem(1 To MAX_QUESTS_ITEMS) As QuestGiveItemRec
    TakeItem(1 To MAX_QUESTS_ITEMS) As QuestTakeItemRec

    RequiredLevel As Integer
    RequiredQuest As Integer
    RequiredClass(1 To 5) As Integer
    RequiredItem(1 To MAX_QUESTS_ITEMS) As QuestRequiredItemRec

    RewardExp As Long
    RewardLevel As Integer
    RewardSpell As Integer
    RewardItem(1 To MAX_QUESTS_ITEMS) As QuestRewardItemRec

    Task(1 To MAX_TASKS) As TaskRec
    '/Alatar v1.2

End Type

' ////////////
' // Editor //
' ////////////

Public Sub QuestEditorInit()
    Dim i As Long

    If frmEditor_Quest.visible = False Then Exit Sub
    EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1

    With frmEditor_Quest
        'Alatar v1.2
        .txtName = Trim$(Quest(EditorIndex).Name)
        
        .optRepeat(Quest(EditorIndex).Repeat).Value = True
        .txtSegs = Quest(EditorIndex).Time
        
        .txtQuestLog = Trim$(Quest(EditorIndex).QuestLog)
        .txtSpeech.Text = Trim$(Quest(EditorIndex).Speech)

        .scrlReqLevel.Value = Quest(EditorIndex).RequiredLevel
        .scrlReqQuest.Value = Quest(EditorIndex).RequiredQuest
        For i = 1 To 5
            .scrlReqClass.Value = Quest(EditorIndex).RequiredClass(i)
        Next

        .txtExp.Text = Quest(EditorIndex).RewardExp
        .txtLevel.Text = Quest(EditorIndex).RewardLevel

        'Update the lists
        UpdateQuestGiveItems
        UpdateQuestTakeItems
        UpdateQuestRewardItems
        UpdateQuestRequirementItems
        UpdateQuestClass

        '/Alatar v1.2

        'load task nº1
        .scrlTotalTasks.Value = 1
        LoadTask EditorIndex, 1

    End With

    Quest_Changed(EditorIndex) = True

End Sub

'Alatar v1.2
Public Sub UpdateQuestGiveItems()
    Dim i As Long

    frmEditor_Quest.lstGiveItem.Clear

    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).GiveItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstGiveItem.AddItem "-"
            Else
                frmEditor_Quest.lstGiveItem.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestTakeItems()
    Dim i As Long

    frmEditor_Quest.lstTakeItem.Clear

    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).TakeItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstTakeItem.AddItem "-"
            Else
                frmEditor_Quest.lstTakeItem.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRewardItems()
    Dim i As Long

    frmEditor_Quest.lstItemRew.Clear

    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).RewardItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstItemRew.AddItem "-"
            Else
                frmEditor_Quest.lstItemRew.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRequirementItems()
    Dim i As Long

    frmEditor_Quest.lstReqItem.Clear

    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).RequiredItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstReqItem.AddItem "-"
            Else
                frmEditor_Quest.lstReqItem.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestClass()
    Dim i As Long

    frmEditor_Quest.lstReqClass.Clear

    For i = 1 To 5
        If Quest(EditorIndex).RequiredClass(i) = 0 Then
            frmEditor_Quest.lstReqClass.AddItem "-"
        Else
            frmEditor_Quest.lstReqClass.AddItem Trim$(Trim$(Class(Quest(EditorIndex).RequiredClass(i)).Name))
        End If
    Next
End Sub
'/Alatar v1.2

Public Sub QuestEditorOk()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If Quest_Changed(i) Then
            Call SendSaveQuest(i)
        End If
    Next

    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest

End Sub

Public Sub QuestEditorCancel()
    Editor = 0
    Unload frmEditor_Quest
    ClearChanged_Quest
    ClearQuests
    SendRequestQuests
End Sub

Public Sub ClearChanged_Quest()
    ZeroMemory Quest_Changed(1), MAX_QUESTS * 2    ' 2 = boolean length
End Sub

' //////////////
' // DATABASE //
' //////////////

Sub ClearQuest(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = vbNullString
    Quest(Index).QuestLog = vbNullString
    Quest(Index).Speech = vbNullString
End Sub

Sub ClearQuests()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next
End Sub

' ////////////////////
' // C&S PROCEDURES //
' ////////////////////

Public Sub SendRequestEditQuest()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditQuest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Public Sub SendSaveQuest(ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong CSaveQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendRequestQuests()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestQuests
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Public Sub UpdateQuestLog()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuestLogUpdate
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Public Sub PlayerCancelQuest()
Dim QuestName As String
    
    With Windows(GetWindowIndex("winQuest"))
    
    If QuestSelect = 0 Then Exit Sub
    
    If Not .Controls(GetControlIndex("winQuest", "lblList" & QuestSelect)).visible Then Exit Sub
    
    QuestName = .Controls(GetControlIndex("winQuest", "lblList" & QuestSelect)).Text

    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong CPlayerHandleQuest
    Buffer.WriteLong FindQuestIndex(QuestName)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    End With
End Sub

' ///////////////
' // Functions //
' ///////////////

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal QuestNum As Long) As Boolean
    QuestInProgress = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function

    If Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_STARTED Then    'Status=1 means started
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal QuestNum As Long) As Boolean
    QuestCompleted = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function

    If Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Or Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

'Subroutine that load the desired task in the form
Public Sub LoadTask(ByVal QuestNum As Long, ByVal TaskNum As Long)
    Dim TaskToLoad As TaskRec
    TaskToLoad = Quest(QuestNum).Task(TaskNum)

    With frmEditor_Quest
        'Load the task type
        .optTask(TaskToLoad.Order).Value = True
        'Load textboxes
        .txtTaskLog.Text = "" & Trim$(TaskToLoad.TaskLog)
        'Set scrolls to 0 and disable them so they can be enabled when needed
        .scrlNPC.Value = 0
        .scrlItem.Value = 0
        .scrlMap.Value = 0
        .scrlResource.Value = 0
        .scrlAmount.Value = 0
        .scrlNPC.enabled = False
        .scrlItem.enabled = False
        .scrlMap.enabled = False
        .scrlResource.enabled = False
        .scrlAmount.enabled = False
        
        ' Quest Timer
        .chkTaskTimer.Value = TaskToLoad.TaskTimer.Active
        .optTaskTimer(TaskToLoad.TaskTimer.TimerType).Value = True
        .txtTaskTimer.Text = CLng(TaskToLoad.TaskTimer.Timer)
        .chkTaskTeleport = TaskToLoad.TaskTimer.Teleport
        .txtTaskTeleport.Text = CInt(TaskToLoad.TaskTimer.Teleport)
        .optReset(TaskToLoad.TaskTimer.ResetType).Value = True
        .txtTaskTeleport = CInt(TaskToLoad.TaskTimer.MapNum)
        .txtTaskX.Text = CByte(TaskToLoad.TaskTimer.x)
        .txtTaskY.Text = CByte(TaskToLoad.TaskTimer.Y)
        .txtMsg.Text = Trim$(CStr(TaskToLoad.TaskTimer.Msg))
        
        If TaskToLoad.QuestEnd = True Then
            .chkEnd.Value = 1
        Else
            .chkEnd.Value = 0
        End If

        Select Case TaskToLoad.Order
        Case 0    'Nothing

        Case QUEST_TYPE_GOSLAY    '1
            .scrlNPC.enabled = True
            .scrlNPC.Value = TaskToLoad.NPC
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount

        Case QUEST_TYPE_GOGATHER    '2
            .scrlItem.enabled = True
            .scrlItem.Value = TaskToLoad.Item
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount

        Case QUEST_TYPE_GOTALK    '3
            .scrlNPC.enabled = True
            .scrlNPC.Value = TaskToLoad.NPC

        Case QUEST_TYPE_GOREACH    '4
            .scrlMap.enabled = True
            .scrlMap.Value = TaskToLoad.Map

        Case QUEST_TYPE_GOGIVE    '5
            .scrlItem.enabled = True
            .scrlItem.Value = TaskToLoad.Item
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount
            .scrlNPC.enabled = True
            .scrlNPC.Value = TaskToLoad.NPC

        Case QUEST_TYPE_GOKILL    '6
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount

        Case QUEST_TYPE_GOTRAIN    '7
            .scrlResource.enabled = True
            .scrlResource.Value = TaskToLoad.Resource
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount

        Case QUEST_TYPE_GOGET    '8
            .scrlNPC.enabled = True
            .scrlNPC.Value = TaskToLoad.NPC
            .scrlItem.enabled = True
            .scrlItem.Value = TaskToLoad.Item
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount

        End Select
    End With
End Sub
