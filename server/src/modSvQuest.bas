Attribute VB_Name = "modSvQuest"
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////
Option Explicit
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Integer = 70
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

Private Enum TaskType
    Day = 0
    Hour
    Minutes
    Seconds
End Enum

' //////////////
' // DATABASE //
' //////////////

Sub SaveQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next
End Sub

Sub SaveQuest(ByVal QuestNum As Long)
    Dim FileName As String
    Dim F As Long, i As Long
    FileName = App.Path & "\data\quests\quest" & QuestNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    'Alatar v1.2
    Put #F, , Quest(QuestNum).Name
    Put #F, , Quest(QuestNum).Repeat
    Put #F, , Quest(QuestNum).QuestLog
    Put #F, , Quest(QuestNum).Speech
    For i = 1 To MAX_QUESTS_ITEMS
        Put #F, , Quest(QuestNum).GiveItem(i)
    Next
    For i = 1 To MAX_QUESTS_ITEMS
        Put #F, , Quest(QuestNum).TakeItem(i)
    Next
    Put #F, , Quest(QuestNum).RequiredLevel
    Put #F, , Quest(QuestNum).RequiredQuest
    For i = 1 To 5
        Put #F, , Quest(QuestNum).RequiredClass(i)
    Next
    For i = 1 To MAX_QUESTS_ITEMS
        Put #F, , Quest(QuestNum).RequiredItem(i)
    Next
    Put #F, , Quest(QuestNum).RewardExp
    For i = 1 To MAX_QUESTS_ITEMS
        Put #F, , Quest(QuestNum).RewardItem(i)
    Next
    For i = 1 To MAX_TASKS
        Put #F, , Quest(QuestNum).Task(i)
    Next
    '/Alatar v1.2
    Close #F
End Sub

Sub LoadQuests()
    Dim FileName As String
    Dim i As Integer
    Dim F As Long, n As Long
    Dim sLen As Long

    Call CheckQuests

    For i = 1 To MAX_QUESTS
        ' Clear
        Call ClearQuest(i)
        'Load
        FileName = App.Path & "\data\quests\quest" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F

        'Alatar v1.2
        Get #F, , Quest(i).Name
        Get #F, , Quest(i).Repeat
        Get #F, , Quest(i).QuestLog
        Get #F, , Quest(i).Speech
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).GiveItem(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).TakeItem(n)
        Next
        Get #F, , Quest(i).RequiredLevel
        Get #F, , Quest(i).RequiredQuest
        For n = 1 To 5
            Get #F, , Quest(i).RequiredClass(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).RequiredItem(n)
        Next
        Get #F, , Quest(i).RewardExp
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(i).RewardItem(n)
        Next
        For n = 1 To MAX_TASKS
            Get #F, , Quest(i).Task(n)
        Next
        '/Alatar v1.2
        Close #F
    Next
End Sub

Sub CheckQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        If Not FileExist("\Data\quests\quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If
    Next
End Sub

Sub ClearQuest(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = vbNullString
    Quest(Index).QuestLog = vbNullString
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

Sub SendQuests(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(Index, i)
        End If
    Next
End Sub

Public Sub SendPlayerQuests(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest

    For i = 1 To MAX_QUESTS

        If Player(Index).PlayerQuest(i).Status > 0 Then
            Buffer.WriteLong i
            Buffer.WriteLong Player(Index).PlayerQuest(i).Status
            Buffer.WriteLong Player(Index).PlayerQuest(i).ActualTask
            Buffer.WriteLong Player(Index).PlayerQuest(i).CurrentCount


            Buffer.WriteByte Player(Index).PlayerQuest(i).TaskTimer.Active
            Buffer.WriteLong Player(Index).PlayerQuest(i).TaskTimer.Timer
        End If
    Next

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing

End Sub

Private Sub SendPlayerQuest(ByVal Index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest

    Buffer.WriteLong QuestNum
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).Status
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).ActualTask
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).CurrentCount

    Buffer.WriteByte Player(Index).PlayerQuest(QuestNum).TaskTimer.Active
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).TaskTimer.Timer

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub SendQuestCancel(ByVal Index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestCancel

    Buffer.WriteLong QuestNum
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).Status
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).ActualTask
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).CurrentCount

    Buffer.WriteByte Player(Index).PlayerQuest(QuestNum).TaskTimer.Active
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).TaskTimer.Timer

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

'sends a message to the client that is shown on the screen
Public Sub QuestMessage(ByVal Index As Long, ByVal QuestNum As Long, ByVal message As String, Optional ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SQuestMessage
    Buffer.WriteLong QuestNum
    Buffer.WriteString Trim$(message)
    Buffer.WriteLong saycolour
    Buffer.WriteString "[Quest] "
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

' ///////////////
' // Functions //
' ///////////////

Public Function CanStartQuest(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    Dim i As Long, n As Long
    CanStartQuest = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function

    If QuestInProgress(Index, QuestNum) Then
        Call QuestMessage(Index, QuestNum, "Você ja iniciou a quest, precisa termina-la!", BrightRed)
        Exit Function
    End If

    'check if now a completed quest can be repeated
    Select Case Player(Index).PlayerQuest(QuestNum).Status
    Case QUEST_COMPLETED    ' Normal?
        If Quest(QuestNum).Repeat = 1 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
        ElseIf Quest(QuestNum).Repeat = 2 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        ElseIf Quest(QuestNum).Repeat = 3 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        End If
    Case QUEST_COMPLETED_BUT    ' Repetível?
        If Quest(QuestNum).Repeat = 0 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        ElseIf Quest(QuestNum).Repeat = 2 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        ElseIf Quest(QuestNum).Repeat = 3 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        End If
    Case QUEST_COMPLETED_DIARY    ' Diaria?
        If Quest(QuestNum).Repeat = 0 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        ElseIf Quest(QuestNum).Repeat = 1 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
        ElseIf Quest(QuestNum).Repeat = 3 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        End If
    Case QUEST_COMPLETED_TIME    ' Tempo pra refazer?
        If Quest(QuestNum).Repeat = 0 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        ElseIf Quest(QuestNum).Repeat = 1 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
        ElseIf Quest(QuestNum).Repeat = 2 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        End If
    End Select

    ' Fazer o processamento da quest diaria e quest por tempo!
    Select Case Player(Index).PlayerQuest(QuestNum).Status
    Case QUEST_COMPLETED_DIARY
        If Format(Player(Index).PlayerQuest(QuestNum).Data, "dd/mm/yyyy") <> CStr(Date) Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        Else
            PlayerMsg Index, "Você ja realizou essa missão hoje, volte novamente amanhã!", BrightRed
            Exit Function
        End If
    Case QUEST_COMPLETED_TIME
        If DateDiff("s", Player(Index).PlayerQuest(QuestNum).Data, Now) >= Quest(QuestNum).Time Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        Else
            PlayerMsg Index, "Aguarde: " & SecondsToHMS(Quest(QuestNum).Time - DateDiff("s", Player(Index).PlayerQuest(QuestNum).Data, Now)), BrightRed
            Exit Function
        End If
    End Select

    'Check if player has the quest 0 (not started) or 3 (completed but it can be started again)
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED Or Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        'Check if player's level is right
        If Quest(QuestNum).RequiredLevel <= Player(Index).Level Then

            'Check if item is needed
            For i = 1 To MAX_QUESTS_ITEMS
                If Quest(QuestNum).RequiredItem(i).Item > 0 Then
                    'if we don't have it at all then
                    If HasItem(Index, Quest(QuestNum).RequiredItem(i).Item) = 0 Then
                        PlayerMsg Index, "You need " & Trim$(Item(Quest(QuestNum).RequiredItem(i).Item).Name) & " to take this quest!", BrightRed
                        Exit Function
                    End If
                End If
            Next

            'Check if previous quest is needed
            If Quest(QuestNum).RequiredQuest > 0 And Quest(QuestNum).RequiredQuest <= MAX_QUESTS Then
                If Player(Index).PlayerQuest(Quest(QuestNum).RequiredQuest).Status = QUEST_NOT_STARTED Or Player(Index).PlayerQuest(Quest(QuestNum).RequiredQuest).Status = QUEST_STARTED Then
                    PlayerMsg Index, "You need to complete the " & Trim$(Quest(Quest(QuestNum).RequiredQuest).Name) & " quest in order to take this quest!", BrightRed
                    Exit Function
                End If
            End If
            'Go on :)
            CanStartQuest = True
        Else
            PlayerMsg Index, "You need to be a higher level to take this quest!", BrightRed
        End If
    Else
        PlayerMsg Index, "You can't start that quest again!", BrightRed
    End If
End Function

Public Function CanEndQuest(ByVal Index As Long, QuestNum As Long) As Boolean
    CanEndQuest = False
    If Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).QuestEnd = True Then
        CanEndQuest = True
    End If
End Function

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    QuestInProgress = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function

    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_STARTED Then
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    QuestCompleted = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function

    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Or Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

'Gets the quest reference num (id) from the quest name (it shall be unique)
Public Function GetQuestNum(ByVal QuestName As String) As Long
    Dim i As Long
    GetQuestNum = 0

    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) = Trim$(QuestName) Then
            GetQuestNum = i
            Exit For
        End If
    Next
End Function

Public Function GetItemNum(ByVal ItemName As String) As Long
    Dim i As Long
    GetItemNum = 0

    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) = Trim$(ItemName) Then
            GetItemNum = i
            Exit For
        End If
    Next
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

Public Sub CheckTasks(ByVal Index As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If QuestInProgress(Index, i) Then
            If TaskType = Quest(i).Task(Player(Index).PlayerQuest(i).ActualTask).Order Then
                Call CheckTask(Index, i, TaskType, TargetIndex)
            End If
        End If
    Next
End Sub

Public Sub CheckTask(ByVal Index As Long, ByVal QuestNum As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim ActualTask As Long, i As Long
    ActualTask = Player(Index).PlayerQuest(QuestNum).ActualTask

    Select Case TaskType
    Case QUEST_TYPE_GOSLAY    'Kill X amount of X npc's.

        'is npc's defeated id is the same as the npc i have to kill?
        If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
            'Count +1
            Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
            'show msg
            QuestMessage Index, QuestNum, Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(NPC(TargetIndex).Name) + " killed.", Yellow
            'did i finish the work?
            If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                QuestMessage Index, QuestNum, "Task completed", LightGreen
                'is the quest's end?
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    'otherwise continue to the next task
                    Call ResetPlayerTaskTimer(Index, QuestNum)
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    Call SetPlayerTaskTimer(Index, QuestNum)
                    'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                    SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
                End If
            End If
        End If

    Case QUEST_TYPE_GOGATHER    'Gather X amount of X item.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).Item Then

            'reset the count first
            Player(Index).PlayerQuest(QuestNum).CurrentCount = 0

            'Check inventory for the items
            For i = 1 To MAX_INV
                If GetPlayerInvItemNum(Index, i) = TargetIndex Then
                    If Item(i).Stackable > 0 Then
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = GetPlayerInvItemValue(Index, i)
                    Else
                        'If is the correct item add it to the count
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                    End If
                End If
            Next

            QuestMessage Index, QuestNum, "You have " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Item(TargetIndex).Name), Yellow

            If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                QuestMessage Index, QuestNum, "Task completed", LightGreen
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Call ResetPlayerTaskTimer(Index, QuestNum)
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    Call SetPlayerTaskTimer(Index, QuestNum)
                    'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                    SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
                End If
            End If
        End If

    Case QUEST_TYPE_GOTALK    'Interact with X npc.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
            QuestMessage Index, QuestNum, "Task completed", LightGreen
            If CanEndQuest(Index, QuestNum) Then
                EndQuest Index, QuestNum
            Else
                Call ResetPlayerTaskTimer(Index, QuestNum)
                Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                Call SetPlayerTaskTimer(Index, QuestNum)
                'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
            End If
        End If

    Case QUEST_TYPE_GOREACH    'Reach X map.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).Map Then
            QuestMessage Index, QuestNum, "Task completed", LightGreen
            If CanEndQuest(Index, QuestNum) Then
                EndQuest Index, QuestNum
            Else

                Call ResetPlayerTaskTimer(Index, QuestNum)
                Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                Call SetPlayerTaskTimer(Index, QuestNum)
                'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
            End If
        End If

    Case QUEST_TYPE_GOGIVE    'Give X amount of X item to X npc.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then

            Player(Index).PlayerQuest(QuestNum).CurrentCount = 0

            For i = 1 To MAX_INV
                If GetPlayerInvItemNum(Index, i) = Quest(QuestNum).Task(ActualTask).Item Then
                    If Item(i).Stackable > 0 Then
                        If GetPlayerInvItemValue(Index, i) >= Quest(QuestNum).Task(ActualTask).Amount Then
                            Player(Index).PlayerQuest(QuestNum).CurrentCount = GetPlayerInvItemValue(Index, i)
                        End If
                    Else
                        'If is the correct item add it to the count
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                    End If
                End If
            Next

            If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                'if we have enough items, then remove them and finish the task
                If Item(Quest(QuestNum).Task(ActualTask).Item).Stackable > 0 Then
                    TakeInvItem Index, Quest(QuestNum).Task(ActualTask).Item, Quest(QuestNum).Task(ActualTask).Amount
                Else
                    'If it's not a currency then remove all the items
                    For i = 1 To Quest(QuestNum).Task(ActualTask).Amount
                        TakeInvItem Index, Quest(QuestNum).Task(ActualTask).Item, 1
                    Next
                End If

                QuestMessage Index, QuestNum, "You gave " + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Item(TargetIndex).Name), Yellow
                QuestMessage Index, QuestNum, "Task completed", LightGreen

                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Call ResetPlayerTaskTimer(Index, QuestNum)
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    Call SetPlayerTaskTimer(Index, QuestNum)
                    'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                    SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
                End If
            End If
        End If

    Case QUEST_TYPE_GOKILL    'Kill X amount of players.
        Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
        QuestMessage Index, QuestNum, Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " players killed.", Yellow
        If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
            QuestMessage Index, QuestNum, "Task completed", LightGreen
            If CanEndQuest(Index, QuestNum) Then
                EndQuest Index, QuestNum
            Else
                Call ResetPlayerTaskTimer(Index, QuestNum)
                Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                Call SetPlayerTaskTimer(Index, QuestNum)
                'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
            End If
        End If

    Case QUEST_TYPE_GOTRAIN    'Hit X amount of times X resource.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).Resource Then
            Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
            QuestMessage Index, QuestNum, Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " hits.", Yellow
            If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                QuestMessage Index, QuestNum, "Task completed", LightGreen
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Call ResetPlayerTaskTimer(Index, QuestNum)
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    Call SetPlayerTaskTimer(Index, QuestNum)
                    'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                    SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
                End If
            End If
        End If

    Case QUEST_TYPE_GOGET    'Get X amount of X item from X npc.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
            If GiveInvItem(Index, Quest(QuestNum).Task(ActualTask).Item, Quest(QuestNum).Task(ActualTask).Amount, 0) Then
                QuestMessage Index, QuestNum, Quest(QuestNum).Task(ActualTask).TaskLog, Yellow
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else

                    Call ResetPlayerTaskTimer(Index, QuestNum)
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    Call SetPlayerTaskTimer(Index, QuestNum)
                    'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                    SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
                End If
            End If
        End If

    End Select
    SavePlayer Index
    SendPlayerQuest Index, QuestNum
End Sub

Public Sub EndQuest(ByVal Index As Long, ByVal QuestNum As Long)
    Dim i As Long, n As Long

    ' Reseta os dados da data pra ser somente usado onde necessitar!
    Player(Index).PlayerQuest(QuestNum).Data = vbNullString

    If Quest(QuestNum).Repeat = 0 Then    ' Normal?
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED
    ElseIf Quest(QuestNum).Repeat = 1 Then    ' Repetível?
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
    ElseIf Quest(QuestNum).Repeat = 2 Then    ' Diaria?
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_DIARY
        Player(Index).PlayerQuest(QuestNum).Data = Now
    ElseIf Quest(QuestNum).Repeat = 3 Then    ' Tempo pra refazer?
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_TIME
        Player(Index).PlayerQuest(QuestNum).Data = Now
    End If

    'reset counters to 0
    Call ResetPlayerTaskTimer(Index, QuestNum)
    Player(Index).PlayerQuest(QuestNum).ActualTask = 0
    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0

    'give experience
    GivePlayerEXP Index, Quest(QuestNum).RewardExp

    'give levels
    If Quest(QuestNum).RewardLevel > 0 Then
        CheckPlayerLevelUp Index, Quest(QuestNum).RewardLevel
    End If

    'remove items on the end
    For i = 1 To MAX_QUESTS_ITEMS
        If Quest(QuestNum).TakeItem(i).Item > 0 Then
            If HasItem(Index, Quest(QuestNum).TakeItem(i).Item) > 0 Then
                If Item(Quest(QuestNum).TakeItem(i).Item).Stackable > 0 Then
                    TakeInvItem Index, Quest(QuestNum).TakeItem(i).Item, Quest(QuestNum).TakeItem(i).Value
                Else
                    For n = 1 To Quest(QuestNum).TakeItem(i).Value
                        TakeInvItem Index, Quest(QuestNum).TakeItem(i).Item, 1
                    Next
                End If
            End If
        End If
    Next

    'give rewards
    For i = 1 To MAX_QUESTS_ITEMS
        If Quest(QuestNum).RewardItem(i).Item <> 0 Then
            'check if we have space
            If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(i).Item) = 0 Then
                PlayerMsg Index, "You have no inventory space.", BrightRed
                Exit For
            Else
                'if so, check if it's a currency stack the item in one slot
                If Item(Quest(QuestNum).RewardItem(i).Item).Stackable > 0 Then
                    GiveInvItem Index, Quest(QuestNum).RewardItem(i).Item, Quest(QuestNum).RewardItem(i).Value, 0
                Else
                    'if not, create a new loop and store the item in a new slot if is possible
                    For n = 1 To Quest(QuestNum).RewardItem(i).Value
                        If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(i).Item) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit For
                        Else
                            GiveInvItem Index, Quest(QuestNum).RewardItem(i).Item, 1, 0
                        End If
                    Next
                End If
            End If
        End If
    Next

    ' Give Spell Reward
    If Quest(QuestNum).RewardSpell > 0 Then
        Call GivePlayerSpell(Index, Quest(QuestNum).RewardSpell)
    End If

    'show ending message
    'QuestMessage Index, QuestNum, "Parabens, Você concluiu a missão!", LightGreen
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_DIARY Then
        SendMessageTo Index, Trim$(Quest(QuestNum).Name), "Parabens, Voce concluiu a missao, volte amanha para completar novamente!"
    ElseIf Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_TIME Then
        SendMessageTo Index, Trim$(Quest(QuestNum).Name), "Parabens, Voce concluiu a missao, volte daqui: " & SecondsToHMS(Quest(QuestNum).Time) & " e complete novamente!"
    End If

    SavePlayer Index
    SendEXP Index
    SendStats Index
    SendPlayerQuest Index, QuestNum
End Sub

Sub HandleRequestEditQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleSaveQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong    'CLng(Parse(1))

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
    Call QuestCache_Create(n)
    Call SendQuestAll(n)
    Call SaveQuest(n)
    Call AddLog(GetPlayerName(Index) & " saved Quest #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestQuests(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendQuests Index
End Sub

Sub HandlePlayerCancelQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim QuestNum As Long, i As Long, n As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    QuestNum = Buffer.ReadLong

    Call ResetPlayerTaskTimer(Index, QuestNum)
    Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED    '2
    Player(Index).PlayerQuest(QuestNum).ActualTask = 1
    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0

    PlayerMsg Index, Trim$(Quest(QuestNum).Name) & " has been canceled!", BrightGreen
    For i = 1 To MAX_QUESTS_ITEMS
        If Quest(QuestNum).GiveItem(i).Item > 0 Then
            If HasItem(Index, Quest(QuestNum).GiveItem(i).Item) > 0 Then
                If Item(Quest(QuestNum).GiveItem(i).Item).Stackable > 0 Then
                    TakeInvItem Index, Quest(QuestNum).GiveItem(i).Item, Quest(QuestNum).GiveItem(i).Value
                Else
                    For n = 1 To Quest(QuestNum).GiveItem(i).Value
                        TakeInvItem Index, Quest(QuestNum).GiveItem(i).Item, 1
                    Next
                End If
            End If
        End If
    Next

    SavePlayer Index
    SendQuestCancel Index, QuestNum

    Set Buffer = Nothing
End Sub

Sub HandleQuestLogUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerQuests Index
End Sub

Public Sub StartQuest(ByVal Index As Long, ByVal QuestNum As Long, ByVal Order As Byte)
    Dim i As Long, n As Long
    Dim RemoveStartItems As Boolean

    If Order = 1 Then    'Iniciar
        RemoveStartItems = False
        For i = 1 To MAX_QUESTS_ITEMS

            If Quest(QuestNum).RewardItem(i).Item > 0 Then
                If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(i).Item) = 0 Then
                    QuestMessage Index, QuestNum, "Você não tem espaço na mochila, drope algo para pegar a quest.", Red
                    Exit For
                End If
            End If

            If Quest(QuestNum).GiveItem(i).Item > 0 Then
                If FindOpenInvSlot(Index, Quest(QuestNum).GiveItem(i).Item) = 0 Then
                    QuestMessage Index, QuestNum, "Você não tem espaço na mochila, drope algo para pegar a quest.", Red
                    RemoveStartItems = True
                    Exit For
                Else
                    If Item(Quest(QuestNum).GiveItem(i).Item).Stackable > 0 Then
                        GiveInvItem Index, Quest(QuestNum).GiveItem(i).Item, Quest(QuestNum).GiveItem(i).Value, 0
                    Else
                        GiveInvItem Index, Quest(QuestNum).GiveItem(i).Item, 1, 0
                    End If
                End If
            End If


        Next

        If RemoveStartItems = False Then    'this means everything went ok
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_STARTED    '1
            Player(Index).PlayerQuest(QuestNum).ActualTask = 1
            Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
            QuestMessage Index, QuestNum, "Nova missão aceita, olhe seu QuestLog!", BrightGreen

            Call SetPlayerTaskTimer(Index, QuestNum)
        End If

    ElseIf Order = 2 Then
        Call ResetPlayerTaskTimer(Index, QuestNum)
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED    '2
        Player(Index).PlayerQuest(QuestNum).ActualTask = 1
        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0

        RemoveStartItems = True    'avoid exploits
        QuestMessage Index, QuestNum, " foi cancelada!", Yellow
    End If

    If RemoveStartItems = True Then
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).GiveItem(i).Item > 0 Then
                If HasItem(Index, Quest(QuestNum).GiveItem(i).Item) > 0 Then
                    If Item(Quest(QuestNum).GiveItem(i).Item).Stackable > 0 Then
                        TakeInvItem Index, Quest(QuestNum).GiveItem(i).Item, Quest(QuestNum).GiveItem(i).Value
                    Else
                        For n = 1 To Quest(QuestNum).GiveItem(i).Value
                            TakeInvItem Index, Quest(QuestNum).GiveItem(i).Item, 1
                        Next
                    End If
                End If
            End If
        Next
    End If

    SavePlayer Index
    SendPlayerQuest Index, QuestNum
End Sub

Public Sub ResetPlayerTaskTimer(ByVal Index As Long, ByVal QuestNum As Integer)
    Player(Index).PlayerQuest(QuestNum).TaskTimer.Active = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.MapNum = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.ResetType = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.Teleport = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.Timer = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.TimerType = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.x = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.Y = 0
End Sub

Public Sub SetPlayerTaskTimer(ByVal Index As Long, ByVal QuestNum As Integer)
    With Player(Index).PlayerQuest(QuestNum).TaskTimer
        .Active = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Active
        .Teleport = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Teleport
        .MapNum = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.MapNum
        .x = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.x
        .Y = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Y
        .ResetType = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.ResetType


        .TimerType = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.TimerType

        ' Converter o tipo de contador pelo menor pra ter um melhor processamento pelo loop
        If .TimerType = TaskType.Day Then
            .Timer = (((Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Timer * 24) * 60) * 60)
        ElseIf .TimerType = TaskType.Hour Then
            .Timer = ((Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Timer * 60) * 60)
        ElseIf .TimerType = TaskType.Minutes Then
            .Timer = (Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Timer * 60)
        Else    ' segundos já pré configurado no editor
            .Timer = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Timer
        End If
    End With
End Sub

Public Sub CheckPlayerTaskTimer(ByVal Index As Long)
    Dim i As Integer

    If IsPlaying(Index) Then
        For i = 1 To MAX_QUESTS
            If LenB(Trim$(Quest(i).Name)) > 0 Then
                With Player(Index).PlayerQuest(i).TaskTimer
                    If .Active = YES Then
                        If .Timer > 0 Then
                            .Timer = .Timer - 1
                        End If

                        If .Timer <= 0 Then
                            If .Teleport = YES Then
                                If .MapNum > 0 And .MapNum <= MAX_MAPS Then
                                    Call PlayerWarp(Index, .MapNum, .x, .Y)
                                Else
                                    Call PlayerWarp(Index, Class(GetPlayerClass(Index)).START_MAP, Class(GetPlayerClass(Index)).START_X, Class(GetPlayerClass(Index)).START_Y)
                                End If
                            End If

                            ' 0=Resetar Task ; 1=Resetar Quest.
                            If .ResetType = 0 Then
                                Player(Index).PlayerQuest(i).CurrentCount = 0    ' Retornar a zero a contagem do objetivo da task.
                                .Timer = Quest(i).Task(Player(Index).PlayerQuest(i).ActualTask).TaskTimer.Timer    ' Resetar o tempo que o jogador vai refazê-lá.
                            ElseIf .ResetType = 1 Then
                                Call ResetPlayerTaskTimer(Index, i)    ' Resetar todo os dados da task das variaveis do jogador!
                                Call StartQuest(Index, i, 2)    ' Cancelar a quest toda!
                            End If

                            ' enviar a mensagem do editor de task
                            If Trim$(Quest(i).Task(Player(Index).PlayerQuest(i).ActualTask).TaskTimer.Msg) <> vbNullString Then
                                Call SendMessageTo(Index, Trim$(Quest(i).Name), Trim$(Quest(i).Task(Player(Index).PlayerQuest(i).ActualTask).TaskTimer.Msg))
                            End If

                            Call SendPlayerQuest(Index, i)
                        End If
                    Else
                        If .Teleport = YES Then
                            If .MapNum > 0 And .MapNum <= MAX_MAPS Then
                                Call PlayerWarp(Index, .MapNum, .x, .Y)
                            Else
                                Call PlayerWarp(Index, Class(GetPlayerClass(Index)).START_MAP, Class(GetPlayerClass(Index)).START_X, Class(GetPlayerClass(Index)).START_Y)
                            End If
                            
                            Call ResetPlayerTaskTimer(Index, i)
                        End If
                    End If
                End With
            End If
        Next i
    End If

End Sub

Function SecondsToHMS(ByRef Segundos As Long) As String
    Dim HR As Long, ms As Long, SS As Long, MM As Long
    Dim Total As Long, Count As Long

    If Segundos = 0 Then
        SecondsToHMS = "0s "
        Exit Function
    End If

    HR = (Segundos \ 3600)
    MM = (Segundos \ 60)
    SS = Segundos
    'ms = (Segundos * 10)

    ' Pega o total de segundos pra trabalharmos melhor na variavel!
    Total = Segundos

    ' Verifica se tem mais de 1 hora em segundos!
    If HR > 0 Then
        '// Horas
        Do While (Total >= 3600)
            Total = Total - 3600
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = Count & "h "
            Count = 0
        End If
        '// Minutos
        Do While (Total >= 60)
            Total = Total - 60
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "m "
            Count = 0
        End If
        '// Segundos
        Do While (Total > 0)
            Total = Total - 1
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "s "
            Count = 0
        End If
    ElseIf MM > 0 Then
        '// Minutos
        Do While (Total >= 60)
            Total = Total - 60
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "m "
            Count = 0
        End If
        '// Segundos
        Do While (Total > 0)
            Total = Total - 1
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "s "
            Count = 0
        End If
    ElseIf SS > 0 Then
        ' Joga na função esse segundo.
        SecondsToHMS = SS & "s "
        Total = Total - SS
    End If
End Function




