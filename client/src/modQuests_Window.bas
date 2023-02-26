Attribute VB_Name = "modQuests_Window"
Option Explicit

Private Const QuestOffsetX As Long = 20
Private Const QuestOffsetY As Long = 10

Private Const ListOffsetY As Integer = 25
Private Const RewardOffsetX As Integer = 21
Private Const ListX As Integer = 20
Private Const ListY As Integer = 25

Private Const DescriptionX As Integer = 180
Private Const DescriptionY As Integer = 26

' Quantidade de quests mostradas na janela
Private Const MAX_QUESTS_WINDOW As Byte = 14

Private Const QuestMouseMoveColour = Brown
Private Const QuestMouseDownColour = DarkBrown
Private Const QuestDefaultColour = White

Public QuestSelect As Byte

Public QuestTimeToFinish As String
Public QuestNameToFinish As String

Public Sub CreateWindow_Quest()
    Dim i As Byte

    CreateWindow "winQuest", "Quests em Andamento...", zOrder_Win, 0, 0, 436, 414, Tex_Item(1), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, , , , , , , GetAddress(AddressOf lblList_ClearColour)
    ' Centralise it
    CentraliseWindow WindowCount
    ' Set the index for spawning controls
    zOrder_Con = 1

    'Close Btn
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Quest)

    ' Parchment
    CreatePictureBox WindowCount, "picList", ListX - 14, ListY + 1, 175, 380, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment, , , , GetAddress(AddressOf lblList_ClearColour)
    CreatePictureBox WindowCount, "picDescription", DescriptionX, DescriptionY, 250, 358, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment, , , , GetAddress(AddressOf lblList_ClearColour)

    ' Shadow
    CreatePictureBox WindowCount, "picShadow_1", ListX + 1, ListY + 10, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblQuestList", ListX - 14, ListY + 7, 175, 25, "Quest's Name", rockwellDec_15, White, Alignment.alignCentre

    ' Shadow descrição
    CreatePictureBox WindowCount, "picShadow_1", ListX + 215, ListY + 10, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblQuestDes", ListX - -200, ListY + 7, 175, 25, "Description", rockwellDec_15, White, Alignment.alignCentre
    CreatePictureBox WindowCount, "picBackground", ListX - -175, ListY + 25, 219, 124, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblQuestDescription1", ListX - -175, ListY + 25, 219, 124, "", rockwellDec_15, White, Alignment.alignCentre

    ' Shadow Objective
    CreatePictureBox WindowCount, "picShadow_1", ListX + 215, ListY + 158, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblQuestObj", ListX - -200, ListY + 155, 175, 25, "Objective", rockwellDec_15, Yellow, Alignment.alignCentre
    CreatePictureBox WindowCount, "picBackgroun2", ListX - -175, ListY + 175, 219, 78, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblQuestDescription2", ListX - -175, ListY + 175, 219, 124, "", rockwellDec_15, Yellow, Alignment.alignCentre

    ' Text Rewards
    CreatePictureBox WindowCount, "picShadow_1", ListX + 215, ListY + 260, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblQuestRew", ListX - -200, ListY + 257, 175, 25, "Rewards", rockwellDec_15, BrightGreen, Alignment.alignCentre
    CreatePictureBox WindowCount, "picBackground3", ListX - -175, ListY + 276, 219, 70, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblQuestDescription3", ListX - -175, ListY + 276, 219, 70, "", rockwellDec_15, BrightGreen, Alignment.alignCentre

    For i = 1 To MAX_QUESTS_ITEMS
        CreatePictureBox WindowCount, "picReward" & i, (ListX + 160) + (RewardOffsetX * i), ListY + 320, 20, 20, True, , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval, , GetAddress(AddressOf ShowRewardDesc), , GetAddress(AddressOf ShowRewardDesc), , GetAddress(AddressOf ShowRewardItem)
    Next i

    For i = 1 To MAX_QUESTS_WINDOW
        CreatePictureBox WindowCount, "picList" & i, ListX, ListY + (ListOffsetY * i), 130, 20, False, , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , 0
        CreateLabel WindowCount, "lblList" & i, ListX, ListY + (ListOffsetY * i) + 3, 130, 20, "Vazio", , , Alignment.alignCentre, False, , , , , GetAddress(AddressOf lblList_MouseDown), GetAddress(AddressOf lblList_MouseMove)
    Next i

    ' Btns
    CreateButton WindowCount, "btnCancel", 238, 385, 134, 20, "Cancel Quest", rockwellDec_15, White, , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf PlayerCancelQuest)
End Sub

Private Sub ShowRewardDesc()
    Dim x As Integer, Y As Integer, Width As Integer, Height As Integer, i As Integer
    Dim itemNum As Long
    With Windows(GetWindowIndex("winQuest"))
        For i = 1 To MAX_QUESTS_ITEMS
            If .Controls(GetControlIndex("winQuest", "picReward" & i)).visible Then
                x = .Window.Left + .Controls(GetControlIndex("winQuest", "picReward" & i)).Left
                Y = .Window.top + .Controls(GetControlIndex("winQuest", "picReward" & i)).top
                Width = .Controls(GetControlIndex("winQuest", "picReward" & i)).Width
                Height = .Controls(GetControlIndex("winQuest", "picReward" & i)).Height
                If GlobalX >= x And GlobalX <= x + Width And GlobalY >= Y And GlobalY <= Y + Height Then
                    itemNum = .Controls(GetControlIndex("winQuest", "picReward" & i)).Value
                    ShowItemDesc GlobalX, GlobalY, itemNum, False
                End If
            End If
        Next i
    End With
End Sub

Private Sub ShowRewardItem()
    Dim x As Integer, Y As Integer, Width As Integer, Height As Integer, i As Integer
    Dim itemNum As Long
    With Windows(GetWindowIndex("winQuest"))
        For i = 1 To MAX_QUESTS_ITEMS
            If .Controls(GetControlIndex("winQuest", "picReward" & i)).visible Then
                itemNum = .Controls(GetControlIndex("winQuest", "picReward" & i)).Value
                If itemNum > 0 And itemNum <= Count_Item Then
                    x = .Window.Left + .Controls(GetControlIndex("winQuest", "picReward" & i)).Left
                    Y = .Window.top + .Controls(GetControlIndex("winQuest", "picReward" & i)).top
                    Width = .Controls(GetControlIndex("winQuest", "picReward" & i)).Width
                    Height = .Controls(GetControlIndex("winQuest", "picReward" & i)).Height
                    RenderTexture Tex_Item(Item(itemNum).Pic), x, Y, 0, 0, Width, Height, PIC_X, PIC_Y
                End If
            End If
        Next i
    End With
End Sub

Private Sub lblList_MouseMove()
    Dim i As Byte, x As Long, Y As Long, Width As Long, Height As Long
    With Windows(GetWindowIndex("winQuest"))

        For i = 1 To MAX_QUESTS_WINDOW

            x = .Window.Left + .Controls(GetControlIndex("winQuest", "lblList" & i)).Left
            Y = .Window.top + .Controls(GetControlIndex("winQuest", "lblList" & i)).top
            Width = .Controls(GetControlIndex("winQuest", "lblList" & i)).Width
            Height = .Controls(GetControlIndex("winQuest", "lblList" & i)).Height


            If QuestSelect <> i Then
                If GlobalX >= x And GlobalX <= x + Width And GlobalY >= Y And GlobalY <= Y + Height Then
                    .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestMouseMoveColour
                Else
                    .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestDefaultColour
                End If
            End If

        Next i
    End With
End Sub

Private Sub lblList_MouseDown()
    Dim i As Byte, x As Long, Y As Long, Width As Long, Height As Long
    With Windows(GetWindowIndex("winQuest"))

        For i = 1 To MAX_QUESTS_WINDOW

            x = .Window.Left + .Controls(GetControlIndex("winQuest", "lblList" & i)).Left
            Y = .Window.top + .Controls(GetControlIndex("winQuest", "lblList" & i)).top
            Width = .Controls(GetControlIndex("winQuest", "lblList" & i)).Width
            Height = .Controls(GetControlIndex("winQuest", "lblList" & i)).Height

            If GlobalX >= x And GlobalX <= x + Width And GlobalY >= Y And GlobalY <= Y + Height Then
                If QuestSelect = i Then
                    .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestDefaultColour
                    QuestSelect = 0
                    QuestTimeToFinish = vbNullString
                    QuestNameToFinish = vbNullString
                    ClearQuestLogBox
                    Exit For
                End If
                QuestSelect = i
                If Player(MyIndex).PlayerQuest(FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & i)).Text)).TaskTimer.Timer > 0 Then
                    QuestNameToFinish = "Quest: " & Trim$(Quest(FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & i)).Text)).Name)
                    QuestTimeToFinish = "Tempo da Task: " & SecondsToHMS(CLng(Player(MyIndex).PlayerQuest(FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & i)).Text)).TaskTimer.Timer))
                End If
                LoadQuestLogBox QuestSelect
                .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestMouseDownColour
            Else
                .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestDefaultColour
            End If

        Next i
    End With
End Sub

Private Sub lblList_ClearColour()
    Dim i As Byte
    With Windows(GetWindowIndex("winQuest"))
        For i = 1 To MAX_QUESTS_WINDOW
            If .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestMouseMoveColour Then
                .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestDefaultColour
            End If
        Next i
    End With
End Sub

Public Sub RefreshQuestWindow()
    Dim i As Long, n As Long, LastQuest As Integer

    With Windows(GetWindowIndex("winQuest"))

        For n = 1 To MAX_QUESTS_WINDOW

            'clear
            .Controls(GetControlIndex("winQuest", "lblList" & n)).Text = "Vazio"
            .Controls(GetControlIndex("winQuest", "lblList" & n)).textColour = QuestDefaultColour
            .Controls(GetControlIndex("winQuest", "lblList" & n)).visible = False
            .Controls(GetControlIndex("winQuest", "picList" & n)).visible = False
            ClearQuestLogBox

            For i = 1 To MAX_QUESTS
                If QuestInProgress(i) Then
                    If LastQuest < i Then
                        .Controls(GetControlIndex("winQuest", "lblList" & n)).Text = Trim$(Quest(i).Name)
                        .Controls(GetControlIndex("winQuest", "lblList" & n)).visible = True
                        .Controls(GetControlIndex("winQuest", "picList" & n)).visible = True
                        LastQuest = i
                        QuestSelect = n
                        Exit For
                    End If
                End If
            Next i

            If .Controls(GetControlIndex("winQuest", "lblList" & n)).visible = False Then Exit For
        Next n

    End With
End Sub

Public Function FindQuestIndex(ByVal QuestName As String) As Integer
    Dim i As Integer

    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) = Trim$(QuestName) Then
            If QuestInProgress(i) Then
                FindQuestIndex = i
                Exit Function
            End If
        End If
    Next i

End Function

Private Sub ClearQuestLogBox()
    Dim i As Byte

    With Windows(GetWindowIndex("winQuest"))

        For i = 1 To 3
            .Controls(GetControlIndex("winQuest", "lblQuestDescription" & i)).Text = ""
        Next i
        
        For i = 1 To MAX_QUESTS_ITEMS
            .Controls(GetControlIndex("winQuest", "picReward" & i)).visible = False
        Next i

    End With
End Sub

Private Sub LoadQuestLogBox(ByVal QuestSelected As Byte)
    Dim QuestNum As Long, i As Long
    Dim QuestString As String

    ' Clear window first
    ClearQuestLogBox

    With Windows(GetWindowIndex("winQuest"))

        QuestNum = FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & QuestSelected)).Text)

        If QuestNum = 0 Then Exit Sub

        'Descrição da quest
        QuestString = Trim$(Quest(QuestNum).Speech)
        .Controls(GetControlIndex("winQuest", "lblQuestDescription1")).Text = QuestString

        'Objetivo da Task
        If Player(MyIndex).PlayerQuest(QuestNum).ActualTask > 0 Then
            QuestString = GetQuestObjetiveCurrent(QuestNum) & GetQuestObjetives(QuestNum)
        End If

        .Controls(GetControlIndex("winQuest", "lblQuestDescription2")).Text = QuestString

        'Recompensa da quest
        QuestString = "Exp: " & Quest(QuestNum).RewardExp & vbNewLine & "Level(s): " & Quest(QuestNum).RewardLevel
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).RewardItem(i).Item > 0 Then
                .Controls(GetControlIndex("winQuest", "picReward" & i)).Value = Quest(QuestNum).RewardItem(i).Item
                .Controls(GetControlIndex("winQuest", "picReward" & i)).visible = True
            End If
        Next i
        .Controls(GetControlIndex("winQuest", "lblQuestDescription3")).Text = QuestString

    End With
End Sub

Private Function GetQuestObjetives(ByVal QuestNum As Integer) As String
    Dim i As Byte
    Dim sString As String

    If Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        GetQuestObjetives = "Objetivos ja foram concluidos, voce pode iniciar a missão novamente!"
        Exit Function
    ElseIf Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Then
        GetQuestObjetives = "Objetivos ja foram concluidos, siga para proxima missao!"
        Exit Function
    End If

    For i = 1 To MAX_TASKS
        If i > Player(MyIndex).PlayerQuest(QuestNum).ActualTask Then

            If Quest(QuestNum).Task(i).Order <> 0 Then
                If i = (Player(MyIndex).PlayerQuest(QuestNum).ActualTask + 1) Then
                    sString = "PROX.:" & Space(1)
                End If
            End If

            Select Case Quest(QuestNum).Task(i).Order
            Case 0    'None

            Case QUEST_TYPE_GOSLAY
                sString = sString & "Derrotar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(NPC(Quest(QuestNum).Task(i).NPC).Name) & "/"
            Case QUEST_TYPE_GOGATHER
                sString = sString & "Obter" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(Item(Quest(QuestNum).Task(i).Item).Name) & "/"
            Case QUEST_TYPE_GOTALK
                sString = sString & "Falar com" & Space(1) & Trim$(NPC(Quest(QuestNum).Task(i).NPC).Name) & "/"
            Case QUEST_TYPE_GOREACH
                sString = sString & Quest(QuestNum).Task(i).TaskLog & "/"
            Case QUEST_TYPE_GOGIVE
                sString = sString & "Entregar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(Item(Quest(QuestNum).Task(i).Item).Name) & Space(1) & "Ao NPC" & Space(1) & Trim$(NPC(Quest(QuestNum).Task(i).NPC).Name) & "/"
            Case QUEST_TYPE_GOKILL
                sString = sString & "Derrotar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Jogadores" & "/"
            Case QUEST_TYPE_GOTRAIN
                sString = sString & "Treinar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Vezes na Resource" & Space(1) & Trim$(Resource(Quest(QuestNum).Task(i).Resource).Name) & "/"
            Case QUEST_TYPE_GOGET
                sString = sString & "Obter" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Item(s) do NPC" & Space(1) & Trim$(NPC(Quest(QuestNum).Task(i).NPC).Name) & "/"
            End Select
        End If
    Next i

    GetQuestObjetives = sString

End Function

Private Function GetQuestObjetiveCurrent(ByVal QuestNum As Integer) As String
    Dim i As Byte

    If Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        Exit Function
    ElseIf Player(MyIndex).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Then
        Exit Function
    End If

    For i = 1 To MAX_TASKS
        If i = Player(MyIndex).PlayerQuest(QuestNum).ActualTask Then

            Select Case Quest(QuestNum).Task(i).Order
            Case 0    'None
                GetQuestObjetiveCurrent = "ATUAL: Nenhum(a)"
            Case QUEST_TYPE_GOSLAY
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Derrotar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(NPC(Quest(QuestNum).Task(i).NPC).Name) & Space(1)
            Case QUEST_TYPE_GOGATHER
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Obter" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(Item(Quest(QuestNum).Task(i).Item).Name) & Space(1)
            Case QUEST_TYPE_GOTALK
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Falar com" & Space(1) & Trim$(NPC(Quest(QuestNum).Task(i).NPC).Name) & Space(1)
            Case QUEST_TYPE_GOREACH
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & Quest(QuestNum).Task(i).TaskLog & Space(1)
            Case QUEST_TYPE_GOGIVE
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Entregar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(Item(Quest(QuestNum).Task(i).Item).Name) & Space(1) & "Ao NPC" & Space(1) & Trim$(NPC(Quest(QuestNum).Task(i).NPC).Name) & Space(1)
            Case QUEST_TYPE_GOKILL
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Derrotar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Jogadores" & Space(1)
            Case QUEST_TYPE_GOTRAIN
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Treinar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Vezes na Resource" & Space(1) & Trim$(Resource(Quest(QuestNum).Task(i).Resource).Name) & Space(1)
            Case QUEST_TYPE_GOGET
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Obter" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Item(s) do NPC" & Space(1) & Trim$(NPC(Quest(QuestNum).Task(i).NPC).Name) & Space(1)
            End Select

            Exit Function
        End If
    Next i

End Function

Public Sub CalculateQuestTimer()
    Dim i As Integer

    With Windows(GetWindowIndex("winQuest"))

        If QuestSelect > 0 Then
            i = FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & QuestSelect)).Text)
            If i > 0 And i <= MAX_QUESTS Then
                If LenB(Trim$(Quest(i).Name)) > 0 Then
                    If Player(MyIndex).PlayerQuest(i).Status = QUEST_STARTED Then
                        If Player(MyIndex).PlayerQuest(i).TaskTimer.Active = YES Then
                            If Player(MyIndex).PlayerQuest(i).TaskTimer.Timer > 0 Then
                                Player(MyIndex).PlayerQuest(i).TaskTimer.Timer = Player(MyIndex).PlayerQuest(i).TaskTimer.Timer - 1
                                QuestNameToFinish = "Quest: " & Trim$(Quest(i).Name)
                                QuestTimeToFinish = "Tempo da Task: " & SecondsToHMS(CLng(Player(MyIndex).PlayerQuest(i).TaskTimer.Timer))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub
