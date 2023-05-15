Attribute VB_Name = "modQuests_Handle"
Option Explicit

Public Sub HandleQuestEditor()
    Dim i As Long

    With frmEditor_Quest
        Editor = EDITOR_TASKS
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_QUESTS
            .lstIndex.AddItem i & ": " & Trim$(Quest(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        QuestEditorInit
    End With

End Sub

Public Sub HandleUpdateQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Dim DecompData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData

    n = buffer.ReadLong
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set buffer = Nothing
End Sub

Public Sub HandlePlayerQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long, QuestNum As Long, QSelected As Integer

    Set buffer = New clsBuffer

    buffer.WriteBytes data()
    
    ' Recebe se começou a quest e seleciona ela na lista
    QSelected = buffer.ReadInteger

    For i = 1 To MAX_QUESTS
        QuestNum = buffer.ReadLong

        If QuestNum > 0 Then
            Player(MyIndex).PlayerQuest(QuestNum).Status = buffer.ReadLong
            Player(MyIndex).PlayerQuest(QuestNum).ActualTask = buffer.ReadLong
            Player(MyIndex).PlayerQuest(QuestNum).CurrentCount = buffer.ReadLong

            Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Active = buffer.ReadByte
            Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Timer = buffer.ReadLong

            QuestTimeToFinish = vbNullString
            QuestNameToFinish = vbNullString
            QuestSelect = QuestNum
        End If
    Next

    RefreshQuestWindow
    
    If QSelected > 0 Then
        SelectLastQuest QSelected
    End If

    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleQuestMessage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long, QuestNum As Long, header As String, saycolour As Long
    Dim message As String

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    QuestNum = buffer.ReadLong
    message = Trim$(buffer.ReadString)
    saycolour = buffer.ReadLong
    header = buffer.ReadString

    ' remove the colour char from the message
    message = Replace$(message, ColourChar, vbNullString)

    AddText ColourChar & GetColStr(Gold) & header & Trim$(Quest(QuestNum).Name) & " : " & ColourChar & GetColStr(saycolour) & message, Grey, , ChatChannel.chQuest

    Set buffer = Nothing
End Sub

Public Sub HandleQuestCancel(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim QuestNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    QuestNum = buffer.ReadLong
    Player(MyIndex).PlayerQuest(QuestNum).Status = buffer.ReadLong
    Player(MyIndex).PlayerQuest(QuestNum).ActualTask = buffer.ReadLong
    Player(MyIndex).PlayerQuest(QuestNum).CurrentCount = buffer.ReadLong

    Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Active = buffer.ReadByte
    Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Timer = buffer.ReadLong

    QuestTimeToFinish = vbNullString
    QuestNameToFinish = vbNullString

    RefreshQuestWindow

    Set buffer = Nothing
End Sub

