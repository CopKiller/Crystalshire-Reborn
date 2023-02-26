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

Public Sub HandleUpdateQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Dim DecompData()   As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData
    
    n = Buffer.ReadLong
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set Buffer = Nothing
End Sub

Public Sub HandlePlayerQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, QuestNum As Long

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()

    For i = 1 To MAX_QUESTS
        QuestNum = Buffer.ReadLong

        If QuestNum > 0 Then
            Player(MyIndex).PlayerQuest(QuestNum).Status = Buffer.ReadLong
            Player(MyIndex).PlayerQuest(QuestNum).ActualTask = Buffer.ReadLong
            Player(MyIndex).PlayerQuest(QuestNum).CurrentCount = Buffer.ReadLong
            
            Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Active = Buffer.ReadByte
            Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Timer = Buffer.ReadLong
            
            QuestTimeToFinish = vbNullString
            QuestNameToFinish = vbNullString
            QuestSelect = QuestNum
        End If
    Next

    RefreshQuestWindow

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleQuestMessage(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, QuestNum As Long, header As String, saycolour As Long
    Dim message As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    QuestNum = Buffer.ReadLong
    message = Trim$(Buffer.ReadString)
    saycolour = Buffer.ReadLong
    header = Buffer.ReadString
    
    ' remove the colour char from the message
    message = Replace$(message, ColourChar, vbNullString)
    
    AddText ColourChar & GetColStr(Gold) & header & Trim$(Quest(QuestNum).Name) & " : " & ColourChar & GetColStr(saycolour) & message, Grey, , ChatChannel.chQuest

    Set Buffer = Nothing
End Sub

Public Sub HandleQuestCancel(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim QuestNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    QuestNum = Buffer.ReadLong
    Player(MyIndex).PlayerQuest(QuestNum).Status = Buffer.ReadLong
    Player(MyIndex).PlayerQuest(QuestNum).ActualTask = Buffer.ReadLong
    Player(MyIndex).PlayerQuest(QuestNum).CurrentCount = Buffer.ReadLong
    
    Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Active = Buffer.ReadByte
    Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Timer = Buffer.ReadLong
    
    QuestTimeToFinish = vbNullString
    QuestNameToFinish = vbNullString

    RefreshQuestWindow
    
    Set Buffer = Nothing
End Sub

