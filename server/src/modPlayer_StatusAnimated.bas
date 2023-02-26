Attribute VB_Name = "modPlayer_StatusAnimated"
Option Explicit

Public Type StatusRec
    Ativo As Byte
End Type

Public Sub HandlePlayerStatus(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim StatusType As Byte
    Dim StatusOn As Byte
    
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    StatusType = Buffer.ReadByte
    StatusOn = Buffer.ReadByte

    Buffer.Flush: Set Buffer = Nothing

    SendStatusPlayer Index, StatusType, StatusOn
End Sub

Public Sub CheckPlayerStatus(ByVal Index As Integer)
    Dim Tick As Currency
    Tick = getTime

    ' Ativa player afk no balão
    If TempPlayer(Index).StatusNum(Status.Afk).Ativo = NO Then
        If Tick > TempPlayer(Index).AFKTimer + 300000 Then    ' 10 Minutos ativa
            SendStatusPlayer Index, Status.Afk, YES
        End If
    End If

End Sub

Public Sub SendStatusPlayer(ByVal Index As Long, ByVal StatusNum As Status, ByVal OnOff As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SStatus
    Buffer.WriteLong Index
    Buffer.WriteByte StatusNum
    Buffer.WriteByte OnOff

    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing

    Select Case StatusNum

    Case Status.Afk
        TempPlayer(Index).StatusNum(Status.Afk).Ativo = OnOff
        If OnOff = NO Then TempPlayer(Index).AFKTimer = getTime

    Case Status.Typing
        TempPlayer(Index).StatusNum(Status.Typing).Ativo = OnOff
        
    End Select
End Sub

