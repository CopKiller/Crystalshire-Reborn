Attribute VB_Name = "modEvent"
Option Explicit

Public Sub SendEventMsgTo(ByVal Index As Long, Header As String, Message As String, saycolour As Long)
    Dim Buffer As clsBuffer, Tmr As Currency

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventMsg

    Buffer.WriteString Header
    Buffer.WriteString Message
    Buffer.WriteLong saycolour

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendEventMsgAll(Header As String, Message As String, saycolour As Long)
    Dim Buffer As clsBuffer, Tmr As Currency

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventMsg

    Buffer.WriteString Header
    Buffer.WriteString Message
    Buffer.WriteLong saycolour

    SendDataToAll Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

