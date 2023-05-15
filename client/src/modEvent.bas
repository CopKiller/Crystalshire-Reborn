Attribute VB_Name = "modEvent"
Option Explicit

Public Sub HandleEventMessage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long, header As String, saycolour As Long
    Dim message As String

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    header = buffer.ReadString
    message = Trim$(buffer.ReadString)
    saycolour = buffer.ReadLong

    ' remove the colour char from the message
    message = Replace$(message, ColourChar, vbNullString)

    AddText ColourChar & GetColStr(Gold) & header & ": " & ColourChar & GetColStr(saycolour) & message, Grey, , ChatChannel.chEvent

    Set buffer = Nothing
End Sub

