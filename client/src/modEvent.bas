Attribute VB_Name = "modEvent"
Option Explicit

Public Sub HandleEventMessage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, header As String, saycolour As Long
    Dim message As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    header = Buffer.ReadString
    message = Trim$(Buffer.ReadString)
    saycolour = Buffer.ReadLong

    ' remove the colour char from the message
    message = Replace$(message, ColourChar, vbNullString)

    AddText ColourChar & GetColStr(Gold) & header & ": " & ColourChar & GetColStr(saycolour) & message, Grey, , ChatChannel.chEvent

    Set Buffer = Nothing
End Sub

