Attribute VB_Name = "modDiscord"
Option Explicit

Public Enum DiscordMsgType
    Entrou = 1
    Levelup
    Chat
    Death
End Enum

Private Const JoinedString = " Entrou no jogo! :rocket:"
Private Const DeathString = " Morreu pra "

Public Sub SendDiscordMsg(ByVal DscMsgType As DiscordMsgType, _
                          Optional ByVal Index As Long, _
                          Optional ByVal Txt As String)

    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SeDiscordMsg

    Buffer.WriteByte DscMsgType
    If DscMsgType = Entrou Then
        Buffer.WriteString GetPlayerName(Index) & JoinedString
    ElseIf DscMsgType = Death Then
        Buffer.WriteString GetPlayerName(Index) & DeathString & Txt & " :boom:"
    Else
        Buffer.WriteString GetPlayerName(Index)
    End If

    Buffer.WriteLong GetPlayerLevel(Index)
    Buffer.WriteByte GetPlayerPremium(Index)

    If DscMsgType = Chat Or Levelup Then Buffer.WriteString Txt

    SendToEventServer Buffer.ToArray

    Buffer.Flush: Set Buffer = Nothing
End Sub

