Attribute VB_Name = "modGuild_Handle"
Option Explicit

Public Sub HandleGuildWindow(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ShowWindow GetWindowIndex("winGuildMaker")
End Sub

Public Sub HandleUpdateGuild(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim DecompData() As Byte
    Dim buffer As clsBuffer


    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData
    n = buffer.ReadLong

    Guild(n).Name = buffer.ReadString
    Guild(n).Motd = buffer.ReadString
    Guild(n).Color = buffer.ReadByte
    Guild(n).Honra = buffer.ReadLong
    Guild(n).GuildID = buffer.ReadByte
    Guild(n).GuildDisponivel = buffer.ReadByte
    Guild(n).Capacidade = buffer.ReadByte
    Guild(n).Boost = buffer.ReadByte
    Guild(n).Kills = buffer.ReadLong
    Guild(n).Victory = buffer.ReadLong
    Guild(n).Lose = buffer.ReadLong
    Guild(n).Icon = buffer.ReadByte

    ReDim GuildMembers(n).Membro(1 To Guild(n).Capacidade)

    For i = 1 To Guild(n).Capacidade
        With GuildMembers(n).Membro(i)
            .Login = Trim$(buffer.ReadString)
            .Name = Trim$(buffer.ReadString)
            .Level = buffer.ReadLong
            .Online = buffer.ReadByte
            .Dono = buffer.ReadByte
            .Admin = buffer.ReadByte
            .MembroID = buffer.ReadLong
            .MembroDisponivel = buffer.ReadByte
        End With
    Next

    buffer.Flush: Set buffer = Nothing

    If Player(MyIndex).Guild_ID <> 0 And n = Player(MyIndex).Guild_ID Then
        Call UpdateWindowGuild
    End If

End Sub
