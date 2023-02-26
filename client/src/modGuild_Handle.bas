Attribute VB_Name = "modGuild_Handle"
Option Explicit

Public Sub HandleGuildWindow(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ShowWindow GetWindowIndex("winGuildMaker")
End Sub

Public Sub HandleUpdateGuild(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim DecompData()   As Byte
    Dim Buffer As clsBuffer
    
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData
    n = Buffer.ReadLong

    Guild(n).Name = Buffer.ReadString
    Guild(n).Motd = Buffer.ReadString
    Guild(n).Color = Buffer.ReadByte
    Guild(n).Honra = Buffer.ReadLong
    Guild(n).GuildID = Buffer.ReadByte
    Guild(n).GuildDisponivel = Buffer.ReadByte
    Guild(n).Capacidade = Buffer.ReadByte
    Guild(n).Boost = Buffer.ReadByte
    Guild(n).Kills = Buffer.ReadLong
    Guild(n).Victory = Buffer.ReadLong
    Guild(n).Lose = Buffer.ReadLong
    Guild(n).Icon = Buffer.ReadByte

    ReDim GuildMembers(n).Membro(1 To Guild(n).Capacidade)

    For i = 1 To Guild(n).Capacidade
        With GuildMembers(n).Membro(i)
            .Login = Trim$(Buffer.ReadString)
            .Name = Trim$(Buffer.ReadString)
            .Level = Buffer.ReadLong
            .Online = Buffer.ReadByte
            .Dono = Buffer.ReadByte
            .Admin = Buffer.ReadByte
            .MembroID = Buffer.ReadLong
            .MembroDisponivel = Buffer.ReadByte
        End With
    Next

    Buffer.Flush: Set Buffer = Nothing

    If Player(MyIndex).Guild_ID <> 0 And n = Player(MyIndex).Guild_ID Then
        Call UpdateWindowGuild
    End If

End Sub
