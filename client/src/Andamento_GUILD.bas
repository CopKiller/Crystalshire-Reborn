Attribute VB_Name = "modGuild"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''' SISTEMA DE GUILD ''''''''''''''''''
''''''''''''''''''   ESCRITO POR    ''''''''''''''''''
''''''''''''''''''   Filipe Bispo   ''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Const MAX_GUILDS As Byte = 20    ' Máximo de guilds (Valor Cliente & Server)
Private Const GUILD_CAPACIDADE_INICIAL As Byte = 5    ' Capacidade de membros inicial

Public GuildMembers(1 To MAX_GUILDS) As GuildMembersRec

' Declaração principal
Public Guild(1 To MAX_GUILDS) As GuildRec

Type GuildMemberRec
    Login As String    ' Login do membro
    Name As String    ' Nome do membro
    Level As Long    ' Level do membro
    Online As Boolean    ' Estaria ele online?
    Dono As Boolean    ' Seria ele dono da guild?
    Admin As Boolean    ' Seria ele admin da guild?
    MembroID As Long    ' ID do membro
    MembroDisponivel As Boolean    ' Slot de membro disponível?
End Type

Type GuildMembersRec
    Membro() As GuildMemberRec
End Type

Type GuildRec
    Name As String
    Motd As String    ' Mensagem do dia da guild
    Color As Long    ' Cor da guild
    Honra As Long    ' Honra da Guild
    Capacidade As Byte    ' Capacidade de membros na guild
    GuildID As Long    ' ID da Guild nas pastas
    Boost As Long
    GuildDisponivel As Boolean    ' Guild disponível para uso?
    Kills As Long
    Victory As Long
    Lose As Long
    Icon As Byte
End Type

' // TCP //

Public Sub GuildAccept_MouseDown()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CGuildInviteResposta
    buffer.WriteByte 1
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing

    'GUIWindow(GUI_GUILDINVITE).visible = False
End Sub
Public Sub GuildDecline_MouseDown()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CGuildInviteResposta
    buffer.WriteByte 0
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing

    'GUIWindow(GUI_GUILDINVITE).visible = False
End Sub
Public Sub SendCriarGuild(ByVal Nome As String)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CCriarGuild
    buffer.WriteString Nome
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendGuildInvite(ByVal Nome As String)
    Dim buffer As clsBuffer

    ' Proteção
    If Player(MyIndex).Guild_ID = 0 Then
        Exit Sub
    End If

    If GuildMembers(Player(MyIndex).Guild_ID).Membro(Player(MyIndex).Guild_MembroID).Admin = False Then
        AddText "Apenas admins da guild podem fazer isso!", BrightRed
        HideWindow GetWindowIndex("winGuildMenu")
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CGuildInvite
    buffer.WriteString Nome
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendGuildKick(ByVal Index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CGuildKick
    buffer.WriteLong Index
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendGuildDestroy()
    Dim buffer As clsBuffer

    ' Proteção
    If GuildMembers(Player(MyIndex).Guild_ID).Membro(Player(MyIndex).Guild_MembroID).Dono = False Then
        AddText "Apenas donos da guild podem fazer isso!", BrightRed
        HideWindow GetWindowIndex("winGuildMenu")
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CGuildDestroy
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendLeaveGuild()
    Dim buffer As clsBuffer

    If Player(MyIndex).Guild_ID > 0 Then
        If MsgBox("Tem certeza que deseja sair da guild?", vbYesNo) = vbYes Then
            Set buffer = New clsBuffer
            buffer.WriteLong CLeaveGuild
            SendData buffer.ToArray()
            buffer.Flush: Set buffer = Nothing
        End If
    End If
End Sub

Public Sub SendGuildPromote(ByVal Promover As Byte, ByVal MemberID As Byte)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CGuildPromote
    buffer.WriteByte Promover
    buffer.WriteByte MemberID
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub
' // FIM TCP //
