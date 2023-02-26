Attribute VB_Name = "modSvGuild"
Option Explicit

'For Clear functions
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''' SISTEMA DE GUILD ''''''''''''''''''
''''''''''''''''''   ESCRITO POR    ''''''''''''''''''
''''''''''''''''''   Filipe Bispo   ''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Const MAX_GUILDS As Byte = 20    ' Máximo de guilds (Valor Cliente & Server)
Private Const GUILD_CAPACIDADE_INICIAL As Byte = 5    ' Capacidade de membros inicial
Public Const MAX_GUILD_MEMBERS As Byte = 10

Public GuildMembers(1 To MAX_GUILDS) As GuildMembersRec

' Declaração principal
Public Guild(1 To MAX_GUILDS) As GuildRec

Private Type GuildMemberRec
    Login As String    ' Login do membro
    Name As String    ' Nome do membro
    Level As Long    ' Level do membro
    Online As Boolean    ' Estaria ele online?
    Dono As Boolean    ' Seria ele dono da guild?
    Admin As Boolean    ' Seria ele admin da guild?
    MembroID As Long    ' ID do membro
    MembroDisponivel As Boolean    ' Slot de membro disponível?
End Type

Private Type GuildMembersRec
    Membro() As GuildMemberRec
End Type

Private Type GuildRec
    Name As String
    MOTD As String    ' Mensagem do dia da guild
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

' // INICIO LOGICA //

Private Sub CriarGuild(ByVal Index As Long, ByVal Nome As String)
    Dim GuildSlot As Long
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Sub

    ' Checar se o usuário já tem uma guild
    If Player(Index).Guild_ID > 0 Then
        PlayerMsg Index, "Você já possui uma guild. (" & Guild(Player(Index).Guild_ID).Name & ")", Red
        Exit Sub
    End If

    If Len(Nome) < 4 Then
        PlayerMsg Index, "O nome da guild deve possuir ao menos 4 letras!", Red
        Exit Sub
    End If

    'If GetPlayerLevel(index) < 50 Then
    '   PlayerMsg index, "É preciso ser level 50 para criar uma guild.", Red
    '  Exit Sub
    ' End If

    For i = 1 To MAX_GUILDS
        If Guild(i).GuildDisponivel = False Then
            If Trim$(Guild(i).Name) = Trim$(Nome) Then
                PlayerMsg Index, "Nome de guild já em uso!", Red
                Exit Sub
            End If
        End If
    Next

    GuildSlot = FindOpenGuildSlot

    If GuildSlot > 0 Then
        ' Cria a guild:
        Guild(GuildSlot).GuildDisponivel = False
        Guild(GuildSlot).GuildID = GuildSlot
        Guild(GuildSlot).Name = Nome
        Guild(GuildSlot).Color = White
        Guild(GuildSlot).Honra = 0
        Guild(GuildSlot).MOTD = "Bem vindo à " & Nome & ". O fundador pode editar essa mensagem em seu painel!"
        Guild(GuildSlot).Capacidade = GUILD_CAPACIDADE_INICIAL
        Guild(GuildSlot).Boost = 0
        Guild(GuildSlot).Kills = 0
        Guild(GuildSlot).Victory = 0
        Guild(GuildSlot).Lose = 0
        Guild(GuildSlot).Icon = 0


        ReDim GuildMembers(GuildSlot).Membro(1 To Guild(GuildSlot).Capacidade)

        ' Torna o fundador:
        With GuildMembers(GuildSlot).Membro(1)
            .Login = Player(Index).Login
            .Name = GetPlayerName(Index)
            .Level = GetPlayerLevel(Index)
            .MembroID = 1
            .MembroDisponivel = False
            .Online = True
            .Dono = True
            .Admin = True
        End With

        For i = 2 To Guild(GuildSlot).Capacidade
            With GuildMembers(GuildSlot).Membro(i)
                .Login = vbNullString
                .Name = vbNullString
                .Level = 0
                .MembroID = 0
                .MembroDisponivel = True
                .Online = False
                .Dono = False
                .Admin = False
            End With
        Next

        Player(Index).Guild_ID = Guild(GuildSlot).GuildID
        Player(Index).Guild_MembroID = 1

        ' Salva tudo
        SaveGuild GuildSlot
        SavePlayer Index
        GuildCache_Create GuildSlot
        SendGuildAll GuildSlot
        SendPlayerData Index

        Call PlayerMsg(Index, Guild(GuildSlot).MOTD, Yellow)
        
        Call AddLog(GetPlayerName(Index) & " Criou uma guild nº" & Guild(GuildSlot).GuildID & ": " & Guild(GuildSlot).Name, PLAYER_LOG)
    Else
        PlayerMsg Index, "Ocorreu um erro ao criar a guild. Contatar administrador.", Red
        Exit Sub
    End If

End Sub

Private Sub GuildInviteResposta(ByVal Index As Long, ByVal Resposta As Byte)
    Dim GuildSlot As Long
    Dim MembroSlot As Long
    Dim Inviter As Long
    Dim Convidado As Long

    Convidado = Index
    Inviter = TempPlayer(Index).guildInvite

    If Not IsPlaying(Convidado) Then Exit Sub

    If Inviter = 0 Or Not IsPlaying(Inviter) Then
        PlayerMsg Index, "O convite expirou.", Red
        TempPlayer(Convidado).guildInvite = 0
        Exit Sub
    End If

    If Not Player(Inviter).Guild_ID > 0 Then
        PlayerMsg Index, "O convite expirou.", Red
        TempPlayer(Convidado).guildInvite = 0
        Exit Sub
    End If

    GuildSlot = Player(Inviter).Guild_ID

    ' 1 = Aceitar, 0 = Recusar
    If Resposta = 1 Then
        MembroSlot = FindOpenGuildMemberSlot(GuildSlot)

        If MembroSlot > 0 Then

            With GuildMembers(GuildSlot).Membro(MembroSlot)
                .Login = GetPlayerLogin(Convidado)
                .Name = GetPlayerName(Convidado)
                .Level = GetPlayerLevel(Convidado)
                .Dono = False
                .Admin = False
                .MembroID = MembroSlot
                .MembroDisponivel = False
                .Online = True
            End With

            TempPlayer(Convidado).guildInvite = 0
            Player(Convidado).Guild_ID = GuildSlot
            Player(Convidado).Guild_MembroID = MembroSlot
            
            SaveGuild GuildSlot
            SavePlayer Convidado
            GuildCache_Create GuildSlot
            SendGuildAll GuildSlot
            SendPlayerData Convidado

            ' Enviar bem vindo
            If LenB(Guild(GuildSlot).MOTD) > 0 Then
                Call PlayerMsg(Convidado, Guild(GuildSlot).MOTD, Yellow)
            End If
        Else
            TempPlayer(Convidado).guildInvite = 0
            PlayerMsg Convidado, "Ocorreu um erro ao entrar na guild.", Red
            Exit Sub
        End If

    Else
        TempPlayer(Convidado).guildInvite = 0
        PlayerMsg Inviter, GetPlayerName(Convidado) & " recusou seu convite.", Red
    End If
End Sub
Private Sub GuildLeave(ByVal Index As Long)
    Dim GuildSlot As Long
    Dim MemberSlot As Long
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Sub
    GuildSlot = Player(Index).Guild_ID
    MemberSlot = Player(Index).Guild_MembroID
    If Not GuildSlot > 0 Then Exit Sub

    If GuildMembers(GuildSlot).Membro(MemberSlot).Dono = True Then
        PlayerMsg Index, "O Fundador não pode sair da guild, apenas destruí-la no seu painel!", Red
        Exit Sub
    End If

    Player(Index).Guild_ID = 0
    Player(Index).Guild_MembroID = 0
    With GuildMembers(GuildSlot).Membro(MemberSlot)
        .Login = vbNullString
        .Name = vbNullString
        .Level = 0
        .MembroID = 0
        .MembroDisponivel = True
        .Online = False
        .Dono = False
        .Admin = False
    End With

    ' Salva
    SaveGuild GuildSlot
    SavePlayer Index
    GuildCache_Create GuildSlot
    SendUpdateGuildTo Index, GuildSlot
    SendPlayerData Index

    PlayerMsg Index, "Você saiu da guild!", White

    ' Atualiza a guild
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Guild_ID = GuildSlot Then
                SendUpdateGuildTo i, GuildSlot
                PlayerMsg i, GetPlayerName(Index) & " saiu da guild!", Grey
            End If
        End If
    Next
End Sub
Private Sub GuildDestroy(ByVal Index As Long)
    Dim GuildSlot As Long
    Dim MemberSlot As Long
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Sub
    GuildSlot = Player(Index).Guild_ID
    MemberSlot = Player(Index).Guild_MembroID

    If Not GuildSlot > 0 Then Exit Sub

    ' Apenas donos podem destruir suas guilds.
    If GuildMembers(GuildSlot).Membro(MemberSlot).Dono = False Then
        PlayerMsg Index, "É o doido né, só o fundador da guild pode fazer isso!", Red
        Exit Sub
    End If

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Guild_ID = GuildSlot Then
                Player(i).Guild_ID = 0
                Player(i).Guild_MembroID = 0
                SendPlayerData i
                PlayerMsg i, "Sua guild foi desfeita pelo fundador.", White
            End If
        End If
    Next
    
    Call AddLog(GetPlayerName(Index) & " Destruiu a Guild nº" & Guild(GuildSlot).GuildID & ": " & Guild(GuildSlot).Name, PLAYER_LOG)

    ClearGuild GuildSlot
    SaveGuild GuildSlot
    GuildCache_Create GuildSlot
    SendGuildAll GuildSlot
End Sub
Private Sub GuildKick(ByVal Index As Long, ByVal toKick As Long)
    Dim GuildSlot As Long
    Dim Kicked As Long
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Sub
    GuildSlot = Player(Index).Guild_ID

    If Not GuildSlot > 0 Then Exit Sub
    If GuildMembers(GuildSlot).Membro(toKick).MembroDisponivel = True Then Exit Sub

    ' Apenas donos de guild podem kickar os outros
    If GuildMembers(GuildSlot).Membro(Player(Index).Guild_MembroID).Admin = False Then
        PlayerMsg Index, "Apenas admins de guild podem kickar.", Red
        Exit Sub
    End If

    ' Não pode kickar o fundador da guild.
    If GuildMembers(GuildSlot).Membro(toKick).Dono = True Then
        PlayerMsg Index, "Você não pode kickar o fundador da guild.", Red
        Exit Sub
    End If

    ' Só o fundador pode kickar outro admin de guild.
    If GuildMembers(GuildSlot).Membro(toKick).Admin = True And GuildMembers(GuildSlot).Membro(Player(Index).Guild_MembroID).Dono = False Then
        PlayerMsg Index, "Apenas o fundador pode kickar um admin de guild.", Red
        Exit Sub
    End If

    Kicked = FindPlayer(GuildMembers(GuildSlot).Membro(toKick).Name)

    ' Checar se a vítima está online
    If GuildMembers(GuildSlot).Membro(toKick).Online = True And Kicked > 0 Then

        ' Remove o membro
        Player(Kicked).Guild_ID = 0
        Player(Kicked).Guild_MembroID = 0
        With GuildMembers(GuildSlot).Membro(toKick)
            .Login = vbNullString
            .Name = vbNullString
            .Level = 0
            .MembroID = 0
            .MembroDisponivel = True
            .Online = False
            .Dono = False
            .Admin = False
        End With

        ' Salva
        SaveGuild GuildSlot
        SavePlayer Kicked
        GuildCache_Create GuildSlot
        SendUpdateGuildTo Kicked, GuildSlot
        SendPlayerData Kicked

        ' Atualiza a guild
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Player(i).Guild_ID = GuildSlot Then
                    SendUpdateGuildTo i, GuildSlot
                End If
            End If
        Next

        ' Avisa que tudo deu certo
        PlayerMsg Index, "Membro kickado!", White
        PlayerMsg Kicked, "Você foi kickado de sua guild!", White
    Else
        ' Nesse caso precisamos kickar um offline
        With GuildMembers(GuildSlot).Membro(toKick)
            .Login = vbNullString
            .Name = vbNullString
            .Level = 0
            .MembroID = 0
            .MembroDisponivel = True
            .Online = False
            .Dono = False
            .Admin = False
        End With

        SaveGuild GuildSlot
        GuildCache_Create GuildSlot

        ' Atualiza a guild
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Player(i).Guild_ID = GuildSlot Then
                    SendUpdateGuildTo i, GuildSlot
                End If
            End If
        Next

        PlayerMsg Index, "Membro kickado!", White
    End If

End Sub
Public Sub guildInvite(ByVal Index As Long, ByVal Nome As String)
    Dim GuildSlot As Long
    Dim Inviter As Long
    Dim Convidado As Long

    GuildSlot = Player(Index).Guild_ID

    Inviter = Index

    If Not IsPlaying(Inviter) Then Exit Sub

    ' Checar se o usuário tem uma guild
    If Not Player(Index).Guild_ID > 0 Then
        PlayerMsg Index, "Você nem tem guild lek.", Red
        Exit Sub
    End If

    Convidado = FindPlayer(Nome)

    ' Apenas admins de guild podem invitar os outros
    If GuildMembers(GuildSlot).Membro(Player(Index).Guild_MembroID).Admin = False Then
        PlayerMsg Index, "Apenas admins de guild podem convidar.", Red
        Exit Sub
    End If

    ' Checar se o convidado existe/está online
    If Not IsPlaying(Convidado) Then
        PlayerMsg Inviter, "Usuário offline ou inexistente.", Red
        Exit Sub
    End If

    ' Checar se o convidado já tem uma guild
    If Player(Convidado).Guild_ID > 0 Then
        PlayerMsg Inviter, "Usuário já possui uma guild.", Red
        Exit Sub
    End If

    ' Checar se o convidado já está durante um convite
    If TempPlayer(Convidado).guildInvite > 0 Then
        PlayerMsg Inviter, "Usuário já está decidindo outro convite no momento.", Red
        Exit Sub
    End If

    ' Tudo certo
    TempPlayer(Convidado).guildInvite = Inviter
    Call SendGuildInvite(Inviter, Convidado)
    PlayerMsg Inviter, "Convite à guild enviado!", Yellow
End Sub

Sub GuildPromote(ByVal Index As Long, ByVal MembroID As Byte)
    Dim GuildSlot As Long
    Dim Inviter As Long
    Dim Ss As String
    Dim i As Byte

    Inviter = Index

    If Not IsPlaying(Inviter) Then Exit Sub

    GuildSlot = Player(Index).Guild_ID

    ' Checar se o usuário tem uma guild
    If Not GuildSlot > 0 Then
        PlayerMsg Index, "Você nem tem guild lek.", Red
        Exit Sub
    End If

    ' Apenas donos de guild podem promover os outros
    If GuildMembers(GuildSlot).Membro(Player(Index).Guild_MembroID).Dono = False Then
        PlayerMsg Index, "Apenas o Líder da guild pode promover.", Red
        Exit Sub
    End If

    ' Checar se o jogador existe na guild
    If GuildMembers(GuildSlot).Membro(MembroID).MembroDisponivel = True Then
        PlayerMsg Inviter, "O jogador nao existe na sua guild.", Red
        Exit Sub
    End If

    ' Checar se não é o próprio lider tentando promover ele mesmo
    If Player(Index).Guild_MembroID = MembroID Then
        PlayerMsg Inviter, "Você não pode se promover.", Red
        Exit Sub
    End If

    ' Tudo certo, pode promover
    GuildMembers(GuildSlot).Membro(MembroID).Admin = True

    ' Mensagem da promoção pra todos da Guild
    Ss = GuildMembers(GuildSlot).Membro(MembroID).Name & " Foi Promovido à Admin da Guild!"
    Call SayMsg_Guild(Index, Ss, QBColor(Yellow))

    ' Save and Update guild to members
    Call SaveGuild(GuildSlot)
    GuildCache_Create GuildSlot

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Guild_ID = Player(Index).Guild_ID Then
                SendUpdateGuildTo i, Player(Index).Guild_ID
            End If
        End If
    Next

End Sub

Sub GuildRebaixar(ByVal Index As Long, ByVal MembroID As Byte)
    Dim GuildSlot As Long
    Dim Inviter As Long
    Dim Ss As String
    Dim i As Byte

    Inviter = Index

    If Not IsPlaying(Inviter) Then Exit Sub

    GuildSlot = Player(Index).Guild_ID

    ' Checar se o usuário tem uma guild
    If Not GuildSlot > 0 Then
        PlayerMsg Index, "Você nem tem guild lek.", Red
        Exit Sub
    End If

    ' Apenas donos de guild podem rebaixar os outros
    If GuildMembers(GuildSlot).Membro(Player(Index).Guild_MembroID).Dono = False Then
        PlayerMsg Index, "Apenas o Líder da guild pode rebaixar.", Red
        Exit Sub
    End If

    ' Checar se o jogador existe na guild
    If GuildMembers(GuildSlot).Membro(MembroID).MembroDisponivel = True Then
        PlayerMsg Inviter, "O jogador nao existe na sua guild.", Red
        Exit Sub
    End If

    ' Checar se não é o próprio lider tentando rebaixar ele mesmo
    If Player(Index).Guild_MembroID = MembroID Then
        PlayerMsg Inviter, "Você não pode se rebaixar.", Red
        Exit Sub
    End If

    ' Tudo certo, pode rebaixar
    GuildMembers(GuildSlot).Membro(MembroID).Admin = False

    ' Mensagem da promoção pra todos da Guild
    Ss = GuildMembers(GuildSlot).Membro(MembroID).Name & " Foi rebaixado a membro novamente!"
    Call SayMsg_Guild(Index, Ss, QBColor(Yellow))

    ' Save and Update guild to members
    Call SaveGuild(GuildSlot)
    GuildCache_Create GuildSlot

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Guild_ID = Player(Index).Guild_ID Then
                SendUpdateGuildTo i, Player(Index).Guild_ID
            End If
        End If
    Next

End Sub

' // FIM LOGICA //

' // INICIO TCP //

Sub SendGuildInvite(ByVal Inviter As Long, ByVal Convidado As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SGuildInvite
    Buffer.WriteString GetPlayerName(Inviter)
    Buffer.WriteString Guild(Player(Inviter).Guild_ID).Name

    SendDataTo Convidado, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub
Sub SendGuilds(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_GUILDS
        If Guild(i).GuildDisponivel = False Then
            Call SendUpdateGuildTo(Index, i)
        End If
    Next
End Sub

Sub SendGuildWindow(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SGuildWindow
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Function FindOpenGuildSlot() As Long
    Dim i As Integer

    For i = 1 To MAX_GUILDS
        If Guild(i).GuildDisponivel = True Then
            FindOpenGuildSlot = i
            Exit Function
        End If

        FindOpenGuildSlot = 0
    Next i
End Function

Public Function FindOpenGuildMemberSlot(ByVal Guilda As Long) As Long
    Dim i As Integer

    For i = 1 To Guild(Guilda).Capacidade
        If GuildMembers(Guilda).Membro(i).MembroDisponivel = True Then
            FindOpenGuildMemberSlot = i
            Exit Function
        End If

        FindOpenGuildMemberSlot = 0
    Next i
End Function
' // FIM DATABASE //

Public Sub HandleCriarGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Nome As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Nome = Buffer.ReadString

    Buffer.Flush: Set Buffer = Nothing

    Call CriarGuild(Index, Nome)
End Sub

Public Sub HandleGuildInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Nome As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Nome = Buffer.ReadString

    Buffer.Flush: Set Buffer = Nothing

    Call guildInvite(Index, Nome)
End Sub

Public Sub HandleGuildInviteResposta(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Resposta As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Resposta = Buffer.ReadByte

    Buffer.Flush: Set Buffer = Nothing

    Call GuildInviteResposta(Index, Resposta)
End Sub

Public Sub HandleSaveGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim GuildSlot As Long

    GuildSlot = Player(Index).Guild_ID

    If Not GuildSlot > 0 Then Exit Sub
    If GuildMembers(GuildSlot).Membro(Player(Index).Guild_MembroID).Admin = False Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Guild(GuildSlot).MOTD = Buffer.ReadString
    Guild(GuildSlot).Color = Buffer.ReadByte
    Guild(GuildSlot).Icon = Buffer.ReadByte

    SaveGuild GuildSlot
    GuildCache_Create GuildSlot
    SendGuildAll GuildSlot    '
    PlayerMsg Index, "Guild salva!", Yellow

    Buffer.Flush: Set Buffer = Nothing

End Sub

Public Sub HandleGuildKick(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim He As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    He = Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing

    Call GuildKick(Index, He)
End Sub

Public Sub HandleGuildPromote(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim ID As Byte
    Dim Promote As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    Promote = Buffer.ReadByte
    ID = Buffer.ReadByte

    Buffer.Flush: Set Buffer = Nothing

    If Promote = YES Then
        Call GuildPromote(Index, ID)
    Else
        Call GuildRebaixar(Index, ID)
    End If
End Sub

Public Sub HandleGuildDestroy(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call GuildDestroy(Index)
End Sub

Public Sub HandleLeaveGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call GuildLeave(Index)
End Sub

'////////////////////
'///GUILD DATABASE///
'////////////////////
Sub ClearGuilds()
    Dim i As Long
    For i = 1 To MAX_GUILDS
        Call ClearGuild(i)
    Next
End Sub
Private Sub ClearGuild(ByVal Index As Long)
    Dim n As Long

    Call ZeroMemory(ByVal VarPtr(Guild(Index)), LenB(Guild(Index)))
    Guild(Index).Name = vbNullString
    Guild(Index).MOTD = vbNullString
    Guild(Index).Color = 0
    Guild(Index).Honra = 0
    Guild(Index).GuildDisponivel = True
    Guild(Index).GuildID = 0
    Guild(Index).Capacidade = GUILD_CAPACIDADE_INICIAL

    Call ZeroMemory(ByVal VarPtr(GuildMembers(Index)), LenB(GuildMembers(Index)))
    ReDim GuildMembers(Index).Membro(1 To Guild(Index).Capacidade)

    For n = 1 To Guild(Index).Capacidade
        GuildMembers(Index).Membro(n).MembroDisponivel = True
    Next
End Sub

Private Sub CheckGuilds()
    Dim i As Long

    For i = 1 To MAX_GUILDS

        If Not FileExist("\Data\Guilds\Guild" & i & ".dat") Then
            Call SaveGuild(i)
        End If

    Next

End Sub

Sub LoadGuilds()
    Dim filename As String
    Dim i As Long
    Dim n As Long
    Dim F As Long
    Call CheckGuilds

    For i = 1 To MAX_GUILDS
        filename = App.Path & "\data\Guilds\Guild" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Guild(i)
        Close #F
        ReDim GuildMembers(i).Membro(1 To Guild(i).Capacidade)

        filename = App.Path & "\data\Guilds\guildmembers\guild" & i & ".dat"
        F = FreeFile

        Open filename For Binary As #F
        For n = 1 To Guild(i).Capacidade
            Get #F, , GuildMembers(i).Membro(n)
            GuildMembers(i).Membro(n).Online = False
        Next
        Close #F
    Next
End Sub

Private Sub SaveGuilds()
    Dim i As Long

    For i = 1 To MAX_GUILDS
        Call SaveGuild(i)
    Next

End Sub

Public Sub SaveGuild(ByVal Num As Long)
    Dim filename As String
    Dim F As Long
    Dim n As Long
    filename = App.Path & "\data\Guilds\Guild" & Num & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Guild(Num)
    Close #F
    filename = App.Path & "\data\guilds\guildmembers\guild" & Num & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    For n = 1 To Guild(Num).Capacidade
        Put #F, , GuildMembers(Num).Membro(n)
    Next
    Close #F
End Sub

