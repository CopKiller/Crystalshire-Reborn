Attribute VB_Name = "modHandleData"
Option Explicit

Public HandleDataSub(CMSG_COUNT) As Long

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CGuildMsg) = GetAddress(AddressOf HandleGuildMsg)
    HandleDataSub(CPartyMsg) = GetAddress(AddressOf HandlePartyMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CTarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CTradeGold) = GetAddress(AddressOf HandleTradeGold)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CChatOption) = GetAddress(AddressOf HandleChatOption)
    HandleDataSub(CRequestEditConv) = GetAddress(AddressOf HandleRequestEditConv)
    HandleDataSub(CSaveConv) = GetAddress(AddressOf HandleSaveConv)
    HandleDataSub(CRequestConvs) = GetAddress(AddressOf HandleRequestConvs)
    HandleDataSub(CFinishTutorial) = GetAddress(AddressOf HandleFinishTutorial)
    'Guild
    HandleDataSub(CCriarGuild) = GetAddress(AddressOf HandleCriarGuild)
    HandleDataSub(CGuildInvite) = GetAddress(AddressOf HandleGuildInvite)
    HandleDataSub(CGuildInviteResposta) = GetAddress(AddressOf HandleGuildInviteResposta)
    HandleDataSub(CSaveGuild) = GetAddress(AddressOf HandleSaveGuild)
    HandleDataSub(CGuildKick) = GetAddress(AddressOf HandleGuildKick)
    HandleDataSub(CGuildDestroy) = GetAddress(AddressOf HandleGuildDestroy)
    HandleDataSub(CLeaveGuild) = GetAddress(AddressOf HandleLeaveGuild)
    HandleDataSub(CGuildPromote) = GetAddress(AddressOf HandleGuildPromote)
    'Serial
    HandleDataSub(CRequestEditSerial) = GetAddress(AddressOf HandleEditSerial)
    HandleDataSub(CSaveSerial) = GetAddress(AddressOf HandleSaveSerial)
    HandleDataSub(CRequestSerial) = GetAddress(AddressOf HandleRequestSerial)
    HandleDataSub(CSendSerial) = GetAddress(AddressOf HandleSendSerial)
    'Premium
    HandleDataSub(CRequestEditPremium) = GetAddress(AddressOf HandleRequestEditPremium)
    HandleDataSub(CChangePremium) = GetAddress(AddressOf HandleChangePremium)
    HandleDataSub(CRemovePremium) = GetAddress(AddressOf HandleRemovePremium)
    'Quest
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CRequestQuests) = GetAddress(AddressOf HandleRequestQuests)
    HandleDataSub(CPlayerHandleQuest) = GetAddress(AddressOf HandlePlayerCancelQuest)
    HandleDataSub(CQuestLogUpdate) = GetAddress(AddressOf HandleQuestLogUpdate)
    'Status Animated
    HandleDataSub(CStatus) = GetAddress(AddressOf HandlePlayerStatus)
    'Conjuntos
    HandleDataSub(CRequestEditConjunto) = GetAddress(AddressOf HandleRequestEditConjunto)
    HandleDataSub(CSaveConjunto) = GetAddress(AddressOf HandleSaveConjunto)
    HandleDataSub(CRequestConjuntos) = GetAddress(AddressOf HandleRequestConjuntos)
    HandleDataSub(CCheckIn) = GetAddress(AddressOf HandleCheckIn)
    HandleDataSub(CSendBet) = GetAddress(AddressOf HandleBet)
End Sub

Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
    Dim TempBuffer() As Byte

    TempBuffer = DecryptPacket(Data, (UBound(Data) - LBound(Data)) + 1)

    Dim Buffer As clsBuffer
    Dim MsgType As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes TempBuffer

    MsgType = Buffer.ReadLong

    If MsgType <= 0 Then
        Exit Sub
    End If

    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    ' Desativar o afk caso receba alguma packet do client
    If TempPlayer(Index).InGame = True And MsgType <> ClientPackets.CCheckPing Then
        If TempPlayer(Index).StatusNum(Status.Afk).Ativo = YES Then
        SendStatusPlayer Index, Afk, NO
    Else
        TempPlayer(Index).AFKTimer = getTime
    End If
    End If

    ' Quando o usuario está autenticado, processa os dados recebidos.
    If LoginTokenAccepted(Index) Then
        CallWindowProc HandleDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
    Else
        If MsgType = CLogin Then CallWindowProc HandleDataSub(CLogin), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
    End If
End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, Name As String, i As Long, n As Long, Password As String
    Dim token As String

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer

            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            token = Buffer.ReadString

            ' Check versions
            If Buffer.ReadLong <> CLIENT_MAJOR Or Buffer.ReadLong <> CLIENT_MINOR Or Buffer.ReadLong <> CLIENT_REVISION Then
                Call AlertMsg(Index, DIALOGUE_MSG_OUTDATED)
                Buffer.Flush: Set Buffer = Nothing
                Exit Sub
            End If

            Buffer.Flush: Set Buffer = Nothing

            If token = vbNullString Or Name = vbNullString Then
                Call AlertMsg(Index, DIALOGUE_MSG_CONNECTION, MENU_LOGIN)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(Index, DIALOGUE_MSG_REBOOTING, MENU_LOGIN)
                Exit Sub
            End If

            'EventSv Is Connected?
            If Not IsEventServerConnected Then
                Call AlertMsg(Index, DIALOGUE_MSG_REBOOTING, MENU_LOGIN)
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Then
                Call AlertMsg(Index, DIALOGUE_MSG_USERLENGTH, MENU_LOGIN)
                Exit Sub
            End If

            If IsMultiAccounts(Index, Name) Then
                ' Expulsa quem está logado, e libera pra login
            End If

            If Not loginTokenOk(Index, Name, token) Then
                Call AlertMsg(Index, DIALOGUE_MSG_CONNECTION, MENU_LOGIN)
                Exit Sub
            End If
            ' Indica que o token acabou de ser ativado.
            LoginTokenAccepted(Index) = True

            ' we have a char!
            UseChar Index

            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextLoginAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
        End If
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(Index), Index, Msg, QBColor(White))
    Call SendChatBubble(GetPlayerMap(Index), Index, TARGET_TYPE_PLAYER, Msg, White)

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    If Player(Index).isMuted Then
        PlayerMsg Index, "You have been muted and cannot talk in global.", BrightRed
        Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    s = "[Global]" & GetPlayerName(Index) & ": " & Msg
    Call SayMsg_Global(Index, Msg, QBColor(White))
    Call SendDiscordMsg(Chat, Index, Msg)
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleGuildMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    If Player(Index).Guild_ID = 0 Then
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Buffer.Flush: Set Buffer = Nothing: Exit Sub
            Exit Sub
        End If

    Next

    s = "[Guild]" & GetPlayerName(Index) & ": " & Msg
    Call SayMsg_Guild(Index, Msg, QBColor(Yellow))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)

    Buffer.Flush: Set Buffer = Nothing: Exit Sub
End Sub

Private Sub HandlePartyMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    If TempPlayer(Index).inParty = 0 Then
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Buffer.Flush: Set Buffer = Nothing: Exit Sub
            Exit Sub
        End If
    Next

    s = "[Party]" & GetPlayerName(Index) & ": " & Msg
    Call PartyMsg(TempPlayer(Index).inParty, s, QBColor(Yellow))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)

    Buffer.Flush: Set Buffer = Nothing: Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> Index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Msg & "'", TellColor)
            Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(GetPlayerName(Index), "Cannot message yourself.", BrightRed)
    End If

    Buffer.Flush: Set Buffer = Nothing

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    dir = Buffer.ReadLong    'CLng(Parse(1))
    movement = Buffer.ReadLong    'CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        If Spell(TempPlayer(Index).spellBuffer.Spell).CanRun = NO Then
            Call SendPlayerXY(Index)
            Exit Sub
        End If
    End If

    'Cant move if in the bank!
    If TempPlayer(Index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(Index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(Index).StunDuration > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' Prever player from moving if in shop
    If TempPlayer(Index).InShop > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(Index) <> tmpX Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    If GetPlayerY(Index) <> tmpY Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    ' cant move if chatting
    If TempPlayer(Index).inChatWith > 0 Then
        ClosePlayerChat Index
    End If

    Call PlayerMove(Index, dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    dir = Buffer.ReadLong    'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim InvNum As Long
    Dim Buffer As clsBuffer

    ' get inventory slot number
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    InvNum = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing

    UseItem Index, InvNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, n As Long, Damage As Long, TempIndex As Long, X As Long, Y As Long, MapNum As Long, dirReq As Long

    ' can't attack whilst casting
    If TempPlayer(Index).spellBuffer.Spell > 0 Then Exit Sub

    ' can't attack whilst stunned
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack Index

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> Index Then
            TryPlayerAttackPlayer Index, i
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc Index, i
    Next

    ' check if we've got a remote chat tile
    MapNum = GetPlayerMap(Index)
    X = GetPlayerX(Index)
    Y = GetPlayerY(Index)
    If Map(MapNum).TileData.Tile(X, Y).Type = TILE_TYPE_CHAT Then
        dirReq = Map(MapNum).TileData.Tile(X, Y).Data2
        If Player(Index).dir = dirReq Then
            InitChat Index, MapNum, Map(MapNum).TileData.Tile(X, Y).Data1, True
            Exit Sub
        End If
    End If

    ' Check tradeskills
    Select Case GetPlayerDir(Index)
    Case DIR_UP

        If GetPlayerY(Index) = 0 Then Exit Sub
        X = GetPlayerX(Index)
        Y = GetPlayerY(Index) - 1
    Case DIR_DOWN

        If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MapData.MaxY Then Exit Sub
        X = GetPlayerX(Index)
        Y = GetPlayerY(Index) + 1
    Case DIR_LEFT

        If GetPlayerX(Index) = 0 Then Exit Sub
        X = GetPlayerX(Index) - 1
        Y = GetPlayerY(Index)
    Case DIR_RIGHT

        If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MapData.MaxX Then Exit Sub
        X = GetPlayerX(Index) + 1
        Y = GetPlayerY(Index)
    End Select

    CheckResource Index, X, Y
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim PointType As Byte
    Dim Buffer As clsBuffer
    Dim sMes As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PointType = Buffer.ReadByte    'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(Index) > 0 Then
        ' make sure they're not maxed
        If GetPlayerRawStat(Index, PointType) >= 255 Then
            PlayerMsg Index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If

        ' make sure they're not spending too much
        If GetPlayerRawStat(Index, PointType) - Class(GetPlayerClass(Index)).Stat(PointType) >= (GetPlayerLevel(Index) * 2) - 1 Then
            PlayerMsg Index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If

        ' Take away a stat point
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)

        ' Everything is ok
        Select Case PointType
        Case Stats.Strength
            Call SetPlayerStat(Index, Stats.Strength, GetPlayerRawStat(Index, Stats.Strength) + 1)
            sMes = "Strength"
        Case Stats.Endurance
            Call SetPlayerStat(Index, Stats.Endurance, GetPlayerRawStat(Index, Stats.Endurance) + 1)
            sMes = "Endurance"
        Case Stats.Intelligence
            Call SetPlayerStat(Index, Stats.Intelligence, GetPlayerRawStat(Index, Stats.Intelligence) + 1)
            sMes = "Intelligence"
        Case Stats.Agility
            Call SetPlayerStat(Index, Stats.Agility, GetPlayerRawStat(Index, Stats.Agility) + 1)
            sMes = "Agility"
        Case Stats.Willpower
            Call SetPlayerStat(Index, Stats.Willpower, GetPlayerRawStat(Index, Stats.Willpower) + 1)
            sMes = "Willpower"
        End Select
        
        Call AllocateConjuntoBonus(Index)

        SendActionMsg GetPlayerMap(Index), "+1 " & sMes, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
    Else
        Exit Sub
    End If

    ' Send the update
    SendPlayerData Index
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString    'Parse(1)
    Buffer.Flush: Set Buffer = Nothing
    i = FindPlayer(Name)
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString)    'Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(Index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(Index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp to yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString)    'Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(Index) & ".", BrightBlue)
            Call PlayerMsg(Index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp yourself to yourself!", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong    'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
    Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The sprite
    n = Buffer.ReadLong    'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing
    Call SetPlayerSprite(Index, n)
    Call SendPlayerData(Index)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    dir = Buffer.ReadLong    'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(Index, dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim MapNum As Long
    Dim X As Long
    Dim Y As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Index)

    Call ClearMap(MapNum)

    With Map(MapNum).MapData
        .Name = Buffer.ReadString
        .Music = Buffer.ReadString
        .Moral = Buffer.ReadByte
        .Up = Buffer.ReadLong
        .Down = Buffer.ReadLong
        .left = Buffer.ReadLong
        .Right = Buffer.ReadLong
        .BootMap = Buffer.ReadLong
        .BootX = Buffer.ReadByte
        .BootY = Buffer.ReadByte
        .MaxX = Buffer.ReadByte
        .MaxY = Buffer.ReadByte
        .BossNpc = Buffer.ReadLong
        .Panorama = Buffer.ReadByte

        .Weather = Buffer.ReadByte
        .WeatherIntensity = Buffer.ReadByte

        .Fog = Buffer.ReadByte
        .FogSpeed = Buffer.ReadByte
        .FogOpacity = Buffer.ReadByte

        .Red = Buffer.ReadByte
        .Green = Buffer.ReadByte
        .Blue = Buffer.ReadByte
        .Alpha = Buffer.ReadByte
        
        .Sun = Buffer.ReadByte
        
        .DAYNIGHT = Buffer.ReadByte

        For i = 1 To MAX_MAP_NPCS
            .NPC(i) = Buffer.ReadLong
            Call ClearMapNpc(i, MapNum)
        Next
    End With

    Map(MapNum).TileData.EventCount = Buffer.ReadLong
    If Map(MapNum).TileData.EventCount > 0 Then
        ReDim Preserve Map(MapNum).TileData.Events(1 To Map(MapNum).TileData.EventCount)
        For i = 1 To Map(MapNum).TileData.EventCount
            With Map(MapNum).TileData.Events(i)
                .Name = Buffer.ReadString
                .X = Buffer.ReadLong
                .Y = Buffer.ReadLong
                .PageCount = Buffer.ReadLong
            End With
            If Map(MapNum).TileData.Events(i).PageCount > 0 Then
                ReDim Preserve Map(MapNum).TileData.Events(i).EventPage(1 To Map(MapNum).TileData.Events(i).PageCount)
                For X = 1 To Map(MapNum).TileData.Events(i).PageCount
                    With Map(MapNum).TileData.Events(i).EventPage(X)
                        .chkPlayerVar = Buffer.ReadByte
                        .chkSelfSwitch = Buffer.ReadByte
                        .chkHasItem = Buffer.ReadByte
                        .PlayerVarNum = Buffer.ReadLong
                        .SelfSwitchNum = Buffer.ReadLong
                        .HasItemNum = Buffer.ReadLong
                        .PlayerVariable = Buffer.ReadLong
                        .GraphicType = Buffer.ReadByte
                        .Graphic = Buffer.ReadLong
                        .GraphicX = Buffer.ReadLong
                        .GraphicY = Buffer.ReadLong
                        .MoveType = Buffer.ReadByte
                        .MoveSpeed = Buffer.ReadByte
                        .MoveFreq = Buffer.ReadByte
                        .WalkAnim = Buffer.ReadByte
                        .StepAnim = Buffer.ReadByte
                        .DirFix = Buffer.ReadByte
                        .WalkThrough = Buffer.ReadByte
                        .Priority = Buffer.ReadByte
                        .Trigger = Buffer.ReadByte
                        .CommandCount = Buffer.ReadLong
                    End With
                    If Map(MapNum).TileData.Events(i).EventPage(X).CommandCount > 0 Then
                        ReDim Preserve Map(MapNum).TileData.Events(i).EventPage(X).Commands(1 To Map(MapNum).TileData.Events(i).EventPage(X).CommandCount)
                        For Y = 1 To Map(MapNum).TileData.Events(i).EventPage(X).CommandCount
                            With Map(MapNum).TileData.Events(i).EventPage(X).Commands(Y)
                                .Type = Buffer.ReadByte
                                .Text = Buffer.ReadString
                                .colour = Buffer.ReadLong
                                .Channel = Buffer.ReadByte
                                .TargetType = Buffer.ReadByte
                                .target = Buffer.ReadLong
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If

    ReDim Map(MapNum).TileData.Tile(0 To Map(MapNum).MapData.MaxX, 0 To Map(MapNum).MapData.MaxY)

    For X = 0 To Map(MapNum).MapData.MaxX
        For Y = 0 To Map(MapNum).MapData.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).TileData.Tile(X, Y).Layer(i).X = Buffer.ReadLong
                Map(MapNum).TileData.Tile(X, Y).Layer(i).Y = Buffer.ReadLong
                Map(MapNum).TileData.Tile(X, Y).Layer(i).Tileset = Buffer.ReadLong
                Map(MapNum).TileData.Tile(X, Y).Autotile(i) = Buffer.ReadByte
            Next
            Map(MapNum).TileData.Tile(X, Y).Type = Buffer.ReadByte
            Map(MapNum).TileData.Tile(X, Y).Data1 = Buffer.ReadLong
            Map(MapNum).TileData.Tile(X, Y).Data2 = Buffer.ReadLong
            Map(MapNum).TileData.Tile(X, Y).Data3 = Buffer.ReadLong
            Map(MapNum).TileData.Tile(X, Y).Data4 = Buffer.ReadLong
            Map(MapNum).TileData.Tile(X, Y).Data5 = Buffer.ReadLong
            Map(MapNum).TileData.Tile(X, Y).DirBlock = Buffer.ReadByte
        Next
    Next

    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    ' Save the map
    Call MapCache_Create(MapNum)
    Call SaveMap(MapNum)
    Call ClearTempTile(MapNum)
    Call CacheResources(MapNum)
    Call GetMapCRC32(MapNum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i

    Buffer.Flush: Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Get yes/no value
    s = Buffer.ReadLong    'Parse(1)
    Buffer.Flush: Set Buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If

    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)

    'send Resource cache
    For i = 0 To MapResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next

    TempPlayer(Index).GettingMap = NO
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapDone
    SendDataTo Index, Buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(Index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim InvNum As Long
    Dim Amount As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    InvNum = Buffer.ReadLong    'CLng(Parse(1))
    Amount = Buffer.ReadLong    'CLng(Parse(2))
    Buffer.Flush: Set Buffer = Nothing

    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub

    If GetPlayerInvItemNum(Index, InvNum) < 1 Or GetPlayerInvItemNum(Index, InvNum) > MAX_ITEMS Then Exit Sub

    If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable > 0 Then
        If Amount < 1 Or Amount > GetPlayerInvItemValue(Index, InvNum) Then Exit Sub
    End If

    ' everything worked out fine
    Call PlayerMapDropItem(Index, InvNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next

    CacheResources GetPlayerMap(Index)
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim i As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1

    For i = 1 To MAX_MAPS

        If LenB(Trim$(Map(i).MapData.Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = i + 1
            tMapEnd = i + 1
        End If

    Next

    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    Call PlayerMsg(Index, s, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString)    'Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.GAME_NAME & " by " & GetPlayerName(Index) & "!", White)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, DIALOGUE_MSG_KICKED)
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot kick yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString)    'Parse(1))
    Buffer.Flush: Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n)
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot ban yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEditMap
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SItemEditor
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong    'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Buffer.Flush: Set Buffer = Nothing

    ' Save it
    Call SaveItem(n)
    Call GetItemCRC32(n)
    Call ItemCache_Create(n)
    Call SendItemsAll(n)
    Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimationEditor
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong    'CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Buffer.Flush: Set Buffer = Nothing

    ' Save it
    Call AnimationCache_Create(n)
    Call SendAnimationsAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(Index) & " saved Animation #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcEditor
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim NpcNum As Long
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    NpcNum = Buffer.ReadLong

    ' Prevent hacking
    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Exit Sub
    End If

    NpcSize = LenB(NPC(NpcNum))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(NPC(NpcNum)), ByVal VarPtr(NpcData(0)), NpcSize
    ' Save it
    
    Call SaveNpc(NpcNum)
    Call GetNpcCRC32(NpcNum)
    Call NpcCache_Create(NpcNum)
    Call SendNpcsAll(NpcNum)
    Call AddLog(GetPlayerName(Index) & " saved Npc #" & NpcNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceEditor
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call ResourceCache_Create(ResourceNum)
    Call SendResourcesAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(Index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SShopEditor
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ShopNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    ShopNum = Buffer.ReadLong

    ' Prevent hacking
    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Buffer.Flush: Set Buffer = Nothing
    ' Save it
    Call ShopCache_Create(ShopNum)
    Call SendShopsAll(ShopNum)
    Call SaveShop(ShopNum)
    Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellEditor
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    SpellNum = Buffer.ReadLong

    ' Prevent hacking
    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SpellCache_Create(SpellNum)
    Call SendSpellsAll(SpellNum)
    Call SaveSpell(SpellNum)
    Call AddLog(GetPlayerName(Index) & " saved Spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(Buffer.ReadString)    'Parse(1))
    ' The access
    i = Buffer.ReadLong    'CLng(Parse(2))
    Buffer.Flush: Set Buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
                Call PlayerMsg(Index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Invalid access level.", Red)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(Buffer.ReadString)    'Parse(1))
    SaveOptions
    Buffer.Flush: Set Buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, target As Long, TargetType As Long

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()

    target = Buffer.ReadLong
    TargetType = Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing

    ' set player's target - no need to send, it's client side
    TempPlayer(Index).target = target
    TempPlayer(Index).TargetType = TargetType
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(Index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong    'CLng(Parse(1))
    Buffer.Flush: Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(Index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim OldSlot As Long, NewSlot As Long

    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    OldSlot = Buffer.ReadLong
    NewSlot = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    PlayerSwitchInvSlots Index, OldSlot, NewSlot
End Sub

Sub HandleSwapSpellSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim OldSlot As Long, NewSlot As Long, n As Long

    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    OldSlot = Buffer.ReadLong
    NewSlot = Buffer.ReadLong
    
    ' Check for subscript out of range
    If OldSlot < 1 Or OldSlot > MAX_PLAYER_SPELLS Then
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If
    ' Check for subscript out of range
    If NewSlot < 1 Or NewSlot > MAX_PLAYER_SPELLS Then
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If

    ' Proteções
    If TempPlayer(Index).spellBuffer.Spell = Player(Index).Spell(OldSlot).Spell Then
        PlayerMsg Index, "You cannot swap spells whilst casting.", BrightRed
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    ElseIf TempPlayer(Index).spellBuffer.Spell = Player(Index).Spell(NewSlot).Spell Then
        PlayerMsg Index, "You cannot swap spells whilst casting.", BrightRed
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    ElseIf Player(Index).SpellCD(OldSlot) > getTime Then
        PlayerMsg Index, "You cannot swap spells whilst they're cooling down.", BrightRed
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    ElseIf Player(Index).SpellCD(NewSlot) > getTime Then
        PlayerMsg Index, "You cannot swap spells whilst they're cooling down.", BrightRed
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If

    Buffer.Flush: Set Buffer = Nothing

    PlayerSwitchSpellSlots Index, OldSlot, NewSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendPing
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleUnequip(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerUnequipItem Index, Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData Index
End Sub

Sub HandleRequestItems(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems Index
End Sub

Sub HandleRequestAnimations(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations Index
End Sub

Sub HandleRequestNPCS(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs Index
End Sub

Sub HandleRequestResources(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources Index
End Sub

Sub HandleRequestSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells Index
End Sub

Sub HandleRequestShops(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops Index
End Sub

Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' item
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong

    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub

    SpawnItem tmpItem, tmpAmount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index)
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess(Index) < 4 Then Exit Sub
    SetPlayerExp Index, GetPlayerNextLevel(Index)
    CheckPlayerLevelUp Index
End Sub

Sub HandleForgetSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim spellSlot As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    spellSlot = Buffer.ReadLong

    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If

    ' dont let them forget a spell which is in CD
    If Player(Index).SpellCD(spellSlot) > getTime Then
        PlayerMsg Index, "Cannot forget a spell which is cooling down!", BrightRed
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If

    ' dont let them forget a spell which is buffered
    If TempPlayer(Index).spellBuffer.Spell = spellSlot Then
        PlayerMsg Index, "Cannot forget a spell which you are casting!", BrightRed
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If

    Call RemovePlayerSpell(Index, Player(Index).Spell(spellSlot).Spell)

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(Index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopslot As Long
    Dim ShopNum As Long
    Dim itemAmount As Long
    Dim SString As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    shopslot = Buffer.ReadLong

    ' not in shop, exit out
    ShopNum = TempPlayer(Index).InShop
    If ShopNum < 1 Or ShopNum > MAX_SHOPS Then Exit Sub

    With Shop(ShopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub

        ' make sure they have inventory space
        If FindOpenInvSlot(Index, .Item) = 0 Then
            PlayerMsg Index, "You do not have enough inventory space.", BrightRed
            ResetShopAction Index
            Exit Sub
        End If

        ' check has the cost item
        If .costitem = 0 Then    ' is gold
            itemAmount = GetPlayerGold(Index)
        Else
            itemAmount = HasItem(Index, .costitem)
        End If

        If itemAmount = 0 Or itemAmount < .costvalue Then
            PlayerMsg Index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction Index
            Exit Sub
        End If

        ' it's fine, let's go ahead
        If .costitem = 0 Then    ' is gold
            Call SetPlayerGold(Index, GetPlayerGold(Index) - .costvalue)
            Call SendGoldUpdate(Index)
            SString = "You successfully bought " & Trim$(Item(.Item).Name) & " for " & .costvalue & " Golds."
        Else
            TakeInvItem Index, .costitem, .costvalue
            SString = "You successfully bought " & Trim$(Item(.Item).Name) & " for " & .costvalue & " " & Trim$(Item(.costitem).Name) & "."
        End If
        
        If GiveInvItem(Index, .Item, .ItemValue, 0) Then
            PlayerMsg Index, SString, BrightGreen
        End If

    End With

    ' send confirmation message & reset their shop action
    'PlayerMsg index, "Trade successful.", BrightGreen

    ResetShopAction Index

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleSellItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim InvSlot As Long
    Dim ItemNum As Long
    Dim price As Long
    Dim multiplier As Double
    Dim Amount As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    InvSlot = Buffer.ReadLong

    If TempPlayer(Index).InShop = 0 Then Exit Sub
    
    If TempPlayer(Index).InTrade > 0 Then Exit Sub
    
    If TempPlayer(Index).InBank Then Exit Sub

    ' if invalid, exit out
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Sub

    ' has item?
    If GetPlayerInvItemNum(Index, InvSlot) < 1 Or GetPlayerInvItemNum(Index, InvSlot) > MAX_ITEMS Then Exit Sub

    ' seems to be valid
    ItemNum = GetPlayerInvItemNum(Index, InvSlot)

    ' work out price
    multiplier = Shop(TempPlayer(Index).InShop).BuyRate / 100
    price = Item(ItemNum).price * multiplier

    ' item has cost?
    If price <= 0 Then
        PlayerMsg Index, "The shop doesn't want that item.", BrightRed
        ResetShopAction Index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem Index, ItemNum, 1
    SetPlayerGold Index, price
    Call SendGoldUpdate(Index)

    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Trade successful.", BrightGreen
    ResetShopAction Index

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long
    Dim Y As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    X = Buffer.ReadLong
    Y = Buffer.ReadLong

    If X < 0 Then X = 0
    If Y < 0 Then Y = 0

    If GetPlayerAccess(Index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX Index, X
        SetPlayerY Index, Y
        SendPlayerXYToMap Index
    End If

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long, Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' find the target
    tradeTarget = Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing

    If Not IsConnected(Index) Or Not IsPlaying(Index) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If

    If Not IsConnected(tradeTarget) Or Not IsPlaying(tradeTarget) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If

    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > Player_HighIndex Then Exit Sub

    ' can't trade with yourself..
    If tradeTarget = Index Then
        PlayerMsg Index, "You can't trade with yourself.", BrightRed
        Exit Sub
    End If

    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(Index).Map Then Exit Sub

    ' make sure they're stood next to each other
    tX = Player(tradeTarget).X
    tY = Player(tradeTarget).Y
    sX = Player(Index).X
    sY = Player(Index).Y

    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If

    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg Index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = Index
    SendTradeRequest tradeTarget, Index
End Sub

Sub HandleAcceptTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long

    tradeTarget = TempPlayer(Index).TradeRequest

    If Not IsConnected(Index) Or Not IsPlaying(Index) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        Exit Sub
    End If

    If Not IsConnected(tradeTarget) Or Not IsPlaying(tradeTarget) Then
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If

    If TempPlayer(Index).TradeRequest <= 0 Or TempPlayer(Index).TradeRequest > Player_HighIndex Then Exit Sub
    ' let them know they're trading
    PlayerMsg Index, "You have accepted " & Trim$(GetPlayerName(tradeTarget)) & "'s trade request.", BrightGreen
    PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has accepted your trade request.", BrightGreen
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
    TempPlayer(tradeTarget).TradeRequest = 0
    ' set that they're trading with each other
    TempPlayer(Index).InTrade = tradeTarget
    TempPlayer(tradeTarget).InTrade = Index
    ' clear out their trade offers
    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).value = 0
    Next
    TempPlayer(Index).TradeGold = 0
    TempPlayer(tradeTarget).TradeGold = 0
    ' Used to init the trade window clientside
    SendTrade Index, tradeTarget
    SendTrade tradeTarget, Index
    ' Send the offer data - Used to clear their client
    SendTradeUpdate Index, 0
    SendTradeUpdate Index, 1
    SendTradeUpdate tradeTarget, 0
    SendTradeUpdate tradeTarget, 1
End Sub

Sub HandleDeclineTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(Index).TradeRequest, GetPlayerName(Index) & " has declined your trade request.", BrightRed
    PlayerMsg Index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long, X As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeGold As Long, tmpTradeGold2 As Long
    Dim ItemNum As Long
    Dim theirInvSpace As Long, yourInvSpace As Long
    Dim theirItemCount As Long, yourItemCount As Long

    If TempPlayer(Index).InTrade = 0 Then Exit Sub

    TempPlayer(Index).AcceptTrade = True
    tradeTarget = TempPlayer(Index).InTrade

    If Not IsConnected(Index) Or Not IsPlaying(Index) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If

    If Not IsConnected(tradeTarget) Or Not IsPlaying(tradeTarget) Then
        TempPlayer(tradeTarget).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If

    ' if not both of them accept, then exit
    If Not TempPlayer(tradeTarget).AcceptTrade Then
        SendTradeStatus Index, 2
        SendTradeStatus tradeTarget, 1
        Exit Sub
    End If

    ' get inventory spaces
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, i) > 0 Then
            ' check if we're offering it
            For X = 1 To MAX_INV
                If TempPlayer(Index).TradeOffer(X).Num = i Then
                    ItemNum = Player(Index).Inv(TempPlayer(Index).TradeOffer(X).Num).Num
                    ' if it's a currency then make sure we're offering all of it
                    If Item(ItemNum).Stackable > 0 Then
                        If TempPlayer(Index).TradeOffer(X).value = GetPlayerInvItemNum(Index, i) Then
                            yourInvSpace = yourInvSpace + 1
                        End If
                    Else
                        yourInvSpace = yourInvSpace + 1
                    End If
                End If
            Next
        Else
            yourInvSpace = yourInvSpace + 1
        End If
        If GetPlayerInvItemNum(tradeTarget, i) > 0 Then
            ' check if we're offering it
            For X = 1 To MAX_INV
                If TempPlayer(tradeTarget).TradeOffer(X).Num = i Then
                    ItemNum = Player(tradeTarget).Inv(TempPlayer(tradeTarget).TradeOffer(X).Num).Num
                    ' if it's a currency then make sure we're offering all of it
                    If Item(ItemNum).Stackable > 0 Then
                        If TempPlayer(tradeTarget).TradeOffer(X).value = GetPlayerInvItemNum(tradeTarget, i) Then
                            theirInvSpace = theirInvSpace + 1
                        End If
                    Else
                        theirInvSpace = theirInvSpace + 1
                    End If
                End If
            Next
        Else
            theirInvSpace = theirInvSpace + 1
        End If
    Next

    ' get item count
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).Num > 0 Then
            ItemNum = Player(Index).Inv(TempPlayer(Index).TradeOffer(i).Num).Num
            If ItemNum > 0 Then
                If Item(ItemNum).Stackable > 0 Then
                    ' check if the other player has the item
                    If HasItem(tradeTarget, ItemNum) = 0 Then
                        yourItemCount = yourItemCount + 1
                    End If
                Else
                    yourItemCount = yourItemCount + 1
                End If
            End If
        End If
        If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
            ItemNum = Player(tradeTarget).Inv(TempPlayer(tradeTarget).TradeOffer(i).Num).Num
            If ItemNum > 0 Then
                If Item(ItemNum).Stackable > 0 Then
                    ' check if the other player has the item
                    If HasItem(Index, ItemNum) = 0 Then
                        theirItemCount = theirItemCount + 1
                    End If
                Else
                    theirItemCount = theirItemCount + 1
                End If
            End If
        End If
    Next

    ' make sure they have enough space
    If yourInvSpace < theirItemCount Then
        PlayerMsg Index, "You don't have enough inventory space.", BrightRed
        PlayerMsg tradeTarget, "They don't have enough inventory space.", BrightRed
        TempPlayer(Index).AcceptTrade = False
        TempPlayer(tradeTarget).AcceptTrade = False
        SendTradeUpdate Index, 0
        SendTradeUpdate tradeTarget, 0
        SendTradeStatus Index, 3
        SendTradeStatus tradeTarget, 3
        Exit Sub
    End If
    If theirInvSpace < yourItemCount Then
        PlayerMsg Index, "They don't have enough inventory space.", BrightRed
        PlayerMsg tradeTarget, "You don't have enough inventory space.", BrightRed
        TempPlayer(Index).AcceptTrade = False
        TempPlayer(tradeTarget).AcceptTrade = False
        SendTradeUpdate Index, 0
        SendTradeUpdate tradeTarget, 0
        SendTradeStatus Index, 3
        SendTradeStatus tradeTarget, 3
        Exit Sub
    End If

    ' take their items
    For i = 1 To MAX_INV
        ' player
        If TempPlayer(Index).TradeOffer(i).Num > 0 Then
            ItemNum = Player(Index).Inv(TempPlayer(Index).TradeOffer(i).Num).Num
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem(i).Num = ItemNum
                tmpTradeItem(i).value = TempPlayer(Index).TradeOffer(i).value
                ' take item
                TakeInvSlot Index, TempPlayer(Index).TradeOffer(i).Num, tmpTradeItem(i).value
            End If
        End If
        ' target
        If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
            ItemNum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem2(i).Num = ItemNum
                tmpTradeItem2(i).value = TempPlayer(tradeTarget).TradeOffer(i).value
                ' take item
                TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, tmpTradeItem2(i).value
            End If
        End If
    Next
    
    ' Take Golds
    tmpTradeGold = TempPlayer(Index).TradeGold
    tmpTradeGold2 = TempPlayer(tradeTarget).TradeGold
    Call SetPlayerGold(Index, GetPlayerGold(Index) - tmpTradeGold)
    Call SetPlayerGold(tradeTarget, GetPlayerGold(tradeTarget) - tmpTradeGold2)

    ' taken all items. now they can't not get items because of no inventory space.
    For i = 1 To MAX_INV
        ' player
        If tmpTradeItem2(i).Num > 0 Then
            ' give away!
            GiveInvItem Index, tmpTradeItem2(i).Num, tmpTradeItem2(i).value, 0, False
        End If
        ' target
        If tmpTradeItem(i).Num > 0 Then
            ' give away!
            GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).value, 0, False
        End If
    Next
    
    ' Get Golds
    Call SetPlayerGold(Index, GetPlayerGold(Index) + tmpTradeGold2)
    Call SetPlayerGold(tradeTarget, GetPlayerGold(tradeTarget) + tmpTradeGold)

    SendGoldUpdate Index
    SendGoldUpdate tradeTarget
    SendInventory Index
    SendInventory tradeTarget

    ' they now have all the items. Clear out values + let them out of the trade.
    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).value = 0
    Next
    TempPlayer(Index).TradeGold = 0
    TempPlayer(tradeTarget).TradeGold = 0

    TempPlayer(Index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0

    PlayerMsg Index, "Trade completed.", BrightGreen
    PlayerMsg tradeTarget, "Trade completed.", BrightGreen

    SendCloseTrade Index
    SendCloseTrade tradeTarget
End Sub

Sub HandleDeclineTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim tradeTarget As Long

    tradeTarget = TempPlayer(Index).InTrade

    If tradeTarget = 0 Then
        SendCloseTrade Index
        Exit Sub
    End If

    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).value = 0
    Next
    
    TempPlayer(Index).TradeGold = 0
    TempPlayer(tradeTarget).TradeGold = 0

    TempPlayer(Index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0

    PlayerMsg Index, "You declined the trade.", BrightRed
    PlayerMsg tradeTarget, GetPlayerName(Index) & " has declined the trade.", BrightRed

    SendCloseTrade Index
    SendCloseTrade tradeTarget
End Sub

Sub HandleTradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim InvSlot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim ItemNum As Long
    Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    InvSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing

    If InvSlot <= 0 Or InvSlot > MAX_INV Then Exit Sub

    ItemNum = GetPlayerInvItemNum(Index, InvSlot)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub

    If TempPlayer(Index).InTrade <= 0 Or TempPlayer(Index).InTrade > Player_HighIndex Then Exit Sub

    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, InvSlot) Then
        PlayerMsg Index, "You do not have that many.", BrightRed
        Exit Sub
    End If

    ' make sure it's not soulbound
    If Item(ItemNum).BindType > 0 Then
        If Player(Index).Inv(InvSlot).Bound > 0 Then
            PlayerMsg Index, "Cannot trade a soulbound item.", BrightRed
            Exit Sub
        End If
    End If

    If Item(ItemNum).Stackable > 0 Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = InvSlot Then
                ' add amount
                TempPlayer(Index).TradeOffer(i).value = TempPlayer(Index).TradeOffer(i).value + Amount
                ' clamp to limits
                If TempPlayer(Index).TradeOffer(i).value > GetPlayerInvItemValue(Index, InvSlot) Then
                    TempPlayer(Index).TradeOffer(i).value = GetPlayerInvItemValue(Index, InvSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(Index).AcceptTrade = False
                TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False

                SendTradeStatus Index, 0
                SendTradeStatus TempPlayer(Index).InTrade, 0

                SendTradeUpdate Index, 0
                SendTradeUpdate TempPlayer(Index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = InvSlot Then
                PlayerMsg Index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If

    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(Index).TradeOffer(EmptySlot).Num = InvSlot
    TempPlayer(Index).TradeOffer(EmptySlot).value = Amount

    ' cancel any trade agreement and send new data
    TempPlayer(Index).AcceptTrade = False
    TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False

    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0

    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleTradeGold(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Amount As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    Amount = Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing

    If TempPlayer(Index).InTrade <= 0 Or TempPlayer(Index).InTrade > Player_HighIndex Then Exit Sub

    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerGold(Index) Then
        PlayerMsg Index, "You do not have that many.", BrightRed
        Exit Sub
    End If

    TempPlayer(Index).TradeGold = Amount

    ' cancel any trade agreement and send new data
    TempPlayer(Index).AcceptTrade = False
    TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False

    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0

    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    tradeSlot = Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing

    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(Index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub

    TempPlayer(Index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(Index).TradeOffer(tradeSlot).value = 0

    If TempPlayer(Index).AcceptTrade Then TempPlayer(Index).AcceptTrade = False
    If TempPlayer(TempPlayer(Index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False

    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0

    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim sType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    Dim tmpSlot As Byte, tmpType As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    sType = Buffer.ReadLong
    Slot = Buffer.ReadLong
    hotbarNum = Buffer.ReadLong

    Select Case sType
    Case 0    ' clear
        Player(Index).Hotbar(hotbarNum).Slot = 0
        Player(Index).Hotbar(hotbarNum).sType = 0
    Case 1    ' inventory
        If Slot > 0 And Slot <= MAX_INV Then
            If Player(Index).Inv(Slot).Num > 0 Then
                If Len(Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name)) > 0 Then
                    Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Inv(Slot).Num
                    Player(Index).Hotbar(hotbarNum).sType = sType
                End If
            End If
        End If
    Case 2    ' spell
        If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
            If Player(Index).Spell(Slot).Spell > 0 Then
                If Len(Trim$(Spell(Player(Index).Spell(Slot).Spell).Name)) > 0 Then
                    Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Spell(Slot).Spell
                    Player(Index).Hotbar(hotbarNum).sType = sType
                End If
            End If
        End If

    Case 3    ' hotbar change
        If Slot > 0 And Slot <= MAX_HOTBAR Then
            tmpSlot = Player(Index).Hotbar(hotbarNum).Slot
            tmpType = Player(Index).Hotbar(hotbarNum).sType
            Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Hotbar(Slot).Slot
            Player(Index).Hotbar(hotbarNum).sType = Player(Index).Hotbar(Slot).sType

            Player(Index).Hotbar(Slot).Slot = tmpSlot
            Player(Index).Hotbar(Slot).sType = tmpType
        End If
    End Select

    SendHotbar Index

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    Slot = Buffer.ReadLong

    Select Case Player(Index).Hotbar(Slot).sType
    Case 1    ' inventory
        For i = 1 To MAX_INV
            If Player(Index).Inv(i).Num > 0 Then
                If Player(Index).Inv(i).Num = Player(Index).Hotbar(Slot).Slot Then
                    UseItem Index, i
                    Exit Sub
                End If
            End If
        Next
    Case 2    ' spell
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Index).Spell(i).Spell > 0 Then
                If Player(Index).Spell(i).Spell = Player(Index).Hotbar(Slot).Slot Then
                    BufferSpell Index, i
                    Exit Sub
                End If
            End If
        Next
    End Select

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, TargetIndex As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    TargetIndex = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing

    ' make sure it's a valid target
    If TargetIndex = Index Then
        PlayerMsg Index, "You can't invite yourself. That would be weird.", BrightRed
        Exit Sub
    End If

    ' make sure they're connected and on the same map
    If Not IsConnected(TargetIndex) Or Not IsPlaying(TargetIndex) Then Exit Sub
    If GetPlayerMap(TargetIndex) <> GetPlayerMap(Index) Then Exit Sub

    ' init the request
    Party_Invite Index, TargetIndex
End Sub

Sub HandleAcceptParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteAccept TempPlayer(Index).partyInvite, Index
End Sub

Sub HandleDeclineParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(Index).partyInvite, Index
End Sub

Sub HandlePartyLeave(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave Index
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Conv packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditConv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SConvEditor
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save Conv packet ::
' :::::::::::::::::::::::
Sub HandleSaveConv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim convNum As Long
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim X As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    convNum = Buffer.ReadLong

    ' Prevent hacking
    If convNum < 0 Or convNum > MAX_CONVS Then
        Exit Sub
    End If

    With Conv(convNum)
        .Name = Buffer.ReadString
        .chatCount = Buffer.ReadLong
        ReDim .Conv(1 To .chatCount)
        For i = 1 To .chatCount
            .Conv(i).Conv = Buffer.ReadString
            For X = 1 To 4
                .Conv(i).rText(X) = Buffer.ReadString
                .Conv(i).rTarget(X) = Buffer.ReadLong
            Next
            .Conv(i).Event = Buffer.ReadLong
            .Conv(i).Data1 = Buffer.ReadLong
            .Conv(i).Data2 = Buffer.ReadLong
            .Conv(i).Data3 = Buffer.ReadLong
        Next
    End With

    Buffer.Flush: Set Buffer = Nothing

    ' Save it
    Call ConvCache_Create(convNum)
    Call SendConvAll(convNum)
    Call SaveConv(convNum)
    Call AddLog(GetPlayerName(Index) & " salvou a Conversa #" & convNum & ".", ADMIN_LOG)
End Sub

Sub HandleRequestConvs(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendConvs Index
End Sub

Sub HandleChatOption(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    chatOption Index, Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleFinishTutorial(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Player(Index).TutorialState = 1
    SavePlayer Index
End Sub
