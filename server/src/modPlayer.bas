Attribute VB_Name = "modPlayer"
Option Explicit

Public Sub InitChat(ByVal Index As Long, ByVal MapNum As Long, ByVal mapnpcnum As Long, Optional ByVal remoteChat As Boolean = False)
    Dim NpcNum As Long
    NpcNum = MapNpc(MapNum).NPC(mapnpcnum).Num

    ' check if we can chat
    If NPC(NpcNum).Conv = 0 Then Exit Sub
    If Len(Trim$(Conv(NPC(NpcNum).Conv).Name)) = 0 Then Exit Sub

    If Not remoteChat Then
        With MapNpc(MapNum).NPC(mapnpcnum)
            .c_inChatWith = Index
            .c_lastDir = .dir
            If GetPlayerY(Index) = .Y - 1 Then
                .dir = DIR_UP
            ElseIf GetPlayerY(Index) = .Y + 1 Then
                .dir = DIR_DOWN
            ElseIf GetPlayerX(Index) = .X - 1 Then
                .dir = DIR_LEFT
            ElseIf GetPlayerX(Index) = .X + 1 Then
                .dir = DIR_RIGHT
            End If
            ' send NPC's dir to the map
            NpcDir MapNum, mapnpcnum, .dir
        End With
    End If

    ' Set chat value to Npc
    TempPlayer(Index).inChatWith = NpcNum
    TempPlayer(Index).c_mapNpcNum = mapnpcnum
    TempPlayer(Index).c_mapNum = MapNum
    ' set to the root chat
    TempPlayer(Index).curChat = 1
    ' send the root chat
    sendChat Index
End Sub

Public Sub chatOption(ByVal Index As Long, ByVal chatOption As Long)
    Dim exitChat As Boolean
    Dim convNum As Long
    Dim curChat As Long

    If TempPlayer(Index).inChatWith = 0 Then Exit Sub

    convNum = NPC(TempPlayer(Index).inChatWith).Conv
    curChat = TempPlayer(Index).curChat

    exitChat = False

    ' follow route
    If Conv(convNum).Conv(curChat).rTarget(chatOption) = 0 Then
        exitChat = True
    Else
        TempPlayer(Index).curChat = Conv(convNum).Conv(curChat).rTarget(chatOption)
    End If

    ' if exiting chat, clear temp values
    If exitChat Then
        TempPlayer(Index).inChatWith = 0
        TempPlayer(Index).curChat = 0
        ' send chat update
        sendChat Index
        ' send npc dir
        With MapNpc(TempPlayer(Index).c_mapNum).NPC(TempPlayer(Index).c_mapNpcNum)
            If .c_inChatWith = Index Then
                .c_inChatWith = 0
                .dir = .c_lastDir
                NpcDir TempPlayer(Index).c_mapNum, TempPlayer(Index).c_mapNpcNum, .dir
            End If
        End With
        ' clear last of data
        TempPlayer(Index).c_mapNpcNum = 0
        TempPlayer(Index).c_mapNum = 0
        ' exit out early so we don't send chat update twice
        Exit Sub
    End If

    ' send update to the client
    sendChat Index
End Sub

Public Sub chat_Unique(ByVal Index As Long)
    Dim convNum As Long
    Dim curChat As Long
    Dim itemAmount As Long

    If TempPlayer(Index).inChatWith > 0 Then
        convNum = NPC(TempPlayer(Index).inChatWith).Conv
        curChat = TempPlayer(Index).curChat

        ' is unique?
        If Conv(convNum).Conv(curChat).Event = 4 Then    ' unique
            ' which unique event?
            Select Case Conv(convNum).Conv(curChat).Data1
            Case 1    ' Little Boy
                ' check has the gold
                itemAmount = GetPlayerGold(Index)
                If itemAmount = 0 Or itemAmount < 50 Then
                    PlayerMsg Index, "You do not have enough $.", BrightRed
                    Exit Sub
                Else
                    PlayerWarp Index, 15, 33, 32
                    SetPlayerDir Index, DIR_LEFT
                    Call SetPlayerGold(Index, GetPlayerGold(Index) - 50)
                    Call SendGoldUpdate(Index)
                    PlayerMsg Index, "The boy takes your money then pushes you head first through the hole.", BrightGreen
                    Exit Sub
                End If
            End Select
        End If
    End If
End Sub

Public Sub sendChat(ByVal Index As Long)
    Dim convNum As Long
    Dim curChat As Long
    Dim mainText As String
    Dim optText(1 To 4) As String
    Dim P_GENDER As String
    Dim P_NAME As String
    Dim P_CLASS As String
    Dim I As Long

    If TempPlayer(Index).inChatWith > 0 Then
        convNum = NPC(TempPlayer(Index).inChatWith).Conv
        curChat = TempPlayer(Index).curChat

        ' check for unique events and trigger them early
        If Conv(convNum).Conv(curChat).Event > 0 Then
            Select Case Conv(convNum).Conv(curChat).Event
            Case 1    ' Open Shop
                If Conv(convNum).Conv(curChat).Data1 > 0 Then    ' shop exists?
                    SendOpenShop Index, Conv(convNum).Conv(curChat).Data1
                    TempPlayer(Index).InShop = Conv(convNum).Conv(curChat).Data1    ' stops movement and the like
                End If
                ' exit out early so we don't send chat update twice
                ClosePlayerChat Index
                Exit Sub
            Case 2    ' Open Bank
                SendBank Index
                TempPlayer(Index).InBank = True
                ' exit out early
                ClosePlayerChat Index
                Exit Sub
            Case 3    ' Give Item
                ' exit out early
                ClosePlayerChat Index
                Exit Sub
            Case 4    ' Unique event
                chat_Unique Index
                ClosePlayerChat Index
                Exit Sub
            Case 5    ' Give Start Quest
                If Conv(convNum).Conv(curChat).Data1 > 0 Then
                    If QuestInProgress(Index, Conv(convNum).Conv(curChat).Data1) Then
                        'if the quest is in progress show the meanwhile message (speech2)
                        mainText = Trim$(Quest(Conv(convNum).Conv(curChat).Data1).Task(Player(Index).PlayerQuest(Conv(convNum).Conv(curChat).Data1).ActualTask).TaskLog)
                        SendChatUpdate Index, TempPlayer(Index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
                        Exit Sub
                    End If
                    If CanStartQuest(Index, Conv(convNum).Conv(curChat).Data1) Then
                        'if can start show the request message (speech1)
                        StartQuest Index, Conv(convNum).Conv(curChat).Data1, 1
                        mainText = Trim$(Quest(Conv(convNum).Conv(curChat).Data1).QuestLog)
                        SendChatUpdate Index, TempPlayer(Index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
                        Exit Sub
                    End If
                    mainText = "Voce nao cumpre algum requisito, verifique o log no chat!"
                    SendChatUpdate Index, TempPlayer(Index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
                    Exit Sub
                End If

            Case 6   ' Get ProtectDrop
                If GetPlayerProtectDrop(Index) = NO Then
                    SetPlayerProtectDrop Index, YES
                    GoTo Continue
                Else
                    mainText = "Voce ja possui sua bencao"
                    SendChatUpdate Index, TempPlayer(Index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
                    Exit Sub
                End If
            Case 7   ' Create Guild
                SendGuildWindow Index
                ClosePlayerChat Index
                Exit Sub
            Case 8   ' Claim Serial
                SendSerialWindow Index
                ClosePlayerChat Index
                Exit Sub
            Case 9  ' Lottery
                If VerifyLotteryStatus Then
                    SendLotteryWindow Index
                    ClosePlayerChat Index
                Else
                    mainText = "Loteria nao esta funcionando ainda, volte em " & SecondsToHMS(((LOTTERY_START_HOURS * 60) * 60) - ((getTime - Lottery.Ended) / 1000))
                    SendChatUpdate Index, TempPlayer(Index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
                    Exit Sub
                End If
                Exit Sub
            End Select
        End If

Continue:
        ' cache player's details
        If Player(Index).Sex = SEX_MALE Then
            P_GENDER = "man"
        Else
            P_GENDER = "woman"
        End If
        P_NAME = Trim$(Player(Index).Name)
        P_CLASS = Trim$(Class(Player(Index).Class).Name)

        mainText = Conv(convNum).Conv(curChat).Conv
        For I = 1 To 4
            optText(I) = Conv(convNum).Conv(curChat).rText(I)
        Next
    End If

    SendChatUpdate Index, TempPlayer(Index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
    Exit Sub
End Sub

Public Sub ClosePlayerChat(ByVal Index As Long)
' exit the chat
    TempPlayer(Index).inChatWith = 0
    TempPlayer(Index).curChat = 0
    ' send chat update
    sendChat Index
    ' send npc dir
    With MapNpc(TempPlayer(Index).c_mapNum).NPC(TempPlayer(Index).c_mapNpcNum)
        If .c_inChatWith = Index Then
            .c_inChatWith = 0
            .dir = .c_lastDir
            NpcDir TempPlayer(Index).c_mapNum, TempPlayer(Index).c_mapNpcNum, .dir
        End If
    End With
    ' clear last of data
    TempPlayer(Index).c_mapNpcNum = 0
    TempPlayer(Index).c_mapNum = 0
    Exit Sub
End Sub

Sub UseChar(ByVal Index As Long)
    If Not IsPlaying(Index) Then
        Call JoinGame(Index)
        Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.GAME_NAME & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.GAME_NAME & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal Index As Long)
    Dim I As Long

    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
    
    ' send the CRC32 keys.
    Call SendItemsCRC32(Index)
    Call SendNpcsCRC32(Index)

    ' send the login ok
    SendLoginOk Index

    TotalPlayersOnline = TotalPlayersOnline + 1

    ' Send some more little goodies, no need to explain these
    Call CheckPremium(Index)
    Call CheckEquippedItems(Index)

    Call SendClasses(Index)
    'Call SendItems(Index)
    'Call SendNpcs(Index)
    Call SendAnimations(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendConvs(Index)
    Call SendResources(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    Call SendPlayerSpells(Index)
    Call SendHotbar(Index)
    Call SendPlayerVariables(Index)
    Call SendSerial(Index)
    Call SendDataPremium(Index)
    Call SendQuests(Index)
    Call SendPlayerQuests(Index)
    Call SendClientTimeTo(Index)
    Call SendConjuntos(Index)

    ' Faz a verificação, se o jogador está usando algum conjunto, manda a msg e coloca os bonus como temporários!
    Call CheckConjunto(Index)

    ' Send Spells Cooldown
    For I = 1 To MAX_PLAYER_SPELLS
        If Player(Index).Spell(I).Spell > 0 Then
            If Player(Index).SpellCD(I) > 0 Then
                Call SendPlayerSpellsCD(Index, I)
            End If
        End If
    Next


    ' Atualiza ele como online na guild
    If Player(Index).Guild_ID > 0 Then
        ' Primeiro vemos se ele não foi kickado da guild enquanto estava offline
        If Player(Index).Guild_MembroID > UBound(GuildMembers(Player(Index).Guild_ID).Membro) Then
            Player(Index).Guild_ID = 0
            Player(Index).Guild_MembroID = 0
            PlayerMsg Index, "Sua guild foi desfeita enquanto estava offline.", White
            SavePlayer Index
        Else
            If GuildMembers(Player(Index).Guild_ID).Membro(Player(Index).Guild_MembroID).MembroDisponivel = True Then
                If Guild(Player(Index).Guild_ID).GuildDisponivel = True Then
                    Player(Index).Guild_ID = 0
                    Player(Index).Guild_MembroID = 0
                    PlayerMsg Index, "Sua guild foi desfeita enquanto estava offline.", White
                    SavePlayer Index
                Else
                    Player(Index).Guild_ID = 0
                    Player(Index).Guild_MembroID = 0
                    PlayerMsg Index, "Você foi kickado de sua guild enquanto estava offline!", White
                    SavePlayer Index
                End If
            Else
                ' Seta ele como online na guild
                If Guild(Player(Index).Guild_ID).GuildDisponivel = False Then
                    GuildMembers(Player(Index).Guild_ID).Membro(Player(Index).Guild_MembroID).Online = True

                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) And I <> Index Then
                            If Player(I).Guild_ID = Player(Index).Guild_ID Then
                                SendUpdateGuildTo I, Player(Index).Guild_ID
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End If

    Call SendGuilds(Index)

    ' send vitals, exp + stats
    For I = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, I)
    Next
    SendEXP Index
    Call SendStats(Index)

    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))

    ' Send a global message that he/she joined
    Call GlobalMsg(GetPlayerName(Index) & " has joined " & Options.GAME_NAME & "!", White)

    ' Send welcome messages
    Call SendWelcome(Index)

    ' Send Resource cache
    If GetPlayerMap(Index) > 0 And GetPlayerMap(Index) <= MAX_MAPS Then
        For I = 0 To MapResourceCache(GetPlayerMap(Index)).Resource_Count
            SendResourceCacheTo Index, I
        Next
    End If

    ' Send the flag so they know they can start doing stuff
    SendInGame Index

    SendTimeToBirthday Index
    
    ' Faz parte do CheckIn Diario, envia o dia a ser mostrado na ordem do checkin do jogador!
    ' Verifica se o checkin seguido do jogador não passou do limite máximo de dias do mês!
    If Player(Index).CheckIn < UBound(MonthReward(Month(Date)).DayReward) Then
        ' Verifica se o ultimo checkin do jogador é em uma data diferente!
        If GetPlayerLastCheckIn(Index) <> Date Then
     
            If DateDiff("d", GetPlayerLastCheckIn(Index), Date) > 1 Then 'if (Day(Date) - Day(GetPlayerLastCheckIn(Index))) > 1 Then
                Player(Index).CheckIn = 1
            Else
                Player(Index).CheckIn = Player(Index).CheckIn + 1
            End If

            SendDayReward Index
        End If
    Else
        Player(Index).CheckIn = 0
    End If


    ' tell them to do the damn tutorial
    If Player(Index).TutorialState = 0 Then SendStartTutorial Index
End Sub

Public Sub SendTimeToBirthday(ByVal Index As Long)
    Dim Dia As Integer, Mes As Integer, Ano As Integer
    Dim targetYear As Integer
    Dim DateResult As Date

    DateResult = GetPlayerBirthDay(Index)

    If Not IsDate(Player(Index).BirthDay) Then
        Call PlayerMsg(Index, "Houve um erro com a sua data de aniversário, contate um admin!", BrightRed)
        Exit Sub
    End If

    'Has the birthday already passed this year?
    If Month(Now) > Month(DateResult) Or _
       (Month(Now) = Month(DateResult) And Day(Now) > Day(DateResult)) Then
        'Then use next year.
        targetYear = Year(Now) + 1
    Else
        targetYear = Year(Now)
    End If

    If Ceil(DateSerial(targetYear, Month(DateResult), Day(DateResult)) - Now) = 0 Then
        Call SendMessageTo(Index, "Happy Birthday!!!", "Parabéns, hoje é o dia do seu aniversário, como reconhecimento vamos te disponibilizar um serial, procure o npc que recebe o serial e resgate o seu pacote! " & GetBirthDaySerialNum(Index))
    Else
        Call PlayerMsg(Index, "Faltam: " & CInt(DateSerial(targetYear, Month(DateResult), Day(DateResult)) - Now) & " Dia(s) para o seu aniversário!", Yellow)
    End If
End Sub

Public Function Ceil(valor As Double)

    Ceil = CDbl(CLng(valor + 0.5))

End Function

Public Function GetBirthDaySerialNum(ByVal Index As Long) As String
    Dim I As Integer
    Dim Count As Integer

    For I = 1 To MAX_SERIAL_NUMBER
        If Trim$(Serial(I).Name) <> vbNullString Then
            If Serial(I).BirthDay > 0 Then
                Count = Count + 1
                If Count = 1 Then
                    GetBirthDaySerialNum = "Serial " & Count & ":" & Trim$(Serial(I).Serial)
                Else
                    GetBirthDaySerialNum = GetBirthDaySerialNum & ", Serial " & Count & ":" & Trim$(Serial(I).Serial)
                End If
            End If
        End If
    Next I
End Function

Public Function GetPlayerBirthDay(ByVal Index As Long) As Date
    Dim Dia As Integer, Mes As Integer, Ano As Integer
    
    If Not IsDate(Player(Index).BirthDay) Then Exit Function
    
    GetPlayerBirthDay = Player(Index).BirthDay
    
End Function

Sub LeftGame(ByVal Index As Long)
    Dim n As Long, I As Long
    Dim tradeTarget As Long

    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetPlayerMap(Index) <= 0 Or GetPlayerMap(Index) > MAX_MAPS Then Exit Sub
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If

        ' cancel any trade they're in
        If TempPlayer(Index).InTrade > 0 Then
            tradeTarget = TempPlayer(Index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For I = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(I).Num = 0
                TempPlayer(tradeTarget).TradeOffer(I).Value = 0
            Next
            TempPlayer(tradeTarget).TradeGold = 0
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If

        ' leave party.
        Party_PlayerLeave Index

        ' Atualiza ele como offline na guild
        If Player(Index).Guild_ID > 0 Then
            If Guild(Player(Index).Guild_ID).GuildDisponivel = False Then
                GuildMembers(Player(Index).Guild_ID).Membro(Player(Index).Guild_MembroID).Online = False

                For I = 1 To Player_HighIndex
                    If IsPlaying(I) And I <> Index Then
                        If Player(I).Guild_ID = Player(Index).Guild_ID Then
                            SendUpdateGuildTo I, Player(Index).Guild_ID
                        End If
                    End If
                Next
            End If
        End If

        ' save and clear data.
        Call SavePlayer(Index)

        ' Send a global message that he/she left
        Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.GAME_NAME & "!", White)

        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & Options.GAME_NAME & ".")
        Call SendLeftGame(Index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(Index)
End Sub

Function GetPlayerProtection(ByVal Index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipmentNum(Index, Armor)
    Helm = GetPlayerEquipmentNum(Index, Helmet)
    GetPlayerProtection = (GetPlayerStat(Index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

End Function

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim ShopNum As Long
    Dim OldMap As Long
    Dim I As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If X > Map(MapNum).MapData.MaxX Then X = Map(MapNum).MapData.MaxX
    If Y > Map(MapNum).MapData.MaxY Then Y = Map(MapNum).MapData.MaxY
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0

    ' if same map then just send their co-ordinates
    If MapNum = GetPlayerMap(Index) Then
        SendPlayerXYToMap Index
        ' check tasks
        Call CheckTasks(Index, QUEST_TYPE_GOREACH, MapNum)
    End If

    ' clear target
    TempPlayer(Index).target = 0
    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
    SendTarget Index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If OldMap <> MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    ' send player's equipment to new map
    SendMapEquipment Index

    ' send equipment of all people on new map
    If GetTotalMapPlayers(MapNum) > 0 Then
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                If GetPlayerMap(I) = MapNum Then
                    SendMapEquipmentTo I, Index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
        ' Regenerate all NPCs' Health and Spirit
        For I = 1 To MAX_MAP_NPCS
            If MapNpc(OldMap).NPC(I).Num > 0 Then
                MapNpc(OldMap).NPC(I).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).NPC(I).Num, Vitals.HP)
                MapNpc(OldMap).NPC(I).Vital(Vitals.MP) = GetNpcMaxVital(MapNpc(OldMap).NPC(I).Num, Vitals.MP)
            End If
        Next
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    TempPlayer(Index).GettingMap = YES
    Call CheckTasks(Index, QUEST_TYPE_GOREACH, MapNum)
    SendCheckForMap Index, MapNum
End Sub

Function CanMove(Index As Long, dir As Long) As Byte
    Dim warped As Boolean, newMapX As Long, newMapY As Long

    CanMove = 1
    Select Case dir
    Case DIR_UP
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(Index) > 0 Then
            If CheckDirection(Index, DIR_UP) Then
                CanMove = 0
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(Index)).MapData.Up > 0 Then
                newMapY = Map(Map(GetPlayerMap(Index)).MapData.Up).MapData.MaxY
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Up, GetPlayerX(Index), newMapY)
                warped = True
                CanMove = 2
            End If
            CanMove = 0
            Exit Function
        End If
    Case DIR_DOWN
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(Index) < Map(GetPlayerMap(Index)).MapData.MaxY Then
            If CheckDirection(Index, DIR_DOWN) Then
                CanMove = 0
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(Index)).MapData.Down > 0 Then
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Down, GetPlayerX(Index), 0)
                warped = True
                CanMove = 2
            End If
            CanMove = False
            Exit Function
        End If
    Case DIR_LEFT
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(Index) > 0 Then
            If CheckDirection(Index, DIR_LEFT) Then
                CanMove = 0
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(Index)).MapData.left > 0 Then
                newMapX = Map(Map(GetPlayerMap(Index)).MapData.left).MapData.MaxX
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.left, newMapX, GetPlayerY(Index))
                warped = True
                CanMove = 2
            End If
            CanMove = False
            Exit Function
        End If
    Case DIR_RIGHT
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(Index) < Map(GetPlayerMap(Index)).MapData.MaxX Then
            If CheckDirection(Index, DIR_RIGHT) Then
                CanMove = False
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(Index)).MapData.Right > 0 Then
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).MapData.Right, 0, GetPlayerY(Index))
                warped = True
                CanMove = 2
            End If
            CanMove = False
            Exit Function
        End If
    End Select
    ' check if we've warped
    If warped Then
        ' clear their target
        TempPlayer(Index).target = 0
        TempPlayer(Index).TargetType = TARGET_TYPE_NONE
        SendTarget Index
    End If
End Function

Function CheckDirection(Index As Long, direction As Long) As Boolean
    Dim X As Long, Y As Long, I As Long, EventCount As Long, MapNum As Long, page As Long

    CheckDirection = False

    Select Case direction
    Case DIR_UP
        X = GetPlayerX(Index)
        Y = GetPlayerY(Index) - 1
    Case DIR_DOWN
        X = GetPlayerX(Index)
        Y = GetPlayerY(Index) + 1
    Case DIR_LEFT
        X = GetPlayerX(Index) - 1
        Y = GetPlayerY(Index)
    Case DIR_RIGHT
        X = GetPlayerX(Index) + 1
        Y = GetPlayerY(Index)
    End Select

    ' Check to see if the map tile is blocked or not
    If Map(GetPlayerMap(Index)).TileData.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map(GetPlayerMap(Index)).TileData.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to make sure that any events on that space aren't blocked
    MapNum = GetPlayerMap(Index)
    EventCount = Map(MapNum).TileData.EventCount
    For I = 1 To EventCount
        With Map(MapNum).TileData.Events(I)
            If .X = X And .Y = Y Then
                ' Get the active event page
                page = ActiveEventPage(Index, I)
                If page > 0 Then
                    If Map(MapNum).TileData.Events(I).EventPage(page).WalkThrough = 0 Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End With
    Next

    ' Check to see if a player is already on that tile
    If Map(GetPlayerMap(Index)).MapData.Moral = 0 Then
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(Index) Then
                If GetPlayerX(I) = X Then
                    If GetPlayerY(I) = Y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next I
    End If

    ' Check to see if a npc is already on that tile
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index)).NPC(I).Num > 0 Then
            If MapNpc(GetPlayerMap(Index)).NPC(I).X = X Then
                If MapNpc(GetPlayerMap(Index)).NPC(I).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Sub PlayerMove(ByVal Index As Long, ByVal dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, MapNum As Long, X As Long, Y As Long, moved As Byte, MovedSoFar As Boolean, newMapX As Byte, newMapY As Byte
    Dim TileType As Long, vitalType As Long, colour As Long, Amount As Long, canMoveResult As Long, I As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or dir < DIR_UP Or dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, dir)
    moved = NO
    MapNum = GetPlayerMap(Index)

    If MapNum = 0 Then Exit Sub

    ' check if they're casting a spell
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        If Spell(TempPlayer(Index).spellBuffer.Spell).CanRun = NO Then
            SendCancelAnimation Index
            SendClearSpellBuffer Index
            TempPlayer(Index).spellBuffer.Spell = 0
            TempPlayer(Index).spellBuffer.target = 0
            TempPlayer(Index).spellBuffer.Timer = 0
            TempPlayer(Index).spellBuffer.tType = 0
        End If
    End If

    ' check directions
    canMoveResult = CanMove(Index, dir)
    If canMoveResult = 1 Then
        Select Case dir
        Case DIR_UP
            Call SetPlayerY(Index, GetPlayerY(Index) - 1)
            SendPlayerMove Index, movement, sendToSelf
            moved = YES
        Case DIR_DOWN
            Call SetPlayerY(Index, GetPlayerY(Index) + 1)
            SendPlayerMove Index, movement, sendToSelf
            moved = YES
        Case DIR_LEFT
            Call SetPlayerX(Index, GetPlayerX(Index) - 1)
            SendPlayerMove Index, movement, sendToSelf
            moved = YES
        Case DIR_RIGHT
            Call SetPlayerX(Index, GetPlayerX(Index) + 1)
            SendPlayerMove Index, movement, sendToSelf
            moved = YES
        End Select
    End If

    With Map(GetPlayerMap(Index)).TileData.Tile(GetPlayerX(Index), GetPlayerY(Index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            MapNum = .Data1
            X = .Data2
            Y = .Data3
            Call PlayerWarp(Index, MapNum, X, Y)
            moved = YES
        End If

        ' Check to see if the tile is a door tile, and if so warp them
        If .Type = TILE_TYPE_DOOR Then
            MapNum = .Data1
            X = .Data2
            Y = .Data3
            ' send the animation to the map
            SendDoorAnimation GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
            Call PlayerWarp(Index, MapNum, X, Y)
            moved = YES
        End If

        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            X = .Data1
            Y = .Data2

            If Map(GetPlayerMap(Index)).TileData.Tile(X, Y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = getTime
                SendMapKey Index, X, Y, 1
                'Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
            End If
        End If

        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            X = .Data1
            If X > 0 Then    ' shop exists?
                If Len(Trim$(Shop(X).Name)) > 0 Then    ' name exists?
                    SendOpenShop Index, X
                    TempPlayer(Index).InShop = X    ' stops movement and the like
                End If
            End If
        End If

        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank Index
            TempPlayer(Index).InBank = True
            moved = YES
        End If

        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            vitalType = .Data1
            Amount = .Data2
            If Not GetPlayerVital(Index, vitalType) = GetPlayerMaxVital(Index, vitalType) Then
                If vitalType = Vitals.HP Then
                    colour = BrightGreen
                Else
                    colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(Index), "+" & Amount, colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
                SetPlayerVital Index, vitalType, GetPlayerVital(Index, vitalType) + Amount
                PlayerMsg Index, "You feel rejuvinating forces flowing through your boy.", BrightGreen
                Call SendMapVitals(Index)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            moved = YES
        End If

        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            Amount = .Data1
            SendActionMsg GetPlayerMap(Index), "-" & Amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
            If GetPlayerVital(Index, HP) - Amount <= 0 Then
                KillPlayer Index
                PlayerMsg Index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital Index, HP, GetPlayerVital(Index, HP) - Amount
                PlayerMsg Index, "You're injured by a trap.", BrightRed
                Call SendMapVitals(Index)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            moved = YES
        End If
    End With

    ' check for events
    If Map(GetPlayerMap(Index)).TileData.EventCount > 0 Then
        For I = 1 To Map(GetPlayerMap(Index)).TileData.EventCount
            CheckPlayerEvent Index, I
        Next
    End If

    ' They tried to hack
    If moved = NO And canMoveResult <> 2 Then
        PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
    End If
End Sub

Sub CheckPlayerEvent(Index As Long, eventNum As Long)
    Dim Count As Long, MapNum As Long, I As Long
    ' find the page to process
    MapNum = GetPlayerMap(Index)
    ' make sure it's in the same spot
    If Map(MapNum).TileData.Events(eventNum).X <> GetPlayerX(Index) Then Exit Sub
    If Map(MapNum).TileData.Events(eventNum).Y <> GetPlayerY(Index) Then Exit Sub
    ' loop
    Count = Map(MapNum).TileData.Events(eventNum).PageCount
    ' get the active page
    I = ActiveEventPage(Index, eventNum)
    ' exit out early
    If I = 0 Then Exit Sub
    ' make sure the page has actual commands
    If Map(MapNum).TileData.Events(eventNum).EventPage(I).CommandCount = 0 Then Exit Sub
    ' set event
    TempPlayer(Index).inEvent = True
    TempPlayer(Index).eventNum = eventNum
    TempPlayer(Index).pageNum = I
    TempPlayer(Index).commandNum = 1
    ' send it to the player
    SendEvent Index
End Sub

Sub EventLogic(Index As Long)
    Dim eventNum As Long, pageNum As Long, commandNum As Long
    eventNum = TempPlayer(Index).eventNum
    pageNum = TempPlayer(Index).pageNum
    commandNum = TempPlayer(Index).commandNum
    ' carry out the command
    With Map(GetPlayerMap(Index)).TileData.Events(eventNum).EventPage(pageNum)
        ' server-side processing
        Select Case .Commands(commandNum).Type
        Case EventType.evPlayerVar
            If .Commands(commandNum).target > 0 Then Player(Index).Variable(.Commands(commandNum).target) = .Commands(commandNum).colour
        End Select
        ' increment commands
        If commandNum < .CommandCount Then
            TempPlayer(Index).commandNum = TempPlayer(Index).commandNum + 1
            Exit Sub
        End If
    End With
    ' we're done - close event
    TempPlayer(Index).eventNum = 0
    TempPlayer(Index).pageNum = 0
    TempPlayer(Index).commandNum = 0
    TempPlayer(Index).inEvent = False
    ' send it to the player
    'SendEvent index
End Sub

Sub ForcePlayerMove(ByVal Index As Long, ByVal movement As Long, ByVal direction As Long)
    If direction < DIR_UP Or direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub

    Select Case direction
    Case DIR_UP
        If GetPlayerY(Index) = 0 Then Exit Sub
    Case DIR_LEFT
        If GetPlayerX(Index) = 0 Then Exit Sub
    Case DIR_DOWN
        If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MapData.MaxY Then Exit Sub
    Case DIR_RIGHT
        If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MapData.MaxX Then Exit Sub
    End Select

    PlayerMove Index, direction, movement, True
End Sub

Sub CheckEquippedItems(ByVal Index As Long)
    Dim Slot As Long
    Dim ItemNum As Long
    Dim I As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For I = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipmentNum(Index, I)

        If ItemNum > 0 Then

            Select Case I
            Case Equipment.Weapon

                If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment Index, 0, I
            Case Equipment.Armor

                If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment Index, 0, I
            Case Equipment.Helmet

                If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment Index, 0, I
            Case Equipment.Shield

                If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment Index, 0, I
            Case Equipment.Legs

                If Item(ItemNum).Type <> ITEM_TYPE_LEGS Then SetPlayerEquipment Index, 0, I
            Case Equipment.Boots

                If Item(ItemNum).Type <> ITEM_TYPE_BOOTS Then SetPlayerEquipment Index, 0, I
            Case Equipment.Amulet

                If Item(ItemNum).Type <> ITEM_TYPE_AMULET Then SetPlayerEquipment Index, 0, I
            Case Equipment.RingLeft

                If Item(ItemNum).Type <> ITEM_TYPE_RINGLEFT Then SetPlayerEquipment Index, 0, I
            Case Equipment.RingLeft

                If Item(ItemNum).Type <> ITEM_TYPE_RINGRIGHT Then SetPlayerEquipment Index, 0, I
            End Select

        Else
            SetPlayerEquipment Index, 0, I
        End If

    Next

End Sub

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim I As Long

    For I = 1 To MAX_PLAYER_SPELLS

        If Player(Index).Spell(I).Spell = SpellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim I As Long

    For I = 1 To MAX_PLAYER_SPELLS

        If Player(Index).Spell(I).Spell = 0 Then
            FindOpenSpellSlot = I
            Exit Function
        End If

    Next

End Function

Sub CheckPlayerLevelUp(ByVal Index As Long, Optional ByVal level_count As Long)
    Dim I As Long, PontosPorLevel As Byte
    Dim expRollover As Long

    PontosPorLevel = 3

    ' Caso queira adicionar levels diretamente!
    If level_count > 0 Then
        ' can level up?
        If Not SetPlayerLevel(Index, GetPlayerLevel(Index) + level_count) Then
            Exit Sub
        End If

        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + (level_count * PontosPorLevel))
        GoTo Continue
    End If

    ' Adiciona level pela experiência, método normal de um rpg
    level_count = 0
    Do While GetPlayerExp(Index) >= GetPlayerNextLevel(Index)
        expRollover = GetPlayerExp(Index) - GetPlayerNextLevel(Index)

        ' can level up?
        If Not SetPlayerLevel(Index, GetPlayerLevel(Index) + 1) Then
            Exit Sub
        End If

        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + PontosPorLevel)
        Call SetPlayerExp(Index, expRollover)
        level_count = level_count + 1
    Loop

Continue:
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " level!", Brown
        Else
            'plural
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " levels!", Brown
        End If
        SendEXP Index
        SendPlayerData Index
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerName(ByVal Index As Long) As String

    If Index > Player_HighIndex Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    If Index <= 0 Or Index > Player_HighIndex Then Exit Function
    GetPlayerLevel = Player(Index).Level
End Function

Function SetPlayerLevel(ByVal Index As Long, ByVal Level As Long) As Boolean
    If Index <= 0 Or Index > Player_HighIndex Then Exit Function
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then
        Player(Index).Level = MAX_LEVELS
        Exit Function
    End If
    Player(Index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = 100 + (((GetPlayerLevel(Index) ^ 2) * 10) * 2)
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    If Index <= 0 Or Index > Player_HighIndex Then Exit Function
    GetPlayerExp = Player(Index).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    If Index <= 0 Or Index > Player_HighIndex Then Exit Sub
    Player(Index).EXP = EXP
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > Player_HighIndex Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    If GetPlayerVital(Index, Vital) < 0 Then
        Player(Index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Dim X As Long, I As Long
    If Index > Player_HighIndex Then Exit Function

    X = Player(Index).Stat(Stat)

    For I = 1 To Equipment.Equipment_Count - 1
        If Player(Index).Equipment(I).Num > 0 Then
            If Item(Player(Index).Equipment(I).Num).Add_Stat(Stat) > 0 Then
                If Item(Player(Index).Equipment(I).Num).Stat_Percent(Stat) > 0 Then
                    X = X + ((Player(Index).Stat(Stat) / 100) * Item(Player(Index).Equipment(I).Num).Add_Stat(Stat))
                Else
                    X = X + Item(Player(Index).Equipment(I).Num).Add_Stat(Stat)
                End If
            End If
        End If
    Next

    If TempPlayer(Index).Bonus.Add_Stat(Stat) > 0 Then
        X = X + TempPlayer(Index).Bonus.Add_Stat(Stat)
    End If

    GetPlayerStat = X
End Function

Public Function GetPlayerRawStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    If Index > Player_HighIndex Then Exit Function

    GetPlayerRawStat = Player(Index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(Index).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    If POINTS <= 0 Then POINTS = 0
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long

    If Index <= 0 Or Index > Player_HighIndex Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Map = MapNum
    End If

End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    If Index <= 0 Or Index > Player_HighIndex Then Exit Function
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    If Index <= 0 Or Index > Player_HighIndex Then Exit Sub
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    If Index <= 0 Or Index > Player_HighIndex Then Exit Function
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    If Index <= 0 Or Index > Player_HighIndex Then Exit Sub
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerDir = Player(Index).dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal dir As Long)
    Player(Index).dir = dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String

    If Index > Player_HighIndex Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

' ToDo
Sub OnDeath(ByVal Index As Long)
    Dim I As Long
    Dim Count As Long
    Dim Random As Byte

    ' Set HP to nothing
    Call SetPlayerVital(Index, Vitals.HP, 0)
    SendVital Index, HP

    ' Drop all worn items
    ' Verifica se tem proteção divina!
    If GetPlayerProtectDrop(Index) = NO Then
        ' dropa os items equipados
        For I = 1 To Equipment.Equipment_Count - 1
            If GetPlayerEquipmentNum(Index, I) > 0 Then
                If Item(GetPlayerEquipmentNum(Index, I)).DropDead > 0 Then
                    ' Random Drop Chance
                    Random = Rand(0, 100)
                    If Item(GetPlayerEquipmentNum(Index, I)).DropDeadChance >= Random Then
                        DropItemOnDead Index, GetPlayerEquipmentNum(Index, I), 1, True
                        Count = Count + 1
                    End If
                End If
            End If
        Next

        ' Verifica se algum item do conjunto foi removido e então retira bonus!
        Call CheckConjunto(Index)

        ' dropa os items da bolsa
        For I = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, I) > 0 Then
                If Item(GetPlayerInvItemNum(Index, I)).DropDead > 0 Then
                    ' Random Drop Chance
                    Random = Rand(0, 100)
                    If Item(GetPlayerInvItemNum(Index, I)).DropDeadChance >= Random Then
                        DropItemOnDead Index, GetPlayerInvItemNum(Index, I), GetPlayerInvItemValue(Index, I)
                        Count = Count + 1
                    End If
                End If
            End If
        Next I

        Call PlayerMsg(Index, "Você estava sem proteção divina e perdeu " & Count & " pertences!", BrightRed)

    ElseIf GetPlayerProtectDrop(Index) = YES Then
        Call SetPlayerProtectDrop(Index, NO)
        Call PlayerMsg(Index, "Seus pertences foram protegidos e sua benção consumida!", BrightGreen)

    End If
    ' Verifica se tem proteção divina!
    If GetPlayerProtectDrop(Index) = YES Then
        Call SetPlayerProtectDrop(Index, NO)
        Call PlayerMsg(Index, "Seus pertences foram protegidos e sua benção consumida!", BrightGreen)
    End If

    ' Warp player away
    Call SetPlayerDir(Index, DIR_DOWN)

    With Map(GetPlayerMap(Index)).MapData
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp Index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(Index, Options.START_MAP, Options.START_X, Options.START_Y)
        End If
    End With

    ' clear all DoTs and HoTs
    For I = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(I)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With

        With TempPlayer(Index).HoT(I)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next

    ' Clear spell casting
    TempPlayer(Index).spellBuffer.Spell = 0
    TempPlayer(Index).spellBuffer.Timer = 0
    TempPlayer(Index).spellBuffer.target = 0
    TempPlayer(Index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(Index)

    ' Restore vitals
    Call SetPlayerVital(Index, Vitals.HP, GetPlayerMaxVital(Index, Vitals.HP))
    Call SetPlayerVital(Index, Vitals.MP, GetPlayerMaxVital(Index, Vitals.MP))
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    Call SendMapVitals(Index)

    ' send vitals to party if in one
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
    End If

    Call SavePlayer(Index)
End Sub

Sub CheckResource(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim I As Long
    Dim damage As Long

    If Map(GetPlayerMap(Index)).TileData.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(Index)).TileData.Tile(X, Y).Data1

        ' Get the cache number
        For I = 0 To MapResourceCache(GetPlayerMap(Index)).Resource_Count

            If MapResourceCache(GetPlayerMap(Index)).ResourceData(I).X = X Then
                If MapResourceCache(GetPlayerMap(Index)).ResourceData(I).Y = Y Then
                    Resource_num = I
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipmentNum(Index, Weapon) > 0 Then
                If Item(GetPlayerEquipmentNum(Index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Or Item(GetPlayerEquipmentNum(Index, Weapon)).Data3 = 0 Then

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(Index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If MapResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 0 Then

                        rX = MapResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).X
                        rY = MapResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).Y

                        damage = Item(GetPlayerEquipmentNum(Index, Weapon)).Data2

                        ' check if damage is more than health
                        If damage > 0 Then
                            ' cut it down!
                            If MapResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - damage <= 0 Then
                                SendActionMsg GetPlayerMap(Index), "-" & MapResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                MapResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 1    ' Cut
                                MapResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceTimer = getTime
                                SendResourceCacheToMap GetPlayerMap(Index), Resource_num
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                End If
                                ' carry on
                                GiveInvItem Index, Resource(Resource_index).ItemReward, 1, 0
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            Else
                                ' just do the damage
                                MapResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health = MapResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - damage
                                SendActionMsg GetPlayerMap(Index), "-" & damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            End If
                            ' send the sound
                            SendMapSound Index, rX, rY, SoundEntity.seResource, Resource_index
                            ' check tasks
                            Call CheckTasks(Index, QUEST_TYPE_GOTRAIN, Resource_index)
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(Index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                        End If
                    End If

                Else
                    PlayerMsg Index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                PlayerMsg Index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Public Sub KillPlayer(ByVal Index As Long)
    Dim EXP As Long

    ' Calculate exp to give attacker
    EXP = GetPlayerExp(Index) \ 3

    ' Make sure we dont get less then 0
    If EXP < 0 Then EXP = 0
    If EXP = 0 Then
        Call PlayerMsg(Index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(Index, GetPlayerExp(Index) - EXP)
        SendEXP Index
        Call PlayerMsg(Index, "You lost " & EXP & " exp.", BrightRed)
    End If

    Call OnDeath(Index)
End Sub
