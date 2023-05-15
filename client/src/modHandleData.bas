Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SSendMapHpMp) = GetAddress(AddressOf HandleMapHpMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SChatUpdate) = GetAddress(AddressOf HandleChatUpdate)
    HandleDataSub(SConvEditor) = GetAddress(AddressOf HandleConvEditor)
    HandleDataSub(SUpdateConv) = GetAddress(AddressOf HandleUpdateConv)
    HandleDataSub(SStartTutorial) = GetAddress(AddressOf HandleStartTutorial)
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    HandleDataSub(SSetPlayerLoginToken) = GetAddress(AddressOf HandleSetPlayerLoginToken)
    HandleDataSub(SCancelAnimation) = GetAddress(AddressOf HandleCancelAnimation)
    HandleDataSub(SPlayerVariables) = GetAddress(AddressOf HandlePlayerVariables)
    HandleDataSub(SEvent) = GetAddress(AddressOf HandleEvent)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(SPlayerBankUpdate) = GetAddress(AddressOf HandlePlayerBankUpdate)
    'Guild
    HandleDataSub(SGuildWindow) = GetAddress(AddressOf HandleGuildWindow)
    HandleDataSub(SUpdateGuild) = GetAddress(AddressOf HandleUpdateGuild)
    HandleDataSub(SGuildInvite) = GetAddress(AddressOf HandleGuildInvite)
    'Serial
    HandleDataSub(SSerialWindow) = GetAddress(AddressOf HandleSerialWindow)
    HandleDataSub(SUpdateSerial) = GetAddress(AddressOf HandleUpdateSerial)
    HandleDataSub(SSerialEditor) = GetAddress(AddressOf HandleSerialEditor)
    'Premium
    HandleDataSub(SPlayerDPremium) = GetAddress(AddressOf HandlePlayerDPremium)
    HandleDataSub(SPremiumEditor) = GetAddress(AddressOf HandlePremiumEditor)
    'Quest
    HandleDataSub(SQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(SUpdateQuest) = GetAddress(AddressOf HandleUpdateQuest)
    HandleDataSub(SPlayerQuest) = GetAddress(AddressOf HandlePlayerQuest)
    HandleDataSub(SQuestMessage) = GetAddress(AddressOf HandleQuestMessage)
    HandleDataSub(SQuestCancel) = GetAddress(AddressOf HandleQuestCancel)
    'Status Animated
    HandleDataSub(SStatus) = GetAddress(AddressOf HandleStatusPlayer)
    ' Client Time
    HandleDataSub(SClientTime) = GetAddress(AddressOf HandleClientTime)
    ' Message Window
    HandleDataSub(SMessage) = GetAddress(AddressOf HandleMessageWindow)
    ' Conjuntos
    HandleDataSub(SConjuntoEditor) = GetAddress(AddressOf HandleConjuntoEditor)
    HandleDataSub(SUpdateConjunto) = GetAddress(AddressOf HandleUpdateConjunto)
    HandleDataSub(SUpdateConjuntoWindow) = GetAddress(AddressOf HandleUpdateConjuntoWindow)
    ' Day Rewards
    HandleDataSub(SSendDayReward) = GetAddress(AddressOf HandleUpdateDayReward)
    ' Cache CRC
    HandleDataSub(SCheckItemCRC) = GetAddress(AddressOf HandleItemsCRC)
    HandleDataSub(SCheckNpcCRC) = GetAddress(AddressOf HandleNpcsCRC)
    ' Lottery
    HandleDataSub(SLotteryWindow) = GetAddress(AddressOf HandleLotteryWindow)
    HandleDataSub(SGoldUpdate) = GetAddress(AddressOf HandleGoldUpdate)
    HandleDataSub(SLotteryInfo) = GetAddress(AddressOf HandleLotteryInfo)

    ' Event Msg
    HandleDataSub(SEventMsg) = GetAddress(AddressOf HandleEventMessage)
End Sub

Sub HandleData(ByRef data() As Byte)
    Dim TempBuffer() As Byte

    TempBuffer = DecryptPacket(data, (UBound(data) - LBound(data)) + 1)

    Dim buffer As clsBuffer
    Dim MsgType As Long
    Set buffer = New clsBuffer

    buffer.WriteBytes TempBuffer

    MsgType = buffer.ReadLong

    If MsgType <= 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If

    CallWindowProc HandleDataSub(MsgType), 1, buffer.ReadBytes(buffer.length), 0, 0
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, dialogue_index As Long, menuReset As Long, kick As Long

    Set buffer = New clsBuffer

    buffer.WriteBytes data()
    dialogue_index = buffer.ReadLong
    menuReset = buffer.ReadLong
    kick = buffer.ReadLong

    Set buffer = Nothing

    If menuReset > 0 Or inMenu = True Then
        HideWindows
        Select Case menuReset
        Case MenuCount.menuLogin
            ShowWindow GetWindowIndex("winLogin")
        Case MenuCount.menuClass
            ShowWindow GetWindowIndex("winClasses")
        Case MenuCount.menuNewChar
            ShowWindow GetWindowIndex("winNewChar")
        Case MenuCount.menuMain
            ShowWindow GetWindowIndex("winLogin")
        Case MenuCount.menuRegister
            ShowWindow GetWindowIndex("winRegister")
        End Select
    Else
        If kick > 0 Then
            HideWindows
            ShowWindow GetWindowIndex("winLogin")
            DialogueAlert dialogue_index
            logoutGame
            Exit Sub
        End If
    End If

    DialogueAlert dialogue_index
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ' Now we can receive game data
    MyIndex = buffer.ReadLong
    ' player high index
    Player_HighIndex = buffer.ReadLong
    Set buffer = Nothing
    Call SetStatus("Receiving game data.")

    diaIndex = 0

    ' close the reconnect window
    If Windows(GetWindowIndex("winReconnect")).Window.visible Then HideWindow GetWindowIndex("winReconnect")
    isReconnect = False
End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, X As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = buffer.ReadString
            .Vital(Vitals.HP) = buffer.ReadLong
            .Vital(Vitals.MP) = buffer.ReadLong
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next

            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    ShowClasses

    Set buffer = Nothing
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, X As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong    'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = buffer.ReadString    'Trim$(Parse(n))
            .Vital(Vitals.HP) = buffer.ReadLong    'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = buffer.ReadLong    'CLng(Parse(n + 2))
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next

            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    Set buffer = Nothing
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InGame = True
    inMenu = False
    SetStatus vbNullString
    ' show gui
    ShowWindow GetWindowIndex("winBars"), , False
    ShowWindow GetWindowIndex("winMenu"), , False
    ShowWindow GetWindowIndex("winHotbar"), , False
    ShowWindow GetWindowIndex("winChatSmall"), , False
    ' enter loop
    GameLoop
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, buffer.ReadLong)
        PlayerInv(i).bound = buffer.ReadByte
    Next

    SetGoldLabel

    Set buffer = Nothing
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadByte    'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, buffer.ReadInteger)    'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, buffer.ReadLong)    'CLng(Parse(3)))
    PlayerInv(n).bound = buffer.ReadByte
    Set buffer = Nothing
    SetGoldLabel
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Call SetPlayerEquipment(MyIndex, buffer.ReadInteger, Armor)
    Call SetPlayerEquipment(MyIndex, buffer.ReadInteger, Weapon)
    Call SetPlayerEquipment(MyIndex, buffer.ReadInteger, Helmet)
    Call SetPlayerEquipment(MyIndex, buffer.ReadInteger, Shield)
    Call SetPlayerEquipment(MyIndex, buffer.ReadInteger, Legs)
    Call SetPlayerEquipment(MyIndex, buffer.ReadInteger, Boots)
    Call SetPlayerEquipment(MyIndex, buffer.ReadInteger, Amulet)
    Call SetPlayerEquipment(MyIndex, buffer.ReadInteger, RingLeft)
    Call SetPlayerEquipment(MyIndex, buffer.ReadInteger, RingRight)
    Set buffer = Nothing
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim playerNum As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    playerNum = buffer.ReadLong
    Call SetPlayerEquipment(playerNum, buffer.ReadInteger, Armor)
    Call SetPlayerEquipment(playerNum, buffer.ReadInteger, Weapon)
    Call SetPlayerEquipment(playerNum, buffer.ReadInteger, Helmet)
    Call SetPlayerEquipment(playerNum, buffer.ReadInteger, Shield)
    Call SetPlayerEquipment(playerNum, buffer.ReadInteger, Legs)
    Call SetPlayerEquipment(playerNum, buffer.ReadInteger, Boots)
    Call SetPlayerEquipment(playerNum, buffer.ReadInteger, Amulet)
    Call SetPlayerEquipment(playerNum, buffer.ReadInteger, RingLeft)
    Call SetPlayerEquipment(playerNum, buffer.ReadInteger, RingRight)
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    If MyIndex = 0 Then Exit Sub
    buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.HP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, buffer.ReadLong)
    ' set max width
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 Then
        BarWidth_GuiHP_Max = ((GetPlayerVital(MyIndex, Vitals.HP) / 209) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / 209)) * 209
    Else
        BarWidth_GuiHP_Max = 0
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.MP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, buffer.ReadLong)
    ' set max width
    If GetPlayerVital(MyIndex, Vitals.MP) > 0 Then
        BarWidth_GuiSP_Max = ((GetPlayerVital(MyIndex, Vitals.MP) / 209) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / 209)) * 209
    Else
        BarWidth_GuiSP_Max = 0
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandleMapHpMp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim i As Byte
    Dim pIndex As Byte

    buffer.WriteBytes data()

    pIndex = buffer.ReadByte

    For i = 1 To Vitals.Vital_Count - 1
        Player(pIndex).MaxVital(i) = buffer.ReadLong
        Call SetPlayerVital(pIndex, i, buffer.ReadLong)
    Next i

    ' Update enemy bars
    If myTargetType = TARGET_TYPE_PLAYER And myTarget = pIndex Then
        UpdateEnemyBars
    End If
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, i, buffer.ReadLong
    Next

    UpdateStatsWindow
End Sub

Private Sub UpdateStatsWindow()
    Dim i As Long, X As Long

    i = MyIndex
    With Windows(GetWindowIndex("winCharacter"))
        For X = 1 To Stats.Stat_Count - 1
            If CLng(.Controls(GetControlIndex("winCharacter", "lblStat_" & X)).text) <> GetPlayerStat(i, X) Then
                .Controls(GetControlIndex("winCharacter", "lblStat_" & X)).text = GetPlayerStat(i, X)
            End If
        Next X
    End With
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    SetPlayerExp MyIndex, buffer.ReadLong
    TNL = buffer.ReadLong
    ' set max width
    If GetPlayerLevel(MyIndex) <= MAX_LEVELS Then
        If GetPlayerExp(MyIndex) > 0 Then
            BarWidth_GuiEXP_Max = ((GetPlayerExp(MyIndex) / 209) / (TNL / 209)) * 209
        Else
            BarWidth_GuiEXP_Max = 0
        End If
    Else
        BarWidth_GuiEXP_Max = 209
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, X As Long, StatusOn As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    Call SetPlayerName(i, buffer.ReadString)
    Call SetPlayerLevel(i, buffer.ReadLong)
    Call SetPlayerPOINTS(i, buffer.ReadLong)
    Call SetPlayerSprite(i, buffer.ReadLong)
    Call SetPlayerMap(i, buffer.ReadLong)
    Call SetPlayerX(i, buffer.ReadLong)
    Call SetPlayerY(i, buffer.ReadLong)
    Call SetPlayerDir(i, buffer.ReadLong)
    Call SetPlayerAccess(i, buffer.ReadLong)
    Call SetPlayerPK(i, buffer.ReadLong)
    Call SetPlayerClass(i, buffer.ReadLong)
    Player(i).Guild_ID = buffer.ReadInteger
    Player(i).Guild_MembroID = buffer.ReadByte
    Player(i).Premium = buffer.ReadByte
    Call SetPlayerGold(i, buffer.ReadLong)

    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, buffer.ReadLong
    Next X

    For X = 1 To Vitals.Vital_Count - 1
        SetPlayerMaxVital i, X, buffer.ReadLong
        SetPlayerVital i, X, buffer.ReadLong
    Next X

    For X = 1 To (status_count - 1)
        StatusOn = buffer.ReadByte
        If StatusOn > 0 Then
            Player(i).StatusNum(StatusOn).Ativo = YES
        End If
    Next X

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        upDown = False
        leftDown = False
        downDown = False
        rightDown = False

        SetaUp = False
        SetaDown = False
        SetaLeft = False
        SetaRight = False
        ' set form
        With Windows(GetWindowIndex("winCharacter"))
            .Controls(GetControlIndex("winCharacter", "lblName")).text = "Name: " & Trim$(GetPlayerName(i))
            .Controls(GetControlIndex("winCharacter", "lblClass")).text = "Class: " & Trim$(Class(GetPlayerClass(i)).Name)
            .Controls(GetControlIndex("winCharacter", "lblLevel")).text = "Level: " & GetPlayerLevel(i)
            .Controls(GetControlIndex("winCharacter", "lblHealth")).text = "Health: " & GetPlayerVital(i, HP) & "/" & GetPlayerMaxVital(i, HP)
            .Controls(GetControlIndex("winCharacter", "lblSpirit")).text = "Spirit: " & GetPlayerVital(i, MP) & "/" & GetPlayerMaxVital(i, MP)
            .Controls(GetControlIndex("winCharacter", "lblExperience")).text = "Exp: " & Player(i).EXP & "/" & TNL
            .Controls(GetControlIndex("winCharacter", "lblVip")).text = "Vip: " & PPremium
            .Controls(GetControlIndex("winCharacter", "lblVipD")).text = "Days: " & RPremium

            ' Att Golds
            Call SetGoldLabel

            ' Guild
            If Player(i).Guild_ID > 0 Then
                .Controls(GetControlIndex("winCharacter", "lblGuild")).text = "Guild: " & Trim$(Guild(Player(i).Guild_ID).Name)
            Else
                .Controls(GetControlIndex("winCharacter", "lblGuild")).text = "Guild: " & "Nenhuma"
            End If
            ' Update the window guild
            Call UpdateWindowGuild

            ' stats
            For X = 1 To Stats.Stat_Count - 1
                .Controls(GetControlIndex("winCharacter", "lblStat_" & X)).text = GetPlayerStat(i, X)
            Next
            ' points
            .Controls(GetControlIndex("winCharacter", "lblPoints")).text = GetPlayerPOINTS(i)
            ' grey out buttons
            If GetPlayerPOINTS(i) = 0 Then

                If .Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = 0 Then
                    For X = 1 To Stats.Stat_Count - 1
                        .Controls(GetControlIndex("winCharacter", "btnGreyStat_" & X)).visible = True
                    Next
                End If
            Else
                For X = 1 To Stats.Stat_Count - 1
                    .Controls(GetControlIndex("winCharacter", "btnGreyStat_" & X)).visible = False
                Next
            End If
        End With
    End If

    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).xOffset = 0
    Player(i).yOffset = 0
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim dir As Long
    Dim n As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    dir = buffer.ReadLong
    n = buffer.ReadLong
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, dir)
    Player(i).xOffset = 0
    Player(i).yOffset = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)

    Case DIR_UP
        Player(i).yOffset = PIC_Y

    Case DIR_DOWN
        Player(i).yOffset = PIC_Y * -1

    Case DIR_LEFT
        Player(i).xOffset = PIC_X

    Case DIR_RIGHT
        Player(i).xOffset = PIC_X * -1
    End Select
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MapNpcNum As Long
    Dim X As Long
    Dim Y As Long
    Dim dir As Long
    Dim Movement As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapNpcNum = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    dir = buffer.ReadLong
    Movement = buffer.ReadLong

    With MapNpc(MapNpcNum)
        .X = X
        .Y = Y
        .dir = dir
        .xOffset = 0
        .yOffset = 0
        .Moving = Movement

        Select Case .dir

        Case DIR_UP
            .yOffset = PIC_Y

        Case DIR_DOWN
            .yOffset = PIC_Y * -1

        Case DIR_LEFT
            .xOffset = PIC_X

        Case DIR_RIGHT
            .xOffset = PIC_X * -1
        End Select

    End With

End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim dir As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    dir = buffer.ReadLong
    Call SetPlayerDir(i, dir)

    With Player(i)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim dir As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    dir = buffer.ReadLong

    With MapNpc(i)
        .dir = dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    dir = buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).xOffset = 0
    Player(MyIndex).yOffset = 0
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim dir As Long
    Dim buffer As clsBuffer
    Dim thePlayer As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    dir = buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).xOffset = 0
    Player(thePlayer).yOffset = 0
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = getTime

    If i = MyIndex Then
        TimeSinceAttack = Tick
    End If
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = getTime
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, NeedMap As Byte, buffer As clsBuffer, MapDataCRC As Long, MapTileCRC As Long, MapNum As Long

    GettingMap = True
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Erase all players except self
    For i = 1 To Player_HighIndex
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap

    ' Clear the blood
    For i = 1 To MAX_BYTE
        Call ClearBlood(i, False)
    Next
    Blood_HighIndex = 0

    ' Get map num
    MapNum = buffer.ReadLong
    MapDataCRC = buffer.ReadLong
    MapTileCRC = buffer.ReadLong

    ' check against our own CRC32s
    NeedMap = 0
    If MapDataCRC <> MapCRC32(MapNum).MapDataCRC Then
        NeedMap = 1
    End If
    If MapTileCRC <> MapCRC32(MapNum).MapTileCRC Then
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong NeedMap
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If Not applyingMap Then
        If InMapEditor Then
            InMapEditor = False
            frmEditor_Map.visible = False
            ClearAttributeDialogue

            If frmEditor_MapProperties.visible Then
                frmEditor_MapProperties.visible = False
            End If
        End If
    End If

    ' load the map if we don't need it
    If NeedMap = 0 Then
        LoadMap MapNum
        applyingMap = False
    End If
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, MapNum As Long, i As Long, X As Long, Y As Long
    Dim DecompData() As Byte


    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData

    MapNum = buffer.ReadLong

    With Map.MapData
        .Name = buffer.ReadString
        .Music = buffer.ReadString
        .Moral = buffer.ReadByte
        .Up = buffer.ReadLong
        .Down = buffer.ReadLong
        .Left = buffer.ReadLong
        .Right = buffer.ReadLong
        .BootMap = buffer.ReadLong
        .BootX = buffer.ReadByte
        .BootY = buffer.ReadByte
        .MaxX = buffer.ReadByte
        .MaxY = buffer.ReadByte
        .BossNpc = buffer.ReadLong
        .Panorama = buffer.ReadByte

        .Weather = buffer.ReadByte
        .WeatherIntensity = buffer.ReadByte

        .Fog = buffer.ReadByte
        .FogSpeed = buffer.ReadByte
        .FogOpacity = buffer.ReadByte

        .Red = buffer.ReadByte
        .Green = buffer.ReadByte
        .Blue = buffer.ReadByte
        .Alpha = buffer.ReadByte

        .Sun = buffer.ReadByte

        .DayNight = buffer.ReadByte

        For i = 1 To MAX_MAP_NPCS
            .NPC(i) = buffer.ReadLong
        Next
    End With

    Map.TileData.EventCount = buffer.ReadLong
    If Map.TileData.EventCount > 0 Then
        ReDim Preserve Map.TileData.Events(1 To Map.TileData.EventCount)
        For i = 1 To Map.TileData.EventCount
            With Map.TileData.Events(i)
                .Name = buffer.ReadString
                .X = buffer.ReadLong
                .Y = buffer.ReadLong
                .pageCount = buffer.ReadLong
            End With
            If Map.TileData.Events(i).pageCount > 0 Then
                ReDim Preserve Map.TileData.Events(i).EventPage(1 To Map.TileData.Events(i).pageCount)
                For X = 1 To Map.TileData.Events(i).pageCount
                    With Map.TileData.Events(i).EventPage(X)
                        .chkPlayerVar = buffer.ReadByte
                        .chkSelfSwitch = buffer.ReadByte
                        .chkHasItem = buffer.ReadByte
                        .PlayerVarNum = buffer.ReadLong
                        .SelfSwitchNum = buffer.ReadLong
                        .HasItemNum = buffer.ReadLong
                        .PlayerVariable = buffer.ReadLong
                        .GraphicType = buffer.ReadByte
                        .Graphic = buffer.ReadLong
                        .GraphicX = buffer.ReadLong
                        .GraphicY = buffer.ReadLong
                        .MoveType = buffer.ReadByte
                        .MoveSpeed = buffer.ReadByte
                        .MoveFreq = buffer.ReadByte
                        .WalkAnim = buffer.ReadByte
                        .StepAnim = buffer.ReadByte
                        .DirFix = buffer.ReadByte
                        .WalkThrough = buffer.ReadByte
                        .Priority = buffer.ReadByte
                        .Trigger = buffer.ReadByte
                        .CommandCount = buffer.ReadLong
                    End With
                    If Map.TileData.Events(i).EventPage(X).CommandCount > 0 Then
                        ReDim Preserve Map.TileData.Events(i).EventPage(X).Commands(1 To Map.TileData.Events(i).EventPage(X).CommandCount)
                        For Y = 1 To Map.TileData.Events(i).EventPage(X).CommandCount
                            With Map.TileData.Events(i).EventPage(X).Commands(Y)
                                .Type = buffer.ReadByte
                                .text = buffer.ReadString
                                .Colour = buffer.ReadLong
                                .channel = buffer.ReadByte
                                .TargetType = buffer.ReadByte
                                .Target = buffer.ReadLong
                                .X = buffer.ReadLong
                                .Y = buffer.ReadLong
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If

    ReDim Map.TileData.Tile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.TileData.Tile(X, Y).Layer(i).X = buffer.ReadLong
                Map.TileData.Tile(X, Y).Layer(i).Y = buffer.ReadLong
                Map.TileData.Tile(X, Y).Layer(i).tileSet = buffer.ReadLong
                Map.TileData.Tile(X, Y).Autotile(i) = buffer.ReadByte
            Next
            Map.TileData.Tile(X, Y).Type = buffer.ReadByte
            Map.TileData.Tile(X, Y).Data1 = buffer.ReadLong
            Map.TileData.Tile(X, Y).Data2 = buffer.ReadLong
            Map.TileData.Tile(X, Y).Data3 = buffer.ReadLong
            Map.TileData.Tile(X, Y).Data4 = buffer.ReadLong
            Map.TileData.Tile(X, Y).Data5 = buffer.ReadLong
            Map.TileData.Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    ClearTempTile
    initAutotiles
    Set buffer = Nothing
    ' Save the map
    Call SaveMap(MapNum)
    GetMapCRC32 MapNum
    AddText "Downloaded new map.", BrightGreen

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If Not applyingMap Then
        If InMapEditor Then
            InMapEditor = False
            frmEditor_Map.visible = False
            ClearAttributeDialogue
            If frmEditor_MapProperties.visible Then
                frmEditor_MapProperties.visible = False
            End If
        End If
    End If
    applyingMap = False

End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_MAP_ITEMS

        With MapItem(i)
            .playerName = buffer.ReadString
            .Num = buffer.ReadLong
            .Value = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .bound = buffer.ReadByte

        End With
    Next

End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Dim IdDeath As Integer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_MAP_NPCS

        With MapNpc(i)
            .Num = buffer.ReadInteger
            .X = buffer.ReadByte
            .Y = buffer.ReadByte
            .dir = buffer.ReadByte
            .Vital(Vitals.HP) = buffer.ReadLong
            .Vital(Vitals.MP) = buffer.ReadLong
            .StunDuration = buffer.ReadLong
            .Dead = buffer.ReadByte

            ' Obtem o ID temporário, caso o npc esteja morto. Então coloca o número dele pra renderizar os dados e sprite!
            IdDeath = buffer.ReadInteger
            If .Dead = YES Then
                .Num = IdDeath
            End If
        End With

    Next

End Sub

Private Sub HandleMapDone()
    Dim i As Long
    Dim musicFile As String

    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i

    Action_HighIndex = 1

    ' player music
    If InGame Then
        musicFile = Trim$(Map.MapData.Music)

        If Not musicFile = "None." Then
            Play_Music musicFile
        Else
            Stop_Music
        End If
    End If

    ' get the npc high index
    For i = MAX_MAP_NPCS To 1 Step -1

        If MapNpc(i).Num > 0 Then
            Npc_HighIndex = i + 1
            Exit For
        End If

    Next

    ' make sure we're not overflowing
    If Npc_HighIndex > MAX_MAP_NPCS Then Npc_HighIndex = MAX_MAP_NPCS
    ' now cache the positions
    initAutotiles
    GettingMap = False
    CanMoveNow = True
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong

    With MapItem(n)
        .playerName = buffer.ReadString
        .Num = buffer.ReadLong
        .Value = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .bound = buffer.ReadByte
        .Gravity = -10
    End With

End Sub

Private Sub HandleItemEditor()
    Dim i As Long

    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

End Sub

Private Sub HandleAnimationEditor()
    Dim i As Long

    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Dim DecompData() As Byte
    Dim ItemCRC As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData

    n = buffer.ReadLong
    ItemCRC = buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize

    Set buffer = Nothing

    If ItemCRC <> ItemCRC32(n).ItemDataCRC Then
        Call SaveItem(n)
        Call GetItemCRC32(n)
    End If
End Sub


Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Dim DecompData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData

    n = buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong

    With MapNpc(n)
        .Num = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .dir = buffer.ReadLong
        ' Client use only
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
        .Dead = NO
    End With

End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim NpcNum As Long, X As Integer, Y As Integer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong

    ' unload it if we're not in target
    If myTargetType = TARGET_TYPE_NPC And myTarget = n Then
        HideWindow GetWindowIndex("winEnemyBars")
    End If

    NpcNum = MapNpc(n).Num
    X = MapNpc(n).X
    Y = MapNpc(n).Y

    ' Erase Data
    Call ClearMapNpc(n)

    ' Set Dead
    MapNpc(n).Num = NpcNum
    MapNpc(n).Dead = YES
    MapNpc(n).X = X
    MapNpc(n).Y = Y
End Sub

Private Sub HandleNpcEditor()
    Dim i As Long

    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(NPC(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With

End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Dim DecompData() As Byte
    Dim NpcCRC As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData

    n = buffer.ReadLong
    NpcCRC = buffer.ReadLong

    NpcSize = LenB(NPC(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(NPC(n)), ByVal VarPtr(NpcData(0)), NpcSize
    Set buffer = Nothing

    If NpcCRC <> NpcCRC32(n).NpcDataCRC Then
        Call SaveNpc(n)
        Call GetNpcCRC32(n)
    End If
End Sub

Private Sub HandleResourceEditor()
    Dim i As Long

    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    Dim DecompData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData

    ResourceNum = buffer.ReadLong
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    ClearResource ResourceNum
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    Set buffer = Nothing
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    n = buffer.ReadByte
    TempTile(X, Y).DoorOpen = n

    ' re-cache rendering
    If Not GettingMap Then cacheRenderState X, Y, MapLayer.Mask
End Sub

Private Sub HandleEditMap()
    Call MapEditorInit
End Sub

Private Sub HandleShopEditor()
    Dim i As Long

    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With

End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte

    Dim DecompData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData

    shopNum = buffer.ReadLong
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize
    Set buffer = Nothing
End Sub

Private Sub HandleSpellEditor()
    Dim i As Long

    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim spellnum As Long
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    Dim DecompData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData

    spellnum = buffer.ReadLong
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set buffer = Nothing
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i).Spell = buffer.ReadLong
        PlayerSpells(i).Uses = buffer.ReadLong
    Next

    Set buffer = Nothing
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Call ClearPlayer(buffer.ReadLong)
    Set buffer = Nothing
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Resource_Index = buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = buffer.ReadByte
            MapResource(i).X = buffer.ReadLong
            MapResource(i).Y = buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set buffer = Nothing
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PingEnd = getTime
    Ping = PingEnd - PingStart
End Sub

Private Sub HandleDoorAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    Y = buffer.ReadLong

    With TempTile(X, Y)
        .DoorFrame = 1
        .DoorAnimate = 1    ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = getTime
    End With

    Set buffer = Nothing
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long, message As String, Color As Long, tmpType As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    message = buffer.ReadString
    Color = buffer.ReadLong
    tmpType = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Set buffer = Nothing
    CreateActionMsg message, Color, tmpType, X, Y
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    Dim X As Long, Y As Long, Sprite As Long, i As Long

    Set buffer = New clsBuffer

    buffer.WriteBytes data()

    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing

    Call CreateBlood(X, Y)

End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, X As Long, Y As Long, isCasting As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    AnimationIndex = AnimationIndex + 1

    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1

    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .LockType = buffer.ReadByte
        .lockindex = buffer.ReadLong
        .isCasting = buffer.ReadByte
        .Used(0) = True
        .Used(1) = True
    End With

    Set buffer = Nothing

    ' play the sound if we've got one
    With AnimInstance(AnimationIndex)

        If .LockType = 0 Then
            X = AnimInstance(AnimationIndex).X
            Y = AnimInstance(AnimationIndex).Y
        ElseIf .LockType = TARGET_TYPE_PLAYER Then
            X = GetPlayerX(.lockindex)
            Y = GetPlayerY(.lockindex)
        ElseIf .LockType = TARGET_TYPE_NPC Then
            X = MapNpc(.lockindex).X
            Y = MapNpc(.lockindex).Y
        End If

    End With

    PlayMapSound X, Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim MapNpcNum As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapNpcNum = buffer.ReadByte

    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = buffer.ReadLong
    Next

    ' Update enemy bars
    If myTargetType = TARGET_TYPE_NPC And myTarget = MapNpcNum Then
        UpdateEnemyBars
    End If

    Set buffer = Nothing
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Slot As Long
    Dim CDTime As Integer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Slot = buffer.ReadLong
    CDTime = buffer.ReadInteger
    SpellCD(Slot) = CDTime
    Set buffer = Nothing
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SpellBuffer = 0
    SpellBufferTimer = 0
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Access As Long, Name As String, message As String, Colour As Long, header As String, PK As Long, saycolour As Long
    Dim channel As Byte, colStr As String

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Name = buffer.ReadString
    Access = buffer.ReadLong
    PK = buffer.ReadLong
    message = buffer.ReadString
    header = buffer.ReadString
    saycolour = buffer.ReadLong
    Set buffer = Nothing

    ' Check access level
    Colour = White

    If Access > 0 Then Colour = Pink
    If PK > 0 Then Colour = BrightRed

    ' find channel
    channel = 0
    Select Case header
    Case "[Map] "
        channel = ChatChannel.chMap
    Case "[Global] "
        channel = ChatChannel.chGlobal
    Case "[Guild] "
        channel = ChatChannel.chGuild
    Case "[Party] "
        channel = ChatChannel.chParty
    End Select

    ' remove the colour char from the message
    message = Replace$(message, ColourChar, vbNullString)
    ' add to the chat box
    AddText ColourChar & GetColStr(Colour) & header & Name & ": " & ColourChar & GetColStr(Grey) & message, Grey, , channel
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim shopNum As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    shopNum = buffer.ReadLong
    OpenShop shopNum
    Set buffer = Nothing
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, MapNum As Long, i As Long, TargetType As Byte, StunDuration As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    MapNum = buffer.ReadLong
    i = buffer.ReadLong
    TargetType = buffer.ReadByte
    StunDuration = buffer.ReadLong

    If MapNum <> GetPlayerMap(MyIndex) Then
        buffer.Flush: Set buffer = Nothing: Exit Sub
    End If

    If TargetType = TARGET_TYPE_PLAYER Then
        Player(i).StunDuration = StunDuration
        If StunDuration > 0 Then
            Player(i).StatusNum(Status.Confused).Ativo = 1
        Else
            Player(i).StatusNum(Status.Confused).Ativo = 0
        End If
    ElseIf TargetType = TARGET_TYPE_NPC Then
        MapNpc(i).StunDuration = StunDuration
    End If

    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    InTrade = buffer.ReadLong
    Set buffer = Nothing

    ShowTrade
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InTrade = 0
    HideWindow GetWindowIndex("winTrade")
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, dataType As Byte, i As Long, yourWorth As Long, theirWorth As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    dataType = buffer.ReadByte

    If dataType = 0 Then    ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).Num = buffer.ReadLong
            TradeYourOffer(i).Value = buffer.ReadLong
        Next
        yourWorth = buffer.ReadLong

        ' Call SetPlayerGold(Index, GetPlayerGold(Index) - yourWorth)
        ' Call SetGoldLabel
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblYourValue")).text = yourWorth & "g"
    ElseIf dataType = 1 Then    'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).Num = buffer.ReadLong
            TradeTheirOffer(i).Value = buffer.ReadLong
        Next
        theirWorth = buffer.ReadLong

        ' Call SetPlayerGold(Index, GetPlayerGold(Index) - theirWorth)
        ' Call SetGoldLabel
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblTheirValue")).text = theirWorth & "g"
    End If

    Set buffer = Nothing
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tradeStatus As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    tradeStatus = buffer.ReadByte
    Set buffer = Nothing

    Select Case tradeStatus
    Case 0    ' clear
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Choose items to offer."
    Case 1    ' they've accepted
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Other player has accepted."
    Case 2    ' you've accepted
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Waiting for other player to accept."
    Case 3    ' no room
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Not enough inventory space."
    End Select
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    myTarget = buffer.ReadLong
    myTargetType = buffer.ReadLong
    Set buffer = Nothing

    UpdateEnemyInterface
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = buffer.ReadLong
        Hotbar(i).sType = buffer.ReadByte
    Next
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player_HighIndex = buffer.ReadLong
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    UpdateShop
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    entityType = buffer.ReadLong
    entityNum = buffer.ReadLong
    PlayMapSound X, Y, entityType, entityNum
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, theName As String, top As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    theName = buffer.ReadString
    ' cache name and show invitation
    diaDataString = theName
    ShowWindow GetWindowIndex("winInvite_Trade")
    Windows(GetWindowIndex("winInvite_Trade")).Controls(GetControlIndex("winInvite_Trade", "btnInvite")).text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to trade."
    AddText Trim$(theName) & " has invited you to trade.", White
    ' loop through
    top = screenHeight - 80
    If Windows(GetWindowIndex("winInvite_Party")).Window.visible Then
        top = top - 37
    End If
    If Windows(GetWindowIndex("winInvite_Guild")).Window.visible Then
        top = top - 37
    End If
    Windows(GetWindowIndex("winInvite_Trade")).Window.top = top
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, theName As String, top As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    theName = buffer.ReadString
    ' cache name and show invitation popup
    diaDataString = theName
    ShowWindow GetWindowIndex("winInvite_Party")
    Windows(GetWindowIndex("winInvite_Party")).Controls(GetControlIndex("winInvite_Party", "btnInvite")).text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to a party."
    AddText Trim$(theName) & " has invited you to a party.", White
    ' loop through
    top = screenHeight - 80
    If Windows(GetWindowIndex("winInvite_Trade")).Window.visible Then
        top = top - 37
    End If
    If Windows(GetWindowIndex("winInvite_Guild")).Window.visible Then
        top = top - 37
    End If
    Windows(GetWindowIndex("winInvite_Party")).Window.top = top
End Sub

Public Sub HandleGuildInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, theName As String, top As Long, theGuild As String

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    theName = buffer.ReadString
    theGuild = buffer.ReadString
    ' cache name and show invitation popup
    diaDataString = "Guild: " & Trim$(theGuild)
    ShowWindow GetWindowIndex("winInvite_Guild")
    Windows(GetWindowIndex("winInvite_Guild")).Controls(GetControlIndex("winInvite_Guild", "btnInvite")).text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to a Guild."
    AddText Trim$(theName) & " has invited you to a Guild.", White
    ' loop through
    top = screenHeight - 80
    If Windows(GetWindowIndex("winInvite_Trade")).Window.visible Then
        top = top - 37
    End If
    If Windows(GetWindowIndex("winInvite_Party")).Window.visible Then
        top = top - 37
    End If

    Windows(GetWindowIndex("winInvite_Guild")).Window.top = top
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long, inParty As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    inParty = buffer.ReadByte

    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        UpdatePartyInterface
        ' exit out early
        Exit Sub
    End If

    ' carry on otherwise
    Party.Leader = buffer.ReadLong

    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = buffer.ReadLong
    Next

    Party.MemberCount = buffer.ReadLong

    ' update the party interface
    UpdatePartyInterface
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim playerNum As Long, Level As Integer
    Dim buffer As clsBuffer, i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ' which player?
    playerNum = buffer.ReadLong
    Level = buffer.ReadInteger

    Call SetPlayerLevel(playerNum, Level)

    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = buffer.ReadLong
        Player(playerNum).Vital(i) = buffer.ReadLong
    Next

    ' update the party interface
    UpdatePartyBars
End Sub

Private Sub HandleConvEditor(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    With frmEditor_Conv
        Editor = EDITOR_CONV
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_CONVS
            .lstIndex.AddItem i & ": " & Trim$(Conv(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ConvEditorInit
    End With

End Sub

Private Sub HandleUpdateConv(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Convnum As Long
    Dim buffer As clsBuffer
    Dim i As Long
    Dim X As Long
    Dim DecompData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData

    Convnum = buffer.ReadLong
    With Conv(Convnum)
        .Name = buffer.ReadString
        .chatCount = buffer.ReadLong
        ReDim Conv(Convnum).Conv(1 To .chatCount)

        For i = 1 To .chatCount

            .Conv(i).Conv = buffer.ReadString

            For X = 1 To 4
                .Conv(i).rText(X) = buffer.ReadString
                .Conv(i).rTarget(X) = buffer.ReadLong
            Next

            .Conv(i).Event = buffer.ReadLong
            .Conv(i).Data1 = buffer.ReadLong
            .Conv(i).Data2 = buffer.ReadLong
            .Conv(i).Data3 = buffer.ReadLong
        Next
    End With

    buffer.Flush: Set buffer = Nothing
End Sub

Private Sub HandleChatUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, NpcNum As Long, mT As String, o(1 To 4) As String, i As Long

    Set buffer = New clsBuffer

    buffer.WriteBytes data()

    NpcNum = buffer.ReadLong
    mT = buffer.ReadString
    For i = 1 To 4
        o(i) = buffer.ReadString
    Next

    Set buffer = Nothing

    ' if npcNum is 0, exit the chat system
    If NpcNum = 0 Then
        inChat = False
        HideWindow GetWindowIndex("winNpcChat")
        Exit Sub
    End If

    ' set chat going
    OpenNpcChat NpcNum, mT, o
End Sub

Private Sub HandleStartTutorial(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
'inTutorial = True
' set the first message
'SetTutorialState 1
End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, TargetType As Long, Target As Long, message As String, Colour As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Target = buffer.ReadLong
    TargetType = buffer.ReadLong
    message = buffer.ReadString
    Colour = buffer.ReadLong
    AddChatBubble Target, TargetType, message, Colour
    Set buffer = Nothing
End Sub

Private Sub HandleSetPlayerLoginToken(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim user As String
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    loginToken = buffer.ReadString
    Set buffer = Nothing
    ' try and login to game server

    user = Windows(GetWindowIndex("winLogin")).Controls(GetControlIndex("winLogin", "txtUser")).text

    If user = vbNullString Then
        DialogueAlert DialogueMsg.MsgUSERNULL
        ClearUserAndPass
        Exit Sub
    End If

    If Len(Trim$(user)) < 3 Or Len(Trim$(user)) > ACCOUNT_LENGTH Then
        DialogueAlert DialogueMsg.MsgUSERLENGTH
        ClearUserAndPass
        Exit Sub
    End If

    AttemptLogin
End Sub

Private Sub HandlePlayerChars(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

Private Sub HandleCancelAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim theIndex As Long, buffer As clsBuffer, i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    theIndex = buffer.ReadLong
    Set buffer = Nothing
    ' find the casting animation
    For i = 1 To MAX_BYTE
        If AnimInstance(i).LockType = TARGET_TYPE_PLAYER Then
            If AnimInstance(i).lockindex = theIndex Then
                If AnimInstance(i).isCasting = 1 Then
                    ' clear it
                    ClearAnimInstance i
                End If
            End If
        End If
    Next
End Sub

Private Sub HandlePlayerVariables(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_BYTE
        Player(MyIndex).Variable(i) = buffer.ReadLong
    Next

    Set buffer = Nothing
End Sub

Private Sub HandleEvent(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    If buffer.ReadLong = 1 Then
        inEvent = True
    Else
        inEvent = False
    End If
    eventNum = buffer.ReadLong
    eventPageNum = buffer.ReadLong
    eventCommandNum = buffer.ReadLong

    Set buffer = Nothing
End Sub

Private Sub HandleClientTime(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    GameSecondsPerSecond = buffer.ReadByte
    GameMinutesPerMinute = buffer.ReadByte
    GameSeconds = buffer.ReadByte
    GameMinutes = buffer.ReadByte
    GameHours = buffer.ReadByte

    Set buffer = Nothing
End Sub

Private Sub HandleMessageWindow(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim WindowName As String
    Dim message As String

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    WindowName = buffer.ReadString
    message = buffer.ReadString

    Set buffer = Nothing

    ShowMessageWindow WindowName, message
End Sub

Private Sub HandleGoldUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Golds As Long

    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    Golds = buffer.ReadLong

    Set buffer = Nothing

    Call SetPlayerGold(MyIndex, Golds)

    Call SetGoldLabel
End Sub
