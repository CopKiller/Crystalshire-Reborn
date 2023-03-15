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

    Dim Buffer As clsBuffer
    Dim MsgType As Long
    Set Buffer = New clsBuffer

    Buffer.WriteBytes TempBuffer

    MsgType = Buffer.ReadLong

    If MsgType <= 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If

    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.length), 0, 0
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, dialogue_index As Long, menuReset As Long, kick As Long

    Set Buffer = New clsBuffer

    Buffer.WriteBytes data()
    dialogue_index = Buffer.ReadLong
    menuReset = Buffer.ReadLong
    kick = Buffer.ReadLong

    Set Buffer = Nothing

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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    ' player high index
    Player_HighIndex = Buffer.ReadLong
    Set Buffer = Nothing
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString
            .Vital(Vitals.HP) = Buffer.ReadLong
            .Vital(Vitals.MP) = Buffer.ReadLong
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = Buffer.ReadLong
            Next

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = Buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = Buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    ShowClasses

    Set Buffer = Nothing
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong    'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString    'Trim$(Parse(n))
            .Vital(Vitals.HP) = Buffer.ReadLong    'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = Buffer.ReadLong    'CLng(Parse(n + 2))
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = Buffer.ReadLong
            Next

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = Buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = Buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    Set Buffer = Nothing
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
        PlayerInv(i).bound = Buffer.ReadByte
    Next

    SetGoldLabel

    Set Buffer = Nothing
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadByte    'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadInteger)    'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong)    'CLng(Parse(3)))
    PlayerInv(n).bound = Buffer.ReadByte
    Set Buffer = Nothing
    SetGoldLabel
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Call SetPlayerEquipment(MyIndex, Buffer.ReadInteger, Armor)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadInteger, Weapon)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadInteger, Helmet)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadInteger, Shield)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadInteger, Legs)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadInteger, Boots)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadInteger, Amulet)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadInteger, RingLeft)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadInteger, RingRight)
    Set Buffer = Nothing
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim playerNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    playerNum = Buffer.ReadLong
    Call SetPlayerEquipment(playerNum, Buffer.ReadInteger, Armor)
    Call SetPlayerEquipment(playerNum, Buffer.ReadInteger, Weapon)
    Call SetPlayerEquipment(playerNum, Buffer.ReadInteger, Helmet)
    Call SetPlayerEquipment(playerNum, Buffer.ReadInteger, Shield)
    Call SetPlayerEquipment(playerNum, Buffer.ReadInteger, Legs)
    Call SetPlayerEquipment(playerNum, Buffer.ReadInteger, Boots)
    Call SetPlayerEquipment(playerNum, Buffer.ReadInteger, Amulet)
    Call SetPlayerEquipment(playerNum, Buffer.ReadInteger, RingLeft)
    Call SetPlayerEquipment(playerNum, Buffer.ReadInteger, RingRight)
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    If MyIndex = 0 Then Exit Sub
    Buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.HP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, Buffer.ReadLong)
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.MP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, Buffer.ReadLong)
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Byte
    Dim pIndex As Byte

    Buffer.WriteBytes data()

    pIndex = Buffer.ReadByte

    For i = 1 To Vitals.Vital_Count - 1
        Player(pIndex).MaxVital(i) = Buffer.ReadLong
        Call SetPlayerVital(pIndex, i, Buffer.ReadLong)
    Next i

    ' Update enemy bars
    If myTargetType = TARGET_TYPE_PLAYER And myTarget = pIndex Then
        UpdateEnemyBars
    End If
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, i, Buffer.ReadLong
    Next
    
    UpdateStatsWindow
End Sub

Private Sub UpdateStatsWindow()
    Dim i As Long, X As Long

    i = MyIndex
    With Windows(GetWindowIndex("winCharacter"))
        For X = 1 To Stats.Stat_Count - 1
            If CLng(.Controls(GetControlIndex("winCharacter", "lblStat_" & X)).Text) <> GetPlayerStat(i, X) Then
                .Controls(GetControlIndex("winCharacter", "lblStat_" & X)).Text = GetPlayerStat(i, X)
            End If
        Next X
    End With
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    SetPlayerExp MyIndex, Buffer.ReadLong
    TNL = Buffer.ReadLong
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerLevel(i, Buffer.ReadLong)
    Call SetPlayerPOINTS(i, Buffer.ReadLong)
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    Call SetPlayerClass(i, Buffer.ReadLong)
    Player(i).Guild_ID = Buffer.ReadInteger
    Player(i).Guild_MembroID = Buffer.ReadByte
    Player(i).Premium = Buffer.ReadByte
    Call SetPlayerGold(i, Buffer.ReadLong)

    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, Buffer.ReadLong
    Next X

    For X = 1 To Vitals.Vital_Count - 1
        SetPlayerMaxVital i, X, Buffer.ReadLong
        SetPlayerVital i, X, Buffer.ReadLong
    Next X

    For X = 1 To (status_count - 1)
        StatusOn = Buffer.ReadByte
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
            .Controls(GetControlIndex("winCharacter", "lblName")).Text = "Name: " & Trim$(GetPlayerName(i))
            .Controls(GetControlIndex("winCharacter", "lblClass")).Text = "Class: " & Trim$(Class(GetPlayerClass(i)).Name)
            .Controls(GetControlIndex("winCharacter", "lblLevel")).Text = "Level: " & GetPlayerLevel(i)
            .Controls(GetControlIndex("winCharacter", "lblHealth")).Text = "Health: " & GetPlayerVital(i, HP) & "/" & GetPlayerMaxVital(i, HP)
            .Controls(GetControlIndex("winCharacter", "lblSpirit")).Text = "Spirit: " & GetPlayerVital(i, MP) & "/" & GetPlayerMaxVital(i, MP)
            .Controls(GetControlIndex("winCharacter", "lblExperience")).Text = "Exp: " & Player(i).EXP & "/" & TNL
            .Controls(GetControlIndex("winCharacter", "lblVip")).Text = "Vip: " & PPremium
            .Controls(GetControlIndex("winCharacter", "lblVipD")).Text = "Days: " & RPremium
            
            ' Att Golds
            Call SetGoldLabel

            ' Guild
            If Player(i).Guild_ID > 0 Then
                .Controls(GetControlIndex("winCharacter", "lblGuild")).Text = "Guild: " & Trim$(Guild(Player(i).Guild_ID).Name)
            Else
                .Controls(GetControlIndex("winCharacter", "lblGuild")).Text = "Guild: " & "Nenhuma"
            End If
            ' Update the window guild
            Call UpdateWindowGuild

            ' stats
            For X = 1 To Stats.Stat_Count - 1
                .Controls(GetControlIndex("winCharacter", "lblStat_" & X)).Text = GetPlayerStat(i, X)
            Next
            ' points
            .Controls(GetControlIndex("winCharacter", "lblPoints")).Text = GetPlayerPOINTS(i)
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
    Player(i).XOffSet = 0
    Player(i).YOffSet = 0
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim dir As Long
    Dim n As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    dir = Buffer.ReadLong
    n = Buffer.ReadLong
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, dir)
    Player(i).XOffSet = 0
    Player(i).YOffSet = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)

    Case DIR_UP
        Player(i).YOffSet = PIC_Y

    Case DIR_DOWN
        Player(i).YOffSet = PIC_Y * -1

    Case DIR_LEFT
        Player(i).XOffSet = PIC_X

    Case DIR_RIGHT
        Player(i).XOffSet = PIC_X * -1
    End Select
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MapNpcNum As Long
    Dim X As Long
    Dim Y As Long
    Dim dir As Long
    Dim Movement As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    MapNpcNum = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Movement = Buffer.ReadLong

    With MapNpc(MapNpcNum)
        .X = X
        .Y = Y
        .dir = dir
        .XOffSet = 0
        .YOffSet = 0
        .Moving = Movement

        Select Case .dir

        Case DIR_UP
            .YOffSet = PIC_Y

        Case DIR_DOWN
            .YOffSet = PIC_Y * -1

        Case DIR_LEFT
            .XOffSet = PIC_X

        Case DIR_RIGHT
            .XOffSet = PIC_X * -1
        End Select

    End With

End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim dir As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerDir(i, dir)

    With Player(i)
        .XOffSet = 0
        .YOffSet = 0
        .Moving = 0
    End With

End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim dir As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong

    With MapNpc(i)
        .dir = dir
        .XOffSet = 0
        .YOffSet = 0
        .Moving = 0
    End With

End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).XOffSet = 0
    Player(MyIndex).YOffSet = 0
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim dir As Long
    Dim Buffer As clsBuffer
    Dim thePlayer As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    thePlayer = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).XOffSet = 0
    Player(thePlayer).YOffSet = 0
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = getTime
    
    If i = MyIndex Then
        TimeSinceAttack = Tick
    End If
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = getTime
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, NeedMap As Byte, Buffer As clsBuffer, MapDataCRC As Long, MapTileCRC As Long, MapNum As Long

    GettingMap = True
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

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
    MapNum = Buffer.ReadLong
    MapDataCRC = Buffer.ReadLong
    MapTileCRC = Buffer.ReadLong

    ' check against our own CRC32s
    NeedMap = 0
    If MapDataCRC <> MapCRC32(MapNum).MapDataCRC Then
        NeedMap = 1
    End If
    If MapTileCRC <> MapCRC32(MapNum).MapTileCRC Then
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong NeedMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing

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
    Dim Buffer As clsBuffer, MapNum As Long, i As Long, X As Long, Y As Long
    Dim DecompData() As Byte


    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing

    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData

    MapNum = Buffer.ReadLong

    With Map.MapData
        .Name = Buffer.ReadString
        .Music = Buffer.ReadString
        .Moral = Buffer.ReadByte
        .Up = Buffer.ReadLong
        .Down = Buffer.ReadLong
        .Left = Buffer.ReadLong
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
        
        .DayNight = Buffer.ReadByte

        For i = 1 To MAX_MAP_NPCS
            .NPC(i) = Buffer.ReadLong
        Next
    End With

    Map.TileData.EventCount = Buffer.ReadLong
    If Map.TileData.EventCount > 0 Then
        ReDim Preserve Map.TileData.Events(1 To Map.TileData.EventCount)
        For i = 1 To Map.TileData.EventCount
            With Map.TileData.Events(i)
                .Name = Buffer.ReadString
                .X = Buffer.ReadLong
                .Y = Buffer.ReadLong
                .pageCount = Buffer.ReadLong
            End With
            If Map.TileData.Events(i).pageCount > 0 Then
                ReDim Preserve Map.TileData.Events(i).EventPage(1 To Map.TileData.Events(i).pageCount)
                For X = 1 To Map.TileData.Events(i).pageCount
                    With Map.TileData.Events(i).EventPage(X)
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
                    If Map.TileData.Events(i).EventPage(X).CommandCount > 0 Then
                        ReDim Preserve Map.TileData.Events(i).EventPage(X).Commands(1 To Map.TileData.Events(i).EventPage(X).CommandCount)
                        For Y = 1 To Map.TileData.Events(i).EventPage(X).CommandCount
                            With Map.TileData.Events(i).EventPage(X).Commands(Y)
                                .Type = Buffer.ReadByte
                                .Text = Buffer.ReadString
                                .Colour = Buffer.ReadLong
                                .channel = Buffer.ReadByte
                                .TargetType = Buffer.ReadByte
                                .Target = Buffer.ReadLong
                                .X = Buffer.ReadLong
                                .Y = Buffer.ReadLong
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
                Map.TileData.Tile(X, Y).Layer(i).X = Buffer.ReadLong
                Map.TileData.Tile(X, Y).Layer(i).Y = Buffer.ReadLong
                Map.TileData.Tile(X, Y).Layer(i).tileSet = Buffer.ReadLong
                Map.TileData.Tile(X, Y).Autotile(i) = Buffer.ReadByte
            Next
            Map.TileData.Tile(X, Y).Type = Buffer.ReadByte
            Map.TileData.Tile(X, Y).Data1 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data2 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data3 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data4 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data5 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).DirBlock = Buffer.ReadByte
        Next
    Next

    ClearTempTile
    initAutotiles
    Set Buffer = Nothing
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_MAP_ITEMS

        With MapItem(i)
            .playerName = Buffer.ReadString
            .num = Buffer.ReadLong
            .Value = Buffer.ReadLong
            .X = Buffer.ReadLong
            .Y = Buffer.ReadLong
            .bound = Buffer.ReadByte

        End With
    Next

End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim IdDeath As Integer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_MAP_NPCS

        With MapNpc(i)
            .num = Buffer.ReadInteger
            .X = Buffer.ReadByte
            .Y = Buffer.ReadByte
            .dir = Buffer.ReadByte
            .Vital(Vitals.HP) = Buffer.ReadLong
            .Vital(Vitals.MP) = Buffer.ReadLong
            .StunDuration = Buffer.ReadLong
            .Dead = Buffer.ReadByte

            ' Obtem o ID temporário, caso o npc esteja morto. Então coloca o número dele pra renderizar os dados e sprite!
            IdDeath = Buffer.ReadInteger
            If .Dead = YES Then
                .num = IdDeath
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

        If MapNpc(i).num > 0 Then
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
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong

    With MapItem(n)
        .playerName = Buffer.ReadString
        .num = Buffer.ReadLong
        .Value = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .bound = Buffer.ReadByte
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
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Dim DecompData()   As Byte
    Dim ItemCRC As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData
    
    n = Buffer.ReadLong
    ItemCRC = Buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    
    Set Buffer = Nothing

    If ItemCRC <> ItemCRC32(n).ItemDataCRC Then
        Call SaveItem(n)
        Call GetItemCRC32(n)
    End If
End Sub


Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Dim DecompData()   As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData
    
    n = Buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong

    With MapNpc(n)
        .num = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .dir = Buffer.ReadLong
        ' Client use only
        .XOffSet = 0
        .YOffSet = 0
        .Moving = 0
        .Dead = NO
    End With

End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim NpcNum As Long, X As Integer, Y As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong

    ' unload it if we're not in target
    If myTargetType = TARGET_TYPE_NPC And myTarget = n Then
        HideWindow GetWindowIndex("winEnemyBars")
    End If
    
    NpcNum = MapNpc(n).num
    X = MapNpc(n).X
    Y = MapNpc(n).Y

    ' Erase Data
    Call ClearMapNpc(n)
    
    ' Set Dead
    MapNpc(n).num = NpcNum
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
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Dim DecompData()   As Byte
    Dim NpcCRC As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData
    
    n = Buffer.ReadLong
    NpcCRC = Buffer.ReadLong
    
    NpcSize = LenB(NPC(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(NPC(n)), ByVal VarPtr(NpcData(0)), NpcSize
    Set Buffer = Nothing
    
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
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    Dim DecompData()   As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData
    
    ResourceNum = Buffer.ReadLong
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    ClearResource ResourceNum
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    Set Buffer = Nothing
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    n = Buffer.ReadByte
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
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Dim DecompData()   As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData

    shopNum = Buffer.ReadLong
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize
    Set Buffer = Nothing
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
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    Dim DecompData()   As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData
    
    spellnum = Buffer.ReadLong
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set Buffer = Nothing
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i).Spell = Buffer.ReadLong
        PlayerSpells(i).Uses = Buffer.ReadLong
    Next

    Set Buffer = Nothing
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Call ClearPlayer(Buffer.ReadLong)
    Set Buffer = Nothing
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Resource_Index = Buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = Buffer.ReadByte
            MapResource(i).X = Buffer.ReadLong
            MapResource(i).Y = Buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PingEnd = getTime
    Ping = PingEnd - PingStart
End Sub

Private Sub HandleDoorAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong

    With TempTile(X, Y)
        .DoorFrame = 1
        .DoorAnimate = 1    ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = getTime
    End With

    Set Buffer = Nothing
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long, message As String, Color As Long, tmpType As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    message = Buffer.ReadString
    Color = Buffer.ReadLong
    tmpType = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Set Buffer = Nothing
    CreateActionMsg message, Color, tmpType, X, Y
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    Dim X As Long, Y As Long, Sprite As Long, i As Long

    Set Buffer = New clsBuffer

    Buffer.WriteBytes data()

    X = Buffer.ReadLong
    Y = Buffer.ReadLong

    Set Buffer = Nothing

    Call CreateBlood(X, Y)

End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, X As Long, Y As Long, isCasting As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    AnimationIndex = AnimationIndex + 1

    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1

    With AnimInstance(AnimationIndex)
        .Animation = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .LockType = Buffer.ReadByte
        .lockindex = Buffer.ReadLong
        .isCasting = Buffer.ReadByte
        .Used(0) = True
        .Used(1) = True
    End With

    Set Buffer = Nothing

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
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim MapNpcNum As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    MapNpcNum = Buffer.ReadByte

    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = Buffer.ReadLong
    Next

    ' Update enemy bars
    If myTargetType = TARGET_TYPE_NPC And myTarget = MapNpcNum Then
        UpdateEnemyBars
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Dim CDTime As Integer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Slot = Buffer.ReadLong
    CDTime = Buffer.ReadInteger
    SpellCD(Slot) = CDTime
    Set Buffer = Nothing
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SpellBuffer = 0
    SpellBufferTimer = 0
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, Access As Long, Name As String, message As String, Colour As Long, header As String, PK As Long, saycolour As Long
    Dim channel As Byte, colStr As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Name = Buffer.ReadString
    Access = Buffer.ReadLong
    PK = Buffer.ReadLong
    message = Buffer.ReadString
    header = Buffer.ReadString
    saycolour = Buffer.ReadLong
    Set Buffer = Nothing

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
    Dim Buffer As clsBuffer
    Dim shopNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    shopNum = Buffer.ReadLong
    OpenShop shopNum
    Set Buffer = Nothing
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
        Dim Buffer As clsBuffer, MapNum As Long, i As Long, TargetType As Byte, StunDuration As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    MapNum = Buffer.ReadLong
    i = Buffer.ReadLong
    TargetType = Buffer.ReadByte
    StunDuration = Buffer.ReadLong
    
    If MapNum <> GetPlayerMap(MyIndex) Then
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
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

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    InTrade = Buffer.ReadLong
    Set Buffer = Nothing

    ShowTrade
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InTrade = 0
    HideWindow GetWindowIndex("winTrade")
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, dataType As Byte, i As Long, yourWorth As Long, theirWorth As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    dataType = Buffer.ReadByte

    If dataType = 0 Then    ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).num = Buffer.ReadLong
            TradeYourOffer(i).Value = Buffer.ReadLong
        Next
        yourWorth = Buffer.ReadLong
        
       ' Call SetPlayerGold(Index, GetPlayerGold(Index) - yourWorth)
       ' Call SetGoldLabel
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblYourValue")).Text = yourWorth & "g"
    ElseIf dataType = 1 Then    'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).num = Buffer.ReadLong
            TradeTheirOffer(i).Value = Buffer.ReadLong
        Next
        theirWorth = Buffer.ReadLong
        
       ' Call SetPlayerGold(Index, GetPlayerGold(Index) - theirWorth)
       ' Call SetGoldLabel
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblTheirValue")).Text = theirWorth & "g"
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeStatus As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    tradeStatus = Buffer.ReadByte
    Set Buffer = Nothing

    Select Case tradeStatus
    Case 0    ' clear
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).Text = "Choose items to offer."
    Case 1    ' they've accepted
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).Text = "Other player has accepted."
    Case 2    ' you've accepted
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).Text = "Waiting for other player to accept."
    Case 3    ' no room
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).Text = "Not enough inventory space."
    End Select
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    myTarget = Buffer.ReadLong
    myTargetType = Buffer.ReadLong
    Set Buffer = Nothing

    UpdateEnemyInterface
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = Buffer.ReadLong
        Hotbar(i).sType = Buffer.ReadByte
    Next
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Player_HighIndex = Buffer.ReadLong
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    UpdateShop
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    entityType = Buffer.ReadLong
    entityNum = Buffer.ReadLong
    PlayMapSound X, Y, entityType, entityNum
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, theName As String, top As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    theName = Buffer.ReadString
    ' cache name and show invitation
    diaDataString = theName
    ShowWindow GetWindowIndex("winInvite_Trade")
    Windows(GetWindowIndex("winInvite_Trade")).Controls(GetControlIndex("winInvite_Trade", "btnInvite")).Text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to trade."
    AddText Trim$(theName) & " has invited you to trade.", White
    ' loop through
    top = ScreenHeight - 80
    If Windows(GetWindowIndex("winInvite_Party")).Window.visible Then
        top = top - 37
    End If
    If Windows(GetWindowIndex("winInvite_Guild")).Window.visible Then
        top = top - 37
    End If
    Windows(GetWindowIndex("winInvite_Trade")).Window.top = top
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, theName As String, top As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    theName = Buffer.ReadString
    ' cache name and show invitation popup
    diaDataString = theName
    ShowWindow GetWindowIndex("winInvite_Party")
    Windows(GetWindowIndex("winInvite_Party")).Controls(GetControlIndex("winInvite_Party", "btnInvite")).Text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to a party."
    AddText Trim$(theName) & " has invited you to a party.", White
    ' loop through
    top = ScreenHeight - 80
    If Windows(GetWindowIndex("winInvite_Trade")).Window.visible Then
        top = top - 37
    End If
    If Windows(GetWindowIndex("winInvite_Guild")).Window.visible Then
        top = top - 37
    End If
    Windows(GetWindowIndex("winInvite_Party")).Window.top = top
End Sub

Public Sub HandleGuildInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, theName As String, top As Long, theGuild As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    theName = Buffer.ReadString
    theGuild = Buffer.ReadString
    ' cache name and show invitation popup
    diaDataString = "Guild: " & Trim$(theGuild)
    ShowWindow GetWindowIndex("winInvite_Guild")
    Windows(GetWindowIndex("winInvite_Guild")).Controls(GetControlIndex("winInvite_Guild", "btnInvite")).Text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to a Guild."
    AddText Trim$(theName) & " has invited you to a Guild.", White
    ' loop through
    top = ScreenHeight - 80
    If Windows(GetWindowIndex("winInvite_Trade")).Window.visible Then
        top = top - 37
    End If
    If Windows(GetWindowIndex("winInvite_Party")).Window.visible Then
        top = top - 37
    End If
    
    Windows(GetWindowIndex("winInvite_Guild")).Window.top = top
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, inParty As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    inParty = Buffer.ReadByte

    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        UpdatePartyInterface
        ' exit out early
        Exit Sub
    End If

    ' carry on otherwise
    Party.Leader = Buffer.ReadLong

    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = Buffer.ReadLong
    Next

    Party.MemberCount = Buffer.ReadLong

    ' update the party interface
    UpdatePartyInterface
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim playerNum As Long, Level As Integer
    Dim Buffer As clsBuffer, i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ' which player?
    playerNum = Buffer.ReadLong
    Level = Buffer.ReadInteger

    Call SetPlayerLevel(playerNum, Level)

    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = Buffer.ReadLong
        Player(playerNum).Vital(i) = Buffer.ReadLong
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
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim X As Long
    Dim DecompData()   As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData

    Convnum = Buffer.ReadLong
    With Conv(Convnum)
        .Name = Buffer.ReadString
        .chatCount = Buffer.ReadLong
        ReDim Conv(Convnum).Conv(1 To .chatCount)

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
End Sub

Private Sub HandleChatUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, NpcNum As Long, mT As String, o(1 To 4) As String, i As Long

    Set Buffer = New clsBuffer

    Buffer.WriteBytes data()

    NpcNum = Buffer.ReadLong
    mT = Buffer.ReadString
    For i = 1 To 4
        o(i) = Buffer.ReadString
    Next

    Set Buffer = Nothing

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
    Dim Buffer As clsBuffer, TargetType As Long, Target As Long, message As String, Colour As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Target = Buffer.ReadLong
    TargetType = Buffer.ReadLong
    message = Buffer.ReadString
    Colour = Buffer.ReadLong
    AddChatBubble Target, TargetType, message, Colour
    Set Buffer = Nothing
End Sub

Private Sub HandleSetPlayerLoginToken(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim user As String
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    loginToken = Buffer.ReadString
    Set Buffer = Nothing
    ' try and login to game server

    user = Windows(GetWindowIndex("winLogin")).Controls(GetControlIndex("winLogin", "txtUser")).Text

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
    Dim theIndex As Long, Buffer As clsBuffer, i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    theIndex = Buffer.ReadLong
    Set Buffer = Nothing
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
    Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_BYTE
        Player(MyIndex).Variable(i) = Buffer.ReadLong
    Next

    Set Buffer = Nothing
End Sub

Private Sub HandleEvent(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    If Buffer.ReadLong = 1 Then
        inEvent = True
    Else
        inEvent = False
    End If
    eventNum = Buffer.ReadLong
    eventPageNum = Buffer.ReadLong
    eventCommandNum = Buffer.ReadLong

    Set Buffer = Nothing
End Sub

Private Sub HandleClientTime(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    GameSecondsPerSecond = Buffer.ReadByte
    GameMinutesPerMinute = Buffer.ReadByte
    GameSeconds = Buffer.ReadByte
    GameMinutes = Buffer.ReadByte
    GameHours = Buffer.ReadByte

    Set Buffer = Nothing
End Sub

Private Sub HandleMessageWindow(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim WindowName As String
    Dim message As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    WindowName = Buffer.ReadString
    message = Buffer.ReadString

    Set Buffer = Nothing
    
    ShowMessageWindow WindowName, message
End Sub

Private Sub HandleGoldUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Golds As Long
    
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Golds = Buffer.ReadLong

    Set Buffer = Nothing
    
    Call SetPlayerGold(MyIndex, Golds)
    
    Call SetGoldLabel
End Sub
