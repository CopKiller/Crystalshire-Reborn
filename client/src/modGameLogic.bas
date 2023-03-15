Attribute VB_Name = "modGameLogic"
Option Explicit
Dim IsNpc As Boolean

Public Sub GameLoop()
    Dim TickFPS As Currency, FPS As Long, i As Long, WalkTimer As Currency, X As Long, Y As Long
    Dim tmr25 As Currency, tmr10000 As Currency, tmr100 As Currency, tmr1000 As Currency, mapTimer As Currency, chatTmr As Currency, targetTmr As Currency, fogTmr As Currency, barTmr As Currency
    Dim barDifference As Long, stunTmr As Currency, f As Long, ItemPic As Integer, MaxFrames As Byte

    ' *** Start GameLoop ***
    Do While InGame
        Tick = getTime                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.

        '    Debug.Print Tick

        'espera 20 milisegundos pra executar o resto da sub se estiver minimizado
        If frmMain.WindowState = vbMinimized Then
            Sleep 20
        End If

        ' handle input
        If GetForegroundWindow() = frmMain.hWnd Then
            HandleMouseInput
        End If

        For i = 1 To Blood_HighIndex
            With Blood(i)
                If .Alpha <= 0 Then
                    Call ClearBlood(i)
                Else
                    ' Check if we should be seeing it
                    If .Timer + 20000 < Tick Then
                        .Alpha = .Alpha - 1
                    End If
                End If
            End With
        Next i

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < Tick Then
            ' check ping
            Call GetPing
            tmr10000 = Tick + 10000
        End If

        If tmr1000 < Tick Then
            ' check if we need to end the CD icon
            If Count_Spellicon > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i).Spell > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) > 0 Then
                                SpellCD(i) = SpellCD(i) - 1
                            End If
                        End If
                    End If
                Next
            End If

            'clock
            GameSeconds = GameSeconds + GameSecondsPerSecond
            If GameSeconds > 59 Then
                GameSeconds = 0
                GameMinutes = GameMinutes + GameMinutesPerMinute
                If GameMinutes > 59 Then
                    GameMinutes = 0
                    GameHours = GameHours + 1
                    If GameHours > 23 Then
                        GameHours = 0
                    End If
                End If
            End If

            ' Calcular se a quest em andamento tem tempo pra finalizar em segundos
            Call CalculateQuestTimer

            ' Calcular o tempo da loteria pelo client, e apenas receber atualizações do servidor caso haja algum imprevisto!
            If Windows(GetWindowIndex("winLottery")).Window.visible Then
                If LotteryInfo.LotteryTime > 0 And Not LotteryInfo.LotteryOn And Not LotteryInfo.BetOn Then
                    LotteryInfo.LotteryTime = LotteryInfo.LotteryTime - 1

                    If Windows(GetWindowIndex("winLottery")).Window.visible Then
                        Windows(GetWindowIndex("winLottery")).Controls(GetControlIndex("winLottery", "lblNLottery")).Text = "Next Lottery: " & ColourChar & GetColStr(BrightRed) & SecondsToHMS(LotteryInfo.LotteryTime)
                    End If
                End If
            End If

            tmr1000 = Tick + 1000
        End If

        If MyIndex > 0 Then
            If Player(MyIndex).Moving = 0 And Player(MyIndex).LastMoving > 0 And Player(MyIndex).LastMoving + 125 <= Tick Then
                Player(MyIndex).LastMoving = 0
            End If
        End If

        If tmr25 < Tick Then
            InGame = IsConnected
            Call CheckKeys    ' Check to make sure they aren't trying to auto do anything

            ' Check In Diary
            Call Graphic_ArrowInDay_Animated

            If GetForegroundWindow() = frmMain.hWnd Then
                Call CheckInputKeys    ' Check which keys were pressed
            End If

            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer).Spell).CastTime * 1000) < Tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement    ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If

            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next

            ' animate Status balão
            If stunTmr < Tick Then
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
                            With Player(i)
                                For f = 1 To (status_count - 1)
                                    'status animado
                                    If .StatusNum(f).Ativo > 0 Then
                                        .StatusFrame = .StatusFrame + 1
                                        If .StatusFrame >= 8 Then
                                            .StatusFrame = 1
                                        End If
                                    End If
                                Next
                            End With
                        End If
                    End If
                Next i
                For i = 1 To MAX_MAP_NPCS
                    With MapNpc(i)
                        If .num > 0 Then
                            If Map.MapData.NPC(i) > 0 Then
                                If NPC(.num).Balao > 0 Or .StunDuration > 0 Or CheckNpcHaveQuest(.num) Or CheckNpcQuestProgress(.num) Or .Dead > 0 Then
                                    .StatusFrame = .StatusFrame + 1
                                    If .StatusFrame >= 8 Then
                                        .StatusFrame = 1
                                    End If
                                End If
                            End If
                        End If
                    End With
                Next i

                stunTmr = Tick + 125
            End If

            ' ****** Parallax X ******
            If ParallaxX = -ScreenWidth Then
                ParallaxX = 0
            Else
                ParallaxX = ParallaxX - 1
            End If

            ' appear tile logic
            AppearTileFadeLogic
            CheckAppearTiles

            ' handle events
            If inEvent Then
                If eventNum > 0 Then
                    If eventPageNum > 0 Then
                        If eventCommandNum > 0 Then
                            EventLogic
                        End If
                    End If
                End If
            End If

            ' check for map animation changes#
            If tmr100 < Tick Then
                For i = 1 To MAX_MAP_ITEMS
                    If MapItem(i).num > 0 Then
                        ItemPic = Item(MapItem(i).num).Pic
                        If ItemPic > 0 Or ItemPic <= Count_Item Then
                            MaxFrames = (mTexture(Tex_Item(ItemPic)).w) / PIC_X
                            If MaxFrames > 1 Then
                                If MapItem(i).Frame < MaxFrames Then
                                    MapItem(i).Frame = MapItem(i).Frame + 1
                                Else
                                    MapItem(i).Frame = 1
                                End If
                            End If
                        End If
                    End If

                Next i

                ' Inventory visible? play animation item!
                If Windows(GetWindowIndex("winInventory")).Window.visible Then
                    For i = 1 To MAX_INV
                        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
                            ItemPic = Item(GetPlayerInvItemNum(MyIndex, i)).Pic
                            If ItemPic > 0 Or ItemPic <= Count_Item Then
                                MaxFrames = (mTexture(Tex_Item(ItemPic)).w) / PIC_X
                                If MaxFrames > 1 Then
                                    If PlayerInv(i).Frame < MaxFrames Then
                                        PlayerInv(i).Frame = PlayerInv(i).Frame + 1
                                    Else
                                        PlayerInv(i).Frame = 1
                                    End If
                                End If
                            End If
                        End If

                    Next i
                End If

                ' Bank visible? play animation item!
                If Windows(GetWindowIndex("winBank")).Window.visible Then
                    For i = 1 To MAX_BANK
                        If GetBankItemNum(i) > 0 Then
                            ItemPic = Item(GetBankItemNum(i)).Pic
                            If ItemPic > 0 Or ItemPic <= Count_Item Then
                                MaxFrames = (mTexture(Tex_Item(ItemPic)).w) / PIC_X
                                If MaxFrames > 1 Then
                                    If Bank.Item(i).Frame < MaxFrames Then
                                        Bank.Item(i).Frame = Bank.Item(i).Frame + 1
                                    Else
                                        Bank.Item(i).Frame = 1
                                    End If
                                End If
                            End If
                        End If

                    Next i
                End If

                ' Shop visible? play animation item!
                If Windows(GetWindowIndex("winShop")).Window.visible Then
                    For i = 1 To MAX_TRADES
                        If Shop(InShop).TradeItem(i).Item > 0 Then
                            ItemPic = Item(Shop(InShop).TradeItem(i).Item).Pic
                            If ItemPic > 0 Or ItemPic <= Count_Item Then
                                MaxFrames = (mTexture(Tex_Item(ItemPic)).w) / PIC_X
                                If MaxFrames > 1 Then
                                    If Shop(InShop).TradeItem(i).Frame < MaxFrames Then
                                        Shop(InShop).TradeItem(i).Frame = Shop(InShop).TradeItem(i).Frame + 1
                                    Else
                                        Shop(InShop).TradeItem(i).Frame = 1
                                    End If
                                End If
                            End If
                        End If

                    Next i
                End If

                ' In Trade visible? play animation item!
                If Windows(GetWindowIndex("winTrade")).Window.visible Then
                    For i = 1 To MAX_INV
                        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num) > 0 Then
                            ItemPic = Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Pic
                            If ItemPic > 0 Or ItemPic <= Count_Item Then
                                MaxFrames = (mTexture(Tex_Item(ItemPic)).w) / PIC_X
                                If MaxFrames > 1 Then
                                    If TradeYourOffer(i).Frame < MaxFrames Then
                                        TradeYourOffer(i).Frame = TradeYourOffer(i).Frame + 1
                                    Else
                                        TradeYourOffer(i).Frame = 1
                                    End If
                                End If
                            End If
                        End If
                        If TradeTheirOffer(i).num > 0 Then
                            ItemPic = Item(TradeTheirOffer(i).num).Pic
                            If ItemPic > 0 Or ItemPic <= Count_Item Then
                                MaxFrames = (mTexture(Tex_Item(ItemPic)).w) / PIC_X
                                If MaxFrames > 1 Then
                                    If TradeTheirOffer(i).Frame < MaxFrames Then
                                        TradeTheirOffer(i).Frame = TradeTheirOffer(i).Frame + 1
                                    Else
                                        TradeTheirOffer(i).Frame = 1
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If

                ' Equipments visible? play animation item! ##UTILIZA UMA VARIÁVEL GLOBAL# -> EquipmentFrames As Byte
                If Windows(GetWindowIndex("winCharacter")).Window.visible Then
                    If Windows(GetWindowIndex("winCharacter")).Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = 1 Then
                        For i = 1 To Equipment.Equipment_Count - 1
                            If GetPlayerEquipment(MyIndex, i) > 0 Then
                                ItemPic = Item(GetPlayerEquipment(MyIndex, i)).Pic
                                If ItemPic > 0 Or ItemPic <= Count_Item Then
                                    MaxFrames = (mTexture(Tex_Item(ItemPic)).w) / PIC_X
                                    If MaxFrames > 1 Then
                                        If EquipmentFrames(i) < MaxFrames Then
                                            EquipmentFrames(i) = EquipmentFrames(i) + 1
                                        Else
                                            EquipmentFrames(i) = 1
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If

                tmr100 = Tick + 200
            End If

            tmr25 = Tick + 25
        End If

        ' targetting
        If targetTmr < Tick Then
            If TargetDown Then
                FindNearestTarget
            End If

            targetTmr = Tick + 50
        End If

        ' chat timer
        If chatTmr < Tick Then
            ' scrolling
            If ChatButtonUp Then
                ScrollChatBox 0

                If ChatMouseScroll Then ChatButtonUp = False
            End If

            If ChatButtonDown Then
                ScrollChatBox 1

                If ChatMouseScroll Then ChatButtonDown = False
            End If

            ' remove messages
            If chatLastRemove + CHAT_DIFFERENCE_TIMER < getTime Then
                ' remove timed out messages from chat
                For i = Chat_HighIndex To 1 Step -1
                    If Len(Chat(i).Text) > 0 Then
                        If Chat(i).visible Then
                            If Chat(i).Timer + CHAT_TIMER < Tick Then
                                Chat(i).visible = False
                                chatLastRemove = getTime
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If

            chatTmr = Tick + 50
        End If

        ' fog scrolling
        If fogTmr < Tick Then
            If Map.MapData.FogSpeed > 0 Then
                ' move
                fogOffsetX = fogOffsetX - Map.MapData.FogSpeed
                fogOffsetY = fogOffsetY - Map.MapData.FogSpeed
            End If

            fogTmr = Tick + 40 - Map.MapData.FogSpeed
        End If

        ' elastic bars
        If barTmr < Tick Then
            SetBarWidth BarWidth_GuiHP_Max, BarWidth_GuiHP
            SetBarWidth BarWidth_GuiSP_Max, BarWidth_GuiSP
            SetBarWidth BarWidth_GuiEXP_Max, BarWidth_GuiEXP
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    SetBarWidth BarWidth_NpcHP_Max(i), BarWidth_NpcHP(i)
                End If
            Next

            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    SetBarWidth BarWidth_PlayerHP_Max(i), BarWidth_PlayerHP(i)
                End If
            Next

            ' reset timer
            barTmr = Tick + 10
        End If

        ' Animations!
        If mapTimer < Tick Then

            ' animate waterfalls
            Select Case waterfallFrame

            Case 0
                waterfallFrame = 1

            Case 1
                waterfallFrame = 2

            Case 2
                waterfallFrame = 0
            End Select

            ' animate autotiles
            Select Case autoTileFrame

            Case 0
                autoTileFrame = 1

            Case 1
                autoTileFrame = 2

            Case 2
                autoTileFrame = 0
            End Select

            ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = vbNullString
            Else
                chatShowLine = "|"
            End If

            ' re-set timer
            mapTimer = Tick + 500
        End If

        Call ProcessWeather

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < Tick Then

            For i = 1 To Player_HighIndex

                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                End If

            Next i

            ' Process npc movements (actually move them)
            For i = 1 To Npc_HighIndex

                If Map.MapData.NPC(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If

            Next i

            WalkTimer = Tick + 30    ' edit this value to change WalkTimer
        End If

        If LotteryBtnRandom Then
            Call LotteryRand
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics

        If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents

        ' Lock fps
        If Not FPS_Lock Then

            Do While getTime < Tick + 20
                If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
                Sleep 1
            Loop

        End If

        ' Calculate fps
        If Options.FPSConection = YES Then
            If TickFPS < Tick Then
                GameFPS = FPS
                TickFPS = Tick + 1000
                FPS = 0
            Else
                FPS = FPS + 1
            End If
        End If

    Loop

    Call InitReconnect

End Sub

Public Sub MenuLoop()
    Dim TickFPS As Currency, FPS As Long, tmr500 As Currency, fadeTmr As Currency
    Dim tmr10000 As Currency, tmr1000 As Currency, Reconnects As Long

    ' *** Start GameLoop ***
    Do While inMenu
        Tick = getTime                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.

        'espera 20 milisegundos pra executar o resto da sub se estiver minimizado
        If frmMain.WindowState = vbMinimized Then
            Sleep 20
        End If

        ' handle input
        If GetForegroundWindow() = frmMain.hWnd Then
            HandleMouseInput
        End If

        ' Animations!
        If tmr500 < Tick Then
            ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = vbNullString
            Else
                chatShowLine = "|"
            End If

            ' re-set timer
            tmr500 = Tick + 500
        End If

        ' ****** Parallax X ******
        If ParallaxX = -ScreenWidth Then
            ParallaxX = 0
        Else
            ParallaxX = ParallaxX - 1
        End If

        ' trailer
        If videoPlaying Then VideoLoop

        ' fading
        If fadeTmr < Tick Then
            If Not videoPlaying Then
                If fadeAlpha > 5 Then
                    ' lower fade
                    fadeAlpha = fadeAlpha - 5
                Else
                    fadeAlpha = 0
                End If
            End If
            fadeTmr = Tick + 1
        End If


        Call ProccessReconnect

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Menu

        ' do events
        If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents

        ' Lock fps
        If Not FPS_Lock Then

            Do While getTime < Tick + 20
                If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
                Sleep 1
            Loop

        End If

        ' Calculate fps
        If Options.FPSConection = YES Then
            If TickFPS < Tick Then
                GameFPS = FPS
                TickFPS = Tick + 1000
                FPS = 0
            Else
                FPS = FPS + 1
            End If
        End If

    Loop

End Sub

Sub ProcessMovement(ByVal Index As Long)
    Dim MovementSpeed As Long

    ' Check if player is walking, and if so process moving them over
    Select Case Player(Index).Moving

    Case MOVING_WALKING: MovementSpeed = RUN_SPEED

    Case MOVING_RUNNING: MovementSpeed = WALK_SPEED

    Case Else: Exit Sub
    End Select

    Select Case GetPlayerDir(Index)

    Case DIR_UP
        Player(Index).yOffset = Player(Index).yOffset - MovementSpeed

        If Player(Index).yOffset < 0 Then Player(Index).yOffset = 0

    Case DIR_DOWN
        Player(Index).yOffset = Player(Index).yOffset + MovementSpeed

        If Player(Index).yOffset > 0 Then Player(Index).yOffset = 0

    Case DIR_LEFT
        Player(Index).xOffset = Player(Index).xOffset - MovementSpeed

        If Player(Index).xOffset < 0 Then Player(Index).xOffset = 0

    Case DIR_RIGHT
        Player(Index).xOffset = Player(Index).xOffset + MovementSpeed

        If Player(Index).xOffset > 0 Then Player(Index).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).LastMoving = Tick
                Player(Index).Moving = 0

                If Player(Index).step = 0 Then
                    Player(Index).step = 2
                Else
                    Player(Index).step = 0
                End If
            End If

        Else

            If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).LastMoving = Tick
                Player(Index).Moving = 0

                If Player(Index).step = 0 Then
                    Player(Index).step = 2
                Else
                    Player(Index).step = 0
                End If
            End If
        End If
    End If

End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    Dim MovementSpeed As Long

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        MovementSpeed = RUN_SPEED
    Else
        Exit Sub
    End If

    Select Case MapNpc(MapNpcNum).dir

    Case DIR_UP
        MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - MovementSpeed

        If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0

    Case DIR_DOWN
        MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + MovementSpeed

        If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0

    Case DIR_LEFT
        MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - MovementSpeed

        If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0

    Case DIR_RIGHT
        MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + MovementSpeed

        If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If MapNpc(MapNpcNum).Moving > 0 Then
        If MapNpc(MapNpcNum).dir = DIR_RIGHT Or MapNpc(MapNpcNum).dir = DIR_DOWN Then
            If (MapNpc(MapNpcNum).xOffset >= 0) And (MapNpc(MapNpcNum).yOffset >= 0) Then
                MapNpc(MapNpcNum).Moving = 0

                If MapNpc(MapNpcNum).step = 0 Then
                    MapNpc(MapNpcNum).step = 2
                Else
                    MapNpc(MapNpcNum).step = 0
                End If
            End If

        Else

            If (MapNpc(MapNpcNum).xOffset <= 0) And (MapNpc(MapNpcNum).yOffset <= 0) Then
                MapNpc(MapNpcNum).Moving = 0

                If MapNpc(MapNpcNum).step = 0 Then
                    MapNpc(MapNpcNum).step = 2
                Else
                    MapNpc(MapNpcNum).step = 0
                End If
            End If
        End If
    End If

End Sub

Sub CheckMapGetItem()
    Dim Buffer As New clsBuffer, tmpIndex As Long, i As Long, X As Long
    Set Buffer = New clsBuffer

    If getTime > Player(MyIndex).MapGetTimer + 250 Then

        ' find out if we want to pick it up
        For i = 1 To MAX_MAP_ITEMS

            If MapItem(i).X = Player(MyIndex).X And MapItem(i).Y = Player(MyIndex).Y Then
                If MapItem(i).num > 0 Then
                    If Item(MapItem(i).num).BindType = 1 Then

                        ' make sure it's not a party drop
                        If Party.Leader > 0 Then

                            For X = 1 To MAX_PARTY_MEMBERS
                                tmpIndex = Party.Member(X)

                                If tmpIndex > 0 Then
                                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(i).playerName) Then
                                        If Item(MapItem(i).num).ClassReq > 0 Then
                                            If Item(MapItem(i).num).ClassReq <> Player(MyIndex).Class Then
                                                Dialogue "Loot Check", "This item is BoP and is not for your class.", "Are you sure you want to pick it up?", TypeLOOTITEM, StyleYesNo
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If

                            Next

                        End If

                    Else
                        'not bound
                        Exit For
                    End If
                End If
            End If

        Next

        ' nevermind, pick it up
        Player(MyIndex).MapGetTimer = getTime
        Buffer.WriteLong CMapGetItem
        SendData Buffer.ToArray()
    End If

    Set Buffer = Nothing
End Sub

Public Sub CheckAttack()
    Dim Buffer As clsBuffer
    Dim attackspeed As Long

    If ControlDown Then
        If SpellBuffer > 0 Then Exit Sub    ' currently casting a spell, can't attack
        If Player(MyIndex).StunDuration > 0 Then Exit Sub    ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).Speed
        Else
            attackspeed = 1000
        End If
        
        If TimeSinceAttack + attackspeed > Tick Then
            Exit Sub
        End If

        If Player(MyIndex).Attacking = 0 Then
            Player(MyIndex).LastMoving = getTime

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = getTime
                End With

                Set Buffer = New clsBuffer
                Buffer.WriteLong CAttack
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                
                TimeSinceAttack = Tick
        End If
    End If

End Sub

Function IsTryingToMove() As Boolean

    If upDown Or leftDown Or downDown Or rightDown Or SetaUp Or SetaDown Or SetaLeft Or SetaRight Then
        IsTryingToMove = True
    End If

End Function

Function CanMove() As Boolean
    Dim d As Long

    CanMove = True

    If Player(MyIndex).xOffset <> 0 Or Player(MyIndex).yOffset <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they are not getting a map...
    If GettingMap Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer).Spell).CanRun = NO Then
            CanMove = False
            Exit Function
        End If
    End If

    ' make sure they're not stunned
    If Player(MyIndex).StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If

    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If

    ' not in bank
    If InBank Then
        CanMove = False
        Exit Function
    End If

    If inTutorial Then
        CanMove = False
        Exit Function
    End If

    d = GetPlayerDir(MyIndex)

    If upDown Or SetaUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.MapData.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If downDown Or SetaDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MapData.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.MapData.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If leftDown Or SetaLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.MapData.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If rightDown Or SetaRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MapData.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.MapData.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

End Function

Function CheckDirection(ByVal direction As Byte) As Boolean
    Dim X As Long, Y As Long, i As Long, EventCount As Long, page As Long

    CheckDirection = False

    If GettingMap Then Exit Function

    ' check directional blocking
    If isDirBlocked(Map.TileData.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case direction

    Case DIR_UP
        X = GetPlayerX(MyIndex)
        Y = GetPlayerY(MyIndex) - 1

    Case DIR_DOWN
        X = GetPlayerX(MyIndex)
        Y = GetPlayerY(MyIndex) + 1

    Case DIR_LEFT
        X = GetPlayerX(MyIndex) - 1
        Y = GetPlayerY(MyIndex)

    Case DIR_RIGHT
        X = GetPlayerX(MyIndex) + 1
        Y = GetPlayerY(MyIndex)
    End Select

    ' Check to see if the map tile is blocked or not
    If Map.TileData.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map.TileData.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to make sure that any events on that space aren't blocked
    EventCount = Map.TileData.EventCount
    For i = 1 To EventCount
        With Map.TileData.Events(i)
            If .X = X And .Y = Y Then
                ' Get the active event page
                page = ActiveEventPage(i)
                If page > 0 Then
                    If Map.TileData.Events(i).EventPage(page).WalkThrough = 0 Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End With
    Next

    ' Check to see if the key door is open or not
    If Map.TileData.Tile(X, Y).Type = TILE_TYPE_KEY Then
        ' This actually checks if its open or not
        If TempTile(X, Y).DoorOpen = 0 Then
            CheckDirection = True
            Exit Function
        End If
    End If

    ' Check to see if a player is already on that tile
    If Map.MapData.Moral = 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If

    ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 And MapNpc(i).Dead = NO Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next

    ' check if it's a drop warp - avoid if walking
    If ShiftDown Then
        If Map.TileData.Tile(X, Y).Type = TILE_TYPE_WARP Then
            If Map.TileData.Tile(X, Y).Data4 Then
                CheckDirection = True
            End If
        End If
    End If

End Function

Sub CheckMovement()

    If Not GettingMap Then
        If IsTryingToMove Then
            If CanMove Then

                ' Check if player has the shift key down for running
                If ShiftDown Then
                    Player(MyIndex).Moving = MOVING_RUNNING
                Else
                    Player(MyIndex).Moving = MOVING_WALKING
                End If

                Select Case GetPlayerDir(MyIndex)

                Case DIR_UP
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)

                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)

                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)

                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                End Select

                If Map.TileData.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                    GettingMap = True
                End If
            End If
        End If
    End If

End Sub

Public Function isInBounds()

    If (CurX >= 0) Then
        If (CurX <= Map.MapData.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MapData.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MapData.MaxX Then Exit Function
    If Y > Map.MapData.MaxY Then Exit Function
    IsValidMapPoint = True
End Function

Public Function IsItem(startX As Long, startY As Long) As Long
    Dim tempRec As RECT
    Dim i As Long
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) Then
            With tempRec
                .top = startY + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .bottom = .top + PIC_Y
                .Left = startX + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                    IsItem = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsTrade(startX As Long, startY As Long) As Long
    Dim tempRec As RECT
    Dim i As Long

    For i = 1 To MAX_INV
        With tempRec
            .top = startY + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
            .bottom = .top + PIC_Y
            .Left = startX + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))
            .Right = .Left + PIC_X
        End With

        If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
            If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                IsTrade = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsEqItem(startX As Long, startY As Long) As Long
    Dim tempRec As RECT
    Dim i As Long
    Dim xO As Integer, yO As Integer

    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(MyIndex, i) Then
            Select Case i
            Case Equipment.Helmet
                xO = 70
                yO = 80
            Case Equipment.Armor
                xO = 70
                yO = 132
            Case Equipment.Legs
                xO = 70
                yO = 184
            Case Equipment.Boots
                xO = 70
                yO = 236
            Case Equipment.Weapon
                xO = 18
                yO = 132
            Case Equipment.Shield
                xO = 122
                yO = 132
            Case Equipment.Amulet
                xO = 18
                yO = 80
            Case Equipment.RingLeft
                xO = 18
                yO = 184
            Case Equipment.RingRight
                xO = 122
                yO = 184
            End Select

            With tempRec
                .top = startY + yO
                .bottom = .top + PIC_Y
                .Left = startX + xO
                .Right = .Left + PIC_X
            End With

            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsEqSlot(ByVal startX As Long, ByVal startY As Long, ByVal InvSlot As Byte) As Byte
    Dim tempRec As RECT
    Dim i As Long
    Dim xO As Integer, yO As Integer
    Dim EqType As Byte
    Dim ItemID As Long
    
    ItemID = GetPlayerInvItemNum(MyIndex, InvSlot)

    If ItemID <= 0 Or ItemID > MAX_ITEMS Then Exit Function


    For i = 1 To Equipment.Equipment_Count - 1
        Select Case i
        Case Equipment.Helmet
            xO = 70
            yO = 80
            EqType = i
        Case Equipment.Armor
            xO = 70
            yO = 132
            EqType = i
        Case Equipment.Legs
            xO = 70
            yO = 184
            EqType = i
        Case Equipment.Boots
            xO = 70
            yO = 236
            EqType = i
        Case Equipment.Weapon
            xO = 18
            yO = 132
            EqType = i
        Case Equipment.Shield
            xO = 122
            yO = 132
            EqType = i
        Case Equipment.Amulet
            xO = 18
            yO = 80
            EqType = i
        Case Equipment.RingLeft
            xO = 18
            yO = 184
            EqType = i
        Case Equipment.RingRight
            xO = 122
            yO = 184
            EqType = i
        End Select

        With tempRec
            .top = startY + yO
            .bottom = .top + PIC_Y
            .Left = startX + xO
            .Right = .Left + PIC_X
        End With

        If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
            If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                If EqType = Item(ItemID).Type Then
                    IsEqSlot = EqType
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsSkill(startX As Long, startY As Long) As Long
    Dim tempRec As RECT
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If PlayerSpells(i).Spell Then
            With tempRec
                .top = startY + SkillTop + ((SkillOffsetY + 32) * ((i - 1) \ SkillColumns))
                .bottom = .top + PIC_Y
                .Left = startX + SkillLeft + ((SkillOffsetX + 32) * (((i - 1) Mod SkillColumns)))
                .Right = .Left + PIC_X
            End With

            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                    IsSkill = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsHotbar(startX As Long, startY As Long) As Long
    Dim tempRec As RECT
    Dim i As Long

    For i = 1 To MAX_HOTBAR
        If Hotbar(i).Slot Then
            With tempRec
                .top = startY + HotbarTop
                .bottom = .top + PIC_Y
                .Left = startX + HotbarLeft + ((i - 1) * HotbarOffsetX)
                .Right = .Left + PIC_X
            End With

            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                    IsHotbar = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Sub UseItem()

' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
End Sub

Public Sub ForgetSpell(ByVal spellSlot As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    ' dont let them forget a spell which is in CD
    If SpellCD(spellSlot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If

    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If

    If PlayerSpells(spellSlot).Spell > 0 Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CForgetSpell
        Buffer.WriteLong spellSlot
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If

End Sub

Public Sub CastSpell(ByVal spellSlot As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    If SpellCD(spellSlot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    
    If Player(MyIndex).StunDuration > 0 Then
        AddText "Nao pode usar Stunado!", BrightRed
    Exit Sub
    End If

    ' make sure we're not casting same spell
    If SpellBuffer > 0 Then
        If SpellBuffer = spellSlot Then
            ' stop them
            Exit Sub
        End If
    End If

    If PlayerSpells(spellSlot).Spell = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellSlot).Spell).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(PlayerSpells(spellSlot).Spell).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellSlot).Spell > 0 Then
        If getTime > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong CCast
                Buffer.WriteLong spellSlot
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                SpellBuffer = spellSlot
                SpellBufferTimer = getTime
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If

    Else
        Call AddText("No spell here.", BrightRed)
    End If

End Sub

Sub ClearTempTile()
    Dim X As Long
    Dim Y As Long
    ReDim TempTile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            TempTile(X, Y).DoorOpen = 0

            If Not GettingMap Then cacheRenderState X, Y, MapLayer.Mask
        Next
    Next

End Sub

Public Sub DevMsg(ByVal Text As String, ByVal Color As Byte)

    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(Text, Color)
        End If
    End If
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If

End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If

End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String

    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If

End Function

Public Sub CacheResources()
    Dim X As Long, Y As Long, Resource_Count As Long
    Resource_Count = 0

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY

            If Map.TileData.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If

        Next
    Next

    Resource_Index = Resource_Count
End Sub

Public Sub CreateActionMsg(ByVal message As String, ByVal Color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
    Dim i As Long
    ActionMsgIndex = ActionMsgIndex + 1

    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .message = message
        .Color = Color
        .Type = MsgType
        .Created = getTime
        .Scroll = 1
        .X = X
        .Y = Y
        .Alpha = 255
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMsgSCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    End If

    ' find the new high index
    For i = MAX_BYTE To 1 Step -1

        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If

    Next

    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
    Dim i As Long
    ActionMsg(Index).message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).Color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).X = 0
    ActionMsg(Index).Y = 0

    ' find the new high index
    For i = MAX_BYTE To 1 Step -1

        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If

    Next

    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
    Dim looptime As Long
    Dim Layer As Long
    Dim FrameCount As Long

    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub

    For Layer = 0 To 1

        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)

            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)

            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(Layer) = 0 Then AnimInstance(Index).frameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1

            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).Timer(Layer) + looptime <= getTime Then

                ' check if out of range
                If AnimInstance(Index).frameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1

                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).frameIndex(Layer) = 1
                    End If

                Else
                    AnimInstance(Index).frameIndex(Layer) = AnimInstance(Index).frameIndex(Layer) + 1
                End If

                AnimInstance(Index).Timer(Layer) = getTime
            End If
        End If

    Next

    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
End Sub

Public Function GetBankItemNum(ByVal BankSlot As Long) As Long

    If BankSlot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If

    If BankSlot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If

    GetBankItemNum = Bank.Item(BankSlot).num
End Function

Public Sub SetBankItemNum(ByVal BankSlot As Long, ByVal itemNum As Long)
    Bank.Item(BankSlot).num = itemNum
End Sub

Public Function GetBankItemValue(ByVal BankSlot As Long) As Long
    GetBankItemValue = Bank.Item(BankSlot).Value
End Function

Public Sub SetBankItemValue(ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank.Item(BankSlot).Value = ItemValue
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef dir As Byte, ByVal block As Boolean)

    If block Then
        blockvar = blockvar Or (2 ^ dir)
    Else
        blockvar = blockvar And Not (2 ^ dir)
    End If

End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean

    If Not blockvar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If

End Function

Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim soundName As String

    If entityNum <= 0 Then Exit Sub

    ' find the sound
    Select Case entityType

        ' animations
    Case SoundEntity.seAnimation

        If entityNum > MAX_ANIMATIONS Then Exit Sub
        soundName = Trim$(Animation(entityNum).sound)

        ' items
    Case SoundEntity.seItem

        If entityNum > MAX_ITEMS Then Exit Sub
        soundName = Trim$(Item(entityNum).sound)

        ' npcs
    Case SoundEntity.seNpc

        If entityNum > MAX_NPCS Then Exit Sub
        soundName = Trim$(NPC(entityNum).sound)

        ' resources
    Case SoundEntity.seResource

        If entityNum > MAX_RESOURCES Then Exit Sub
        soundName = Trim$(Resource(entityNum).sound)

        ' spells
    Case SoundEntity.seSpell

        If entityNum > MAX_SPELLS Then Exit Sub
        soundName = Trim$(Spell(entityNum).sound)

        ' other
    Case Else
        Exit Sub
    End Select

    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    If X > 0 And Y > 0 Then Play_Sound soundName, X, Y
End Sub

Public Sub CloseDialogue()
    diaIndex = 0
    HideWindow GetWindowIndex("winBlank")
    HideWindow GetWindowIndex("winDialogue")
End Sub

Public Sub Dialogue(ByVal header As String, ByVal body As String, ByVal body2 As String, ByVal Index As Long, Optional ByVal Style As Byte = 1, Optional ByVal Data1 As Long = 0)

' exit out if we've already got a dialogue open
    If diaIndex > 0 Then Exit Sub

    ' set buttons
    With Windows(GetWindowIndex("winDialogue"))
        If Style = StyleYesNo Then
            .Controls(GetControlIndex("winDialogue", "btnYes")).visible = True
            .Controls(GetControlIndex("winDialogue", "btnNo")).visible = True
            .Controls(GetControlIndex("winDialogue", "btnOkay")).visible = False
            .Controls(GetControlIndex("winDialogue", "txtInput")).visible = False
            .Controls(GetControlIndex("winDialogue", "lblBody_2")).visible = True
        ElseIf Style = StyleOKAY Then
            .Controls(GetControlIndex("winDialogue", "btnYes")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnNo")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnOkay")).visible = True
            .Controls(GetControlIndex("winDialogue", "txtInput")).visible = False
            .Controls(GetControlIndex("winDialogue", "lblBody_2")).visible = True
        ElseIf Style = StyleINPUT Then
            .Controls(GetControlIndex("winDialogue", "btnYes")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnNo")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnOkay")).visible = True
            .Controls(GetControlIndex("winDialogue", "txtInput")).visible = True
            .Controls(GetControlIndex("winDialogue", "lblBody_2")).visible = False
        End If

        ' set labels
        .Controls(GetControlIndex("winDialogue", "lblHeader")).Text = header
        .Controls(GetControlIndex("winDialogue", "lblBody_1")).Text = body
        .Controls(GetControlIndex("winDialogue", "lblBody_2")).Text = body2
        .Controls(GetControlIndex("winDialogue", "txtInput")).Text = vbNullString
    End With

    ' set it all up
    diaIndex = Index
    diaData1 = Data1
    diaStyle = Style

    ' make the windows visible
    ShowWindow GetWindowIndex("winBlank"), True
    ShowWindow GetWindowIndex("winDialogue"), True
End Sub

Public Sub dialogueHandler(ByVal Index As Long)
    Dim Value As Long, diaInput As String

    Dim Buffer As New clsBuffer
    Set Buffer = New clsBuffer

    diaInput = Trim$(Windows(GetWindowIndex("winDialogue")).Controls(GetControlIndex("winDialogue", "txtInput")).Text)

    ' find out which button
    If Index = 1 Then    ' okay button

        ' dialogue index
        Select Case diaIndex
        Case TypeTRADEAMOUNT
            Value = Val(diaInput)
            TradeItem diaData1, Value
        Case TypeDROPITEM
            Value = Val(diaInput)
            SendDropItem diaData1, Value
        Case TypeDEPOSITITEM
            Value = Val(diaInput)
            DepositItem diaData1, Value
        Case TypeWITHDRAWITEM
            Value = Val(diaInput)
            WithdrawItem diaData1, Value
        Case TypeTRADEGOLD
            Value = Val(diaInput)
            TradeGold Value
        Case TypeSENDBET
            Value = Val(diaInput)
            SendBet Value
        End Select

    ElseIf Index = 2 Then    ' yes button

        ' dialogue index
        Select Case diaIndex

        Case TypeTRADE
            SendAcceptTradeRequest

        Case TypeFORGET

            ForgetSpell diaData1

        Case TypePARTY
            SendAcceptParty

        Case TypeLOOTITEM
            ' send the packet
            Player(MyIndex).MapGetTimer = getTime
            Buffer.WriteLong CMapGetItem
            SendData Buffer.ToArray()
        Case TypeGUILD
            Buffer.WriteLong CGuildInviteResposta
            Buffer.WriteByte 1
            SendData Buffer.ToArray()
        
        End Select

    ElseIf Index = 3 Then    ' no button

        ' dialogue index
        Select Case diaIndex

        Case TypeTRADE
            SendDeclineTradeRequest

        Case TypePARTY
            SendDeclineParty
        
        Case TypeGUILD
            Buffer.WriteLong CGuildInviteResposta
            Buffer.WriteByte 0
            SendData Buffer.ToArray()
        End Select
    End If

    CloseDialogue
    diaIndex = 0
    diaInput = vbNullString
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ConvertMapX = X - (TileView.Left * PIC_X) - Camera.Left
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ConvertMapY = Y - (TileView.top * PIC_Y) - Camera.top
End Function

Public Sub UpdateCamera()
    Dim offsetX As Long, offsetY As Long, startX As Long, startY As Long, EndX As Long, EndY As Long

    offsetX = Player(MyIndex).xOffset + PIC_X
    offsetY = Player(MyIndex).yOffset + PIC_Y
    startX = GetPlayerX(MyIndex) - ((TileWidth + 1) \ 2) - 1
    startY = GetPlayerY(MyIndex) - ((TileHeight + 1) \ 2) - 1

    If TileWidth + 1 <= Map.MapData.MaxX Then
        If startX < 0 Then
            offsetX = 0

            If startX = -1 Then
                If Player(MyIndex).xOffset > 0 Then
                    offsetX = Player(MyIndex).xOffset
                End If
            End If

            startX = 0
        End If

        EndX = startX + (TileWidth + 1) + 1

        If EndX > Map.MapData.MaxX Then
            offsetX = 32

            If EndX = Map.MapData.MaxX + 1 Then
                If Player(MyIndex).xOffset < 0 Then
                    offsetX = Player(MyIndex).xOffset + PIC_X
                End If
            End If

            EndX = Map.MapData.MaxX
            startX = EndX - TileWidth - 1
        End If
    Else
        EndX = startX + (TileWidth + 1) + 1
    End If

    If TileHeight + 1 <= Map.MapData.MaxY Then
        If startY < 0 Then
            offsetY = 0

            If startY = -1 Then
                If Player(MyIndex).yOffset > 0 Then
                    offsetY = Player(MyIndex).yOffset
                End If
            End If

            startY = 0
        End If

        EndY = startY + (TileHeight + 1) + 1

        If EndY > Map.MapData.MaxY Then
            offsetY = 32

            If EndY = Map.MapData.MaxY + 1 Then
                If Player(MyIndex).yOffset < 0 Then
                    offsetY = Player(MyIndex).yOffset + PIC_Y
                End If
            End If

            EndY = Map.MapData.MaxY
            startY = EndY - TileHeight - 1
        End If
    Else
        EndY = startY + (TileHeight + 1) + 1
    End If

    If TileWidth + 1 = Map.MapData.MaxX Then
        offsetX = 0
    End If

    If TileHeight + 1 = Map.MapData.MaxY Then
        offsetY = 0
    End If

    With TileView
        .top = startY
        .bottom = EndY
        .Left = startX
        .Right = EndX
    End With

    With Camera
        .top = offsetY
        .bottom = .top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With

    CurX = TileView.Left + ((GlobalX + Camera.Left) \ PIC_X)
    CurY = TileView.top + ((GlobalY + Camera.top) \ PIC_Y)
    GlobalX_Map = GlobalX + (TileView.Left * PIC_X) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.top * PIC_Y) + Camera.top
End Sub

Public Function CensorWord(ByVal SString As String) As String
    CensorWord = String$(Len(SString), "*")
End Function

Public Sub placeAutotile(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)

    With Autotile(X, Y).Layer(layernum).QuarterTile(tileQuarter)

        Select Case autoTileLetter

        Case "a"
            .X = autoInner(1).X
            .Y = autoInner(1).Y

        Case "b"
            .X = autoInner(2).X
            .Y = autoInner(2).Y

        Case "c"
            .X = autoInner(3).X
            .Y = autoInner(3).Y

        Case "d"
            .X = autoInner(4).X
            .Y = autoInner(4).Y

        Case "e"
            .X = autoNW(1).X
            .Y = autoNW(1).Y

        Case "f"
            .X = autoNW(2).X
            .Y = autoNW(2).Y

        Case "g"
            .X = autoNW(3).X
            .Y = autoNW(3).Y

        Case "h"
            .X = autoNW(4).X
            .Y = autoNW(4).Y

        Case "i"
            .X = autoNE(1).X
            .Y = autoNE(1).Y

        Case "j"
            .X = autoNE(2).X
            .Y = autoNE(2).Y

        Case "k"
            .X = autoNE(3).X
            .Y = autoNE(3).Y

        Case "l"
            .X = autoNE(4).X
            .Y = autoNE(4).Y

        Case "m"
            .X = autoSW(1).X
            .Y = autoSW(1).Y

        Case "n"
            .X = autoSW(2).X
            .Y = autoSW(2).Y

        Case "o"
            .X = autoSW(3).X
            .Y = autoSW(3).Y

        Case "p"
            .X = autoSW(4).X
            .Y = autoSW(4).Y

        Case "q"
            .X = autoSE(1).X
            .Y = autoSE(1).Y

        Case "r"
            .X = autoSE(2).X
            .Y = autoSE(2).Y

        Case "s"
            .X = autoSE(3).X
            .Y = autoSE(3).Y

        Case "t"
            .X = autoSE(4).X
            .Y = autoSE(4).Y
        End Select

    End With

End Sub

Public Sub initAutotiles()
    Dim X As Long, Y As Long, layernum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).X = 32
    autoInner(1).Y = 0
    ' NE - b
    autoInner(2).X = 48
    autoInner(2).Y = 0
    ' SW - c
    autoInner(3).X = 32
    autoInner(3).Y = 16
    ' SE - d
    autoInner(4).X = 48
    autoInner(4).Y = 16
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).X = 0
    autoNW(1).Y = 32
    ' NE - f
    autoNW(2).X = 16
    autoNW(2).Y = 32
    ' SW - g
    autoNW(3).X = 0
    autoNW(3).Y = 48
    ' SE - h
    autoNW(4).X = 16
    autoNW(4).Y = 48
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).X = 32
    autoNE(1).Y = 32
    ' NE - g
    autoNE(2).X = 48
    autoNE(2).Y = 32
    ' SW - k
    autoNE(3).X = 32
    autoNE(3).Y = 48
    ' SE - l
    autoNE(4).X = 48
    autoNE(4).Y = 48
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).X = 0
    autoSW(1).Y = 64
    ' NE - n
    autoSW(2).X = 16
    autoSW(2).Y = 64
    ' SW - o
    autoSW(3).X = 0
    autoSW(3).Y = 80
    ' SE - p
    autoSW(4).X = 16
    autoSW(4).Y = 80
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).X = 32
    autoSE(1).Y = 64
    ' NE - r
    autoSE(2).X = 48
    autoSE(2).Y = 64
    ' SW - s
    autoSE(3).X = 32
    autoSE(3).Y = 80
    ' SE - t
    autoSE(4).X = 48
    autoSE(4).Y = 80

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            For layernum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                calculateAutotile X, Y, layernum
                ' cache the rendering state of the tiles and set them
                cacheRenderState X, Y, layernum
            Next
        Next
    Next

End Sub

Public Sub cacheRenderState(ByVal X As Long, ByVal Y As Long, ByVal layernum As Long)
    Dim quarterNum As Long
    
    On Error Resume Next

    ' exit out early
    If X < 0 Or X > Map.MapData.MaxX Or Y < 0 Or Y > Map.MapData.MaxY Then Exit Sub

    With Map.TileData.Tile(X, Y)

        ' check if the tile can be rendered
        If .Layer(layernum).tileSet <= 0 Or .Layer(layernum).tileSet > Count_Tileset Then
            Autotile(X, Y).Layer(layernum).renderState = RENDER_STATE_NONE
            Exit Sub
        End If

        ' check if we're a bottom
        If layernum = MapLayer.Ground Then
            ' check if bottom
            If Y > 0 Then
                If Map.TileData.Tile(X, Y - 1).Type = TILE_TYPE_APPEAR Then
                    If Map.TileData.Tile(X, Y - 1).Data2 Then
                        Autotile(X, Y).Layer(layernum).renderState = RENDER_STATE_APPEAR
                        Exit Sub
                    End If
                End If
            End If
        End If

        ' check if it's a key - hide mask if key is closed
        If layernum = MapLayer.Mask Then
            If .Type = TILE_TYPE_KEY Then
                If TempTile(X, Y).DoorOpen = 0 Then
                    Autotile(X, Y).Layer(layernum).renderState = RENDER_STATE_NONE
                    Exit Sub
                End If
            End If
            If .Type = TILE_TYPE_APPEAR Then
                Autotile(X, Y).Layer(layernum).renderState = RENDER_STATE_APPEAR
                Exit Sub
            End If
        End If

        ' check if it needs to be rendered as an autotile
        If .Autotile(layernum) = AUTOTILE_NONE Or .Autotile(layernum) = AUTOTILE_FAKE Or Options.NoAuto = 1 Then
            ' default to... default
            Autotile(X, Y).Layer(layernum).renderState = RENDER_STATE_NORMAL
        Else
            Autotile(X, Y).Layer(layernum).renderState = RENDER_STATE_AUTOTILE

            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(X, Y).Layer(layernum).srcX(quarterNum) = (Map.TileData.Tile(X, Y).Layer(layernum).X * 32) + Autotile(X, Y).Layer(layernum).QuarterTile(quarterNum).X
                Autotile(X, Y).Layer(layernum).srcY(quarterNum) = (Map.TileData.Tile(X, Y).Layer(layernum).Y * 32) + Autotile(X, Y).Layer(layernum).QuarterTile(quarterNum).Y
            Next

        End If

    End With

End Sub

Public Sub calculateAutotile(ByVal X As Long, ByVal Y As Long, ByVal layernum As Long)

' Right, so we've split the tile block in to an easy to remember
' collection of letters. We now need to do the calculations to find
' out which little lettered block needs to be rendered. We do this
' by reading the surrounding tiles to check for matches.
' First we check to make sure an autotile situation is actually there.
' Then we calculate exactly which situation has arisen.
' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
' Exit out if we don't have an auatotile
    If Map.TileData.Tile(X, Y).Autotile(layernum) = 0 Then Exit Sub

    ' Okay, we have autotiling but which one?
    Select Case Map.TileData.Tile(X, Y).Autotile(layernum)

        ' Normal or animated - same difference
    Case AUTOTILE_NORMAL, AUTOTILE_ANIM
        ' North West Quarter
        CalculateNW_Normal layernum, X, Y
        ' North East Quarter
        CalculateNE_Normal layernum, X, Y
        ' South West Quarter
        CalculateSW_Normal layernum, X, Y
        ' South East Quarter
        CalculateSE_Normal layernum, X, Y

        ' Cliff
    Case AUTOTILE_CLIFF
        ' North West Quarter
        CalculateNW_Cliff layernum, X, Y
        ' North East Quarter
        CalculateNE_Cliff layernum, X, Y
        ' South West Quarter
        CalculateSW_Cliff layernum, X, Y
        ' South East Quarter
        CalculateSE_Cliff layernum, X, Y

        ' Waterfalls
    Case AUTOTILE_WATERFALL
        ' North West Quarter
        CalculateNW_Waterfall layernum, X, Y
        ' North East Quarter
        CalculateNE_Waterfall layernum, X, Y
        ' South West Quarter
        CalculateSW_Waterfall layernum, X, Y
        ' South East Quarter
        CalculateSE_Waterfall layernum, X, Y

        ' Anything else
    Case Else
        ' Don't need to render anything... it's fake or not an autotile
    End Select

End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North West
    If checkTileMatch(layernum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True

    ' North
    If checkTileMatch(layernum, X, Y, X, Y - 1) Then tmpTile(2) = True

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

    Case AUTO_INNER
        placeAutotile layernum, X, Y, 1, "e"

    Case AUTO_OUTER
        placeAutotile layernum, X, Y, 1, "a"

    Case AUTO_HORIZONTAL
        placeAutotile layernum, X, Y, 1, "i"

    Case AUTO_VERTICAL
        placeAutotile layernum, X, Y, 1, "m"

    Case AUTO_FILL
        placeAutotile layernum, X, Y, 1, "q"
    End Select

End Sub

Public Sub CalculateNE_Normal(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North
    If checkTileMatch(layernum, X, Y, X, Y - 1) Then tmpTile(1) = True

    ' North East
    If checkTileMatch(layernum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

    Case AUTO_INNER
        placeAutotile layernum, X, Y, 2, "j"

    Case AUTO_OUTER
        placeAutotile layernum, X, Y, 2, "b"

    Case AUTO_HORIZONTAL
        placeAutotile layernum, X, Y, 2, "f"

    Case AUTO_VERTICAL
        placeAutotile layernum, X, Y, 2, "r"

    Case AUTO_FILL
        placeAutotile layernum, X, Y, 2, "n"
    End Select

End Sub

Public Sub CalculateSW_Normal(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile(1) = True

    ' South West
    If checkTileMatch(layernum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True

    ' South
    If checkTileMatch(layernum, X, Y, X, Y + 1) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

    Case AUTO_INNER
        placeAutotile layernum, X, Y, 3, "o"

    Case AUTO_OUTER
        placeAutotile layernum, X, Y, 3, "c"

    Case AUTO_HORIZONTAL
        placeAutotile layernum, X, Y, 3, "s"

    Case AUTO_VERTICAL
        placeAutotile layernum, X, Y, 3, "g"

    Case AUTO_FILL
        placeAutotile layernum, X, Y, 3, "k"
    End Select

End Sub

Public Sub CalculateSE_Normal(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' South
    If checkTileMatch(layernum, X, Y, X, Y + 1) Then tmpTile(1) = True

    ' South East
    If checkTileMatch(layernum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

    Case AUTO_INNER
        placeAutotile layernum, X, Y, 4, "t"

    Case AUTO_OUTER
        placeAutotile layernum, X, Y, 4, "d"

    Case AUTO_HORIZONTAL
        placeAutotile layernum, X, Y, 4, "p"

    Case AUTO_VERTICAL
        placeAutotile layernum, X, Y, 4, "l"

    Case AUTO_FILL
        placeAutotile layernum, X, Y, 4, "h"
    End Select

End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, X, Y, 1, "i"
    Else
        ' Edge
        placeAutotile layernum, X, Y, 1, "e"
    End If

End Sub

Public Sub CalculateNE_Waterfall(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, X, Y, 2, "f"
    Else
        ' Edge
        placeAutotile layernum, X, Y, 2, "j"
    End If

End Sub

Public Sub CalculateSW_Waterfall(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, X, Y, 3, "k"
    Else
        ' Edge
        placeAutotile layernum, X, Y, 3, "g"
    End If

End Sub

Public Sub CalculateSE_Waterfall(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, X, Y, 4, "h"
    Else
        ' Edge
        placeAutotile layernum, X, Y, 4, "l"
    End If

End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North West
    If checkTileMatch(layernum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True

    ' North
    If checkTileMatch(layernum, X, Y, X, Y - 1) Then tmpTile(2) = True

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile(3) = True

    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

    Case AUTO_INNER
        placeAutotile layernum, X, Y, 1, "e"

    Case AUTO_HORIZONTAL
        placeAutotile layernum, X, Y, 1, "i"

    Case AUTO_VERTICAL
        placeAutotile layernum, X, Y, 1, "m"

    Case AUTO_FILL
        placeAutotile layernum, X, Y, 1, "q"
    End Select

End Sub

Public Sub CalculateNE_Cliff(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North
    If checkTileMatch(layernum, X, Y, X, Y - 1) Then tmpTile(1) = True

    ' North East
    If checkTileMatch(layernum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile(3) = True

    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

    Case AUTO_INNER
        placeAutotile layernum, X, Y, 2, "j"

    Case AUTO_HORIZONTAL
        placeAutotile layernum, X, Y, 2, "f"

    Case AUTO_VERTICAL
        placeAutotile layernum, X, Y, 2, "r"

    Case AUTO_FILL
        placeAutotile layernum, X, Y, 2, "n"
    End Select

End Sub

Public Sub CalculateSW_Cliff(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' West
    If checkTileMatch(layernum, X, Y, X - 1, Y) Then tmpTile(1) = True

    ' South West
    If checkTileMatch(layernum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True

    ' South
    If checkTileMatch(layernum, X, Y, X, Y + 1) Then tmpTile(3) = True

    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

    Case AUTO_INNER
        placeAutotile layernum, X, Y, 3, "o"

    Case AUTO_HORIZONTAL
        placeAutotile layernum, X, Y, 3, "s"

    Case AUTO_VERTICAL
        placeAutotile layernum, X, Y, 3, "g"

    Case AUTO_FILL
        placeAutotile layernum, X, Y, 3, "k"
    End Select

End Sub

Public Sub CalculateSE_Cliff(ByVal layernum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' South
    If checkTileMatch(layernum, X, Y, X, Y + 1) Then tmpTile(1) = True

    ' South East
    If checkTileMatch(layernum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, X, Y, X + 1, Y) Then tmpTile(3) = True

    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

    Case AUTO_INNER
        placeAutotile layernum, X, Y, 4, "t"

    Case AUTO_HORIZONTAL
        placeAutotile layernum, X, Y, 4, "p"

    Case AUTO_VERTICAL
        placeAutotile layernum, X, Y, 4, "l"

    Case AUTO_FILL
        placeAutotile layernum, X, Y, 4, "h"
    End Select

End Sub

Public Function checkTileMatch(ByVal layernum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
' we'll exit out early if true
    checkTileMatch = True

    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MapData.MaxX Or Y2 < 0 Or Y2 > Map.MapData.MaxY Then
        checkTileMatch = True
        Exit Function
    End If

    ' fakes ALWAYS return true
    If Map.TileData.Tile(X2, Y2).Autotile(layernum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If

    ' check neighbour is an autotile
    If Map.TileData.Tile(X2, Y2).Autotile(layernum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If

    ' check we're a matching
    If Map.TileData.Tile(X1, Y1).Layer(layernum).tileSet <> Map.TileData.Tile(X2, Y2).Layer(layernum).tileSet Then
        checkTileMatch = False
        Exit Function
    End If

    ' check tiles match
    If Map.TileData.Tile(X1, Y1).Layer(layernum).X <> Map.TileData.Tile(X2, Y2).Layer(layernum).X Then
        checkTileMatch = False
        Exit Function
    End If

    If Map.TileData.Tile(X1, Y1).Layer(layernum).Y <> Map.TileData.Tile(X2, Y2).Layer(layernum).Y Then
        checkTileMatch = False
        Exit Function
    End If

End Function

Public Sub OpenNpcChat(ByVal NpcNum As Long, ByVal mT As String, ByRef o() As String)
    Dim i As Long, X As Long

    ' find out how many options we have
    convOptions = 0
    For i = 1 To 4
        If Len(Trim$(o(i))) > 0 Then convOptions = convOptions + 1
    Next

    ' gui stuff
    With Windows(GetWindowIndex("winNpcChat"))
        ' set main text

        .Window.Text = "Conversation with " & Trim$(NPC(NpcNum).Name)

        .Controls(GetControlIndex("winNpcChat", "lblChat")).Text = mT
        ' make everything visible

        For i = 1 To 4
            .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).top = optPos(i)
            .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).visible = True
        Next

        ' set sizes
        .Window.Height = optHeight
        .Controls(GetControlIndex("winNpcChat", "picParchment")).Height = .Window.Height - 30
        ' move options depending on count
        If convOptions < 4 Then
            For i = convOptions + 1 To 4
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).top = optPos(i)
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).visible = False
            Next
            For i = 1 To convOptions
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).top = optPos(i + (4 - convOptions))
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).visible = True
            Next
            .Window.Height = optHeight - ((4 - convOptions) * 18)
            .Controls(GetControlIndex("winNpcChat", "picParchment")).Height = .Window.Height - 32
        End If
        ' set labels
        X = convOptions
        For i = 1 To 4
            .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).Text = X & ". " & o(i)
            X = X - 1
        Next

        If NpcNum > 0 Then
            For i = 0 To 5
                .Controls(GetControlIndex("winNpcChat", "picFace")).image(i) = Tex_Face(NPC(NpcNum).Sprite)
            Next
        End If

        '.Window. -100
    End With

    ' we're in chat now boy
    inChat = True

    ' show the window
    ShowWindow GetWindowIndex("winNpcChat")
End Sub

Public Sub SetTutorialState(ByVal stateNum As Byte)
    Dim i As Long

    Select Case stateNum

    Case 1    ' introduction
        chatText = "Ah, so you have appeared at last my dear. Please, listen to what I have to say."
        chatOpt(1) = "*sigh* I suppose I should..."

        For i = 2 To 4
            chatOpt(i) = vbNullString
        Next

    Case 2    ' next
        chatText = "There are some important things you need to know. Here they are. To move, use W, A, S and D. To attack or to talk to someone, press CTRL. To initiate chat press ENTER."
        chatOpt(1) = "Go on..."

        For i = 2 To 4
            chatOpt(i) = vbNullString
        Next

    Case 3    ' chatting
        chatText = "When chatting you can talk in different channels. By default you're talking in the map channel. To talk globally append an apostrophe (') to the start of your message. To perform an emote append a hyphen (-) to the start of your message."
        chatOpt(1) = "Wait, what about combat?"

        For i = 2 To 4
            chatOpt(i) = vbNullString
        Next

    Case 4    ' combat
        chatText = "Combat can be done through melee and skills. You can melee an enemy by facing them and pressing CTRL. To use a skill you can double click it in your skill menu, double click it in the hotbar or use the number keys. (1, 2, 3, etc.)"
        chatOpt(1) = "Oh! What do stats do?"

        For i = 2 To 4
            chatOpt(i) = vbNullString
        Next

    Case 5    ' stats
        chatText = "Strength increases damage and allows you to equip better weaponry. Endurance increases your maximum health. Intelligence increases your maximum spirit. Agility allows you to reduce damage received and also increases critical hit chances. Willpower increase regeneration abilities."
        chatOpt(1) = "Thanks. See you later."

        For i = 2 To 4
            chatOpt(i) = vbNullString
        Next

    Case Else    ' goodbye
        chatText = vbNullString

        For i = 1 To 4
            chatOpt(i) = vbNullString
        Next

        SendFinishTutorial
        inTutorial = False
        AddText "Well done, you finished the tutorial.", BrightGreen
        Exit Sub
    End Select

    ' set the state
    tutorialState = stateNum
End Sub

Public Sub ScrollChatBox(ByVal direction As Byte)
    If direction = 0 Then    ' up
        If ChatScroll < ChatLines Then
            ChatScroll = ChatScroll + 1
        End If
    Else
        If ChatScroll > 0 Then
            ChatScroll = ChatScroll - 1
        End If
    End If
End Sub

Public Sub ClearMapCache()
    Dim i As Long, filename As String

    For i = 1 To MAX_MAPS
        filename = App.path & "\data files\maps\map" & i & ".map"

        If FileExist(filename) Then
            Kill filename
        End If

    Next

    AddText "Map cache destroyed.", BrightGreen
End Sub

Public Sub AddChatBubble(ByVal Target As Long, ByVal TargetType As Byte, ByVal Msg As String, ByVal Colour As Long)
    Dim i As Long, Index As Long
    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1

    ' reset to yourself for eventing
    If TargetType = 0 Then
        TargetType = TARGET_TYPE_PLAYER
        If Target = 0 Then Target = MyIndex
    End If

    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    ' default to new bubble
    Index = chatBubbleIndex

    ' loop through and see if that player/npc already has a chat bubble
    For i = 1 To MAX_BYTE
        If chatBubble(i).TargetType = TargetType Then
            If chatBubble(i).Target = Target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                Index = i
                Exit For
            End If
        End If
    Next

    ' set the bubble up
    With chatBubble(Index)
        .Target = Target
        .TargetType = TargetType
        .Msg = Msg
        .Colour = Colour
        .Timer = getTime
        .Active = True
    End With
End Sub

Public Sub FindNearestTarget()
    Dim i As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, xDif As Long, yDif As Long
    Dim bestX As Long, bestY As Long, bestIndex As Long
    X2 = GetPlayerX(MyIndex)
    Y2 = GetPlayerY(MyIndex)
    bestX = 255
    bestY = 255

    IsNpc = Not IsNpc

    If IsNpc Then    ' se for npc
        X2 = GetPlayerX(MyIndex)
        Y2 = GetPlayerY(MyIndex)

        bestX = 255
        bestY = 255

        For i = 1 To MAX_MAP_NPCS
            If MapNpc(i).num > 0 Then
                X = MapNpc(i).X
                Y = MapNpc(i).Y
                ' find the difference - x
                If X < X2 Then
                    xDif = X2 - X
                ElseIf X > X2 Then
                    xDif = X - X2
                Else
                    xDif = 0
                End If
                ' find the difference - y
                If Y < Y2 Then
                    yDif = Y2 - Y
                ElseIf Y > Y2 Then
                    yDif = Y - Y2
                Else
                    yDif = 0
                End If
                ' best so far?
                If (xDif + yDif) < (bestX + bestY) Then
                    bestX = xDif
                    bestY = yDif
                    bestIndex = i
                End If
            End If
        Next

        ' target the best
        If bestIndex > 0 And bestIndex <> myTarget Then
            PlayerTarget bestIndex, TARGET_TYPE_NPC
            UpdateEnemyInterface
        End If

    Else    'se for player

        X2 = GetPlayerX(MyIndex)
        Y2 = GetPlayerY(MyIndex)

        bestX = 255
        bestY = 255

        For i = 1 To Player_HighIndex
            If i <> MyIndex Then
                If Player(i).Map = Player(MyIndex).Map Then
                    X = (Player(i).X)
                    Y = (Player(i).Y)
                    ' find the difference - x
                    If X < X2 Then
                        xDif = X2 - X
                    ElseIf X > X2 Then
                        xDif = X - X2
                    Else
                        xDif = 0
                    End If
                    ' find the difference - y
                    If Y < Y2 Then
                        yDif = Y2 - Y
                    ElseIf Y > Y2 Then
                        yDif = Y - Y2
                    Else
                        yDif = 0
                    End If
                    ' best so far?
                    If (xDif + yDif) < (bestX + bestY) Then
                        bestX = xDif
                        bestY = yDif
                        bestIndex = i
                    End If
                End If
            End If
        Next

        ' target the best
        If bestIndex > 0 And bestIndex <> myTarget Then
            PlayerTarget bestIndex, TARGET_TYPE_PLAYER
            UpdateEnemyInterface
        End If

    End If
End Sub

Public Sub FindTarget()
    Dim i As Long, X As Long, Y As Long

    ' check players
    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            X = (GetPlayerX(i) * 32) + Player(i).xOffset + 32
            Y = (GetPlayerY(i) * 32) + Player(i).yOffset + 32

            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget i, TARGET_TYPE_PLAYER
                    UpdateEnemyInterface
                    Exit Sub
                End If
            End If
        End If

    Next

    ' check npcs
    For i = 1 To MAX_MAP_NPCS

        If MapNpc(i).num > 0 Then
            X = (MapNpc(i).X * 32) + MapNpc(i).xOffset + 32
            Y = (MapNpc(i).Y * 32) + MapNpc(i).yOffset + 32

            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget i, TARGET_TYPE_NPC
                    UpdateEnemyInterface
                    Exit Sub
                End If
            End If
        End If

    Next

End Sub

Public Sub SetBarWidth(ByRef MaxWidth As Long, ByRef Width As Long)
    Dim barDifference As Long

    If MaxWidth < Width Then
        ' find out the amount to increase per loop
        barDifference = ((Width - MaxWidth) / 100) * 10

        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        Width = Width - barDifference
    ElseIf MaxWidth > Width Then
        ' find out the amount to increase per loop
        barDifference = ((MaxWidth - Width) / 100) * 10

        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        Width = Width + barDifference
    End If

End Sub

Public Sub AttemptLogin()
    TcpInit GAME_SERVER_IP, GAME_SERVER_PORT

    ' send login packet
    If ConnectToServer Then
        SendLogin Windows(GetWindowIndex("winLogin")).Controls(GetControlIndex("winLogin", "txtUser")).Text
        Exit Sub
    End If

    If Not IsConnected And Not isReconnect Then
        ShowWindow GetWindowIndex("winLogin")
        Dialogue "Connection Problem", "Cannot connect to game server.", "Please try again later.", TypeALERT
    End If
End Sub

Public Sub DialogueAlert(ByVal Index As Long)
    Dim header As String, body As String, body2 As String

    ' find the body/header
    Select Case Index

    Case MsgCONNECTION
        header = "Problema de Conexao"
        body = "Perdeu a conexao com o servidor."
        body2 = "Tente novamente em instantes."

    Case MsgBANNED
        header = "Banned"
        body = "Voce foi banido e nao pode jogar."
        body2 = "Crie uma nova conta."

    Case MsgKICKED
        header = "Kicked"
        body = "Voce foi expulso."
        body2 = "Entre novamente."

    Case MsgOUTDATED
        header = "Versao Desatualizada"
        body = "Versao nova foi elaborada."
        body2 = "Por favor atualize seu Client."

    Case MsgUSERLENGTH
        header = "Usuario Invalido"
        body = "Seu usuario esta muito grande ou muito pequeno."
        body2 = "Por favor, corrigir."

    Case MsgUSERILlEGAL
        header = "Caracteres Ilegais"
        body = "Your username or password contains illegal characters."
        body2 = "Please enter a valid username and password."

    Case MsgREBOOTING
        header = "Conexao Negada"
        body = "Servidor reiniciando."
        body2 = "Tente novamente em instantes."

    Case MsgNAMETAKEN
        header = "Nome Invalido"
        body = "Nome em uso."
        body2 = "Tente outro nome."

    Case MsgNAMELENGTH
        header = "Nome Invalido"
        body = "Nome muito pequeno ou muito grande."
        body2 = "Tente outro nome."

    Case MsgNAMEILLEGAL
        header = "Nome Invalido"
        body = "O Nome contem caracteres ilegais."
        body2 = "Use outro nome."

    Case MsgWRONGPASS
        header = "Login Invalido"
        body = "Usuario ou senha invalido."
        body2 = "Tente novamente."

    Case MsgCreated
        header = "Conta Criada"
        body = "Sua conta foi criada com sucesso."
        body2 = "Agora, voce pode jogar!"

    Case MsgEMAILINVALID
        header = "Email Invalido"
        body = "Use o modelo ####@email.com."
        body2 = "Por Favor corrigir!"

    Case MsgPASSLENGTH
        header = "Senha muito grande ou muito pequena"
        body = "Senha MIN 3 MAX " & NAME_LENGTH & " Caracteres!"
        body2 = "Por Favor corrigir!"

    Case MsgPASSNULL
        header = "Senha em branco"
        body = "Digite uma senha."
        body2 = "Por Favor corrigir!"

    Case MsgUSERNULL
        header = "Usuario em branco"
        body = "Digite um usuario"
        body2 = "Por Favor corrigir!"

    Case MsgPASSCONFIRM
        header = "Senhas Divergentes"
        body = "Nao estao iguais"
        body2 = "Por Favor corrigir!"

    Case MsgCAPTCHAINCORRECT
        header = "Error"
        body = "Captcha Incorreto"
        body2 = "Por Favor tente novamente!"
    Case MsgSERIALINCORRECT
        header = "Error"
        body = "Serial Incorrect"
        body2 = "Por Favor digite novamente!"
    Case MsgSERIALCLAIMED
        header = "Success"
        body = "O Serial foi resgatado!"
        body2 = "Verifique no chat os prêmios."
    Case MsgINVALIDBIRTHDAY
        header = "Error, Birthday Incorrect"
        body = "Use o formato '##/##/####"
        body2 = "Pode usar '\' ou '/' ou '-' para separar o dia, mes e ano!"
        
    Case MsgLOTTERYMAXBID
        header = "Invalid Bet"
        body = "Exceeded the maximum value"
        body2 = "Max " & LotteryInfo.Max_Bets_Value & "g"
    Case MsgLOTTERYMINBID
        header = "Invalid Bet"
        body = "Exceeded the minimum value"
        body2 = "Min " & LotteryInfo.Min_Bets_Value & "g"
    Case MsgLOTTERYNUMBERS
        header = "Invalid Bet"
        body = "Invalid Number of Bet"
        body2 = "Bets From 1 to " & MAX_BETS
    Case MsgLOTTERYNUMBERALREADY
        header = "Invalid Bet"
        body = "This number already has a bet"
        body2 = "Choose another"
    Case MsgLOTTERYCLOSED
        header = "Invalid Bet"
        body = "The betting period is closed"
        body2 = "Come back later"
    Case MsgLOTTERYGOLD
        header = "Invalid Bet"
        body = "You do not own the amount"
        body2 = "you are trying to bet"
    Case MsgLOTTERYSUCCESS
        header = "SUCCESS!"
        body = "You made your bet"
        body2 = "Good Lucky!!!"
    End Select

    ' set the dialogue up!
    Dialogue header, body, body2, TypeALERT
End Sub

Public Function hasProficiency(ByVal Index As Long, ByVal proficiency As Long) As Boolean

    Select Case proficiency

    Case 0    ' None
        hasProficiency = True
        Exit Function

    Case 1    ' Heavy

        If GetPlayerClass(Index) = 1 Then
            hasProficiency = True
            Exit Function
        End If

    Case 2    ' Light

        If GetPlayerClass(Index) = 2 Or GetPlayerClass(Index) = 3 Then
            hasProficiency = True
            Exit Function
        End If

    End Select

    hasProficiency = False
End Function

Public Function Clamp(ByVal Value As Long, ByVal Min As Long, ByVal max As Long) As Long
    Clamp = Value

    If Value < Min Then Clamp = Min
    If Value > max Then Clamp = max
End Function

Public Sub ShowClasses()
    HideWindows
    newCharClass = 1
    newCharSprite = 1
    newCharGender = SEX_MALE
    Windows(GetWindowIndex("winClasses")).Controls(GetControlIndex("winClasses", "lblClassName")).Text = Trim$(Class(newCharClass).Name)
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "txtName")).Text = vbNullString
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "chkMale")).Value = 1
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "chkFemale")).Value = 0
    ShowWindow GetWindowIndex("winClasses")
End Sub

Public Sub SetGoldLabel()
    Dim i As Long, Amount As Long
    Amount = GetPlayerGold(MyIndex)
    'For i = 1 To MAX_INV
    '    If GetPlayerInvItemNum(MyIndex, i) > 0 Then
    '        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
    '            Amount = Amount + (Item(GetPlayerInvItemNum(MyIndex, i)).price * GetPlayerInvItemValue(MyIndex, i))
    '        End If
    '    End If
    'Next
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "lblGold")).Text = Format$(Amount, "#,###,###,###") & " $"
    Windows(GetWindowIndex("winInventory")).Controls(GetControlIndex("winInventory", "lblGold")).Text = Format$(Amount, "#,###,###,###") & " $"
End Sub

Public Sub ShowInvDesc(X As Long, Y As Long, invNum As Long)
    Dim SoulBound As Boolean

    ' rte9
    If invNum <= 0 Or invNum > MAX_INV Then Exit Sub

    ' show
    If GetPlayerInvItemNum(MyIndex, invNum) Then
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).BindType > 0 And PlayerInv(invNum).bound > 0 Then SoulBound = True
        ShowItemDesc X, Y, GetPlayerInvItemNum(MyIndex, invNum), SoulBound
    End If
End Sub

Public Sub ShowShopDesc(X As Long, Y As Long, itemNum As Long)
    If itemNum <= 0 Or itemNum > MAX_ITEMS Then Exit Sub
    ' show
    ShowItemDesc X, Y, itemNum, False
End Sub

Public Sub ShowEqDesc(X As Long, Y As Long, eqNum As Long)
    Dim SoulBound As Boolean

    ' rte9
    If eqNum <= 0 Or eqNum > Equipment.Equipment_Count - 1 Then Exit Sub

    ' show
    If Player(MyIndex).Equipment(eqNum) Then
        If Item(Player(MyIndex).Equipment(eqNum)).BindType > 0 Then SoulBound = True
        ShowItemDesc X, Y, Player(MyIndex).Equipment(eqNum), SoulBound
    End If
End Sub

Public Sub ShowPlayerSpellDesc(X As Long, Y As Long, slotNum As Long)

' rte9
    If slotNum <= 0 Or slotNum > MAX_PLAYER_SPELLS Then Exit Sub

    ' show
    If PlayerSpells(slotNum).Spell Then
        ShowSpellDesc X, Y, PlayerSpells(slotNum).Spell, slotNum
    End If
End Sub

Public Sub ShowSpellDesc(X As Long, Y As Long, spellnum As Long, spellSlot As Long)
    Dim Colour As Long, theName As String, sUse As String, i As Long, barWidth As Long, tmpWidth As Long

    ' set globals
    descType = 2    ' spell
    descItem = spellnum

    ' set position
    Windows(GetWindowIndex("winDescription")).Window.Left = X
    Windows(GetWindowIndex("winDescription")).Window.top = Y

    ' show the window
    ShowWindow GetWindowIndex("winDescription"), , False

    ' exit out early if last is same
    If (descLastType = descType) And (descLastItem = descItem) Then Exit Sub

    ' clear
    ReDim descText(1 To 1) As TextColourRec

    ' hide req. labels
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblLevel")).visible = False
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "picBar")).visible = True

    ' set variables
    With Windows(GetWindowIndex("winDescription"))
        ' set name
        .Controls(GetControlIndex("winDescription", "lblName")).Text = Trim$(Spell(spellnum).Name)
        .Controls(GetControlIndex("winDescription", "lblName")).textColour = White

        ' find ranks
        If spellSlot > 0 Then
            ' draw the rank bar
            barWidth = 66
            If Spell(spellnum).NextRank > 0 Then
                tmpWidth = ((PlayerSpells(spellSlot).Uses / barWidth) / (Spell(spellnum).NextUses / barWidth)) * barWidth
            Else
                tmpWidth = 66
            End If
            .Controls(GetControlIndex("winDescription", "picBar")).Value = tmpWidth
            ' does it rank up?
            If Spell(spellnum).NextRank > 0 Then
                Colour = White
                sUse = "Uses: " & PlayerSpells(spellSlot).Uses & "/" & Spell(spellnum).NextUses
                If PlayerSpells(spellSlot).Uses = Spell(spellnum).NextUses Then
                    If Not GetPlayerLevel(MyIndex) >= Spell(Spell(spellnum).NextRank).LevelReq Then
                        Colour = BrightRed
                        sUse = "Lvl " & Spell(Spell(spellnum).NextRank).LevelReq & " req."
                    End If
                End If
            Else
                Colour = Grey
                sUse = "Max Rank"
            End If
            ' show controls
            .Controls(GetControlIndex("winDescription", "lblClass")).visible = True
            .Controls(GetControlIndex("winDescription", "picBar")).visible = True
            'set vals
            .Controls(GetControlIndex("winDescription", "lblClass")).Text = sUse
            .Controls(GetControlIndex("winDescription", "lblClass")).textColour = Colour
        Else
            ' hide some controls
            .Controls(GetControlIndex("winDescription", "lblClass")).visible = False
            .Controls(GetControlIndex("winDescription", "picBar")).visible = False
        End If
    End With

    Select Case Spell(spellnum).Type
    Case SPELL_TYPE_DAMAGEHP
        AddDescInfo "Damage HP"
    Case SPELL_TYPE_DAMAGEMP
        AddDescInfo "Damage SP"
    Case SPELL_TYPE_HEALHP
        AddDescInfo "Heal HP"
    Case SPELL_TYPE_HEALMP
        AddDescInfo "Heal SP"
    Case SPELL_TYPE_WARP
        AddDescInfo "Warp"
    End Select

    ' more info
    Select Case Spell(spellnum).Type
    Case SPELL_TYPE_DAMAGEHP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
        ' damage
        AddDescInfo "Vital: " & Spell(spellnum).Vital

        ' mp cost
        AddDescInfo "Cost: " & Spell(spellnum).MPCost & " SP"

        ' cast time
        AddDescInfo "Cast Time: " & Spell(spellnum).CastTime & "s"

        ' cd time
        AddDescInfo "Cooldown: " & Spell(spellnum).CDTime & "s"

        ' aoe
        If Spell(spellnum).AoE > 0 Then
            AddDescInfo "AoE: " & Spell(spellnum).AoE
        End If

        ' stun
        If Spell(spellnum).StunDuration > 0 Then
            AddDescInfo "Stun: " & Spell(spellnum).StunDuration & "s"
        End If

        ' dot
        If Spell(spellnum).Duration > 0 And Spell(spellnum).Interval > 0 Then
            AddDescInfo "DoT: " & (Spell(spellnum).Duration / Spell(spellnum).Interval) & " tick"
        End If
    End Select
End Sub

Public Sub ShowItemDesc(X As Long, Y As Long, itemNum As Long, SoulBound As Boolean)
    Dim Colour As Long, theName As String, className As String, levelTxt As String, i As Long

    ' set globals
    descType = 1    ' inventory
    descItem = itemNum

    ' set position
    Windows(GetWindowIndex("winDescription")).Window.Left = X
    Windows(GetWindowIndex("winDescription")).Window.top = Y

    ' show the window
    ShowWindow GetWindowIndex("winDescription"), , False

    ' exit out early if last is same
    If (descLastType = descType) And (descLastItem = descItem) Then Exit Sub

    ' set last to this
    descLastType = descType
    descLastItem = descItem

    ' show req. labels
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblClass")).visible = True
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblLevel")).visible = True
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "picBar")).visible = False

    ' set variables
    With Windows(GetWindowIndex("winDescription"))
        ' name
        If Not SoulBound Then
            theName = Trim$(Item(itemNum).Name)
        Else
            theName = "(SB) " & Trim$(Item(itemNum).Name)
        End If
        .Controls(GetControlIndex("winDescription", "lblName")).Text = theName

        Colour = GetItemNameColour(Item(itemNum).Rarity)

        .Controls(GetControlIndex("winDescription", "lblName")).textColour = Colour
        ' class req
        If Item(itemNum).ClassReq > 0 Then
            className = Trim$(Class(Item(itemNum).ClassReq).Name)
            ' do we match it?
            If GetPlayerClass(MyIndex) = Item(itemNum).ClassReq Then
                Colour = Green
            Else
                Colour = BrightRed
            End If
        ElseIf Item(itemNum).proficiency > 0 Then
            Select Case Item(itemNum).proficiency
            Case 1    ' Sword/Armour
                If Item(itemNum).Type >= ITEM_TYPE_ARMOR And Item(itemNum).Type <= ITEM_TYPE_RINGRIGHT Then
                    className = "Heavy Armour"
                ElseIf Item(itemNum).Type = ITEM_TYPE_WEAPON Then
                    className = "Heavy Weapon"
                End If
                If hasProficiency(MyIndex, Item(itemNum).proficiency) Then
                    Colour = Green
                Else
                    Colour = BrightRed
                End If
            Case 2    ' Staff/Cloth
                If Item(itemNum).Type >= ITEM_TYPE_ARMOR And Item(itemNum).Type <= ITEM_TYPE_RINGRIGHT Then
                    className = "Cloth Armour"
                ElseIf Item(itemNum).Type = ITEM_TYPE_WEAPON Then
                    className = "Light Weapon"
                End If
                If hasProficiency(MyIndex, Item(itemNum).proficiency) Then
                    Colour = Green
                Else
                    Colour = BrightRed
                End If
            End Select
        Else
            className = "No class req."
            Colour = Green
        End If
        .Controls(GetControlIndex("winDescription", "lblClass")).Text = className
        .Controls(GetControlIndex("winDescription", "lblClass")).textColour = Colour
        ' level
        If Item(itemNum).LevelReq > 0 Then
            levelTxt = "Level " & Item(itemNum).LevelReq
            ' do we match it?
            If GetPlayerLevel(MyIndex) >= Item(itemNum).LevelReq Then
                Colour = Green
            Else
                Colour = BrightRed
            End If
        Else
            levelTxt = "No level req."
            Colour = Green
        End If
        .Controls(GetControlIndex("winDescription", "lblLevel")).Text = levelTxt
        .Controls(GetControlIndex("winDescription", "lblLevel")).textColour = Colour
    End With

    ' clear
    ReDim descText(1 To 1) As TextColourRec

    ' go through the rest of the text
    Select Case Item(itemNum).Type
    Case ITEM_TYPE_NONE
        AddDescInfo "No type"
    Case ITEM_TYPE_WEAPON
        AddDescInfo "Weapon"
    Case ITEM_TYPE_ARMOR
        AddDescInfo "Armour"
    Case ITEM_TYPE_HELMET
        AddDescInfo "Helmet"
    Case ITEM_TYPE_SHIELD
        AddDescInfo "Shield"
    Case ITEM_TYPE_LEGS
        AddDescInfo "Legs"
    Case ITEM_TYPE_BOOTS
        AddDescInfo "Boots"
    Case ITEM_TYPE_AMULET
        AddDescInfo "Amulet"
    Case ITEM_TYPE_RINGLEFT
        AddDescInfo "Ring Left"
    Case ITEM_TYPE_RINGRIGHT
        AddDescInfo "Ring Right"
    Case ITEM_TYPE_CONSUME
        AddDescInfo "Consume"
    Case ITEM_TYPE_KEY
        AddDescInfo "Key"
    Case ITEM_TYPE_CURRENCY
        AddDescInfo "Currency"
    Case ITEM_TYPE_SPELL
        AddDescInfo "Spell"
    Case ITEM_TYPE_FOOD
        AddDescInfo "Food"
    Case ITEM_TYPE_PROTECTDROP
        AddDescInfo "Protect Drop"
    End Select

    ' more info
    Select Case Item(itemNum).Type
    Case ITEM_TYPE_NONE, ITEM_TYPE_KEY, ITEM_TYPE_CURRENCY
        ' binding
        If Item(itemNum).BindType = 1 Then
            AddDescInfo "Bind on Pickup"
        ElseIf Item(itemNum).BindType = 2 Then
            AddDescInfo "Bind on Equip"
        End If
        ' price
        AddDescInfo "Value: " & Item(itemNum).price & "g"
    Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SHIELD, ITEM_TYPE_LEGS, ITEM_TYPE_BOOTS, ITEM_TYPE_AMULET, ITEM_TYPE_RINGLEFT, ITEM_TYPE_RINGRIGHT
        ' damage/defence
        If Item(itemNum).Type = ITEM_TYPE_WEAPON Then
            If Item(itemNum).Data2_Percent > 0 Then
                AddDescInfo "+Damage: " & Item(itemNum).Data2 & "%"
            Else
                AddDescInfo "Damage: " & Item(itemNum).Data2
            End If
            ' speed
            AddDescInfo "Speed: " & (Item(itemNum).Speed / 1000) & "s"
        Else
            If Item(itemNum).Data2 > 0 Then
                If Item(itemNum).Data2_Percent > 0 Then
                    AddDescInfo "+Defence: " & Item(itemNum).Data2 & "%"
                Else
                    AddDescInfo "Defence: " & Item(itemNum).Data2
                End If
            End If
        End If

        If Item(itemNum).Type = ITEM_TYPE_SHIELD Then
            AddDescInfo "Block: " & Item(itemNum).BlockChance & "%"
        End If

        ' binding
        If Item(itemNum).BindType = 1 Then
            AddDescInfo "Bind on Pickup"
        ElseIf Item(itemNum).BindType = 2 Then
            AddDescInfo "Bind on Equip"
        End If
        ' price
        AddDescInfo "Value: " & Item(itemNum).price & "g"
        ' stat bonuses
        If Item(itemNum).Add_Stat(Stats.strength) > 0 Then
            If Item(itemNum).Stat_Percent(Stats.strength) > 0 Then
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.strength) & "% Str"
            Else
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.strength) & " Str"
            End If
        End If
        If Item(itemNum).Add_Stat(Stats.Endurance) > 0 Then
            If Item(itemNum).Stat_Percent(Stats.Endurance) > 0 Then
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Endurance) & "% End"
            Else
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Endurance) & " End"
            End If
        End If
        If Item(itemNum).Add_Stat(Stats.Intelligence) > 0 Then
            If Item(itemNum).Stat_Percent(Stats.Intelligence) > 0 Then
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Intelligence) & "% Int"
            Else
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Intelligence) & " Int"
            End If
        End If
        If Item(itemNum).Add_Stat(Stats.Agility) > 0 Then
            If Item(itemNum).Stat_Percent(Stats.Agility) > 0 Then
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Agility) & "% Agi"
            Else
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Agility) & " Agi"
            End If
        End If
        If Item(itemNum).Add_Stat(Stats.Willpower) > 0 Then
            If Item(itemNum).Stat_Percent(Stats.Willpower) > 0 Then
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Willpower) & "% Will"
            Else
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Willpower) & " Will"
            End If
        End If
    Case ITEM_TYPE_CONSUME
        If Item(itemNum).CastSpell > 0 Then
            AddDescInfo "Casts Spell"
        End If
        If Item(itemNum).AddHP > 0 Then
            AddDescInfo "+" & Item(itemNum).AddHP & " HP"
        End If
        If Item(itemNum).AddMP > 0 Then
            AddDescInfo "+" & Item(itemNum).AddMP & " SP"
        End If
        If Item(itemNum).AddEXP > 0 Then
            AddDescInfo "+" & Item(itemNum).AddEXP & " EXP"
        End If
        ' price
        AddDescInfo "Value: " & Item(itemNum).price & "g"
    Case ITEM_TYPE_SPELL
        ' price
        AddDescInfo "Value: " & Item(itemNum).price & "g"
    Case ITEM_TYPE_FOOD
        If Item(itemNum).HPorSP = 2 Then
            AddDescInfo "Heal: " & (Item(itemNum).FoodPerTick * Item(itemNum).FoodTickCount) & " SP"
        Else
            AddDescInfo "Heal: " & (Item(itemNum).FoodPerTick * Item(itemNum).FoodTickCount) & " HP"
        End If
        ' time
        AddDescInfo "Time: " & (Item(itemNum).FoodInterval * (Item(itemNum).FoodTickCount / 1000)) & "s"
        ' price
        AddDescInfo "Value: " & Item(itemNum).price & "g"
    Case ITEM_TYPE_PROTECTDROP
        ' price
        AddDescInfo "Value: " & Item(itemNum).price & "g"
    End Select
End Sub

Public Function GetItemNameColour(ByVal Rarity As Byte) As Long
    Select Case Rarity
    Case 0    ' white
        GetItemNameColour = White
    Case 1    ' green
        GetItemNameColour = Green
    Case 2    ' blue
        GetItemNameColour = BrightBlue
    Case 3    ' maroon
        GetItemNameColour = Red
    Case 4    ' purple
        GetItemNameColour = Pink
    Case 5    ' orange
        GetItemNameColour = Brown
    End Select
End Function

Public Sub AddDescInfo(Text As String, Optional Colour As Long = White)
    Dim Count As Long
    Count = UBound(descText)
    ReDim Preserve descText(1 To Count + 1) As TextColourRec
    descText(Count + 1).Text = Text
    descText(Count + 1).Colour = Colour
End Sub

Public Sub SwitchHotbar(OldSlot As Long, NewSlot As Long)
    Dim oldSlot_type As Long, oldSlot_value As Long, newSlot_type As Long, newSlot_value As Long

    oldSlot_type = Hotbar(OldSlot).sType
    newSlot_type = Hotbar(NewSlot).sType
    oldSlot_value = Hotbar(OldSlot).Slot
    newSlot_value = Hotbar(NewSlot).Slot

    ' send the changes
    SendHotbarChange 3, OldSlot, NewSlot
    Call PlayerSwitchHotbarSlots(OldSlot, NewSlot)
End Sub

Public Sub ShowChat()
    ShowWindow GetWindowIndex("winChat"), , False
    HideWindow GetWindowIndex("winChatSmall")
    ' Set the active control
    activeWindow = GetWindowIndex("winChat")
    SetActiveControl GetWindowIndex("winChat"), GetControlIndex("winChat", "txtChat")
    inSmallChat = False
    ChatScroll = 0
End Sub

Public Sub HideChat()
    ShowWindow GetWindowIndex("winChatSmall"), , False
    HideWindow GetWindowIndex("winChat")
    inSmallChat = True
    ChatScroll = 0
End Sub

Public Sub SetChatHeight(Height As Long)
    actChatHeight = Height
End Sub

Public Sub SetChatWidth(Width As Long)
    actChatWidth = Width
End Sub

Public Sub UpdateChat()
    SaveOptions
End Sub

Sub OpenShop(shopNum As Long)
' set globals
    InShop = shopNum
    shopSelectedSlot = 1
    shopSelectedItem = Shop(InShop).TradeItem(1).Item
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "chkSelling")).Value = 0
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "chkBuying")).Value = 1
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "btnSell")).visible = False
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "btnBuy")).visible = True
    shopIsSelling = False
    ' set the current item
    UpdateShop
    ' show the window
    ShowWindow GetWindowIndex("winShop")
End Sub

Sub CloseShop()
    SendCloseShop
    HideWindow GetWindowIndex("winShop")
    shopSelectedSlot = 0
    shopSelectedItem = 0
    shopIsSelling = False
    InShop = 0
End Sub

Sub UpdateShop()
    Dim i As Long, CostValue As Long

    If InShop = 0 Then Exit Sub

    ' make sure we have an item selected
    If shopSelectedSlot = 0 Then shopSelectedSlot = 1

    With Windows(GetWindowIndex("winShop"))
        ' buying items
        If Not shopIsSelling Then
            shopSelectedItem = Shop(InShop).TradeItem(shopSelectedSlot).Item
            ' labels
            If shopSelectedItem > 0 Then
                .Controls(GetControlIndex("winShop", "lblName")).Text = Trim$(Item(shopSelectedItem).Name)
                ' check if it's gold
                If Shop(InShop).TradeItem(shopSelectedSlot).CostItem = 1 Then
                    ' it's gold
                    .Controls(GetControlIndex("winShop", "lblCost")).Text = Shop(InShop).TradeItem(shopSelectedSlot).CostValue & "g"
                Else
                    ' if it's one then just print the name
                    If Shop(InShop).TradeItem(shopSelectedSlot).CostValue = 1 Then
                        .Controls(GetControlIndex("winShop", "lblCost")).Text = Trim$(Item(Shop(InShop).TradeItem(shopSelectedSlot).CostItem).Name)
                    Else
                        If Shop(InShop).TradeItem(shopSelectedSlot).CostItem > 0 Then
                            .Controls(GetControlIndex("winShop", "lblCost")).Text = Shop(InShop).TradeItem(shopSelectedSlot).CostValue & " " & Trim$(Item(Shop(InShop).TradeItem(shopSelectedSlot).CostItem).Name)
                        Else
                            .Controls(GetControlIndex("winShop", "lblCost")).Text = Shop(InShop).TradeItem(shopSelectedSlot).CostValue & " $"
                        End If

                    End If
                End If
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = Tex_Item(Item(shopSelectedItem).Pic)
                Next
            Else
                .Controls(GetControlIndex("winShop", "lblName")).Text = "Empty Slot"
                .Controls(GetControlIndex("winShop", "lblCost")).Text = vbNullString
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = 0
                Next
            End If
        Else
            shopSelectedItem = GetPlayerInvItemNum(MyIndex, shopSelectedSlot)
            ' labels
            If shopSelectedItem > 0 Then
                .Controls(GetControlIndex("winShop", "lblName")).Text = Trim$(Item(shopSelectedItem).Name)
                ' calc cost
                CostValue = (Item(shopSelectedItem).price / 100) * Shop(InShop).BuyRate
                .Controls(GetControlIndex("winShop", "lblCost")).Text = CostValue & "g"
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = Tex_Item(Item(shopSelectedItem).Pic)
                Next
            Else
                .Controls(GetControlIndex("winShop", "lblName")).Text = "Empty Slot"
                .Controls(GetControlIndex("winShop", "lblCost")).Text = vbNullString
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = 0
                Next
            End If
        End If
    End With
End Sub

Public Function IsShopSlot(startX As Long, startY As Long) As Long
    Dim tempRec As RECT
    Dim i As Long

    For i = 1 To MAX_TRADES
        With tempRec
            .top = startY + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            .bottom = .top + PIC_Y
            .Left = startX + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            .Right = .Left + PIC_X
        End With

        If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
            If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                IsShopSlot = i
                Exit Function
            End If
        End If
    Next
End Function

Sub ShowPlayerMenu(Index As Long, X As Long, Y As Long)
    PlayerMenuIndex = Index
    If PlayerMenuIndex = 0 Then Exit Sub
    Windows(GetWindowIndex("winPlayerMenu")).Window.Left = X - 5
    Windows(GetWindowIndex("winPlayerMenu")).Window.top = Y - 5
    Windows(GetWindowIndex("winPlayerMenu")).Controls(GetControlIndex("winPlayerMenu", "btnName")).Text = Trim$(GetPlayerName(PlayerMenuIndex))
    ShowWindow GetWindowIndex("winRightClickBG")
    ShowWindow GetWindowIndex("winPlayerMenu"), , False
End Sub

Public Function AryCount(ByRef Ary() As Byte) As Long
    On Error Resume Next

    AryCount = UBound(Ary) + 1
End Function

Public Function ByteToInt(ByVal B1 As Long, ByVal B2 As Long) As Long
    ByteToInt = B1 * 256 + B2
End Function

Sub UpdateStats_UI()
' set the bar labels
    With Windows(GetWindowIndex("winBars"))
        .Controls(GetControlIndex("winBars", "lblHP")).Text = GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP)
        .Controls(GetControlIndex("winBars", "lblMP")).Text = GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP)
        .Controls(GetControlIndex("winBars", "lblEXP")).Text = GetPlayerExp(MyIndex) & "/" & TNL
    End With
    ' update character screen
    With Windows(GetWindowIndex("winCharacter"))
        .Controls(GetControlIndex("winCharacter", "lblHealth")).Text = "Health: " & GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP)
        .Controls(GetControlIndex("winCharacter", "lblSpirit")).Text = "Spirit: " & GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP)
        .Controls(GetControlIndex("winCharacter", "lblExperience")).Text = "Exp: " & Player(MyIndex).EXP & "/" & TNL
    End With
End Sub

Sub UpdatePartyInterface()
    Dim i As Long, image(0 To 5) As Long, X As Long, pIndex As Long, Height As Long, cIn As Long

    ' unload it if we're not in a party
    If Party.Leader = 0 Then
        HideWindow GetWindowIndex("winParty")
        Exit Sub
    End If

    ' load the window
    ShowWindow GetWindowIndex("winParty")
    ' fill the controls
    With Windows(GetWindowIndex("winParty"))
        ' clear controls first
        For i = 1 To 3
            .Controls(GetControlIndex("winParty", "lblName" & i)).Text = vbNullString
            .Controls(GetControlIndex("winParty", "picEmptyBar_HP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picEmptyBar_SP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picBar_HP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picBar_SP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picShadow" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picChar" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picChar" & i)).Value = 0
        Next
        ' labels
        cIn = 1
        For i = 1 To Party.MemberCount
            ' cache the index
            pIndex = Party.Member(i)
            If pIndex > 0 Then
                If pIndex <> MyIndex Then
                    If IsPlaying(pIndex) Then
                        ' name and level
                        .Controls(GetControlIndex("winParty", "lblName" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "lblName" & cIn)).Text = Trim$(GetPlayerName(pIndex)) & " - " & GetPlayerLevel(pIndex)
                        ' picture
                        .Controls(GetControlIndex("winParty", "picShadow" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picChar" & cIn)).visible = True
                        ' store the player's index as a value for later use
                        .Controls(GetControlIndex("winParty", "picChar" & cIn)).Value = pIndex
                        For X = 0 To 5
                            .Controls(GetControlIndex("winParty", "picChar" & cIn)).image(X) = Tex_Char(GetPlayerSprite(pIndex))
                        Next
                        ' bars
                        .Controls(GetControlIndex("winParty", "picEmptyBar_HP" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picEmptyBar_SP" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picBar_HP" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picBar_SP" & cIn)).visible = True
                        ' increment control usage
                        cIn = cIn + 1
                    End If
                End If
            End If
        Next
        ' update the bars
        UpdatePartyBars
        ' set the window size
        Select Case Party.MemberCount
        Case 2: Height = 78
        Case 3: Height = 118
        Case 4: Height = 158
        End Select
        .Window.Height = Height
    End With
End Sub

Sub UpdatePartyBars()
    Dim i As Long, pIndex As Long, barWidth As Long, Width As Long

    ' unload it if we're not in a party
    If Party.Leader = 0 Then
        Exit Sub
    End If

    ' max bar width
    barWidth = 173

    ' make sure we're in a party
    With Windows(GetWindowIndex("winParty"))
        For i = 1 To 3
            ' get the pIndex from the control
            If .Controls(GetControlIndex("winParty", "picChar" & i)).visible = True Then
                pIndex = .Controls(GetControlIndex("winParty", "picChar" & i)).Value
                ' make sure they exist
                If pIndex > 0 Then
                    If IsPlaying(pIndex) Then
                        ' get playername and level atualization
                        If .Controls(GetControlIndex("winParty", "lblName" & i)).Text <> Trim$(GetPlayerName(pIndex)) & " - " & GetPlayerLevel(pIndex) Then
                            .Controls(GetControlIndex("winParty", "lblName" & i)).Text = Trim$(GetPlayerName(pIndex)) & " - " & GetPlayerLevel(pIndex)
                        End If
                        ' get their health
                        If GetPlayerVital(pIndex, HP) > 0 And GetPlayerMaxVital(pIndex, HP) > 0 Then
                            Width = ((GetPlayerVital(pIndex, Vitals.HP) / barWidth) / (GetPlayerMaxVital(pIndex, Vitals.HP) / barWidth)) * barWidth
                            .Controls(GetControlIndex("winParty", "picBar_HP" & i)).Width = Width
                        Else
                            .Controls(GetControlIndex("winParty", "picBar_HP" & i)).Width = 0
                        End If
                        ' get their spirit
                        If GetPlayerVital(pIndex, MP) > 0 And GetPlayerMaxVital(pIndex, MP) > 0 Then
                            Width = ((GetPlayerVital(pIndex, Vitals.MP) / barWidth) / (GetPlayerMaxVital(pIndex, Vitals.MP) / barWidth)) * barWidth
                            .Controls(GetControlIndex("winParty", "picBar_SP" & i)).Width = Width
                        Else
                            .Controls(GetControlIndex("winParty", "picBar_SP" & i)).Width = 0
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub

Sub ShowTrade()
' show the window
    ShowWindow GetWindowIndex("winTrade")
    ' set the controls up
    With Windows(GetWindowIndex("winTrade"))
        .Window.Text = "Trading with " & Trim$(GetPlayerName(InTrade))
        .Controls(GetControlIndex("winTrade", "lblYourTrade")).Text = Trim$(GetPlayerName(MyIndex)) & "'s Offer"
        .Controls(GetControlIndex("winTrade", "lblTheirTrade")).Text = Trim$(GetPlayerName(InTrade)) & "'s Offer"
        .Controls(GetControlIndex("winTrade", "lblYourValue")).Text = "0 $"
        .Controls(GetControlIndex("winTrade", "lblTheirValue")).Text = "0 $"
        .Controls(GetControlIndex("winTrade", "lblStatus")).Text = "Choose items to offer."
    End With
End Sub

Sub CheckResolution()
    Dim Resolution As Byte, Width As Long, Height As Long
    ' find the selected resolution
    Resolution = Options.Resolution
    ' reset
    If Resolution = 0 Then
        Resolution = 12
        ' loop through till we find one which fits
        Do Until ScreenFit(Resolution) Or Resolution > RES_COUNT
            ScreenFit Resolution
            Resolution = Resolution + 1
            If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
        Loop
        ' right resolution
        If Resolution > RES_COUNT Then Resolution = RES_COUNT
        Options.Resolution = Resolution
    End If

    ' size the window
    GetResolutionSize Options.Resolution, Width, Height
    Resize Width, Height

    ' save it
    curResolution = Options.Resolution

    SaveOptions
End Sub

Function ScreenFit(Resolution As Byte) As Boolean
    Dim sWidth As Long, sHeight As Long, Width As Long, Height As Long

    ' exit out early
    If Resolution = 0 Then
        ScreenFit = False
        Exit Function
    End If

    ' get screen size
    sWidth = Screen.Width / Screen.TwipsPerPixelX
    sHeight = Screen.Height / Screen.TwipsPerPixelY

    GetResolutionSize Resolution, Width, Height

    ' check if match
    If Width > sWidth Or Height > sHeight Then
        ScreenFit = False
    Else
        ScreenFit = True
    End If
End Function

Function GetResolutionSize(Resolution As Byte, ByRef Width As Long, ByRef Height As Long)
    Select Case Resolution
    Case 1
        Width = 1920
        Height = 1080
    Case 2
        Width = 1680
        Height = 1050
    Case 3
        Width = 1600
        Height = 900
    Case 4
        Width = 1440
        Height = 900
    Case 5
        Width = 1440
        Height = 1050
    Case 6
        Width = 1366
        Height = 768
    Case 7
        Width = 1360
        Height = 1024
    Case 8
        Width = 1360
        Height = 768
    Case 9
        Width = 1280
        Height = 1024
    Case 10
        Width = 1280
        Height = 800
    Case 11
        Width = 1280
        Height = 768
    Case 12
        Width = 1280
        Height = 720
    Case 13
        Width = 1024
        Height = 768
    Case 14
        Width = 1024
        Height = 576
    Case 15
        Width = 800
        Height = 600
    Case 16
        Width = 800
        Height = 450
    End Select
End Function

Sub Resize(ByVal Width As Long, ByVal Height As Long)
    frmMain.Width = (frmMain.Width \ 15 - frmMain.ScaleWidth + Width) * 15
    frmMain.Height = (frmMain.Height \ 15 - frmMain.ScaleHeight + Height) * 15
    frmMain.Left = (Screen.Width - frmMain.Width) \ 2
    frmMain.top = (Screen.Height - frmMain.Height) \ 2
    If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
End Sub

Sub ResizeGUI()
    Dim top As Long

    ' move hotbar
    Windows(GetWindowIndex("winHotbar")).Window.Left = ScreenWidth - 430
    ' move chat
    Windows(GetWindowIndex("winChat")).Window.top = ScreenHeight - 178
    Windows(GetWindowIndex("winChatSmall")).Window.top = ScreenHeight - 162
    ' move menu
    Windows(GetWindowIndex("winMenu")).Window.Left = ScreenWidth - 236
    Windows(GetWindowIndex("winMenu")).Window.top = ScreenHeight - 37
    ' move invitations
    Windows(GetWindowIndex("winInvite_Party")).Window.Left = ScreenWidth - 234
    Windows(GetWindowIndex("winInvite_Party")).Window.top = ScreenHeight - 80
    ' loop through
    top = ScreenHeight - 80
    If Windows(GetWindowIndex("winInvite_Party")).Window.visible Then
        top = top - 37
    End If
    Windows(GetWindowIndex("winInvite_Trade")).Window.Left = ScreenWidth - 234
    Windows(GetWindowIndex("winInvite_Trade")).Window.top = top
    ' re-size right-click background
    Windows(GetWindowIndex("winRightClickBG")).Window.Width = ScreenWidth
    Windows(GetWindowIndex("winRightClickBG")).Window.Height = ScreenHeight
    ' re-size black background
    Windows(GetWindowIndex("winBlank")).Window.Width = ScreenWidth
    Windows(GetWindowIndex("winBlank")).Window.Height = ScreenHeight
    ' re-size combo background
    Windows(GetWindowIndex("winComboMenuBG")).Window.Width = ScreenWidth
    Windows(GetWindowIndex("winComboMenuBG")).Window.Height = ScreenHeight
    ' centralise windows
    CentraliseWindow GetWindowIndex("winLogin")
    CentraliseWindow GetWindowIndex("winLoading")
    CentraliseWindow GetWindowIndex("winDialogue")
    CentraliseWindow GetWindowIndex("winClasses")
    CentraliseWindow GetWindowIndex("winNewChar")
    CentraliseWindow GetWindowIndex("winEscMenu")
    CentraliseWindow GetWindowIndex("winInventory")
    CentraliseWindow GetWindowIndex("winCharacter")
    CentraliseWindow GetWindowIndex("winSkills")
    CentraliseWindow GetWindowIndex("winOptions")
    CentraliseWindow GetWindowIndex("winChangeControls")
    CentraliseWindow GetWindowIndex("winShop")
    CentraliseWindow GetWindowIndex("winNpcChat")
    CentraliseWindow GetWindowIndex("winTrade")
    CentraliseWindow GetWindowIndex("winGuild")
    CentraliseWindow GetWindowIndex("winGuildMaker")
    CentraliseWindow GetWindowIndex("winGuildMenu")
    CentraliseWindow GetWindowIndex("winBank")
    CentraliseWindow GetWindowIndex("winReconnect")
    CentraliseWindow GetWindowIndex("winSerial")
    CentraliseWindow GetWindowIndex("winQuest")
    CentraliseWindow GetWindowIndex("winMessage")
    CentraliseWindow GetWindowIndex("winCheckIn")
    CentraliseWindow GetWindowIndex("winLottery")
End Sub

Sub SetResolution()
    Dim Width As Long, Height As Long
    curResolution = Options.Resolution
    GetResolutionSize curResolution, Width, Height
    Resize Width, Height
    ScreenWidth = Width
    ScreenHeight = Height
    TileWidth = (Width / 32) - 1
    TileHeight = (Height / 32) - 1
    ScreenX = (TileWidth) * PIC_X
    ScreenY = (TileHeight) * PIC_Y
    ResetGFX
    ResizeGUI
End Sub

Sub ShowComboMenu(curWindow As Long, curControl As Long)
    Dim top As Long
    With Windows(curWindow).Controls(curControl)
        ' linked to
        Windows(GetWindowIndex("winComboMenu")).Window.linkedToWin = curWindow
        Windows(GetWindowIndex("winComboMenu")).Window.linkedToCon = curControl
        ' set the size
        Windows(GetWindowIndex("winComboMenu")).Window.Height = 2 + (UBound(.list) * 16)
        Windows(GetWindowIndex("winComboMenu")).Window.Left = Windows(curWindow).Window.Left + .Left + 2
        top = Windows(curWindow).Window.top + .top + .Height
        If top + Windows(GetWindowIndex("winComboMenu")).Window.Height > ScreenHeight Then top = ScreenHeight - Windows(GetWindowIndex("winComboMenu")).Window.Height
        Windows(GetWindowIndex("winComboMenu")).Window.top = top
        Windows(GetWindowIndex("winComboMenu")).Window.Width = .Width - 4
        ' set the values
        Windows(GetWindowIndex("winComboMenu")).Window.list() = .list()
        Windows(GetWindowIndex("winComboMenu")).Window.Value = .Value
        Windows(GetWindowIndex("winComboMenu")).Window.group = 0
        ' load the menu
        ShowWindow GetWindowIndex("winComboMenuBG"), True, False
        ShowWindow GetWindowIndex("winComboMenu"), True, False
    End With
End Sub

Sub ComboMenu_MouseMove(curWindow As Long)
    Dim Y As Long, i As Long
    With Windows(curWindow).Window
        Y = currMouseY - .top
        ' find the option we're hovering over
        If UBound(.list) > 0 Then
            For i = 1 To UBound(.list)
                If Y >= (16 * (i - 1)) And Y <= (16 * (i)) Then
                    .group = i
                End If
            Next
        End If
    End With
End Sub

Sub ComboMenu_MouseDown(curWindow As Long)
    Dim Y As Long, i As Long
    With Windows(curWindow).Window
        Y = currMouseY - .top
        ' find the option we're hovering over
        If UBound(.list) > 0 Then
            For i = 1 To UBound(.list)
                If Y >= (16 * (i - 1)) And Y <= (16 * (i)) Then
                    Windows(.linkedToWin).Controls(.linkedToCon).Value = i
                    CloseComboMenu
                End If
            Next
        End If
    End With
End Sub

Sub SetOptionsScreen()
' clear the combolists
    Erase Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRes")).list
    ReDim Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRes")).list(0)
    Erase Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRender")).list
    ReDim Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRender")).list(0)

    ' Resolutions
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1920x1080"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1680x1050"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1600x900"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1440x900"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1440x1050"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1366x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1360x1024"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1360x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x1024"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x800"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x720"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1024x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1024x576"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "800x600"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "800x450"

    ' Render Options
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Automatic"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Hardware"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Mixed"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Software"

    ' fill the options screen
    With Windows(GetWindowIndex("winOptions"))
        .Controls(GetControlIndex("winOptions", "chkMusic")).Value = Options.Music
        .Controls(GetControlIndex("winOptions", "chkSound")).Value = Options.sound
        If Options.NoAuto = 1 Then
            .Controls(GetControlIndex("winOptions", "chkAutotiles")).Value = 0
        Else
            .Controls(GetControlIndex("winOptions", "chkAutotiles")).Value = 1
        End If
        .Controls(GetControlIndex("winOptions", "chkFullscreen")).Value = Options.Fullscreen
        .Controls(GetControlIndex("winOptions", "cmbRes")).Value = Options.Resolution
        .Controls(GetControlIndex("winOptions", "cmbRender")).Value = Options.Render + 1
        .Controls(GetControlIndex("winOptions", "chkReconnect")).Value = Options.Reconnect
        .Controls(GetControlIndex("winOptions", "chkItemName")).Value = Options.ItemName
        .Controls(GetControlIndex("winOptions", "chkItemAnimation")).Value = Options.ItemAnimation
        .Controls(GetControlIndex("winOptions", "chkFPSConection")).Value = Options.FPSConection
    End With
End Sub

Sub EventLogic()
    Dim Target As Long
    ' carry out the command
    With Map.TileData.Events(eventNum).EventPage(eventPageNum)
        Select Case .Commands(eventCommandNum).Type
        Case EventType.evAddText
            AddText .Commands(eventCommandNum).Text, .Commands(eventCommandNum).Colour, , .Commands(eventCommandNum).channel
        Case EventType.evShowChatBubble
            If .Commands(eventCommandNum).TargetType = TARGET_TYPE_PLAYER Then Target = MyIndex Else Target = .Commands(eventCommandNum).Target
            AddChatBubble Target, .Commands(eventCommandNum).TargetType, .Commands(eventCommandNum).Text, .Commands(eventCommandNum).Colour
        Case EventType.evPlayerVar
            If .Commands(eventCommandNum).Target > 0 Then Player(MyIndex).Variable(.Commands(eventCommandNum).Target) = .Commands(eventCommandNum).Colour
        End Select
        ' increment commands
        If eventCommandNum < .CommandCount Then
            eventCommandNum = eventCommandNum + 1
            Exit Sub
        End If
    End With
    ' we're done - close event
    eventNum = 0
    eventPageNum = 0
    eventCommandNum = 0
    inEvent = False
End Sub

Function HasItem(ByVal itemNum As Long) As Long
    Dim i As Long

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(MyIndex, i) = itemNum Then
            If Item(itemNum).Stackable > 0 Then
                HasItem = GetPlayerInvItemValue(MyIndex, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next
End Function

Function ActiveEventPage(ByVal eventNum As Long) As Long
    Dim X As Long, process As Boolean
    For X = Map.TileData.Events(eventNum).pageCount To 1 Step -1
        ' check if we match
        With Map.TileData.Events(eventNum).EventPage(X)
            process = True
            ' player var check
            If .chkPlayerVar Then
                If .PlayerVarNum > 0 Then
                    If Player(MyIndex).Variable(.PlayerVarNum) < .PlayerVariable Then
                        process = False
                    End If
                End If
            End If
            ' has item check
            If .chkHasItem Then
                If .HasItemNum > 0 Then
                    If HasItem(.HasItemNum) = 0 Then
                        process = False
                    End If
                End If
            End If
            ' this page
            If process = True Then
                ActiveEventPage = X
                Exit Function
            End If
        End With
    Next
End Function

Sub PlayerSwitchInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim OldNum As Long, OldValue As Long, oldBound As Byte
    Dim NewNum As Long, NewValue As Long, newBound As Byte

    If OldSlot = 0 Or NewSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(MyIndex, OldSlot)
    OldValue = GetPlayerInvItemValue(MyIndex, OldSlot)
    oldBound = PlayerInv(OldSlot).bound
    NewNum = GetPlayerInvItemNum(MyIndex, NewSlot)
    NewValue = GetPlayerInvItemValue(MyIndex, NewSlot)
    newBound = PlayerInv(NewSlot).bound

    SetPlayerInvItemNum MyIndex, NewSlot, OldNum
    SetPlayerInvItemValue MyIndex, NewSlot, OldValue
    PlayerInv(NewSlot).bound = oldBound

    SetPlayerInvItemNum MyIndex, OldSlot, NewNum
    SetPlayerInvItemValue MyIndex, OldSlot, NewValue
    PlayerInv(OldSlot).bound = newBound
End Sub

Sub PlayerSwitchSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim OldNum As Long, NewNum As Long, OldUses As Long, NewUses As Long

    If OldSlot = 0 Or NewSlot = 0 Then
        Exit Sub
    End If

    OldNum = PlayerSpells(OldSlot).Spell
    NewNum = PlayerSpells(NewSlot).Spell
    OldUses = PlayerSpells(OldSlot).Uses
    NewUses = PlayerSpells(NewSlot).Uses

    PlayerSpells(OldSlot).Spell = NewNum
    PlayerSpells(OldSlot).Uses = NewUses
    PlayerSpells(NewSlot).Spell = OldNum
    PlayerSpells(NewSlot).Uses = OldUses
End Sub

Sub CheckAppearTiles()
    Dim X As Long, Y As Long, i As Long
    If GettingMap Then Exit Sub

    ' clear
    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            If Map.TileData.Tile(X, Y).Type = TILE_TYPE_APPEAR Then
                TempTile(X, Y).DoorOpen = 0
            End If
        Next
    Next

    ' set
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                X = GetPlayerX(i)
                Y = GetPlayerY(i)
                CheckAppearTile X, Y
                If Y - 1 >= 0 Then CheckAppearTile X, Y - 1
                If Y + 1 <= Map.MapData.MaxY Then CheckAppearTile X, Y + 1
                If X - 1 >= 0 Then CheckAppearTile X - 1, Y
                If X + 1 <= Map.MapData.MaxX Then CheckAppearTile X + 1, Y
            End If
        End If
    Next
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).num > 0 Then
            If MapNpc(i).Vital(Vitals.HP) > 0 Then
                X = MapNpc(i).X
                Y = MapNpc(i).Y
                CheckAppearTile X, Y
                If Y - 1 >= 0 Then CheckAppearTile X, Y - 1
                If Y + 1 <= Map.MapData.MaxY Then CheckAppearTile X, Y + 1
                If X - 1 >= 0 Then CheckAppearTile X - 1, Y
                If X + 1 <= Map.MapData.MaxX Then CheckAppearTile X + 1, Y
            End If
        End If
    Next

    ' fade out old
    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            If TempTile(X, Y).DoorOpen = 0 Then
                ' exit if our mother is a bottom
                If Y > 0 Then
                    If Map.TileData.Tile(X, Y - 1).Data2 Then
                        If TempTile(X, Y - 1).DoorOpen = 1 Then GoTo continueLoop
                    End If
                End If
                ' not open - fade them out
                For i = 1 To MapLayer.Layer_Count - 1
                    If TempTile(X, Y).fadeAlpha(i) > 0 Then
                        TempTile(X, Y).isFading(i) = True
                        TempTile(X, Y).fadeAlpha(i) = TempTile(X, Y).fadeAlpha(i) - 1
                        TempTile(X, Y).FadeDir(i) = DIR_DOWN
                    End If
                Next
            End If
continueLoop:
        Next
    Next
End Sub

Sub CheckAppearTile(ByVal X As Long, ByVal Y As Long)
    If Y < 0 Or X < 0 Or Y > Map.MapData.MaxY Or X > Map.MapData.MaxX Then Exit Sub

    If Map.TileData.Tile(X, Y).Type = TILE_TYPE_APPEAR Then
        TempTile(X, Y).DoorOpen = 1

        If TempTile(X, Y).fadeAlpha(MapLayer.Mask) = 255 Then Exit Sub
        If TempTile(X, Y).isFading(MapLayer.Mask) Then
            If TempTile(X, Y).FadeDir(MapLayer.Mask) = DIR_DOWN Then
                TempTile(X, Y).FadeDir(MapLayer.Mask) = DIR_UP
                ' check if bottom
                If Y < Map.MapData.MaxY Then
                    If Map.TileData.Tile(X, Y).Data2 Then
                        TempTile(X, Y + 1).FadeDir(MapLayer.Ground) = DIR_UP
                    End If
                End If
                ' / bottom
            End If
            Exit Sub
        End If

        TempTile(X, Y).FadeDir(MapLayer.Mask) = DIR_UP
        TempTile(X, Y).isFading(MapLayer.Mask) = True
        TempTile(X, Y).fadeAlpha(MapLayer.Mask) = TempTile(X, Y).fadeAlpha(MapLayer.Mask) + 1

        ' check if bottom
        If Y < Map.MapData.MaxY Then
            If Map.TileData.Tile(X, Y).Data2 Then
                TempTile(X, Y + 1).FadeDir(MapLayer.Ground) = DIR_UP
                TempTile(X, Y + 1).isFading(MapLayer.Ground) = True
                TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) = TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) + 1
            End If
        End If
        ' / bottom
    End If
End Sub

Public Sub AppearTileFadeLogic()
    Dim X As Long, Y As Long
    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            If Map.TileData.Tile(X, Y).Type = TILE_TYPE_APPEAR Then
                ' check if it's fading
                If TempTile(X, Y).isFading(MapLayer.Mask) Then
                    ' fading in
                    If TempTile(X, Y).FadeDir(MapLayer.Mask) = DIR_UP Then
                        If TempTile(X, Y).fadeAlpha(MapLayer.Mask) < 255 Then
                            TempTile(X, Y).fadeAlpha(MapLayer.Mask) = TempTile(X, Y).fadeAlpha(MapLayer.Mask) + 20
                            ' check if bottom
                            If Y < Map.MapData.MaxY Then
                                If Map.TileData.Tile(X, Y).Data2 Then
                                    TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) = TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) + 20
                                End If
                            End If
                            ' / bottom
                        End If
                        If TempTile(X, Y).fadeAlpha(MapLayer.Mask) >= 255 Then
                            TempTile(X, Y).fadeAlpha(MapLayer.Mask) = 255
                            TempTile(X, Y).isFading(MapLayer.Mask) = False
                            ' check if bottom
                            If Y < Map.MapData.MaxY Then
                                If Map.TileData.Tile(X, Y).Data2 Then
                                    TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) = 255
                                    TempTile(X, Y + 1).isFading(MapLayer.Ground) = False
                                End If
                            End If
                            ' / bottom
                        End If
                    ElseIf TempTile(X, Y).FadeDir(MapLayer.Mask) = DIR_DOWN Then
                        If TempTile(X, Y).fadeAlpha(MapLayer.Mask) > 0 Then
                            TempTile(X, Y).fadeAlpha(MapLayer.Mask) = TempTile(X, Y).fadeAlpha(MapLayer.Mask) - 20
                            ' check if bottom
                            If Y < Map.MapData.MaxY Then
                                If Map.TileData.Tile(X, Y).Data2 Then
                                    TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) = TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) - 20
                                End If
                            End If
                            ' / bottom
                        End If
                        If TempTile(X, Y).fadeAlpha(MapLayer.Mask) <= 0 Then
                            TempTile(X, Y).fadeAlpha(MapLayer.Mask) = 0
                            TempTile(X, Y).isFading(MapLayer.Mask) = False
                            ' check if bottom
                            If Y < Map.MapData.MaxY Then
                                If Map.TileData.Tile(X, Y).Data2 Then
                                    TempTile(X, Y + 1).fadeAlpha(MapLayer.Ground) = 0
                                    TempTile(X, Y + 1).isFading(MapLayer.Ground) = False
                                End If
                            End If
                            ' / bottom
                        End If
                    End If
                End If
            End If
        Next
    Next
End Sub

Public Sub CreateBlood(ByVal X As Long, ByVal Y As Long)
    Dim i As Long, Sprite As Long

    BloodIndex = 0

    ' Randomize sprite
    Sprite = Rand(1, BloodCount)

    ' Make sure tile doesn't already have blood
    For i = 1 To Blood_HighIndex
        ' Already have blood
        If Blood(i).X = X And Blood(i).Y = Y Then
            ' Refresh the timer
            Blood(i).Timer = Tick
            Exit Sub
        End If
    Next

    ' Carry on with the set
    For i = 1 To MAX_BYTE
        If Blood(i).Timer = 0 Then
            BloodIndex = i
            Exit For
        End If
    Next

    If BloodIndex = 0 Then
        Call ClearBlood(1)
        BloodIndex = 1
    End If

    ' Set the blood up
    With Blood(BloodIndex)
        .X = X
        .Y = Y
        .Sprite = Sprite
        .Timer = getTime
        .Alpha = 255
    End With

    SetBloodHighIndex

End Sub

Public Sub ClearBlood(ByVal Index As Long, Optional ByVal SetHighIndex As Boolean = True)

    With Blood(Index)
        .X = 0
        .Y = 0
        .Sprite = 0
        .Timer = 0
        .Alpha = 0
    End With

    If SetHighIndex Then
        SetBloodHighIndex
    End If

End Sub

Public Sub SetBloodHighIndex()
    Dim i As Long

    Blood_HighIndex = 0

    ' Find the new high index
    For i = MAX_BYTE To 1 Step -1
        If Blood(i).Timer > 0 Then
            Blood_HighIndex = i
            Exit For
        End If
    Next

End Sub

Public Function KeepTwoDigit(num As Byte)
    If (num < 10) Then
        KeepTwoDigit = "0" & num
    Else
        KeepTwoDigit = num
    End If
End Function

Public Function IsDay() As Boolean
    If Map.MapData.DayNight = 2 Then
        IsDay = True
        Exit Function
    ElseIf Map.MapData.DayNight = 1 Then
        IsDay = False
        Exit Function
    ElseIf Map.MapData.DayNight = 0 Then
        If GameHours >= 7 And GameHours < 18 Then
            IsDay = True
        Else
            IsDay = False
        End If
        Exit Function
    End If
End Function

Public Sub ShowMessageWindow(ByRef WindowName As String, ByRef message As String)
    
    With Windows(GetWindowIndex("winMessage"))
        .Window.Text = WindowName
        .Controls(GetControlIndex("winMessage", "lblChat")).Text = message
    End With
    
    ShowWindow GetWindowIndex("winMessage")
End Sub

Public Sub SetPlayerGold(ByVal Index As Long, ByVal Value As Long)
    
    If Not IsPlaying(Index) Then Exit Sub
    
    Player(Index).Gold = Value
End Sub

Public Function GetPlayerGold(ByVal Index As Long) As Long
    If Not IsPlaying(Index) Then Exit Function
    
    GetPlayerGold = Player(Index).Gold
End Function

Public Function ConvertByteToBool(Variavel As Byte) As Boolean
    If Variavel = YES Then
        ConvertByteToBool = True
    Else
        ConvertByteToBool = False
    End If
End Function








