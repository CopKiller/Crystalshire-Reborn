Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, X As Long
    Dim TickCPS As Currency, CPS As Long, FrameTime As Currency
    Dim tmr25 As Currency, tmr500 As Currency, tmr1000 As Currency
    Dim LastUpdateSavePlayers As Currency, LastUpdateMapSpawnItems As Currency, LastUpdatePlayerVitals As Currency

    ServerOnline = True

    Do While ServerOnline
        Tick = getTime
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick

        If Tick > tmr25 Then
            ' loops
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If getTime > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell).Spell).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If getTime > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned GetPlayerMap(i), i, TARGET_TYPE_PLAYER
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < getTime Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    'Status do Player
                    CheckPlayerStatus i

                    ' HoT and DoT logic
                    For X = 1 To MAX_DOTS
                        HandleDoT_Player i, X
                        HandleHoT_Player i, X
                    Next
                    ' food processing
                    UpdatePlayerFood i
                    ' event logic
                    If TempPlayer(i).inEvent Then
                        If TempPlayer(i).pageNum > 0 Then
                            If TempPlayer(i).eventNum > 0 Then
                                If TempPlayer(i).commandNum > 0 Then
                                    EventLogic i
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            ' update entity logic
            UpdateMapEntities
            ' update label
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            tmr25 = getTime + 25
        End If

        ' Check conections to close for need auth login
        CheckConnectionTime

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS

                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next

            UpdateMapLogic
            tmr500 = getTime + 500
        End If

        If Tick > tmr1000 Then
            ' check if shutting down
            If isShuttingDown Then
                Call HandleShutdown
            End If
            ' disable login tokens
            For i = 1 To MAX_PLAYERS
                If LoginToken(i).Active Then
                    If LoginToken(i).TimeCreated + LoginTimer < getTime Then
                        ClearLoginToken i
                    End If
                End If

                ' retira o cooldown das spells
                If IsPlaying(i) Then
                    For X = 1 To MAX_PLAYER_SPELLS
                        If Player(i).Spell(X).Spell > 0 Then
                            If Player(i).SpellCD(X) > 0 Then
                                Player(i).SpellCD(X) = Player(i).SpellCD(X) - 1
                            End If
                        End If
                    Next X

                    ' Verificar se o jogador tem alguma task com tempo!
                    Call CheckPlayerTaskTimer(i)

                End If
            Next i

            ' GameSeconds = 0
            ' GameMinutes = 55
            ' GameHours = 6
            ' SendClientTime

            ' Are we using the time system?
            If Options.DAYNIGHT = YES Then
                ' Change the game time.
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
                    SendClientTime
                End If

                ' See if we need to switch to day or night.
                If DayTime = True Then
                    If GameHours >= 20 Or GameHours < 6 Then
                        DayTime = False
                        GlobalMsg "The Night has shrouded this land in darkness!", Yellow
                        SendClientTime
                    End If
                ElseIf DayTime = False Then
                    If GameHours >= 6 And GameHours < 20 Then
                        DayTime = True
                        GlobalMsg "Light shines forth with the new Day!", Yellow
                        SendClientTime
                    End If
                End If

                ' Update the label
                If DayTime = True Then
                    frmServer.lblGameTime.Caption = "(Day) " & KeepTwoDigit(GameHours) & ":" & KeepTwoDigit(GameMinutes)
                Else
                    frmServer.lblGameTime.Caption = "(Night) " & KeepTwoDigit(GameHours) & ":" & KeepTwoDigit(GameMinutes)
                End If

            End If
            
            ' Lottery
            CheckBetLoop

            ' reset timer
            tmr1000 = getTime + 1000
        End If

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = getTime + 5000
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            CheckPremiumLoop
            LastUpdateSavePlayers = getTime + 300000
        End If

        If Not CPSUnlock Then Sleep 1
        If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents

        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Sub UpdateMapEntities()
    Dim MapNum As Long, i As Long, x1 As Long, y1 As Long, X As Long, Y As Long, Resource_index As Long

    Tick = getTime
    For MapNum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(MapNum, i).Num > 0 Then
                If MapItem(MapNum, i).PlayerName <> vbNullString Then
                    ' make item public?
                    If Not MapItem(MapNum, i).Bound Then
                        If MapItem(MapNum, i).playerTimer < Tick Then
                            ' make it public
                            MapItem(MapNum, i).PlayerName = vbNullString
                            MapItem(MapNum, i).playerTimer = 0
                            ' send updates to everyone
                            SendMapItemsToAll MapNum
                        End If
                    End If
                    ' despawn item?
                    If MapItem(MapNum, i).canDespawn Then
                        If MapItem(MapNum, i).despawnTimer < Tick Then
                            ' despawn it
                            ClearMapItem i, MapNum
                            ' send updates to everyone
                            SendMapItemsToAll MapNum
                        End If
                    End If
                End If
            End If
        Next

        '  Close the doors
        If Tick > TempTile(MapNum).DoorTimer + 5000 Then
            For x1 = 0 To Map(MapNum).MapData.MaxX
                For y1 = 0 To Map(MapNum).MapData.MaxY
                    If Map(MapNum).TileData.Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(MapNum).DoorOpen(x1, y1) = YES Then
                        TempTile(MapNum).DoorOpen(x1, y1) = NO
                        SendMapKeyToMap MapNum, x1, y1, 0
                    End If
                Next
            Next
        End If

        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(MapNum).NPC(i).Num > 0 Then
                For X = 1 To MAX_DOTS
                    HandleDoT_Npc MapNum, i, X
                    HandleHoT_Npc MapNum, i, X
                Next
            End If
        Next

        ' Respawning Resources
        If MapResourceCache(MapNum).Resource_Count > 0 Then
            For i = 0 To MapResourceCache(MapNum).Resource_Count
                Resource_index = Map(MapNum).TileData.Tile(MapResourceCache(MapNum).ResourceData(i).X, MapResourceCache(MapNum).ResourceData(i).Y).Data1

                If Resource_index > 0 Then
                    If MapResourceCache(MapNum).ResourceData(i).ResourceState = 1 Or MapResourceCache(MapNum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If MapResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < Tick Then
                            MapResourceCache(MapNum).ResourceData(i).ResourceTimer = Tick
                            MapResourceCache(MapNum).ResourceData(i).ResourceState = 0    ' normal
                            ' re-set health to resource root
                            MapResourceCache(MapNum).ResourceData(i).cur_health = Resource(Resource_index).health
                            SendResourceCacheToMap MapNum, i
                        End If
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, X As Long, MapNum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Currency, Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long
    Dim target As Long, TargetType As Byte, DidWalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim TargetX As Long, TargetY As Long, target_verify As Boolean

    For MapNum = 1 To MAX_MAPS
        If PlayersOnMap(MapNum) = YES Then
            TickCount = getTime

            For X = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(MapNum).NPC(X).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).MapData.NPC(X) > 0 And MapNpc(MapNum).NPC(X).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then

                        ' make sure it's not stunned
                        If Not MapNpc(MapNum).NPC(X).StunDuration > 0 Then

                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum And MapNpc(MapNum).NPC(X).target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        ' make sure it's within the level range
                                        If (GetPlayerLevel(i) <= NPC(NpcNum).Level - 2) Or (Map(MapNum).MapData.Moral = MAP_MORAL_BOSS) Then
                                            n = NPC(NpcNum).Range
                                            DistanceX = MapNpc(MapNum).NPC(X).X - GetPlayerX(i)
                                            DistanceY = MapNpc(MapNum).NPC(X).Y - GetPlayerY(i)

                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                            If DistanceY < 0 Then DistanceY = DistanceY * -1

                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n And DistanceY <= n Then
                                                If NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    If Len(Trim$(NPC(NpcNum).AttackSay)) > 0 Then
                                                        Call PlayerMsg(i, Trim$(NPC(NpcNum).Name) & " says: " & Trim$(NPC(NpcNum).AttackSay), SayColor)
                                                    End If
                                                    MapNpc(MapNum).NPC(X).TargetType = 1    ' player
                                                    MapNpc(MapNum).NPC(X).target = i
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If

                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).MapData.NPC(X) > 0 And MapNpc(MapNum).NPC(X).Num > 0 Then
                    If MapNpc(MapNum).NPC(X).StunDuration > 0 Then
                        ' check if we can unstun them
                        If getTime > MapNpc(MapNum).NPC(X).StunTimer + (MapNpc(MapNum).NPC(X).StunDuration * 1000) Then
                            MapNpc(MapNum).NPC(X).StunDuration = 0
                            MapNpc(MapNum).NPC(X).StunTimer = 0

                            SendStunned MapNum, X, TARGET_TYPE_NPC
                        End If
                    Else
                        ' check if in conversation
                        If MapNpc(MapNum).NPC(X).c_inChatWith > 0 Then
                            ' check if we can stop having conversation
                            If Not TempPlayer(MapNpc(MapNum).NPC(X).c_inChatWith).inChatWith = NpcNum Then
                                MapNpc(MapNum).NPC(X).c_inChatWith = 0
                                MapNpc(MapNum).NPC(X).dir = MapNpc(MapNum).NPC(X).c_lastDir
                                NpcDir MapNum, X, MapNpc(MapNum).NPC(X).dir
                            End If
                        Else
                            target = MapNpc(MapNum).NPC(X).target
                            TargetType = MapNpc(MapNum).NPC(X).TargetType

                            ' Check to see if its time for the npc to walk
                            If NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then

                                If TargetType = 1 Then    ' player

                                    ' Check to see if we are following a player or not
                                    If target > 0 Then

                                        ' Check if the player is even playing, if so follow'm
                                        If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                            DidWalk = False
                                            target_verify = True
                                            TargetY = GetPlayerY(target)
                                            TargetX = GetPlayerX(target)
                                        Else
                                            MapNpc(MapNum).NPC(X).TargetType = 0    ' clear
                                            MapNpc(MapNum).NPC(X).target = 0
                                        End If
                                    End If

                                ElseIf TargetType = 2 Then    'npc

                                    If target > 0 Then

                                        If MapNpc(MapNum).NPC(target).Num > 0 Then
                                            DidWalk = False
                                            target_verify = True
                                            TargetY = MapNpc(MapNum).NPC(target).Y
                                            TargetX = MapNpc(MapNum).NPC(target).X
                                        Else
                                            MapNpc(MapNum).NPC(X).TargetType = 0    ' clear
                                            MapNpc(MapNum).NPC(X).target = 0
                                        End If
                                    End If
                                End If

                                If target_verify Then

                                    i = Int(Rnd * 5)

                                    ' Lets move the npc
                                    Select Case i
                                    Case 0

                                        ' Up
                                        If MapNpc(MapNum).NPC(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_UP) Then
                                                Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Down
                                        If MapNpc(MapNum).NPC(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Left
                                        If MapNpc(MapNum).NPC(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Right
                                        If MapNpc(MapNum).NPC(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                    Case 1

                                        ' Right
                                        If MapNpc(MapNum).NPC(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Left
                                        If MapNpc(MapNum).NPC(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Down
                                        If MapNpc(MapNum).NPC(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Up
                                        If MapNpc(MapNum).NPC(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_UP) Then
                                                Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                    Case 2

                                        ' Down
                                        If MapNpc(MapNum).NPC(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Up
                                        If MapNpc(MapNum).NPC(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_UP) Then
                                                Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Right
                                        If MapNpc(MapNum).NPC(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Left
                                        If MapNpc(MapNum).NPC(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                    Case 3

                                        ' Left
                                        If MapNpc(MapNum).NPC(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Right
                                        If MapNpc(MapNum).NPC(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Up
                                        If MapNpc(MapNum).NPC(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_UP) Then
                                                Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                        ' Down
                                        If MapNpc(MapNum).NPC(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If

                                    End Select

                                    ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If MapNpc(MapNum).NPC(X).X - 1 = TargetX And MapNpc(MapNum).NPC(X).Y = TargetY Then
                                            If MapNpc(MapNum).NPC(X).dir <> DIR_LEFT Then
                                                Call NpcDir(MapNum, X, DIR_LEFT)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapNpc(MapNum).NPC(X).X + 1 = TargetX And MapNpc(MapNum).NPC(X).Y = TargetY Then
                                            If MapNpc(MapNum).NPC(X).dir <> DIR_RIGHT Then
                                                Call NpcDir(MapNum, X, DIR_RIGHT)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapNpc(MapNum).NPC(X).X = TargetX And MapNpc(MapNum).NPC(X).Y - 1 = TargetY Then
                                            If MapNpc(MapNum).NPC(X).dir <> DIR_UP Then
                                                Call NpcDir(MapNum, X, DIR_UP)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapNpc(MapNum).NPC(X).X = TargetX And MapNpc(MapNum).NPC(X).Y + 1 = TargetY Then
                                            If MapNpc(MapNum).NPC(X).dir <> DIR_DOWN Then
                                                Call NpcDir(MapNum, X, DIR_DOWN)
                                            End If

                                            DidWalk = True
                                        End If

                                        ' We could not move so Target must be behind something, walk randomly.
                                        If Not DidWalk Then
                                            i = Int(Rnd * 2)

                                            If i = 1 Then
                                                i = Int(Rnd * 4)

                                                If CanNpcMove(MapNum, X, i) Then
                                                    Call NpcMove(MapNum, X, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If

                                Else
                                    i = Int(Rnd * 4)

                                    If i = 1 Then
                                        i = Int(Rnd * 4)

                                        If CanNpcMove(MapNum, X, i) Then
                                            Call NpcMove(MapNum, X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).MapData.NPC(X) > 0 And MapNpc(MapNum).NPC(X).Num > 0 Then
                    target = MapNpc(MapNum).NPC(X).target
                    TargetType = MapNpc(MapNum).NPC(X).TargetType

                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                        If TargetType = 1 Then    ' player
                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                ' melee combat
                                TryNpcAttackPlayer X, target
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(MapNum).NPC(X).target = 0
                                MapNpc(MapNum).NPC(X).TargetType = 0    ' clear
                            End If
                        End If
                    End If

                    ' check for spells
                    If MapNpc(MapNum).NPC(X).spellBuffer.Spell = 0 Then
                        ' loop through and try and cast our spells
                        For i = 1 To MAX_NPC_SPELLS
                            If NPC(NpcNum).Spell(i) > 0 Then
                                NpcBufferSpell MapNum, X, i
                            End If
                        Next
                    Else
                        ' check the timer
                        If MapNpc(MapNum).NPC(X).spellBuffer.Timer + (Spell(NPC(NpcNum).Spell(MapNpc(MapNum).NPC(X).spellBuffer.Spell)).CastTime * 1000) < getTime Then
                            ' cast the spell
                            NpcCastSpell MapNum, X, MapNpc(MapNum).NPC(X).spellBuffer.Spell, MapNpc(MapNum).NPC(X).spellBuffer.target, MapNpc(MapNum).NPC(X).spellBuffer.tType
                            ' clear the buffer
                            MapNpc(MapNum).NPC(X).spellBuffer.Spell = 0
                            MapNpc(MapNum).NPC(X).spellBuffer.target = 0
                            MapNpc(MapNum).NPC(X).spellBuffer.Timer = 0
                            MapNpc(MapNum).NPC(X).spellBuffer.tType = 0
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(MapNum).NPC(X).stopRegen Then
                    If MapNpc(MapNum).NPC(X).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(MapNum).NPC(X).Vital(Vitals.HP) > 0 Then
                            MapNpc(MapNum).NPC(X).Vital(Vitals.HP) = MapNpc(MapNum).NPC(X).Vital(Vitals.HP) + GetNpcVitalRegen(NpcNum, Vitals.HP)

                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(MapNum).NPC(X).Vital(Vitals.HP) > GetNpcMaxVital(NpcNum, Vitals.HP) Then
                                MapNpc(MapNum).NPC(X).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
                            End If
                        End If
                    End If
                End If

                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(MapNum).NPC(X).Num = 0 And Map(MapNum).MapData.NPC(X) > 0 Then

                    ' Spawn Variavel or not
                    If MapNpc(MapNum).NPC(X).SecondsToSpawn = 0 Then
                        If NPC(Map(MapNum).MapData.NPC(X)).RndSpawn = YES Then
                            MapNpc(MapNum).NPC(X).SecondsToSpawn = Rand(NPC(Map(MapNum).MapData.NPC(X)).SpawnSecsMin, NPC(Map(MapNum).MapData.NPC(X)).SpawnSecs) * 1000
                        Else
                            MapNpc(MapNum).NPC(X).SecondsToSpawn = NPC(Map(MapNum).MapData.NPC(X)).SpawnSecs * 1000
                        End If
                    End If

                    ' Envia um action msg onde o npc morreu, com o tempo que falta pra ele renascer!
                    If (MapNpc(MapNum).NPC(X).SecondsToSpawn / 1000) > 0 Then
                        If TickCount > MapNpc(MapNum).NPC(X).ActionMsgSpawn Then
                            If (((MapNpc(MapNum).NPC(X).SpawnWait + MapNpc(MapNum).NPC(X).SecondsToSpawn) - TickCount) / 1000) > 0 Then
                                Call SendActionMsg(MapNum, SecondsToHMS(((MapNpc(MapNum).NPC(X).SpawnWait + MapNpc(MapNum).NPC(X).SecondsToSpawn) - TickCount) / 1000), BrightRed, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(X).X * 32, MapNpc(MapNum).NPC(X).Y * 32)
                                MapNpc(MapNum).NPC(X).ActionMsgSpawn = TickCount + 1000
                            End If
                        End If
                    End If

                    If TickCount > MapNpc(MapNum).NPC(X).SpawnWait + MapNpc(MapNum).NPC(X).SecondsToSpawn Then
                        ' if it's a boss chamber then don't let them respawn
                        If Map(MapNum).MapData.Moral = MAP_MORAL_BOSS Then
                            ' make sure the boss is alive
                            If Map(MapNum).MapData.BossNpc > 0 Then
                                If Map(MapNum).MapData.NPC(Map(MapNum).MapData.BossNpc) > 0 Then
                                    If X <> Map(MapNum).MapData.BossNpc Then
                                        If MapNpc(MapNum).NPC(Map(MapNum).MapData.BossNpc).Num > 0 Then
                                            Call SpawnNpc(X, MapNum)
                                        End If
                                    Else
                                        SpawnNpc X, MapNum
                                    End If
                                End If
                            End If
                        Else
                            Call SpawnNpc(X, MapNum)
                        End If
                    End If
                End If

            Next

        End If

        If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If getTime > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = getTime
    End If

    ' Make sure we reset the timer for door closing
    If getTime > KeyTimer + 15000 Then
        KeyTimer = getTime
    End If

End Sub

Private Sub UpdatePlayerFood(ByVal i As Long)
    Dim vitalType As Long, colour As Long, X As Long

    For X = 1 To Vitals.Vital_Count - 1
        If TempPlayer(i).foodItem(X) > 0 Then
            ' make sure not in combat
            If Not TempPlayer(i).stopRegen Then
                ' timer ready?
                If TempPlayer(i).foodTimer(X) + Item(TempPlayer(i).foodItem(X)).FoodInterval < getTime Then
                    ' get vital type
                    If Item(TempPlayer(i).foodItem(X)).HPorSP = 2 Then vitalType = Vitals.MP Else vitalType = Vitals.HP
                    ' make sure we haven't gone over the top
                    If GetPlayerVital(i, vitalType) >= GetPlayerMaxVital(i, vitalType) Then
                        ' bring it back down to normal
                        SetPlayerVital i, vitalType, GetPlayerMaxVital(i, vitalType)
                        SendVital i, vitalType
                        Call SendMapVitals(i)
                        ' remove the food - no point healing when full
                        TempPlayer(i).foodItem(X) = 0
                        TempPlayer(i).foodTick(X) = 0
                        TempPlayer(i).foodTimer(X) = 0
                        Exit Sub
                    End If
                    ' give them the healing
                    SetPlayerVital i, vitalType, GetPlayerVital(i, vitalType) + Item(TempPlayer(i).foodItem(X)).FoodPerTick
                    ' let them know with messages
                    If vitalType = 2 Then colour = BrightBlue Else colour = Green
                    SendActionMsg GetPlayerMap(i), "+" & Item(TempPlayer(i).foodItem(X)).FoodPerTick, colour, ACTIONMSG_SCROLL, GetPlayerX(i) * 32, GetPlayerY(i) * 32
                    ' send vitals
                    SendVital i, vitalType
                    Call SendMapVitals(i)
                    ' increment tick count
                    TempPlayer(i).foodTick(X) = TempPlayer(i).foodTick(X) + 1
                    ' make sure we're not over max ticks
                    If TempPlayer(i).foodTick(X) >= Item(TempPlayer(i).foodItem(X)).FoodTickCount Then
                        ' clear food
                        TempPlayer(i).foodItem(X) = 0
                        TempPlayer(i).foodTick(X) = 0
                        TempPlayer(i).foodTimer(X) = 0
                        Exit Sub
                    End If
                    ' reset the timer
                    TempPlayer(i).foodTimer(X) = getTime
                End If
            Else
                ' remove the food effect
                TempPlayer(i).foodItem(X) = 0
                TempPlayer(i).foodTick(X) = 0
                TempPlayer(i).foodTimer(X) = 0
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub UpdatePlayerVitals()
    Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not TempPlayer(i).stopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, Vitals.HP)
                    Call SendMapVitals(i)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If

                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, Vitals.HP)
                    Call SendMapVitals(i)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextLoginAdd("Saving all online players...")

        For i = 1 To Player_HighIndex

            If IsPlaying(i) Then
                Call SavePlayer(i)
            End If

            If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub
