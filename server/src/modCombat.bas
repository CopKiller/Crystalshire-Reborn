Attribute VB_Name = "modSvCombat"
Option Explicit

' ###################################
' ##      Player Attacking NPC     ##
' ###################################
Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal mapnpcnum As Long)
    Dim blockAmount As Long
    Dim NpcNum As Long
    Dim MapNum As Long
    Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, mapnpcnum) Then

        MapNum = GetPlayerMap(Index)
        NpcNum = MapNpc(MapNum).NPC(mapnpcnum).Num

        ' check if NPC can avoid the attack
        If CanNpcDodge(NpcNum, Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (MapNpc(MapNum).NPC(mapnpcnum).X * 32), (MapNpc(MapNum).NPC(mapnpcnum).Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)

        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(NpcNum)

        If blockAmount > 0 Then
            Damage = Damage - blockAmount
            SendActionMsg MapNum, "Block " & blockAmount, Pink, 1, (MapNpc(MapNum).NPC(mapnpcnum).X * 32), (MapNpc(MapNum).NPC(mapnpcnum).Y * 32)
        End If

        ' take away armour
        'damage = damage - RAND(1, (Npc(NpcNum).Stat(Stats.Agility) * 2))
        Damage = Damage - Rand((GetNpcDefence(NpcNum) / 100) * 10, (GetNpcDefence(NpcNum) / 100) * 10)
        ' randomise from 1 to max hit
        Damage = Rand(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))

        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index, TARGET_TYPE_NPC, mapnpcnum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            
            ' Stuna o npc quando toma crítico
            StunNPCForTimer mapnpcnum, MapNum, 1
        End If

        If Damage > 0 Then
            Call PlayerAttackNpc(Index, mapnpcnum, Damage)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal mapnpcnum As Long, Optional ByVal isSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).NPC(mapnpcnum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).NPC(mapnpcnum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).NPC(mapnpcnum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then

        ' exit out early
        If isSpell Then
            If NpcNum > 0 Then
                If NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_GUILDMAKER And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_CLAIMSERIAL And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_BANK Then
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipmentNum(Attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipmentNum(Attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If NpcNum > 0 And getTime > TempPlayer(Attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                NpcX = MapNpc(MapNum).NPC(mapnpcnum).X
                NpcY = MapNpc(MapNum).NPC(mapnpcnum).Y + 1
            Case DIR_DOWN
                NpcX = MapNpc(MapNum).NPC(mapnpcnum).X
                NpcY = MapNpc(MapNum).NPC(mapnpcnum).Y - 1
            Case DIR_LEFT
                NpcX = MapNpc(MapNum).NPC(mapnpcnum).X + 1
                NpcY = MapNpc(MapNum).NPC(mapnpcnum).Y
            Case DIR_RIGHT
                NpcX = MapNpc(MapNum).NPC(mapnpcnum).X - 1
                NpcY = MapNpc(MapNum).NPC(mapnpcnum).Y
            End Select

            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then
                    If NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_GUILDMAKER And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_CLAIMSERIAL And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_BANK Then
                        CanPlayerAttackNpc = True
                    ElseIf NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                        ' init quest tasks
                        Call CheckTasks(Attacker, QUEST_TYPE_GOTALK, NpcNum)
                        Call CheckTasks(Attacker, QUEST_TYPE_GOGIVE, NpcNum)
                        Call CheckTasks(Attacker, QUEST_TYPE_GOGET, NpcNum)
                        ' init conversation if it's friendly
                        InitChat Attacker, MapNum, mapnpcnum
                    ElseIf NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_GUILDMAKER Then
                        SendGuildWindow Attacker
                    ElseIf NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_CLAIMSERIAL Then
                        SendSerialWindow Attacker
                    ElseIf NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_BANK Then
                        SendBank Attacker
                        TempPlayer(Attacker).InBank = True
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal mapnpcnum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim EXP As Long
    Dim n As Long
    Dim i As Long, R As Long
    Dim Str As Long
    Dim DEF As Long
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim AcumuloDrop As Long, DP As Integer
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).NPC(mapnpcnum).Num
    Name = Trim$(NPC(NpcNum).Name)

    ' Check for weapon
    n = 0

    If GetPlayerEquipmentNum(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipmentNum(Attacker, Weapon)
    End If

    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = getTime

    If Damage >= MapNpc(MapNum).NPC(mapnpcnum).Vital(Vitals.HP) Then

        SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(MapNum).NPC(mapnpcnum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(MapNum).NPC(mapnpcnum).X * 32), (MapNpc(MapNum).NPC(mapnpcnum).Y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y

        ' send the sound
        If SpellNum > 0 Then SendMapSound Attacker, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, SoundEntity.seSpell, SpellNum

        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipmentNum(Attacker, Weapon)).Animation, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y)
            End If
        End If

        ' Calculate exp to give attacker
        EXP = NpcExpCalculate(Attacker, NpcNum)

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, EXP, Attacker, NPC(NpcNum).Level
        Else
            ' no party - keep exp for self
            GivePlayerEXP Attacker, EXP, NPC(NpcNum).Level
        End If

        'Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS

            If NPC(NpcNum).DropItem(n) > 0 Then

                ' Se o jogador for premium, verifica o bonus de chance de drop!
                If GetPlayerPremium(Attacker) = YES Then
                    DP = Options.PREMIUMDROP
                    AcumuloDrop = AcumuloDrop + ((NPC(NpcNum).DropChance(n) / 100) * DP)
                End If

                If TempPlayer(Attacker).Bonus.Drop > 0 Then
                    AcumuloDrop = AcumuloDrop + ((NPC(NpcNum).DropChance(n) / 100) * TempPlayer(Attacker).Bonus.Drop)
                End If

                ' Proteção pra evitar erros...
                If ((NPC(NpcNum).DropChance(n) - AcumuloDrop) + 1) < 1 Then
                    i = 1
                Else
                    i = Int(Rnd * (NPC(NpcNum).DropChance(n) - AcumuloDrop)) + 1
                End If

                R = Int(Rnd * NPC(NpcNum).DropItemValue(n)) + 1

                If (DP + TempPlayer(Attacker).Bonus.Drop) > 0 Then
                    SendActionMsg MapNum, "Drop Chance + " & (DP + TempPlayer(Attacker).Bonus.Drop) & "%", Yellow, ACTIONMSG_SCROLL, GetPlayerX(Attacker) * 32, GetPlayerY(Attacker) * 32 - 16
                End If
                If i = 1 Then
                    Call SpawnItem(NPC(NpcNum).DropItem(n), R, MapNum, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, GetPlayerName(Attacker))
                End If
            End If

        Next

        ' destroy map npcs
        If Map(MapNum).MapData.Moral = MAP_MORAL_BOSS Then
            If mapnpcnum = Map(MapNum).MapData.BossNpc Then
                ' kill all the other npcs
                For i = 1 To MAX_MAP_NPCS
                    If Map(MapNum).MapData.NPC(i) > 0 Then
                        ' only kill dangerous npcs
                        If NPC(Map(MapNum).MapData.NPC(i)).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(Map(MapNum).MapData.NPC(i)).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            ' kill!
                            MapNpc(MapNum).NPC(i).Dead = YES
                            MapNpc(MapNum).NPC(i).tmpNum = MapNpc(MapNum).NPC(i).Num

                            MapNpc(MapNum).NPC(i).Num = 0
                            MapNpc(MapNum).NPC(i).SpawnWait = getTime
                            MapNpc(MapNum).NPC(i).Vital(Vitals.HP) = 0
                            ' send kill command
                            SendNpcDeath MapNum, i
                        End If
                    End If
                Next
            End If
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).NPC(mapnpcnum).Dead = YES
        MapNpc(MapNum).NPC(mapnpcnum).tmpNum = MapNpc(MapNum).NPC(mapnpcnum).Num

        MapNpc(MapNum).NPC(mapnpcnum).Num = 0
        MapNpc(MapNum).NPC(mapnpcnum).SpawnWait = getTime
        MapNpc(MapNum).NPC(mapnpcnum).Vital(Vitals.HP) = 0

        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).NPC(mapnpcnum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With

            With MapNpc(MapNum).NPC(mapnpcnum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next

        ' check task
        Call CheckTasks(Attacker, QUEST_TYPE_GOSLAY, NpcNum)

        ' send death to the map
        SendNpcDeath MapNum, mapnpcnum

        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = MapNum Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = mapnpcnum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).NPC(mapnpcnum).Vital(Vitals.HP) = MapNpc(MapNum).NPC(mapnpcnum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).NPC(mapnpcnum).X * 32), (MapNpc(MapNum).NPC(mapnpcnum).Y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y

        ' send the sound
        If SpellNum > 0 Then SendMapSound Attacker, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, SoundEntity.seSpell, SpellNum

        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipmentNum(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapnpcnum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(MapNum).NPC(mapnpcnum).TargetType = 1    ' player
        MapNpc(MapNum).NPC(mapnpcnum).target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(MapNum).NPC(mapnpcnum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).NPC(i).Num = MapNpc(MapNum).NPC(mapnpcnum).Num Then
                    MapNpc(MapNum).NPC(i).target = Attacker
                    MapNpc(MapNum).NPC(i).TargetType = 1    ' player
                End If
            Next
        End If

        ' set the regen timer
        MapNpc(MapNum).NPC(mapnpcnum).stopRegen = True
        MapNpc(MapNum).NPC(mapnpcnum).stopRegenTimer = getTime

        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunNPC mapnpcnum, MapNum, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Npc MapNum, mapnpcnum, SpellNum, Attacker
            End If
        End If

        SendMapNpcVitals MapNum, mapnpcnum

        ' set the player's target if they don't have one
        TempPlayer(Attacker).TargetType = TARGET_TYPE_NPC
        TempPlayer(Attacker).target = mapnpcnum
        SendTarget Attacker
    End If

    If SpellNum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = getTime
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal mapnpcnum As Long, ByVal Index As Long)
    Dim MapNum As Long, NpcNum As Long, blockAmount As Long, Damage As Long, Defence As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(mapnpcnum, Index) Then
        MapNum = GetPlayerMap(Index)
        NpcNum = MapNpc(MapNum).NPC(mapnpcnum).Num

        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index, TARGET_TYPE_NPC, mapnpcnum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NpcNum)

        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(Index, Damage, mapnpcnum)

        If blockAmount > 0 Then
            Damage = Damage - blockAmount
            SendActionMsg MapNum, "Block " & blockAmount, Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
        End If

        ' take away armour
        Defence = GetPlayerDefence(Index)
        If Defence > 0 Then
            Damage = Damage - Rand(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If

        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = Rand(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))

        ' * 1.5 if crit hit
        If CanNpcCrit(NpcNum, Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).NPC(mapnpcnum).X * 32), (MapNpc(MapNum).NPC(mapnpcnum).Y * 32)
            
            ' stun one second
            StunPlayerForTimer Index, 1
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(mapnpcnum, Index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal mapnpcnum As Long, ByVal Index As Long, Optional ByVal isSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long

    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).NPC(mapnpcnum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Index)
    NpcNum = MapNpc(MapNum).NPC(mapnpcnum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).NPC(mapnpcnum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If
    
    ' Check to make sure they aren't stunned!
    If MapNpc(MapNum).NPC(mapnpcnum).StunDuration > 0 Then Exit Function

    ' exit out early if it's a spell
    If isSpell Then
        If IsPlaying(Index) Then
            If NpcNum > 0 Then
                CanNpcAttackPlayer = True
                Exit Function
            End If
        End If
    End If

    ' Make sure npcs dont attack more then once a second
    If getTime < MapNpc(MapNum).NPC(mapnpcnum).AttackTimer + 1000 Then
        Exit Function
    End If
    MapNpc(MapNum).NPC(mapnpcnum).AttackTimer = getTime

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NpcNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum).NPC(mapnpcnum).Y) And (GetPlayerX(Index) = MapNpc(MapNum).NPC(mapnpcnum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(MapNum).NPC(mapnpcnum).Y) And (GetPlayerX(Index) = MapNpc(MapNum).NPC(mapnpcnum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum).NPC(mapnpcnum).Y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum).NPC(mapnpcnum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum).NPC(mapnpcnum).Y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum).NPC(mapnpcnum).X) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal mapnpcnum As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim EXP As Long
    Dim MapNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).NPC(mapnpcnum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Victim)
    Name = Trim$(NPC(MapNpc(MapNum).NPC(mapnpcnum).Num).Name)

    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong mapnpcnum
    SendDataToMap MapNum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing

    If Damage <= 0 Then
        Exit Sub
    End If

    ' set the regen timer
    MapNpc(MapNum).NPC(mapnpcnum).stopRegen = True
    MapNpc(MapNum).NPC(mapnpcnum).stopRegenTimer = getTime

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)

        ' send the sound
        If SpellNum > 0 Then
            SendMapSound Victim, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, SoundEntity.seSpell, SpellNum
        Else
            SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(MapNum).NPC(mapnpcnum).Num
        End If

        ' send animation
        If Not overTime Then
            If SpellNum = 0 Then Call SendAnimation(MapNum, NPC(MapNpc(MapNum).NPC(mapnpcnum).Num).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
        End If

        ' kill player
        KillPlayer Victim

        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(MapNum).NPC(mapnpcnum).target = 0
        MapNpc(MapNum).NPC(mapnpcnum).TargetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        Call SendMapVitals(Victim)

        ' send the sound
        If SpellNum > 0 Then
            SendMapSound Victim, MapNpc(MapNum).NPC(mapnpcnum).X, MapNpc(MapNum).NPC(mapnpcnum).Y, SoundEntity.seSpell, SpellNum
        Else
            SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(MapNum).NPC(mapnpcnum).Num
        End If

        ' send animation
        If Not overTime Then
            If SpellNum = 0 Then Call SendAnimation(MapNum, NPC(MapNpc(GetPlayerMap(Victim)).NPC(mapnpcnum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
        End If

        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunPlayer Victim, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                ' TODO: Add Npc vs Player DOTs
            End If
        End If

        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim

        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(MapNum).NPC(mapnpcnum).Num

        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)

        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = getTime
    End If

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long)
    Dim blockAmount As Long, NpcNum As Long, MapNum As Long, Damage As Long, Defence As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, Victim) Then

        MapNum = GetPlayerMap(Attacker)

        ' check if NPC can avoid the attack
        If CanPlayerDodge(Victim, TARGET_TYPE_PLAYER, Attacker) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)

        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(Victim, Damage, Attacker)

        If blockAmount > 0 Then
            Damage = Damage - blockAmount
            SendActionMsg MapNum, "Block " & blockAmount, Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        End If

        ' take away armour
        Defence = GetPlayerDefence(Victim)
        If Defence > 0 Then
            Damage = Damage - Rand(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If

        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = Rand(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))

        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker, TARGET_TYPE_PLAYER, Victim) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
            
            ' Stuna o player quando acerta um crítico
            StunPlayerForTimer Victim, 1
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, Victim, Damage)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal isSpell As Boolean = False) As Boolean
    Dim partynum As Long, i As Long

    If Not isSpell Then
        ' Check attack timer
        If GetPlayerEquipmentNum(Attacker, Weapon) > 0 Then
            If getTime < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipmentNum(Attacker, Weapon)).Speed Then Exit Function
        Else
            If getTime < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function

    ' make sure it's not you
    If Victim = Attacker Then
        PlayerMsg Attacker, "Cannot attack yourself.", BrightRed
        Exit Function
    End If

    ' check co-ordinates if not spell
    If Not isSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
        Case DIR_UP

            If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
        Case DIR_DOWN

            If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
        Case DIR_LEFT

            If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
        Case DIR_RIGHT

            If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
        Case Else
            Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).MapData.Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 5 Then
        Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 5 Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    ' make sure not in your party
    partynum = TempPlayer(Attacker).inParty
    If partynum > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partynum).Member(i) > 0 Then
                If Victim = Party(partynum).Member(i) Then
                    PlayerMsg Attacker, "Cannot attack party members.", BrightRed
                    Exit Function
                End If
            End If
        Next
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim EXP As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipmentNum(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipmentNum(Attacker, Weapon)
    End If

    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = getTime

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)

        ' send the sound
        If SpellNum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, SpellNum

        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        EXP = (GetPlayerExp(Victim) \ 10)

        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 0
        End If

        If EXP = 0 Then
            Call PlayerMsg(Victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - EXP)
            SendEXP Victim
            Call PlayerMsg(Victim, "You lost " & EXP & " exp.", BrightRed)

            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, EXP, Attacker, GetPlayerLevel(Victim)
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, EXP, GetPlayerLevel(Victim)
            End If
        End If

        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = Victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
        
        ' check task
        Call CheckTasks(Attacker, QUEST_TYPE_GOKILL, Victim)
        
        ' send death
        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        Call SendMapVitals(Victim)

        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim

        ' send the sound
        If SpellNum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, SpellNum

        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)

        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = getTime

        'if a stunning spell, stun the player
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunPlayer Victim, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Player Victim, SpellNum, Attacker
            End If
        End If

        ' change target if need be
        TempPlayer(Attacker).TargetType = TARGET_TYPE_PLAYER
        TempPlayer(Attacker).target = Victim
        SendTarget Attacker
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = getTime
End Sub

' ############
' ## Spells ##
' ############
Public Sub BufferSpell(ByVal Index As Long, ByVal spellSlot As Long)
    Dim SpellNum As Long, mpCost As Long, LevelReq As Long, MapNum As Long, spellCastType As Long, ClassReq As Long
    Dim AccessReq As Long, Range As Long, HasBuffered As Boolean, TargetType As Byte, target As Long

    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub

    SpellNum = Player(Index).Spell(spellSlot).Spell
    MapNum = GetPlayerMap(Index)

    If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub

    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub
    
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub

    ' make sure we're not buffering already
    If TempPlayer(Index).spellBuffer.Spell = spellSlot Then Exit Sub

    ' see if cooldown has finished
    If Player(Index).SpellCD(spellSlot) <> 0 Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    mpCost = Spell(SpellNum).mpCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < mpCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If

    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If

    AccessReq = Spell(SpellNum).AccessReq

    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If

    ClassReq = Spell(SpellNum).ClassReq

    ' make sure the classreq > 0
    If ClassReq > 0 Then    ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If

    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            spellCastType = 2    ' targetted
        Else
            spellCastType = 3    ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            spellCastType = 0    ' self-cast
        Else
            spellCastType = 1    ' self-cast AoE
        End If
    End If

    TargetType = TempPlayer(Index).TargetType
    target = TempPlayer(Index).target
    Range = Spell(SpellNum).Range
    HasBuffered = False

    Select Case spellCastType
    Case 0, 1    ' self-cast & self-cast AOE
        HasBuffered = True
    Case 2, 3    ' targeted & targeted AOE
        ' check if have target
        If Not target > 0 Then
            PlayerMsg Index, "You do not have a target.", BrightRed
        End If
        If TargetType = TARGET_TYPE_PLAYER Then
            ' if have target, check in range
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(target), GetPlayerY(target)) Then
                PlayerMsg Index, "Target not in range.", BrightRed
            Else
                ' go through spell types
                If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                    HasBuffered = True
                Else
                    If CanPlayerAttackPlayer(Index, target, True) Then
                        HasBuffered = True
                    End If
                End If
            End If
        ElseIf TargetType = TARGET_TYPE_NPC Then
            ' if beneficial magic then self-cast it instead
            If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Or Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                target = Index
                TargetType = TARGET_TYPE_PLAYER
                HasBuffered = True
            Else
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(MapNum).NPC(target).X, MapNpc(MapNum).NPC(target).Y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(Index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
        End If
    End Select

    If HasBuffered Then
        SendAnimation MapNum, Spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index, 1
        TempPlayer(Index).spellBuffer.Spell = spellSlot
        TempPlayer(Index).spellBuffer.Timer = getTime
        TempPlayer(Index).spellBuffer.target = target
        TempPlayer(Index).spellBuffer.tType = TargetType
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If
End Sub

Public Sub NpcBufferSpell(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal npcSpellSlot As Long)
    Dim SpellNum As Long, mpCost As Long, Range As Long, HasBuffered As Boolean, TargetType As Byte, target As Long, spellCastType As Long, i As Long

    ' prevent rte9
    If npcSpellSlot <= 0 Or npcSpellSlot > MAX_NPC_SPELLS Then Exit Sub

    With MapNpc(MapNum).NPC(mapnpcnum)
        ' set the spell number
        SpellNum = NPC(.Num).Spell(npcSpellSlot)

        ' prevent rte9
        If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub

        ' make sure we're not already buffering
        If .spellBuffer.Spell > 0 Then Exit Sub

        ' see if cooldown as finished
        If .SpellCD(npcSpellSlot) > getTime Then Exit Sub

        ' Set the MP Cost
        mpCost = Spell(SpellNum).mpCost

        ' have they got enough mp?
        If .Vital(Vitals.MP) < mpCost Then Exit Sub

        ' find out what kind of spell it is! self cast, target or AOE
        If Spell(SpellNum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not Spell(SpellNum).IsAoE Then
                spellCastType = 2    ' targetted
            Else
                spellCastType = 3    ' targetted aoe
            End If
        Else
            If Not Spell(SpellNum).IsAoE Then
                spellCastType = 0    ' self-cast
            Else
                spellCastType = 1    ' self-cast AoE
            End If
        End If

        TargetType = .TargetType
        target = .target
        Range = Spell(SpellNum).Range
        HasBuffered = False

        ' make sure on the map
        If GetPlayerMap(target) <> MapNum Then Exit Sub

        Select Case spellCastType
        Case 0, 1    ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3    ' targeted & targeted AOE
            ' if it's a healing spell then heal a friend
            If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                ' find a friend who needs healing
                For i = 1 To MAX_MAP_NPCS
                    If MapNpc(MapNum).NPC(i).Num > 0 Then
                        If MapNpc(MapNum).NPC(i).Vital(Vitals.HP) < GetNpcMaxVital(MapNpc(MapNum).NPC(i).Num, Vitals.HP) Then
                            TargetType = TARGET_TYPE_NPC
                            target = i
                            HasBuffered = True
                        End If
                    End If
                Next
            Else
                ' check if have target
                If Not target > 0 Then Exit Sub
                ' make sure it's a player
                If TargetType = TARGET_TYPE_PLAYER Then
                    ' if have target, check in range
                    If Not isInRange(Range, .X, .Y, GetPlayerX(target), GetPlayerY(target)) Then
                        Exit Sub
                    Else
                        If CanNpcAttackPlayer(mapnpcnum, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
        End Select

        If HasBuffered Then
            SendAnimation MapNum, Spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_NPC, mapnpcnum
            .spellBuffer.Spell = npcSpellSlot
            .spellBuffer.Timer = getTime
            .spellBuffer.target = target
            .spellBuffer.tType = TargetType
        End If
    End With
End Sub

Public Sub NpcCastSpell(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal spellSlot As Long, ByVal target As Long, ByVal TargetType As Long)
    Dim SpellNum As Long, mpCost As Long, Vital As Long, DidCast As Boolean, i As Long, AoE As Long, Range As Long, vitalType As Byte, increment As Boolean, X As Long, Y As Long, spellCastType As Long

    DidCast = False

    ' rte9
    If spellSlot <= 0 Or spellSlot > MAX_NPC_SPELLS Then Exit Sub

    With MapNpc(MapNum).NPC(mapnpcnum)
        ' cache spell num
        SpellNum = NPC(.Num).Spell(spellSlot)

        ' cache mp cost
        mpCost = Spell(SpellNum).mpCost

        ' make sure still got enough mp
        If .Vital(Vitals.MP) < mpCost Then Exit Sub

        ' find out what kind of spell it is! self cast, target or AOE
        If Spell(SpellNum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not Spell(SpellNum).IsAoE Then
                spellCastType = 2    ' targetted
            Else
                spellCastType = 3    ' targetted aoe
            End If
        Else
            If Not Spell(SpellNum).IsAoE Then
                spellCastType = 0    ' self-cast
            Else
                spellCastType = 1    ' self-cast AoE
            End If
        End If

        ' get damage
        Vital = GetNpcSpellDamage(.Num, SpellNum)    'GetPlayerSpellDamage(index, spellNum)

        ' store data
        AoE = Spell(SpellNum).AoE
        Range = Spell(SpellNum).Range

        Select Case spellCastType
        Case 0    ' self-cast target
            Select Case Spell(SpellNum).Type
            Case SPELL_TYPE_HEALHP
                SpellNpc_Effect Vitals.HP, True, mapnpcnum, Vital, SpellNum, MapNum
                DidCast = True
            Case SPELL_TYPE_HEALMP
                SpellNpc_Effect Vitals.MP, True, mapnpcnum, Vital, SpellNum, MapNum
                DidCast = True
            End Select
        Case 1, 3    ' self-cast AOE & targetted AOE
            If spellCastType = 1 Then
                X = .X
                Y = .Y
            ElseIf spellCastType = 3 Then
                If TargetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub

                If TargetType = TARGET_TYPE_PLAYER Then
                    X = GetPlayerX(target)
                    Y = GetPlayerY(target)
                Else
                    X = MapNpc(MapNum).NPC(target).X
                    Y = MapNpc(MapNum).NPC(target).Y
                End If

                If Not isInRange(Range, .X, .Y, X, Y) Then Exit Sub
            End If
            Select Case Spell(SpellNum).Type
            Case SPELL_TYPE_DAMAGEHP
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If GetPlayerMap(i) = MapNum Then
                            If isInRange(AoE, .X, .Y, GetPlayerX(i), GetPlayerY(i)) Then
                                If CanNpcAttackPlayer(mapnpcnum, i, True) Then
                                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                    NpcAttackPlayer mapnpcnum, i, Vital, SpellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    End If
                Next
            Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
                If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                    vitalType = Vitals.HP
                    increment = True
                ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                    vitalType = Vitals.MP
                    increment = True
                End If

                If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Or Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).NPC(i).Num > 0 Then
                            If MapNpc(MapNum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                    SpellNpc_Effect vitalType, increment, i, Vital, SpellNum, MapNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                End If
            End Select
        Case 2    ' targetted
            If TargetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub

            If TargetType = TARGET_TYPE_PLAYER Then
                X = GetPlayerX(target)
                Y = GetPlayerY(target)
            Else
                X = MapNpc(MapNum).NPC(target).X
                Y = MapNpc(MapNum).NPC(target).Y
            End If

            If Not isInRange(Range, .X, .Y, X, Y) Then Exit Sub

            Select Case Spell(SpellNum).Type
            Case SPELL_TYPE_DAMAGEHP
                If TargetType = TARGET_TYPE_PLAYER Then
                    If CanNpcAttackPlayer(mapnpcnum, target, True) Then
                        If Vital > 0 Then
                            SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                            NpcAttackPlayer mapnpcnum, target, Vital, SpellNum
                            DidCast = True
                        End If
                    End If
                End If
            Case SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                If Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                    vitalType = Vitals.MP
                    increment = True
                ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                    vitalType = Vitals.HP
                    increment = True
                End If

                If TargetType = TARGET_TYPE_NPC Then
                    SpellNpc_Effect vitalType, increment, target, Vital, SpellNum, MapNum
                    DidCast = True
                End If
            End Select
        End Select

        If DidCast Then
            .Vital(Vitals.MP) = .Vital(Vitals.MP) - mpCost
            .SpellCD(spellSlot) = getTime + (Spell(SpellNum).CDTime * 1000)
            Call SendMapNpcVitals(MapNum, mapnpcnum)
        End If
    End With
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellSlot As Long, ByVal target As Long, ByVal TargetType As Byte)
    Dim SpellNum As Long, mpCost As Long, LevelReq As Long, MapNum As Long, Vital As Long, DidCast As Boolean, ClassReq As Long
    Dim AccessReq As Long, i As Long, AoE As Long, Range As Long, vitalType As Byte, increment As Boolean, X As Long, Y As Long
    Dim Buffer As clsBuffer, spellCastType As Long

    DidCast = False

    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub

    SpellNum = Player(Index).Spell(spellSlot).Spell
    MapNum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub

    mpCost = Spell(SpellNum).mpCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < mpCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If

    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If

    AccessReq = Spell(SpellNum).AccessReq

    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If

    ClassReq = Spell(SpellNum).ClassReq

    ' make sure the classreq > 0
    If ClassReq > 0 Then    ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If

    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            spellCastType = 2    ' targetted
        Else
            spellCastType = 3    ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            spellCastType = 0    ' self-cast
        Else
            spellCastType = 1    ' self-cast AoE
        End If
    End If

    ' get damage
    Vital = GetPlayerSpellDamage(Index, SpellNum)

    ' store data
    AoE = Spell(SpellNum).AoE
    Range = Spell(SpellNum).Range

    Select Case spellCastType
    Case 0    ' self-cast target
        Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_HEALHP
            SpellPlayer_Effect Vitals.HP, True, Index, Vital, SpellNum
            DidCast = True
        Case SPELL_TYPE_HEALMP
            SpellPlayer_Effect Vitals.MP, True, Index, Vital, SpellNum
            DidCast = True
        Case SPELL_TYPE_WARP
            SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
            PlayerWarp Index, Spell(SpellNum).Map, Spell(SpellNum).X, Spell(SpellNum).Y
            SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
            DidCast = True
        End Select
    Case 1, 3    ' self-cast AOE & targetted AOE
        If spellCastType = 1 Then
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index)
        ElseIf spellCastType = 3 Then
            If TargetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub

            If TargetType = TARGET_TYPE_PLAYER Then
                X = GetPlayerX(target)
                Y = GetPlayerY(target)
            Else
                X = MapNpc(MapNum).NPC(target).X
                Y = MapNpc(MapNum).NPC(target).Y
            End If

            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
            End If
        End If
        Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_DAMAGEHP
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If i <> Index Then
                        If GetPlayerMap(i) = GetPlayerMap(Index) Then
                            If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                If CanPlayerAttackPlayer(Index, i, True) Then
                                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                    PlayerAttackPlayer Index, i, Vital, SpellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).NPC(i).Num > 0 Then
                    If MapNpc(MapNum).NPC(i).Vital(HP) > 0 Then
                        If isInRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                            If CanPlayerAttackNpc(Index, i, True) Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                PlayerAttackNpc Index, i, Vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    End If
                End If
            Next
        Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
            If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                vitalType = Vitals.HP
                increment = True
            ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                vitalType = Vitals.MP
                increment = True
            ElseIf Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                vitalType = Vitals.MP
                increment = False
            End If

            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(Index) Then
                        If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                            SpellPlayer_Effect vitalType, increment, i, Vital, SpellNum
                            DidCast = True
                        End If
                    End If
                End If
            Next

            If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                For i = 1 To MAX_MAP_NPCS
                    If MapNpc(MapNum).NPC(i).Num > 0 Then
                        If MapNpc(MapNum).NPC(i).Vital(HP) > 0 Then
                            If isInRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                SpellNpc_Effect vitalType, increment, i, Vital, SpellNum, MapNum
                                DidCast = True
                            End If
                        End If
                    End If
                Next
            End If
        End Select
    Case 2    ' targetted
        If TargetType = 0 Then Exit Sub
        If target = 0 Then Exit Sub

        If TargetType = TARGET_TYPE_PLAYER Then
            X = GetPlayerX(target)
            Y = GetPlayerY(target)
        Else
            X = MapNpc(MapNum).NPC(target).X
            Y = MapNpc(MapNum).NPC(target).Y
        End If

        If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
            PlayerMsg Index, "Target not in range.", BrightRed
            SendClearSpellBuffer Index
            Exit Sub
        End If

        Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_DAMAGEHP
            If TargetType = TARGET_TYPE_PLAYER Then
                If CanPlayerAttackPlayer(Index, target, True) Then
                    If Vital > 0 Then
                        SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                        PlayerAttackPlayer Index, target, Vital, SpellNum
                        DidCast = True
                    End If
                End If
            Else
                If CanPlayerAttackNpc(Index, target, True) Then
                    If Vital > 0 Then
                        SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                        PlayerAttackNpc Index, target, Vital, SpellNum
                        DidCast = True
                    End If
                End If
            End If

        Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
            If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                vitalType = Vitals.MP
                increment = False
            ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                vitalType = Vitals.MP
                increment = True
            ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                vitalType = Vitals.HP
                increment = True
            End If

            If TargetType = TARGET_TYPE_PLAYER Then
                If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                    If CanPlayerAttackPlayer(Index, target, True) Then
                        SpellPlayer_Effect vitalType, increment, target, Vital, SpellNum
                        DidCast = True
                    End If
                Else
                    SpellPlayer_Effect vitalType, increment, target, Vital, SpellNum
                    DidCast = True
                End If
            Else
                If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                    If CanPlayerAttackNpc(Index, target, True) Then
                        SpellNpc_Effect vitalType, increment, target, Vital, SpellNum, MapNum
                        DidCast = True
                    End If
                Else
                    SpellNpc_Effect vitalType, increment, target, Vital, SpellNum, MapNum
                    DidCast = True
                End If
            End If
        End Select
    End Select

    If DidCast Then
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - mpCost)
        Call SendMapVitals(Index)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

        Player(Index).SpellCD(spellSlot) = Spell(SpellNum).CDTime
        Call SendCooldown(Index, spellSlot, Player(Index).SpellCD(spellSlot))

        ' if has a next rank then increment usage
        SetPlayerSpellUsage Index, spellSlot
    End If
End Sub

Public Sub SetPlayerSpellUsage(ByVal Index As Long, ByVal spellSlot As Long)
    Dim SpellNum As Long, i As Long
    SpellNum = Player(Index).Spell(spellSlot).Spell
    ' if has a next rank then increment usage
    If Spell(SpellNum).NextRank > 0 Then
        If Player(Index).Spell(spellSlot).Uses < Spell(SpellNum).NextUses - 1 Then
            Player(Index).Spell(spellSlot).Uses = Player(Index).Spell(spellSlot).Uses + 1
        Else
            If GetPlayerLevel(Index) >= Spell(Spell(SpellNum).NextRank).LevelReq Then
                Player(Index).Spell(spellSlot).Spell = Spell(SpellNum).NextRank
                Player(Index).Spell(spellSlot).Uses = 0
                PlayerMsg Index, "Your spell has ranked up!", Blue
                ' update hotbar
                For i = 1 To MAX_HOTBAR
                    If Player(Index).Hotbar(i).Slot > 0 Then
                        If Player(Index).Hotbar(i).sType = 2 Then    ' spell
                            If Spell(Player(Index).Hotbar(i).Slot).UniqueIndex = Spell(Spell(SpellNum).NextRank).UniqueIndex Then
                                Player(Index).Hotbar(i).Slot = Spell(SpellNum).NextRank
                                SendHotbar Index
                            End If
                        End If
                    End If
                Next
            Else
                Player(Index).Spell(spellSlot).Uses = Spell(SpellNum).NextUses
            End If
        End If
        SendPlayerSpells Index
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long)
    Dim sSymbol As String * 1
    Dim colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then colour = BrightGreen
            If Vital = Vitals.MP Then colour = BrightBlue
        Else
            sSymbol = "-"
            colour = Blue
        End If

        SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32

        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum

        If increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Player Index, SpellNum
            End If
        ElseIf Not increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
        End If

        ' send update
        Call SendMapVitals(Index)
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal MapNum As Long)
    Dim sSymbol As String * 1
    Dim colour As Long
    Dim NpcNum As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then colour = BrightGreen
            If Vital = Vitals.MP Then colour = BrightBlue
        Else
            sSymbol = "-"
            colour = Blue
        End If

        SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
        SendActionMsg MapNum, sSymbol & Damage, colour, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(Index).X * 32, MapNpc(MapNum).NPC(Index).Y * 32

        ' send the sound
        SendMapSound Index, MapNpc(MapNum).NPC(Index).X, MapNpc(MapNum).NPC(Index).Y, SoundEntity.seSpell, SpellNum

        NpcNum = MapNpc(MapNum).NPC(Index).Num
        If increment Then
            MapNpc(MapNum).NPC(Index).Vital(Vital) = MapNpc(MapNum).NPC(Index).Vital(Vital) + Damage
            ' make sure doesn't go over max
            With MapNpc(MapNum).NPC(Index)
                If .Vital(Vital) > GetNpcMaxVital(NpcNum, Vital) Then
                    .Vital(Vital) = GetNpcMaxVital(NpcNum, Vital)
                End If
                Call SendMapNpcVitals(MapNum, Index)
            End With
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Npc MapNum, Index, SpellNum
            End If
        ElseIf Not increment Then
            MapNpc(MapNum).NPC(Index).Vital(Vital) = MapNpc(MapNum).NPC(Index).Vital(Vital) - Damage
            Call SendMapNpcVitals(MapNum, Index)
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = getTime
                .Caster = Caster
                .StartTime = getTime
                Exit Sub
            End If

            If .Used = False Then
                .Spell = SpellNum
                .Timer = getTime
                .Caster = Caster
                .Used = True
                .StartTime = getTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal Index As Long, ByVal SpellNum As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = getTime
                .StartTime = getTime
                Exit Sub
            End If

            If .Used = False Then
                .Spell = SpellNum
                .Timer = getTime
                .Used = True
                .StartTime = getTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).NPC(Index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = getTime
                .Caster = Caster
                .StartTime = getTime
                Exit Sub
            End If

            If .Used = False Then
                .Spell = SpellNum
                .Timer = getTime
                .Caster = Caster
                .Used = True
                .StartTime = getTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal SpellNum As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).NPC(Index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = getTime
                .StartTime = getTime
                Exit Sub
            End If

            If .Used = False Then
                .Spell = SpellNum
                .Timer = getTime
                .Used = True
                .StartTime = getTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal Index As Long, ByVal dotNum As Long)
    With TempPlayer(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If getTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, Index, True) Then
                    PlayerAttackPlayer .Caster, Index, GetPlayerSpellDamage(.Caster, .Spell)
                End If
                .Timer = getTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If getTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal Index As Long, ByVal hotNum As Long)
    With TempPlayer(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If getTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg Player(Index).Map, "+" & GetPlayerSpellDamage(.Caster, .Spell), BrightGreen, ACTIONMSG_SCROLL, Player(Index).X * 32, Player(Index).Y * 32
                Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + GetPlayerSpellDamage(.Caster, .Spell)
                .Timer = getTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If getTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal dotNum As Long)
    With MapNpc(MapNum).NPC(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If getTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, Index, True) Then
                    PlayerAttackNpc .Caster, Index, GetPlayerSpellDamage(.Caster, .Spell), , True
                End If
                .Timer = getTime
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If getTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal hotNum As Long)
    Dim NpcNum As Long

    With MapNpc(MapNum).NPC(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If getTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg MapNum, "+" & GetPlayerSpellDamage(.Caster, .Spell), BrightGreen, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(Index).X * 32, MapNpc(MapNum).NPC(Index).Y * 32
                MapNpc(MapNum).NPC(Index).Vital(Vitals.HP) = MapNpc(MapNum).NPC(Index).Vital(Vitals.HP) + GetPlayerSpellDamage(.Caster, .Spell)
                ' make sure not over max
                NpcNum = MapNpc(MapNum).NPC(Index).Num
                If MapNpc(MapNum).NPC(Index).Vital(Vitals.HP) > GetNpcMaxVital(NpcNum, Vitals.HP) Then
                    MapNpc(MapNum).NPC(Index).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
                End If
                .Timer = getTime
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If getTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal Index As Long, ByVal SpellNum As Long)
' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Spell(SpellNum).StunDuration
        TempPlayer(Index).StunTimer = getTime
        ' send it to the index
        SendStunned GetPlayerMap(Index), Index, TARGET_TYPE_PLAYER
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal MapNum As Long, ByVal SpellNum As Long)
' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(MapNum).NPC(Index).StunDuration = Spell(SpellNum).StunDuration
        MapNpc(MapNum).NPC(Index).StunTimer = getTime
        SendStunned MapNum, Index, TARGET_TYPE_NPC
    End If
End Sub

Public Sub StunPlayerForTimer(ByVal Index As Long, ByVal Secs As Long)
        ' set the values on index
        TempPlayer(Index).StunDuration = Secs
        TempPlayer(Index).StunTimer = getTime
        ' send it to the index
        SendStunned GetPlayerMap(Index), Index, TARGET_TYPE_PLAYER
End Sub

Public Sub StunNPCForTimer(ByVal Index As Long, ByVal MapNum As Long, ByVal Secs As Long)
        ' set the values on index
        MapNpc(MapNum).NPC(Index).StunDuration = Secs
        MapNpc(MapNum).NPC(Index).StunTimer = getTime
        
        SendStunned MapNum, Index, TARGET_TYPE_NPC
End Sub

Public Function NpcExpCalculate(ByVal Attacker As Integer, ByVal NpcNum As Integer) As Long
    Dim ExpOriginal As Long
    
    '//DICA//
    ' Pra adicionar novos bonus, sempre utilize o valor original da exp do npc, pra não acumular bonus em cima de bonus!
    
    ExpOriginal = NPC(NpcNum).EXP

    ' Calculate exp to give attacker
    If NPC(NpcNum).RandExp = 0 Then
        NpcExpCalculate = ExpOriginal
    Else
        'randomize exp within specified value
        If NPC(NpcNum).Percent_5 = 1 Then
            NpcExpCalculate = Rand(ExpOriginal - (ExpOriginal * 0.05), ExpOriginal + (ExpOriginal * 0.05))
        ElseIf NPC(NpcNum).Percent_10 = 1 Then
            NpcExpCalculate = Rand(ExpOriginal - (ExpOriginal * 0.1), ExpOriginal + (ExpOriginal * 0.1))
        ElseIf NPC(NpcNum).Percent_20 = 1 Then
            NpcExpCalculate = Rand(ExpOriginal - (ExpOriginal * 0.2), ExpOriginal + (ExpOriginal * 0.2))
        End If
    End If
    
    ' Caso receba 0 exp, adiciona 1 pra não dividir sobre zero.
    If NpcExpCalculate <= 0 Then NpcExpCalculate = 1
    
    ' Bonus Player Premium
    If GetPlayerPremium(Attacker) = YES Then
        NpcExpCalculate = NpcExpCalculate + ((ExpOriginal / 100) * Options.PREMIUMEXP) ' 50% de exp
    End If
    
    ' Bonus de Conjunto
    If TempPlayer(Attacker).Bonus.EXP > 0 Then
        NpcExpCalculate = NpcExpCalculate + ((ExpOriginal / 100) * TempPlayer(Attacker).Bonus.EXP)
    End If
End Function

