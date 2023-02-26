Attribute VB_Name = "modPlayerUSEITEM"
Option Explicit

Public Sub UseItem(ByVal Index As Long, ByVal InvNum As Long)
    Dim ItemNum As Long

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then
        Exit Sub
    End If

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Find out what kind of item it is

    Select Case Item(ItemNum).Type
    Case ITEM_TYPE_WEAPON
        Call EquipItem(Index, InvNum, Weapon)

    Case ITEM_TYPE_SHIELD
        Call EquipItem(Index, InvNum, Shield)

    Case ITEM_TYPE_HELMET
        Call EquipItem(Index, InvNum, Helmet)

    Case ITEM_TYPE_ARMOR
        Call EquipItem(Index, InvNum, Armor)

    Case ITEM_TYPE_LEGS
        Call EquipItem(Index, InvNum, Legs)

    Case ITEM_TYPE_BOOTS
        Call EquipItem(Index, InvNum, Boots)

    Case ITEM_TYPE_AMULET
        Call EquipItem(Index, InvNum, Amulet)

    Case ITEM_TYPE_RINGLEFT
        Call EquipItem(Index, InvNum, RingLeft)

    Case ITEM_TYPE_RINGRIGHT
        Call EquipItem(Index, InvNum, RingRight)

    Case ITEM_TYPE_CONSUME
        Call UseItem_Consume(Index, InvNum)

    Case ITEM_TYPE_KEY
        Call UseItem_Key(Index, InvNum)

    Case ITEM_TYPE_UNIQUE
        Call Unique_Item(Index, ItemNum)

    Case ITEM_TYPE_SPELL
        Call UseItem_Spell(Index, InvNum)

    Case ITEM_TYPE_FOOD
        Call UseItem_Food(Index, InvNum)
        
    Case ITEM_TYPE_PROTECTDROP
        Call UseItem_ProtectDrop(Index, InvNum)
    End Select
    
    ' Evitar verificar todo tipo de item no conjunto pra melhorar o processamento!
    If Item(ItemNum).Type >= ITEM_TYPE_WEAPON And Item(ItemNum).Type <= ITEM_TYPE_RINGRIGHT Then
        Call CheckConjunto(Index)
    End If
End Sub

Public Sub EquipItem(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Dim ItemNum As Long, ItemBound As Long
    Dim tempItem As Long, tempLevel As Long, tempBound As Long

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    If Not IsPlayerItemRequerimentsOK(Index, ItemNum) Then Exit Sub

    ' tell them if it's soulbound
    If GetPlayerInvItemBound(Index, InvNum) = 0 Then
        If Item(ItemNum).BindType = ITEM_BIND_EQUIPED Then
            Call SetPlayerInvItemBound(Index, InvNum, 1)
            Call PlayerMsg(Index, "Este item agora está ligado a sua alma.", BrightRed)
        End If
    End If

    ItemBound = GetPlayerInvItemBound(Index, InvNum)

    ' Se já há algum item equipado, salva os dados para devolver para o inventario.
    If GetPlayerEquipmentNum(Index, EquipmentSlot) > 0 Then
        tempItem = GetPlayerEquipmentNum(Index, EquipmentSlot)
        tempBound = GetPlayerEquipmentBound(Index, EquipmentSlot)

        ' Remove a magia
        If Item(tempItem).GiveSpellNum > 0 Then
            Call RemovePlayerSpell(Index, Item(tempItem).GiveSpellNum)
        End If
    End If

    SetPlayerEquipment Index, ItemNum, EquipmentSlot
    Call SetPlayerEquipmentBound(Index, ItemBound, EquipmentSlot)

    PlayerMsg Index, "Voce equipou " & Trim$(Item(ItemNum).Name), BrightGreen

    If Item(ItemNum).GiveSpellNum > 0 Then
        Call GivePlayerSpell(Index, Item(ItemNum).GiveSpellNum)
    End If

    TakeInvSlot Index, InvNum, 1
    SendInventoryUpdate Index, InvNum

    If tempItem > 0 Then
        GiveInvItem Index, tempItem, 1, tempBound     ' give back the stored item
        tempItem = 0
        tempLevel = 0
        tempBound = 0
    End If

    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)

    ' send vitals
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    Call SendMapVitals(Index)
    Call SendStats(Index)

    ' send vitals to party if in one
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

    ' send the sound
    SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum

End Sub

Private Function IsPlayerItemRequerimentsOK(ByVal Index As Long, ByVal ItemNum As Long) As Boolean
    IsPlayerItemRequerimentsOK = True
    Dim i As Byte

    ' stat requirements
    For i = 1 To Stats.Stat_Count - 1
        If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
            PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
            IsPlayerItemRequerimentsOK = False
        End If
    Next

    ' level requirement
    If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
        PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
        IsPlayerItemRequerimentsOK = False
    End If

    ' class requirement
    If Item(ItemNum).ClassReq > 0 Then
        If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
            PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
            IsPlayerItemRequerimentsOK = False
        End If
    End If

    ' access requirement
    If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
        PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
        IsPlayerItemRequerimentsOK = False
    End If

    ' prociency requirement
    If Not hasProficiency(Index, Item(ItemNum).proficiency) Then
        PlayerMsg Index, "You do not have the proficiency this item requires.", BrightRed
        IsPlayerItemRequerimentsOK = False
    End If

End Function

Public Sub RemovePlayerSpell(ByVal Index As Long, ByVal SpellNum As Long)
    Dim i As Long

    ' Procura a magia no inventario e remove.
    For i = 1 To MAX_PLAYER_SPELLS
        If Player(Index).Spell(i).Spell = SpellNum Then
            Player(Index).Spell(i).Spell = 0
            Player(Index).Spell(i).Uses = 0

            Call PlayerMsg(Index, "A habilidade " & Trim$(Spell(SpellNum).Name) & " foi removida", BrightGreen)
            Call SendPlayerSpells(Index)

            Exit Sub
        End If
    Next
End Sub

Public Function GivePlayerSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long, FreeSlot As Long
    
    GivePlayerSpell = False

    ' Se o usuário já estiver com a magia, atualiza o level.
    For i = 1 To MAX_PLAYER_SPELLS
        If Player(Index).Spell(i).Spell = SpellNum Then
            Call PlayerMsg(Index, "Você já possui essa habilidade", BrightRed)
            GivePlayerSpell = True
            Exit Function
        End If
    Next

    ' Procura por um slot vazio.
    For i = 1 To MAX_PLAYER_SPELLS
        If Player(Index).Spell(i).Spell = 0 Then
            FreeSlot = i
            Exit For
        End If
    Next

    If FreeSlot <> 0 Then
        Player(Index).Spell(FreeSlot).Spell = SpellNum

        Call PlayerMsg(Index, "A habilidade " & Trim$(Spell(SpellNum).Name) & " foi adquirida", BrightGreen)
        Call SendPlayerSpells(Index)
        GivePlayerSpell = True
    Else
        Call PlayerMsg(Index, "Não há espaço suficiente para novas magias", BrightRed)
    End If

End Function

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

Public Sub Unique_Item(ByVal Index As Long, ByVal ItemNum As Long)
    Dim ClassNum As Long, i As Long

    If Not IsPlayerItemRequerimentsOK(Index, ItemNum) Then Exit Sub

    Select Case Item(ItemNum).Data1
    Case 1    ' Reset Stats
        ClassNum = GetPlayerClass(Index)
        If ClassNum <= 0 Or ClassNum > Max_Classes Then Exit Sub
        ' re-set the actual stats to class defaults
        For i = 1 To Stats.Stat_Count - 1
            SetPlayerStat Index, i, Class(ClassNum).Stat(i)
        Next
        ' give player their points back
        SetPlayerPOINTS Index, (GetPlayerLevel(Index) - 1) * 3
        ' take item
        TakeInvItem Index, ItemNum, 1
        ' let them know we've done it
        PlayerMsg Index, "Your stats have been reset.", BrightGreen
        ' send them their new stats
        SendPlayerData Index
    Case Else    ' Exit out otherwise
        Exit Sub
    End Select
End Sub

Public Sub UseItem_Consume(ByVal Index As Long, ByVal InvNum As Byte)
    Dim ItemNum As Integer

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    ' add hp
    If Item(ItemNum).AddHP > 0 Then
        Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + Item(ItemNum).AddHP
        SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        SendVital Index, HP
        Call SendMapVitals(Index)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
    End If
    ' add mp
    If Item(ItemNum).AddMP > 0 Then
        Player(Index).Vital(Vitals.MP) = Player(Index).Vital(Vitals.MP) + Item(ItemNum).AddMP
        SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        SendVital Index, MP
        Call SendMapVitals(Index)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
    End If
    ' add exp
    If Item(ItemNum).AddEXP > 0 Then
        SetPlayerExp Index, GetPlayerExp(Index) + Item(ItemNum).AddEXP
        CheckPlayerLevelUp Index
        SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        SendEXP Index
    End If
    Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
    Call TakeInvItem(Index, Inv(Index).Item(InvNum).Num, 0)

    ' send the sound
    SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
End Sub

Public Sub UseItem_Key(ByVal Index As Long, ByVal InvNum As Long)
    Dim ItemNum As Long
    Dim x As Byte
    Dim Y As Byte

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    If Not IsPlayerItemRequerimentsOK(Index, ItemNum) Then Exit Sub

    Select Case GetPlayerDir(Index)
    Case DIR_UP

        If GetPlayerY(Index) > 0 Then
            x = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 1
        Else
            Exit Sub
        End If

    Case DIR_DOWN

        If GetPlayerY(Index) < Map(GetPlayerMap(Index)).MapData.MaxY Then
            x = GetPlayerX(Index)
            Y = GetPlayerY(Index) + 1
        Else
            Exit Sub
        End If

    Case DIR_LEFT

        If GetPlayerX(Index) > 0 Then
            x = GetPlayerX(Index) - 1
            Y = GetPlayerY(Index)
        Else
            Exit Sub
        End If

    Case DIR_RIGHT

        If GetPlayerX(Index) < Map(GetPlayerMap(Index)).MapData.MaxX Then
            x = GetPlayerX(Index) + 1
            Y = GetPlayerY(Index)
        Else
            Exit Sub
        End If

    End Select

    ' Check if a key exists
    If Map(GetPlayerMap(Index)).TileData.Tile(x, Y).Type = TILE_TYPE_KEY Then

        ' Check if the key they are using matches the map key
        If ItemNum = Map(GetPlayerMap(Index)).TileData.Tile(x, Y).Data1 Then
            TempTile(GetPlayerMap(Index)).DoorOpen(x, Y) = YES
            TempTile(GetPlayerMap(Index)).DoorTimer = getTime
            SendMapKey Index, x, Y, 1
            'Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)

            Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, x, Y)

            ' Check if we are supposed to take away the item
            If Map(GetPlayerMap(Index)).TileData.Tile(x, Y).Data2 = 1 Then
                Call TakeInvSlot(Index, InvNum, 1)
                Call SendInventoryUpdate(Index, InvNum)
                Call PlayerMsg(Index, "A chave destruiu a fechadura.", Yellow)
            End If
        End If
    End If

    ' send the sound
    SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
End Sub

Public Sub UseItem_Food(ByVal Index As Long, ByVal InvNum As Long)
    Dim ItemNum As Long
    Dim x As Long

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    If Not IsPlayerItemRequerimentsOK(Index, ItemNum) Then
        Exit Sub
    End If

    ' make sure they're not in combat
    If TempPlayer(Index).stopRegen Then
        PlayerMsg Index, "Voce não pode comer enquanto estiver em combate.", BrightRed
        Exit Sub
    End If

    ' make sure not full HP
    x = Item(ItemNum).HPorSP
    If Player(Index).Vital(x) >= GetPlayerMaxVital(Index, x) Then
        PlayerMsg Index, "Voce não precisa comer neste momento.", BrightRed
        Exit Sub
    End If

    ' set the player's food
    If Item(ItemNum).HPorSP = 2 Then    'mp
        If Not TempPlayer(Index).foodItem(Vitals.MP) = ItemNum Then
            TempPlayer(Index).foodItem(Vitals.MP) = ItemNum
            TempPlayer(Index).foodTick(Vitals.MP) = 0
            TempPlayer(Index).foodTimer(Vitals.MP) = getTime
        Else
            PlayerMsg Index, "Voce já está comendo.", BrightRed
            Exit Sub
        End If
    Else    ' HP
        If Not TempPlayer(Index).foodItem(Vitals.HP) = ItemNum Then
            TempPlayer(Index).foodItem(Vitals.HP) = ItemNum
            TempPlayer(Index).foodTick(Vitals.HP) = 0
            TempPlayer(Index).foodTimer(Vitals.HP) = getTime
        Else
            PlayerMsg Index, "Voce já está comendo.", BrightRed
            Exit Sub
        End If
    End If

    ' take the item
    Call TakeInvSlot(Index, InvNum, 1)
    Call SendInventoryUpdate(Index, InvNum)

End Sub

Public Sub UseItem_ProtectDrop(ByVal Index As Long, ByVal InvNum As Long)
    Dim ItemNum As Long
    Dim x As Long

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    If Not IsPlayerItemRequerimentsOK(Index, ItemNum) Then
        Exit Sub
    End If

    ' make sure they're not in combat
    If TempPlayer(Index).stopRegen Then
        PlayerMsg Index, "Voce não pode pegar a proteção enquanto estiver em combate.", BrightRed
        Exit Sub
    End If

    If GetPlayerProtectDrop(Index) >= YES Then
        PlayerMsg Index, "Você já possui a proteção!", BrightRed
        Exit Sub
    End If

    ' adiciona a proteção ao jogador
    Call SetPlayerProtectDrop(Index, YES)
    PlayerMsg Index, "Você recebeu proteção divina e não irá dropar os pertences ao morrer!", BrightGreen

    ' take the item
    Call TakeInvSlot(Index, InvNum, 1)
    Call SendInventoryUpdate(Index, InvNum)

End Sub

Public Sub SetPlayerProtectDrop(ByVal Index As Long, ByVal ProtectDrop As Byte)
    If Index <= 0 Or Index > Player_HighIndex Then Exit Sub
    Player(Index).ProtectDrop = ProtectDrop
End Sub

Public Function GetPlayerProtectDrop(ByVal Index As Long) As Byte
    If Index <= 0 Or Index > Player_HighIndex Then Exit Function
    GetPlayerProtectDrop = Player(Index).ProtectDrop
End Function

Public Sub UseItem_Spell(ByVal Index As Long, ByVal InvNum As Long)
    Dim ItemNum As Long, n As Long, i As Long

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    If Not IsPlayerItemRequerimentsOK(Index, ItemNum) Then
        Exit Sub
    End If

    ' Get the spell num
    n = Item(ItemNum).Data1

    If n > 0 Then

        ' Make sure they are the right class
        If Spell(n).ClassReq = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then

            ' make sure they don't already know it
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(Index).Spell(i).Spell > 0 Then
                    If Player(Index).Spell(i).Spell = n Then
                        PlayerMsg Index, "Voce já aprendeu essa habilidade.", BrightRed
                        Exit Sub
                    End If
                End If
            Next

            ' Make sure they are the right level
            i = Spell(n).LevelReq


            If i <= GetPlayerLevel(Index) Then
                i = FindOpenSpellSlot(Index)

                ' Make sure they have an open spell slot
                If i > 0 Then

                    ' Make sure they dont already have the spell
                    If Not HasSpell(Index, n) Then
                        Player(Index).Spell(i).Spell = n
                        Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)

                        ' take item
                        Call TakeInvSlot(Index, InvNum, 1)
                        Call SendInventoryUpdate(Index, InvNum)

                        Call PlayerMsg(Index, "Você agora pode usar " & Trim$(Spell(n).Name) & ".", BrightGreen)
                        SendPlayerSpells Index
                    Else
                        Call PlayerMsg(Index, "Você já tem conhecimento desta habilidade", BrightRed)
                    End If

                Else
                    Call PlayerMsg(Index, "Você não pode mais aprender habilidades.", BrightRed)
                End If

            Else
                Call PlayerMsg(Index, "Você deve estar no level " & i & " para aprender esta habilidade.", BrightRed)
            End If

        Else
            Call PlayerMsg(Index, "Esta habilidade somente pode ser aprendida por " & GetClassName(Spell(n).ClassReq) & ".", BrightRed)
        End If
    End If

    ' send the sound
    SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
End Sub

Function GetPlayerEquipmentNum(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > Player_HighIndex Then Exit Function
    If EquipmentSlot = 0 Then Exit Function

    GetPlayerEquipmentNum = Player(Index).Equipment(EquipmentSlot).Num
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal ItemNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot).Num = ItemNum
End Sub

Sub SetPlayerEquipmentBound(ByVal Index As Long, ByVal Bound As Byte, ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot).Bound = Bound
End Sub

Function GetPlayerEquipmentBound(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte
    GetPlayerEquipmentBound = Player(Index).Equipment(EquipmentSlot).Bound
End Function

