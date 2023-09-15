Attribute VB_Name = "modPlayerMapItem"
Option Explicit

Sub PlayerMapGetItem(ByVal Index As Long)
    Dim i As Long
    Dim MapNum As Long
    Dim Msg As String

    If Not IsPlaying(Index) Then Exit Sub
    MapNum = GetPlayerMap(Index)

    If MapNum = 0 Then Exit Sub

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(Index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(MapNum, i).X = GetPlayerX(Index)) Then
                    If (MapItem(MapNum, i).Y = GetPlayerY(Index)) Then

                        ' Set item in players inventor
                        If GiveInvItem(Index, MapItem(MapNum, i).Num, MapItem(MapNum, i).value, MapItem(MapNum, i).Bound) Then
                            Msg = MapItem(MapNum, i).value & "x " & Trim$(Item(MapItem(MapNum, i).Num).Name)
                            
                            ' check tasks
                            Call CheckTasks(Index, QUEST_TYPE_GOGATHER, MapItem(MapNum, i).Num)
                            SendActionMsg GetPlayerMap(Index), Msg, GetItemNameColour(Item(MapItem(MapNum, i).Num).Rarity), 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)

                            ' is it bind on pickup?
                            ' Erase item from the map
                            ClearMapItem i, MapNum

                            ' Call SendInventoryUpdate(Index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), 0, 0)
                            Exit For
                        Else
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
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

Function CanPlayerPickupItem(ByVal Index As Long, ByVal mapItemNum As Long)
    Dim MapNum As Long, tmpIndex As Long, i As Long

    MapNum = GetPlayerMap(Index)

    ' no lock or locked to player?
    If MapItem(MapNum, mapItemNum).PlayerName = vbNullString Or MapItem(MapNum, mapItemNum).PlayerName = Trim$(GetPlayerName(Index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If

    ' if in party show their party member's drops
    If TempPlayer(Index).inParty > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            tmpIndex = Party(TempPlayer(Index).inParty).Member(i)
            If tmpIndex > 0 Then
                If Trim$(GetPlayerName(tmpIndex)) = MapItem(MapNum, mapItemNum).PlayerName Then
                    If MapItem(MapNum, mapItemNum).Bound = 0 Then
                        CanPlayerPickupItem = True
                        Exit Function
                    End If
                End If
            End If
        Next
    End If

    ' exit out
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim i As Long
    Dim MapNum As Long
    Dim ItemNum As Long

    ' Check for subscript out of range
    If InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If

    If Amount < 1 Then
        Exit Sub
    End If

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    If Item(ItemNum).Stackable > 0 Then
        ' Check if its more then they have and if so drop it all
        If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
            Amount = GetPlayerInvItemValue(Index, InvNum)
        End If
    Else
        Amount = 1
    End If

    MapNum = GetPlayerMap(Index)

    ' check the player isn't doing something
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Or TempPlayer(Index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then

            ' make sure it's not bound
            If GetPlayerInvItemBound(Index, InvNum) > 0 Then
                PlayerMsg Index, "Este item esta ligado à alma e não pode ser derrubado.", BrightRed
                Exit Sub
            End If

            i = FindOpenMapItemSlot(GetPlayerMap(Index))

            If i <> 0 Then
                MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).Bound = GetPlayerInvItemBound(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), i).Y = GetPlayerY(Index)
                MapItem(GetPlayerMap(Index), i).PlayerName = Trim$(GetPlayerName(Index))
                MapItem(GetPlayerMap(Index), i).playerTimer = getTime + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(Index), i).canDespawn = True
                MapItem(GetPlayerMap(Index), i).despawnTimer = getTime + ITEM_DESPAWN_TIME

                If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable > 0 Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                        Amount = GetPlayerInvItemValue(Index, InvNum)

                        MapItem(GetPlayerMap(Index), i).value = Amount
                        Call SetPlayerInvItemNum(Index, InvNum, 0)
                        Call SetPlayerInvItemValue(Index, InvNum, 0)
                        Call SetPlayerInvItemBound(Index, InvNum, 0)
                    Else
                        MapItem(GetPlayerMap(Index), i).value = Amount
                        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                    End If

                Else
                    MapItem(GetPlayerMap(Index), i).value = 1
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemBound(Index, InvNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(Index, InvNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(MapNum, i).Num, Amount, MapNum, GetPlayerX(Index), GetPlayerY(Index), Trim$(GetPlayerName(Index)), MapItem(MapNum, i).canDespawn, MapItem(MapNum, i).Bound)
            Else
                Call PlayerMsg(Index, "Há muitos itens no chão.", BrightRed)
            End If
        End If
    End If

End Sub

Sub DropItemOnDead(ByVal Index As Long, ByVal ItemNum As Long, ByVal Amount As Long, Optional ByVal IsEquipped As Boolean = False)
    Dim i As Long
    Dim MapNum As Long
    Dim InvNum As Long
    Dim tradeTarget As Long
    Dim Bound As Byte

    ' Check for subscript out of range
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    If Amount < 1 Then
        Exit Sub
    End If

    ' Verify bank, shop and trade
    If TempPlayer(Index).InBank Then
        TempPlayer(Index).InBank = False
    End If
    If TempPlayer(Index).InShop Then
        TempPlayer(Index).InShop = 0
    End If
    If TempPlayer(Index).InTrade > 0 Then
        tradeTarget = TempPlayer(Index).InTrade
        ' cancel any trade they're in
        PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has declined the trade.", BrightRed
        PlayerMsg Index, Trim$(GetPlayerName(tradeTarget)) & " has declined the trade.", BrightRed
        ' clear out trade
        For i = 1 To MAX_INV
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).value = 0
            TempPlayer(Index).TradeOffer(i).Num = 0
            TempPlayer(Index).TradeOffer(i).value = 0
        Next
        TempPlayer(tradeTarget).TradeGold = 0
        TempPlayer(Index).TradeGold = 0
        TempPlayer(tradeTarget).InTrade = 0
        TempPlayer(Index).InTrade = 0
        SendCloseTrade tradeTarget
        SendCloseTrade Index
    End If

    ' Realiza as verificações entre item equipado ou se está na bolsa!
    If Not IsEquipped Then ' Ação na mochila
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                InvNum = i
            End If
        Next i

        Bound = GetPlayerInvItemBound(Index, InvNum)

        ' Retirar o item do jogador e mandar a mensagem!
        Call PlayerMsg(Index, "Você perdeu " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " com " & Item(GetPlayerInvItemNum(Index, InvNum)).DropDeadChance & "% de chance de drop!", BrightRed)
        Call SetPlayerInvItemNum(Index, InvNum, 0)
        Call SetPlayerInvItemValue(Index, InvNum, 0)
        Call SetPlayerInvItemBound(Index, InvNum, 0)
        Call SendInventoryUpdate(Index, InvNum)
    Else ' Ação nos items equipados
        For i = 1 To Equipment.Equipment_Count - 1
            If GetPlayerEquipmentNum(Index, i) = ItemNum Then
                InvNum = i
            End If
        Next i

        Bound = GetPlayerEquipmentBound(Index, InvNum)

        Call PlayerMsg(Index, "Você perdeu " & Trim$(Item(GetPlayerEquipmentNum(Index, InvNum)).Name) & " com " & Item(GetPlayerEquipmentNum(Index, InvNum)).DropDeadChance & "% de chance de drop!", BrightRed)
        SetPlayerEquipment Index, 0, InvNum
        SetPlayerEquipmentBound Index, 0, InvNum
    End If

    MapNum = GetPlayerMap(Index)

    i = FindOpenMapItemSlot(MapNum)

    If i <> 0 Then
        MapItem(MapNum, i).Num = ItemNum
        MapItem(MapNum, i).value = Amount
        MapItem(MapNum, i).Bound = Bound
        MapItem(MapNum, i).X = GetPlayerX(Index)
        MapItem(MapNum, i).Y = GetPlayerY(Index)
        MapItem(MapNum, i).PlayerName = Trim$(GetPlayerName(Index))
        MapItem(MapNum, i).playerTimer = getTime + ITEM_SPAWN_TIME
        MapItem(MapNum, i).canDespawn = True
        MapItem(MapNum, i).despawnTimer = getTime + ITEM_DESPAWN_TIME
        ' Spawn the item before we set the num or we'll get a different free map item slot
        Call SpawnItemSlot(i, MapItem(MapNum, i).Num, MapItem(MapNum, i).value, MapNum, GetPlayerX(Index), GetPlayerY(Index), Trim$(GetPlayerName(Index)), MapItem(MapNum, i).canDespawn, MapItem(MapNum, i).Bound)
    End If



End Sub
