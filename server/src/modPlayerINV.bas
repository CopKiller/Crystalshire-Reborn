Attribute VB_Name = "modPlayerINV"
Option Explicit

'For Clear functions
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim OldNum As Long, OldValue As Long, OldBound As Byte
    Dim NewNum As Long, NewValue As Long, NewBound As Byte
    Dim SameItem As Long, SwapItem As Boolean

    SwapItem = True

    If NewSlot <= 0 Or NewSlot > MAX_INV Then
        Exit Sub
    End If

    If OldSlot <= 0 Or OldSlot > MAX_INV Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(Index, OldSlot)
    OldValue = GetPlayerInvItemValue(Index, OldSlot)
    OldBound = GetPlayerInvItemBound(Index, OldSlot)

    NewNum = GetPlayerInvItemNum(Index, NewSlot)
    NewValue = GetPlayerInvItemValue(Index, NewSlot)
    NewBound = GetPlayerInvItemBound(Index, NewSlot)

    If OldNum = NewNum Then
        SameItem = NewNum
    End If

    If SameItem > 0 Then
        If Item(SameItem).Stackable > 0 Then
            Call SetPlayerInvItemValue(Index, NewSlot, GetPlayerInvItemValue(Index, NewSlot) + OldValue)
            Call TakeInvSlot(Index, OldSlot, OldValue)

            SwapItem = False
        Else
            SwapItem = True
        End If
    End If

    If SwapItem Then
        Call SetPlayerInvItemNum(Index, NewSlot, OldNum)
        Call SetPlayerInvItemValue(Index, NewSlot, OldValue)
        Call SetPlayerInvItemBound(Index, NewSlot, OldBound)

        Call SetPlayerInvItemNum(Index, OldSlot, NewNum)
        Call SetPlayerInvItemValue(Index, OldSlot, NewValue)
        Call SetPlayerInvItemBound(Index, OldSlot, NewBound)
    End If

    SendInventory Index
End Sub

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        HasItem = 0
        Exit Function
    End If

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then

            If Item(ItemNum).Stackable > 0 Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function

        End If
    Next

End Function

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(ItemNum).Stackable > 0 Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long

    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Stackable > 0 Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                Call SetPlayerInvItemBound(Index, i, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Function
            End If
        End If

    Next

End Function

Public Sub TakeInvSlot(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemVal As Long)
    Dim ItemNum As Long

    Dim TakeInvSlot As Boolean
    TakeInvSlot = False

    ' Check for subscript out of range
    If InvSlot <= 0 Or InvSlot > MAX_INV Then
        Exit Sub
    End If

    ItemNum = GetPlayerInvItemNum(Index, InvSlot)

    If Item(ItemNum).Stackable > 0 Then
        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(Index, InvSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(Index, InvSlot, GetPlayerInvItemValue(Index, InvSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(Index, InvSlot, 0)
        Call SetPlayerInvItemValue(Index, InvSlot, 0)
        Call SetPlayerInvItemBound(Index, InvSlot, 0)
    End If

End Sub

Function GiveInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemValue As Long, ByVal ItemBound As Byte, _
                     Optional ByVal sendUpdate As Boolean = True, _
                     Optional ByPending As Boolean = False) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    ' Is Gold? Process first!
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        If Item(ItemNum).price > 0 Then
            If Item(ItemNum).Stackable > 0 Then
                Call SetPlayerGold(Index, GetPlayerGold(Index) + (ItemValue * Item(ItemNum).price))
            Else
                Call SetPlayerGold(Index, GetPlayerGold(Index) + Item(ItemNum).price)
            End If

            Call SendGoldUpdate(Index)
        End If
        GiveInvItem = True
        Exit Function
    End If

    i = FindOpenInvSlot(Index, ItemNum)

    ' Check to see if inventory is full
    If i <> 0 Then

        If ItemBound = 0 And Item(ItemNum).BindType = ITEM_BIND_OBTAINED Then
            ItemBound = 1
        End If

        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemBound(Index, i, ItemBound)

        If Item(ItemNum).Stackable > 0 Then
            Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemValue)
        Else
            Call SetPlayerInvItemValue(Index, i, ItemValue)
        End If

        If sendUpdate Then Call SendInventoryUpdate(Index, i)

        GiveInvItem = True
    Else
        If Not ByPending Then
            Call PlayerMsg(Index, "O inventário esta cheio.", BrightRed)
            GiveInvItem = False
        Else
            If GiveBankItem(Index, ItemNum, ItemValue, ByPending) Then
                Call PlayerMsg(Index, "O inventário e banco esta cheio.", BrightRed)
            Else
                GiveInvItem = False
            End If
        End If
    End If

End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If Index > Player_HighIndex Then Exit Function
    If InvSlot = 0 Then Exit Function

    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemBound(ByVal Index As Long, ByVal InvSlot As Long) As Byte
    If Index > Player_HighIndex Then Exit Function

    If InvSlot <= 0 Or InvSlot > MAX_INV Then
        Exit Function
    End If

    GetPlayerInvItemBound = Player(Index).Inv(InvSlot).Bound
End Function

Sub SetPlayerInvItemBound(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemBound As Byte)
    Player(Index).Inv(InvSlot).Bound = ItemBound
End Sub
