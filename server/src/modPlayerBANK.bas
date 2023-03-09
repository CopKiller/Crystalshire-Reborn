Attribute VB_Name = "modPlayerBANK"
Option Explicit

'For Clear functions
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Sub HandleChangeBankSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim NewSlot As Long
    Dim OldSlot As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    OldSlot = Buffer.ReadLong
    NewSlot = Buffer.ReadLong

    PlayerSwitchBankSlots Index, OldSlot, NewSlot

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim BankSlot As Long
    Dim Amount As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    BankSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong

    TakeBankItem Index, BankSlot, Amount

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim InvSlot As Long
    Dim Amount As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    InvSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong

    GiveBankItem Index, InvSlot, Amount

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    SavePlayer Index

    TempPlayer(Index).InBank = False

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendBankUpdate(ByVal Index As Long, ByVal BankSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SPlayerBankUpdate
    Buffer.WriteByte BankSlot
    Buffer.WriteInteger GetPlayerBankItemNum(Index, BankSlot)
    Buffer.WriteLong GetPlayerBankItemValue(Index, BankSlot)
    Buffer.WriteByte GetPlayerBankItemBound(Index, BankSlot)

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub


Sub PlayerSwitchBankSlots(ByVal Index As Long, ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim OldNum As Long, OldValue As Long, OldBound As Byte
    Dim NewNum As Long, NewValue As Long, NewBound As Byte
    Dim SameItem As Long, SwapItem As Boolean

    SwapItem = True

    If NewSlot <= 0 Or NewSlot > MAX_BANK Then
        Exit Sub
    End If

    If OldSlot <= 0 Or OldSlot > MAX_BANK Then
        Exit Sub
    End If

    OldNum = GetPlayerBankItemNum(Index, OldSlot)
    OldValue = GetPlayerBankItemValue(Index, OldSlot)
    OldBound = GetPlayerBankItemBound(Index, OldSlot)

    NewNum = GetPlayerBankItemNum(Index, NewSlot)
    NewValue = GetPlayerBankItemValue(Index, NewSlot)
    NewBound = GetPlayerBankItemBound(Index, NewSlot)

    If OldNum = NewNum Then
        SameItem = NewNum
    End If

    If SameItem > 0 Then
        If Item(SameItem).Stackable > 0 Then
            Call SetPlayerBankItemValue(Index, NewSlot, GetPlayerBankItemValue(Index, NewSlot) + OldValue)

            Call SetPlayerBankItemNum(Index, OldSlot, 0)
            Call SetPlayerBankItemValue(Index, OldSlot, 0)
            Call SetPlayerBankItemBound(Index, OldSlot, 0)

            SwapItem = False
        Else
            SwapItem = True
        End If
    End If

    If SwapItem Then
        Call SetPlayerBankItemNum(Index, NewSlot, OldNum)
        Call SetPlayerBankItemValue(Index, NewSlot, OldValue)
        Call SetPlayerBankItemBound(Index, NewSlot, OldBound)

        Call SetPlayerBankItemNum(Index, OldSlot, NewNum)
        Call SetPlayerBankItemValue(Index, OldSlot, NewValue)
        Call SetPlayerBankItemBound(Index, OldSlot, NewBound)
    End If

    Call SendBankUpdate(Index, OldSlot)
    Call SendBankUpdate(Index, NewSlot)
End Sub

Function FindOpenBankSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim I As Long

    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(ItemNum).Stackable > 0 Then
        For I = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, I) = ItemNum Then
                FindOpenBankSlot = I
                Exit Function
            End If
        Next I
    End If

    For I = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, I) = 0 Then
            FindOpenBankSlot = I
            Exit Function
        End If
    Next I

End Function

Sub SendBank(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim I As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank

    For I = 1 To MAX_BANK
        Buffer.WriteInteger Player(Index).Bank(I).Num
        Buffer.WriteLong Player(Index).Bank(I).Value
    Next

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Function GiveBankItem(ByVal Index As Long, ByVal InvSlot As Long, ByVal Amount As Long) As Boolean
    Dim BankSlot As Long, ItemNum As Long

    If InvSlot <= 0 Or InvSlot > MAX_INV Then
        Exit Function
    End If

    ItemNum = GetPlayerInvItemNum(Index, InvSlot)

    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' If there's nothing to get, exit
    If Amount < 1 Then Exit Function

    ' check values
    If Item(ItemNum).Stackable > 0 Then
        ' If value is more than in inventory, set to real value
        If Amount > GetPlayerInvItemValue(Index, InvSlot) Then
            Amount = GetPlayerInvItemValue(Index, InvSlot)
        End If
    Else
        Amount = 1
    End If

    BankSlot = FindOpenBankSlot(Index, ItemNum)

    If BankSlot <> 0 Then

        Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, InvSlot))
        Call SetPlayerBankItemBound(Index, BankSlot, GetPlayerInvItemBound(Index, InvSlot))
        Call SetPlayerBankItemValue(Index, BankSlot, Amount)

        Call TakeInvSlot(Index, InvSlot, Amount)

        SendBankUpdate Index, BankSlot
        SendInventoryUpdate Index, InvSlot
    Else
        Call PlayerMsg(Index, "O banco esta cheio", BrightRed)
    End If

End Function

Sub TakeBankItem(ByVal Index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
    Dim ItemNum As Long, ItemBound As Byte

    If BankSlot <= 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If

    ItemNum = GetPlayerBankItemNum(Index, BankSlot)

    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' If there's nothing to get, exit
    If Amount < 1 Then Exit Sub

    ' check values
    If Item(ItemNum).Stackable > 0 Then
        ' If value is more than in inventory, set to real value
        If Amount > GetPlayerBankItemValue(Index, BankSlot) Then
            Amount = GetPlayerBankItemValue(Index, BankSlot)
        End If
    Else
        Amount = 1
    End If

    ItemBound = GetPlayerBankItemBound(Index, BankSlot)

    If GiveInvItem(Index, ItemNum, Amount, ItemBound, True) Then

        If Amount >= GetPlayerBankItemValue(Index, BankSlot) Then
            Call SetPlayerBankItemNum(Index, BankSlot, 0)
            Call SetPlayerBankItemBound(Index, BankSlot, 0)
            Call SetPlayerBankItemValue(Index, BankSlot, 0)
        Else
            Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - Amount)
        End If

        SendBankUpdate Index, BankSlot
    End If

End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    If BankSlot = 0 Then Exit Function
    GetPlayerBankItemNum = Player(Index).Bank(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    If BankSlot = 0 Then Exit Sub
    Player(Index).Bank(BankSlot).Num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    If BankSlot = 0 Then Exit Function
    GetPlayerBankItemValue = Player(Index).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    If BankSlot = 0 Then Exit Sub
    Player(Index).Bank(BankSlot).Value = ItemValue
End Sub

Function GetPlayerBankItemBound(ByVal Index As Long, ByVal BankSlot As Long) As Byte
    If BankSlot <= 0 Or BankSlot > MAX_BANK Then
        Exit Function
    End If

    GetPlayerBankItemBound = Player(Index).Bank(BankSlot).Bound
End Function

Sub SetPlayerBankItemBound(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemBound As Byte)
    If BankSlot <= 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If

   Player(Index).Bank(BankSlot).Bound = ItemBound
End Sub
