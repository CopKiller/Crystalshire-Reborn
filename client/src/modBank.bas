Attribute VB_Name = "modBank"
Option Explicit

' Bank constants
Public Const BankTop As Long = 28
Public Const BankLeft As Long = 9
Public Const BankOffsetY As Long = 6
Public Const BankOffsetX As Long = 6
Public Const BankColumns As Long = 10

Public Sub CreateWindow_Bank()
    CreateWindow "winBank", "Banco", zOrder_Win, 0, 0, 393, 412, Tex_Item(1), True, Fonts.rockwellDec_15, , 2, 5, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , GetAddress(AddressOf Bank_MouseMove), GetAddress(AddressOf Bank_MouseDown), GetAddress(AddressOf Bank_MouseMove), 0, , , GetAddress(AddressOf DrawBank)
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Bank)

End Sub

Public Sub btnMenu_Bank()

    HideWindow GetWindowIndex("winBank")
    CloseBank

End Sub

Sub Bank_MouseMove()
    Dim ItemNum As Long, X As Long, Y As Long, i As Long
    Dim SoulBound As Boolean

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

    ItemNum = IsBankItem(Windows(GetWindowIndex("winBank")).Window.Left, Windows(GetWindowIndex("winBank")).Window.top)

    If ItemNum > 0 Then

        ' make sure we're not dragging the item
        If DragBox.Type = Part_Item And DragBox.Value = ItemNum Then Exit Sub
        ' calc position
        X = Windows(GetWindowIndex("winBank")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        Y = Windows(GetWindowIndex("winBank")).Window.top - 4

        ' offscreen?
        If X < 0 Then
            ' switch to right
            X = Windows(GetWindowIndex("winBank")).Window.Left + Windows(GetWindowIndex("winBank")).Window.Width
        End If

        ' go go go
        If Bank.Item(ItemNum).bound > 0 Then: SoulBound = True
        ShowItemDesc X, Y, Bank.Item(ItemNum).num, SoulBound
    End If
End Sub

Sub Bank_MouseDown()
    Dim BankSlot As Long, winIndex As Long, i As Long

    ' is there an item?
    BankSlot = IsBankItem(Windows(GetWindowIndex("winBank")).Window.Left, Windows(GetWindowIndex("winBank")).Window.top)

    If BankSlot > 0 Then
        ' exit out if we're offering that item

        ' drag it
        With DragBox
            .Type = Part_Item
            .Value = Bank.Item(BankSlot).num
            .Origin = origin_Bank
            .Slot = BankSlot
        End With

        winIndex = GetWindowIndex("winDragBox")
        With Windows(winIndex).Window
            .state = MouseDown
            .Left = lastMouseX - 16
            .top = lastMouseY - 16
            .movedX = clickedX - .Left
            .movedY = clickedY - .top
        End With

        ShowWindow winIndex, , False
        ' stop dragging inventory
        Windows(GetWindowIndex("winBank")).Window.state = Normal
    End If

    ' show desc. if needed
    Bank_MouseMove
End Sub

Public Function IsBankItem(startX As Long, startY As Long) As Long
    Dim tempRec As RECT
    Dim i As Long
    For i = 1 To MAX_BANK

        If Bank.Item(i).num > 0 Then

            With tempRec
                .top = startY + BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .bottom = .top + PIC_Y
                .Left = startX + BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If currMouseX >= tempRec.Left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Sub DrawBank()
    Dim X As Long, Y As Long, xO As Long, yO As Long, Width As Long, Height As Long
    Dim i As Long, ItemNum As Long, ItemPic As Long

    Dim Left As Long, top As Long
    Dim Colour As Long, skipItem As Boolean, Amount As Long, tmpItem As Long

    xO = Windows(GetWindowIndex("winBank")).Window.Left
    yO = Windows(GetWindowIndex("winBank")).Window.top
    Width = Windows(GetWindowIndex("winBank")).Window.Width
    Height = Windows(GetWindowIndex("winBank")).Window.Height

    ' render green
    RenderTexture Tex_GUI(34), xO + 4, yO + 23, 0, 0, Width - 8, Height - 27, 4, 4

    Width = 76
    Height = 76

    Y = yO + 23
    ' render grid - row
    For i = 1 To 5
        If i = 4 Then Height = 38
        RenderTexture Tex_GUI(35), xO + 4, Y, 0, 0, 80, 80, 80, 80
        RenderTexture Tex_GUI(35), xO + 80, Y, 0, 0, 80, 80, 80, 80
        RenderTexture Tex_GUI(35), xO + 156, Y, 0, 0, 80, 80, 80, 80
        RenderTexture Tex_GUI(35), xO + 232, Y, 0, 0, 80, 80, 80, 80
        RenderTexture Tex_GUI(35), xO + 308, Y, 0, 0, 80, 80, 80, 80
        Y = Y + 76
    Next

    ' actually draw the icons
    For i = 1 To MAX_BANK
        ItemNum = Bank.Item(i).num

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ' not dragging?
            If Not (DragBox.Origin = origin_Bank And DragBox.Slot = i) Then
                ItemPic = Item(ItemNum).Pic

                If ItemPic > 0 And ItemPic <= Count_Item Then
                    top = yO + BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    Left = xO + BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))

                    ' draw icon
                    RenderTexture Tex_Item(ItemPic), Left, top, 0, 0, 32, 32, 32, 32

                    ' If item is a stack - draw the amount you have
                    If Bank.Item(i).Value > 1 Then
                        Y = top + 21
                        X = Left + 1
                        Amount = Bank.Item(i).Value

                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Colour = BrightGreen
                        End If

                        RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), X, Y, Colour
                    End If
                End If
            End If
        End If
    Next

End Sub

Public Sub HandleBank(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_BANK
        Bank.Item(i).num = Buffer.ReadInteger
        Bank.Item(i).Value = Buffer.ReadLong
    Next

    Buffer.Flush: Set Buffer = Nothing

    InBank = True

    ShowWindow GetWindowIndex("winBank")
End Sub

Sub HandlePlayerBankUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    n = Buffer.ReadByte

    Bank.Item(n).num = Buffer.ReadInteger
    Bank.Item(n).Value = Buffer.ReadLong
    Bank.Item(n).bound = Buffer.ReadByte

    Buffer.Flush
    Set Buffer = Nothing
End Sub

Public Sub DepositItem(ByVal InvSlot As Long, ByVal Amount As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDepositItem
    Buffer.WriteLong InvSlot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Buffer.Flush
    Set Buffer = Nothing
End Sub

Public Sub WithdrawItem(ByVal BankSlot As Long, ByVal Amount As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWithdrawItem
    Buffer.WriteLong BankSlot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Buffer.Flush
    Set Buffer = Nothing
End Sub

Public Sub CloseBank()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseBank
    SendData Buffer.ToArray()
    Buffer.Flush
    Set Buffer = Nothing

    InBank = False
End Sub

Public Sub ChangeBankSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CChangeBankSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Buffer.Flush
    Set Buffer = Nothing

    PlayerSwitchBankSlots OldSlot, NewSlot
End Sub

Sub PlayerSwitchBankSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim OldNum As Long, OldValue As Long, oldBound As Byte
    Dim NewNum As Long, NewValue As Long, newBound As Byte

    If OldSlot = 0 Or NewSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetBankItemNum(OldSlot)
    OldValue = GetBankItemValue(OldSlot)
    oldBound = Bank.Item(OldSlot).bound
    NewNum = GetBankItemNum(NewSlot)
    NewValue = GetBankItemValue(NewSlot)
    newBound = Bank.Item(NewSlot).bound

    SetBankItemNum NewSlot, OldNum
    SetBankItemValue NewSlot, OldValue
    Bank.Item(NewSlot).bound = oldBound

    SetBankItemNum OldSlot, NewNum
    SetBankItemValue OldSlot, NewValue
    Bank.Item(OldSlot).bound = newBound
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

Public Sub SetBankItemNum(ByVal BankSlot As Long, ByVal ItemNum As Long)
    Bank.Item(BankSlot).num = ItemNum
End Sub

Public Function GetBankItemValue(ByVal BankSlot As Long) As Long
    GetBankItemValue = Bank.Item(BankSlot).Value
End Function

Public Sub SetBankItemValue(ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank.Item(BankSlot).Value = ItemValue
End Sub
