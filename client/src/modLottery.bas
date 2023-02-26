Attribute VB_Name = "modLottery"
Option Explicit

Private Const MAX_BETS As Byte = 100

Private Const MIN_BETS_VALUE As Long = 20    ' min bet value
Private Const MAX_BETS_VALUE As Long = 100000


Public Sub CreateWindow_Lottery()
    Dim i As Byte

    CreateWindow "winLottery", "Apostas em números...", zOrder_Win, 0, 0, 436, 214, Tex_Item(1), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc
    ' Centralise it
    CentraliseWindow WindowCount
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Parchment
    'CreatePictureBox WindowCount, "picParchment", 6, 26, 424, 182, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment

    CreatePictureBox WindowCount, "picMenuHead", 6, 6, 424, 18, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen, DesignTypes.desGreen
    CreatePictureBox WindowCount, "picMenuIcon", 0, 0, 32, 32, , , , , Tex_Item(4), Tex_Item(4), Tex_Item(4)
    CreateLabel WindowCount, "lblMenuName", 6, 8, 424, , "Think about luck and you will be lucky", georgiaDec_16, Yellow, Alignment.alignCentre
    
    'Close Btn
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 22, 9, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Quest)

    CreateLabel WindowCount, "lblBetID", 66, 29, 142, , "Número da aposta (1 ao " & MAX_BETS & ")", rockwellDec_15, White, Alignment.alignCentre
    CreateTextbox WindowCount, "txtBetID", 67, 45, 142, 19, , Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3, , , 3

    CreateLabel WindowCount, "lblBetValue", 66, 79, 142, , "Valor da Aposta", rockwellDec_15, White, Alignment.alignCentre
    CreateTextbox WindowCount, "txtBetValue", 67, 95, 142, 19, , Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3, , , 6
    
    CreateButton WindowCount, "btnSendBet", 238, 155, 134, 20, "Enviar Aposta", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf SendBet)

    ' Btns
    CreateButton WindowCount, "btnCancel", 238, 185, 134, 20, "Sair", rockwellDec_15, White, , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf CloseLottery)
End Sub

Private Sub SendBet()
    Dim Buffer As clsBuffer
    Dim BetValue As Long
    Dim BetID As Integer

    With Windows(GetWindowIndex("winLottery"))
        If IsNumeric(.Controls(GetControlIndex("winLottery", "txtBetValue")).Text) And IsNumeric(.Controls(GetControlIndex("winLottery", "txtBetID")).Text) Then
            BetValue = .Controls(GetControlIndex("winLottery", "txtBetValue")).Text
            BetID = CInt(.Controls(GetControlIndex("winLottery", "txtBetID")).Text)

            If BetValue > MAX_BETS_VALUE Then
                AddText "Max " & Format(MAX_BETS_VALUE, "g"), BrightRed, , ChatChannel.chGame
                Exit Sub
            End If

            If BetValue < MIN_BETS_VALUE Then
                AddText "Min " & Format(MIN_BETS_VALUE, "g"), BrightRed, , ChatChannel.chGame
                Exit Sub
            End If

            If BetID <= 0 Or BetID > MAX_BETS Then
                AddText "Bets From 1 to " & MAX_BETS, BrightRed, , ChatChannel.chGame
                Exit Sub
            End If
        Else
            AddText "Only numbers", BrightRed, , ChatChannel.chGame
            Exit Sub
        End If
    End With

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSendBet
    Buffer.WriteByte BetID
    Buffer.WriteLong BetValue
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Public Sub HandleLotteryWindow(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    OpenLottery
End Sub

Private Sub OpenLottery()
    If Windows(GetWindowIndex("winLottery")).Window.visible = False Then
        ShowWindow GetWindowIndex("winLottery")
    End If
End Sub

Private Sub CloseLottery()
    If Windows(GetWindowIndex("winLottery")).Window.visible = True Then
        HideWindow GetWindowIndex("winLottery")
    End If
End Sub
