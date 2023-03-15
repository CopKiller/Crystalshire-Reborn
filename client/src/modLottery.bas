Attribute VB_Name = "modLottery"
Option Explicit

Public Const MAX_BETS As Byte = 100

Public LotteryBtnRandom As Boolean
Private MaxRndBets As Byte
Private TmrLotteryRnd As Currency
Private tmpNumero As Byte
Private RndBetsCount As Byte
Private SelectedNum As Byte

Public LotteryInfo As LotteryInfo

Private Type LotteryInfo
    LastWinner As String
    LastNumber As Byte
    LotteryOn As Boolean
    BetOn As Boolean
    LotteryTime As Long
    Min_Bets_Value As Long
    Max_Bets_Value As Long
    Accumulated As Long
End Type


Public Sub CreateWindow_Lottery()
    Dim i As Byte, X As Integer, Y As Integer

    CreateWindow "winLottery", "Apostas em números...", zOrder_Win, 0, 0, 436, 400, Tex_Item(1), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc
    ' Centralise it
    CentraliseWindow WindowCount
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Parchment Info
    CreatePictureBox WindowCount, "picParchmentInfo", 6, 24, 424, 100, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Window
    CreatePictureBox WindowCount, "picMenuHead", 6, 6, 424, 18, , , , True, , , , DesignTypes.desGreen, DesignTypes.desGreen, DesignTypes.desGreen
    CreatePictureBox WindowCount, "picMenuIcon", 0, 0, 32, 32, , , , , Tex_Item(4), Tex_Item(4), Tex_Item(4)
    CreateLabel WindowCount, "lblMenuName", 6, 8, 424, , "Think about luck and you will be lucky", georgiaDec_16, Yellow, Alignment.alignCentre
    'Close Btn
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 22, 9, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnCloseLottery)

    ' Infos
    CreatePictureBox WindowCount, "picShadow_1", 142, 34, 150, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblInfo", 0, 30, 436, , "Informations", rockwellDec_15, Pink, Alignment.alignCentre
    
    CreateLabel WindowCount, "lblLWinner", 40, 40, 200, , "Last Winner: - - -", rockwellDec_15, Yellow, Alignment.alignLeft
    CreateLabel WindowCount, "lblBNumber", 40, 60, 200, , "Last Bet Number: - - -", rockwellDec_15, Yellow, Alignment.alignLeft
    CreateLabel WindowCount, "lblLStatus", 40, 80, 200, , "Lottery Status: - - -", rockwellDec_15, Yellow, Alignment.alignLeft
    CreateLabel WindowCount, "lblNLottery", 40, 100, 200, , "Next Lottery: 00:00:00", rockwellDec_15, Yellow, Alignment.alignLeft
    
    CreateLabel WindowCount, "lblAccumulated", 250, 50, 200, , "Accumulated: $$$", rockwellDec_15, Yellow, Alignment.alignLeft
    CreateLabel WindowCount, "lblMinBid", 250, 70, 200, , "Min Bid: $$$", rockwellDec_15, Yellow, Alignment.alignLeft
    CreateLabel WindowCount, "lblMaxBid", 250, 90, 200, , "Max Bid: $$$", rockwellDec_15, Yellow, Alignment.alignLeft

    ' Lucky Number
    CreatePictureBox WindowCount, "picParchmentLucky", 6, 120, 424, 230, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    CreatePictureBox WindowCount, "picTextBackground", 120, 130, 200, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblBetValue", 0, 126, 436, , "Select Your Lucky Number!", rockwellDec_15, Yellow, Alignment.alignCentre
    
    X = 15
    Y = 145
    For i = 1 To MAX_BETS
        CreateButton WindowCount, "btnNumber" & i, CLng(X), CLng(Y), 27, 27, CLng(i), rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnSelectNumber)
        If i Mod 15 = 0 Then Y = Y + 27: X = 15 Else: X = X + 27
    Next i
    
    ' Btns
    CreateButton WindowCount, "btnSendBet", 12, 353, 134, 32, "Send Bet", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnSendBet)
    CreateButton WindowCount, "btnRandom", 151, 353, 134, 32, "Random", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnLotteryRandom)
    CreateButton WindowCount, "btnCancel", 290, 353, 134, 32, "Exit", rockwellDec_15, White, , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnCloseLottery)
End Sub

Public Sub HandleLotteryInfo(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim SString As String
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    LotteryInfo.LastWinner = Trim$(Buffer.ReadString)
    LotteryInfo.LastNumber = Buffer.ReadByte
    LotteryInfo.LotteryOn = ConvertByteToBool(Buffer.ReadByte)
    LotteryInfo.BetOn = ConvertByteToBool(Buffer.ReadByte)
    LotteryInfo.LotteryTime = Buffer.ReadLong
    LotteryInfo.Accumulated = Buffer.ReadLong
    LotteryInfo.Min_Bets_Value = Buffer.ReadLong
    LotteryInfo.Max_Bets_Value = Buffer.ReadLong
    Set Buffer = Nothing

    With Windows(GetWindowIndex("winLottery"))
        ' Last WINNER
        If LenB(Trim$(LotteryInfo.LastWinner)) > 0 Then
            SString = "Last Winner: " & ColourChar & GetColStr(Cyan) & LotteryInfo.LastWinner
        Else
            SString = "Last Winner: " & ColourChar & GetColStr(Cyan) & "- - -"
        End If
        .Controls(GetControlIndex("winLottery", "lblLWinner")).Text = SString
        
        ' Last NUMBER
        If LotteryInfo.LastNumber > 0 Then
            SString = "Last Number: " & ColourChar & GetColStr(Cyan) & LotteryInfo.LastNumber
        Else
            SString = "Last Number: " & ColourChar & GetColStr(Cyan) & "- - -"
        End If
        .Controls(GetControlIndex("winLottery", "lblBNumber")).Text = SString

        ' Lottery STATUS
        If LotteryInfo.BetOn Then
            SString = "Lottery Status: " & ColourChar & GetColStr(Green) & "Bets ON!!!"
        ElseIf LotteryInfo.LotteryOn Then
            SString = "Lottery Status: " & ColourChar & GetColStr(Pink) & "Wait For Winner!!!"
        Else
            SString = "Lottery Status: " & ColourChar & GetColStr(BrightRed) & "Closed!!!"
        End If
        .Controls(GetControlIndex("winLottery", "lblLStatus")).Text = SString

        ' Next Lottery TIME
        If LotteryInfo.BetOn Or LotteryInfo.LotteryOn Then
            SString = "Next Lottery: " & ColourChar & GetColStr(BrightRed) & "- - -"
        Else
            SString = "Next Lottery: " & ColourChar & GetColStr(BrightRed) & SecondsToHMS(LotteryInfo.LotteryTime)
        End If
        .Controls(GetControlIndex("winLottery", "lblNLottery")).Text = SString

        ' Accumulated & Min Bid & Max Bid
        .Controls(GetControlIndex("winLottery", "lblAccumulated")).Text = "Accumulated: " & ColourChar & GetColStr(BrightGreen) & LotteryInfo.Accumulated & " $$"
        .Controls(GetControlIndex("winLottery", "lblMinBid")).Text = "Min Bid: " & ColourChar & GetColStr(BrightGreen) & LotteryInfo.Min_Bets_Value & " $$"
        .Controls(GetControlIndex("winLottery", "lblMaxBid")).Text = "Max Bid: " & ColourChar & GetColStr(BrightGreen) & LotteryInfo.Max_Bets_Value & " $$"

    End With
End Sub

Private Sub btnSelectNumber()
    Dim i As Byte

    With Windows(GetWindowIndex("winLottery"))
        ' Verifica o controle que vai selecionar
        For i = 1 To MAX_BETS
            If GlobalX >= .Window.Left + .Controls(GetControlIndex("winLottery", "btnNumber" & i)).Left And GlobalX <= .Window.Left + .Controls(GetControlIndex("winLottery", "btnNumber" & i)).Left + .Controls(GetControlIndex("winLottery", "btnNumber" & i)).Width Then
                If GlobalY >= .Window.top + .Controls(GetControlIndex("winLottery", "btnNumber" & i)).top And GlobalY <= .Window.top + .Controls(GetControlIndex("winLottery", "btnNumber" & i)).top + .Controls(GetControlIndex("winLottery", "btnNumber" & i)).Height Then
                    ' Limpa o anterior
                    If SelectedNum > 0 Then
                        .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(0) = DesignTypes.desGreen
                        .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(1) = DesignTypes.desGreen_Hover
                        .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(2) = DesignTypes.desGreen_Click
                    End If

                    SelectedNum = i
                End If
            End If
        Next i
        
        .Controls(GetControlIndex("winLottery", "btnSendBet")).Text = "Send Bet(" & SelectedNum & ")"
        .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(0) = DesignTypes.desOrange
        .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(1) = DesignTypes.desOrange_Hover
        .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(2) = DesignTypes.desOrange_Click
    End With

End Sub

Private Sub btnSendBet()
    
    Dialogue "Bet Number Is: " & SelectedNum, "Value of your Bet:", "", TypeSENDBET, StyleINPUT

End Sub

Public Sub SendBet(ByVal BetValue As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSendBet
    Buffer.WriteByte SelectedNum
    Buffer.WriteLong BetValue
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Private Sub btnLotteryRandom()

' Faz a limpeza antes
    Call ClearLotteryValues

    ' Limpa algum controle que ele havia selecionado
    If SelectedNum > 0 Then
    With Windows(GetWindowIndex("winLottery"))
            .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(0) = DesignTypes.desGreen
            .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(1) = DesignTypes.desGreen_Hover
            .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(2) = DesignTypes.desGreen_Click
    End With
    End If

    ' Não deixa o jogador interagir com mais nada na janela, somente sair apartir daqui.
    Call WindowLotteryModeRandom(False)

    LotteryBtnRandom = True
End Sub

Private Sub WindowLotteryModeRandom(ByVal Activate As Boolean)
    Dim i As Byte
    
    ' Desativa os números pra não causar problemas na randomização
    With Windows(GetWindowIndex("winLottery"))
    For i = 1 To MAX_BETS
        .Controls(GetControlIndex("winLottery", "btnNumber" & i)).enabled = Activate
    Next i
    
    .Controls(GetControlIndex("winLottery", "btnSendBet")).enabled = Activate
    .Controls(GetControlIndex("winLottery", "btnRandom")).enabled = Activate
    End With
End Sub

Public Sub LotteryRand()
    Dim AlternarVelocidade As Long
    Dim MsVel As Long
    Dim Velocidades As Byte

    ' Quantas velocidades de alternancia de números aleatorios vai ter?
    Velocidades = 3

    ' Processa um número aleatorio do 1 ao MAX_BETS
    If MaxRndBets = 0 Then
        MaxRndBets = (MAX_BETS * Rnd) + 1
    End If

    ' Verifica se o númbero é maior que a quantidade de velocidades pra poder dividir entre elas
    If MaxRndBets >= Velocidades Then AlternarVelocidade = MaxRndBets / Velocidades

    ' Trabalha a variavel pra alternar entre as velocidades de randomização
    If RndBetsCount < (AlternarVelocidade * 2) Then
        MsVel = 200    ' Variavel pra alternar entre as velocidades de randomização
    ElseIf RndBetsCount >= (AlternarVelocidade * 2) And RndBetsCount < (AlternarVelocidade * 2.8) Then
        MsVel = 500
    ElseIf RndBetsCount >= (AlternarVelocidade * 2.8) And RndBetsCount <= (AlternarVelocidade * 3) Then
        MsVel = 800
    End If

    With Windows(GetWindowIndex("winLottery"))
        ' Realizar a contagem e verificar se está na hora de mudar a cor do numero pra identificação visual do processo!
        If TmrLotteryRnd <= getTime Then

            TmrLotteryRnd = getTime + MsVel

            RndBetsCount = RndBetsCount + 1

            SelectedNum = Int((Rnd * MAX_BETS) + 1)

            ' Limpa a coloração do controle anterior!
            If tmpNumero <> 0 Then
                .Controls(GetControlIndex("winLottery", "btnNumber" & tmpNumero)).design(0) = DesignTypes.desGreen
                .Controls(GetControlIndex("winLottery", "btnNumber" & tmpNumero)).design(1) = DesignTypes.desGreen_Hover
                .Controls(GetControlIndex("winLottery", "btnNumber" & tmpNumero)).design(2) = DesignTypes.desGreen_Click
            End If

            ' Colorir o novo controle e salvar o número dele temporariamente pra posteriormente limpá-lô!
            .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(0) = DesignTypes.desOrange
            .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(1) = DesignTypes.desOrange_Hover
            .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(2) = DesignTypes.desOrange_Click
            tmpNumero = SelectedNum
        End If

        ' Verifica se a quantidade de vezes que foi processado já chegou ao fim da randomização!
        If MaxRndBets = RndBetsCount Then
            .Controls(GetControlIndex("winLottery", "btnSendBet")).Text = "Send Bet(" & SelectedNum & ")"
            Call ClearLotteryValues
            Call WindowLotteryModeRandom(True)
        End If
    End With
End Sub

Private Sub ClearLotteryValues()
    Dim i As Byte

    LotteryBtnRandom = False
    MaxRndBets = 0
    TmrLotteryRnd = 0
    tmpNumero = 0
    RndBetsCount = 0
End Sub

Public Sub HandleLotteryWindow(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    OpenLottery
End Sub

Private Sub OpenLottery()
    If Windows(GetWindowIndex("winLottery")).Window.visible = False Then
        ShowWindow GetWindowIndex("winLottery")
    End If
End Sub

Private Sub btnCloseLottery()
    If Windows(GetWindowIndex("winLottery")).Window.visible = True Then
        HideWindow GetWindowIndex("winLottery")
        Call ClearLotteryValues
        Call WindowLotteryModeRandom(True)
        Windows(GetWindowIndex("winLottery")).Controls(GetControlIndex("winLottery", "btnSendBet")).Text = "Send Bet"
        If SelectedNum > 0 Then
            ' Limpa algum controle que ele havia selecionado
            With Windows(GetWindowIndex("winLottery"))
                .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(0) = DesignTypes.desGreen
                .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(1) = DesignTypes.desGreen_Hover
                .Controls(GetControlIndex("winLottery", "btnNumber" & SelectedNum)).design(2) = DesignTypes.desGreen_Click
            End With
            SelectedNum = 0
        End If
    End If
End Sub
