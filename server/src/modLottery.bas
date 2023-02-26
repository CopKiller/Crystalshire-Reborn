Attribute VB_Name = "modLottery"
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private Const MAX_BETS As Byte = 100

Private Const MIN_BETS_VALUE As Long = 20    ' min bet value
Private Const MAX_BETS_VALUE As Long = 100000    ' max bet value
Public Const LOTTERY_START_HOURS As Byte = 3    ' start in..
Private Const LOTTERY_SECS_DURATION As Long = 300    ' 30 mins duration
Public Const LOTTERY_TIME_BET As Long = 180    'starting time of bets / 3 mins

Private Const MAX_AVISOS As Byte = 3

Public Lottery As LotteryStruct

Private Type BetStruct
    Owner As String * ACCOUNT_LENGTH
    Value As Long
End Type

Private Type LotteryStruct
    Enabled As Boolean
    Started As Currency
    Aviso(1 To MAX_AVISOS) As Boolean
    Ended As Currency
    BetEnabled As Boolean ' Are bets open?
    BetTmr As Currency    ' Comparison time to send the notices
    Bet(1 To MAX_BETS) As BetStruct
End Type

Public Sub AddBet(ByVal Index As Long, ByRef BetID As Byte, ByRef BetValue As Long)
Dim I As Byte

    ' Verify if player have the value and Remove the value
    If GetPlayerGold(Index) >= BetValue Then
        Call SetPlayerGold(Index, GetPlayerGold(Index) - BetValue)
        Call SendGoldUpdate(Index)
    Else
        Call PlayerMsg(Index, "Lottery: You do not own the amount you are trying to bet", BrightRed)
        Exit Sub
    End If

    ' Its OK, GO BET!
    AddBetValue Index, BetID, BetValue
End Sub

Private Sub AddBetValue(ByVal Index As Long, ByVal BetID As Byte, ByVal BetValue As Long)
    
    'Lottery On?
    If Not VerifyBetStatus Then
        Call PlayerMsg(Index, "Lottery: The betting period is closed", BrightRed)
        Exit Sub
    End If

    ' Verify Bet Slot Null
    If Not CheckBetSlot(BetID) Then
        Call PlayerMsg(Index, "Lottery: This number already has a bet, choose another", BrightRed)
        Exit Sub
    End If

    ' Verify if Bet Value have minium value
    If BetValue < MIN_BETS_VALUE Then
        Call PlayerMsg(Index, "Lottery: The minimum bet amount is " & MIN_BETS_VALUE, BrightRed)
        Exit Sub
    End If
    
    ' Verify if Bet Value reached the max value
    If BetValue > MAX_BETS_VALUE Then
        Call PlayerMsg(Index, "Lottery: The maximum bet amount is " & MAX_BETS_VALUE, BrightRed)
        Exit Sub
    End If
    
    Call GlobalMsg("Lottery: " & GetPlayerName(Index) & " bet on the number " & BetID & " value " & BetValue, Yellow)
    
    Lottery.Bet(BetID).Owner = GetPlayerName(Index)
    Lottery.Bet(BetID).Value = Lottery.Bet(BetID).Value + BetValue
End Sub

Public Sub CheckBetLoop()
    Dim Number As Byte
    Dim PlayerID As Integer
    Dim I As Byte
    Dim Accumulated As Long, Tmr As Currency

    If VerifyLotteryStatus Then    'On?
        ' Avisos - 1 (last is diferent message!)
        For I = 1 To MAX_AVISOS - 1
            If Not Lottery.Aviso(I) Then
            Debug.Print "BetTmr " & Lottery.BetTmr + ((LOTTERY_TIME_BET / MAX_AVISOS) * 1000) & " <= " & getTime & " GetTime"
                If Lottery.BetTmr + ((LOTTERY_TIME_BET / MAX_AVISOS) * 1000) <= getTime Then
                    Lottery.Aviso(I) = True
                    Lottery.BetTmr = getTime
                    Call GlobalMsg("Lottery: bets close in " & SecondsToHMS(LOTTERY_TIME_BET - ((getTime - Lottery.Started) / 1000)), Yellow)
                End If
            End If
        Next I

        ' Last Aviso
        If Not Lottery.Aviso(MAX_AVISOS) Then
            If Lottery.Started + (LOTTERY_TIME_BET * 1000) <= getTime Then
                Call GlobalMsg("Lottery: Bets closed, Good Luck!!!", Green)
                Call CloseBets
                Lottery.Aviso(MAX_AVISOS) = True
            End If
        End If

        If ((getTime - Lottery.Started) / 1000) > LOTTERY_SECS_DURATION Then    ' Time End?
            Number = ChooseLoteryNumber
            Accumulated = GetBetsAccumulated

            Call GlobalMsg("Lottery: Drawn Number is " & Number, Yellow)

            If LenB(Trim$(Lottery.Bet(Number).Owner)) > 0 Then
                PlayerID = FindPlayer(Trim$(Lottery.Bet(Number).Owner))
                If PlayerID > 0 Then
                    Call SetPlayerGold(PlayerID, Accumulated)
                    Call SendGoldUpdate(PlayerID)
                    Call GlobalMsg("Lottery: The winner is " & Trim$(Lottery.Bet(Number).Owner) & " Congratulations!!!", Green)
                End If

                Call ClearBets    ' Remove all apostas e all owners
            Else
                Call GlobalMsg("Lottery: There were no winners, jackpot in " & GetBetsAccumulated, Green)
            End If

            Call GlobalMsg("Lottery: The lottery has ended, next lottery starts in " & LOTTERY_START_HOURS & " hours", BrightRed)
            Call ClearBetsOwners    ' Remove apenas os donos das apostas, e deixa livre pra próxima loteria, apostarem novamente nos números anteriores
            Call ClearLottery
        End If
    Else    ' Get Off?? take on
        Tmr = LOTTERY_START_HOURS ' 3 hrs
        Tmr = (Tmr * 60) ' 180 min
        Tmr = (Tmr * 60) ' 10.800 segs
        Tmr = (Tmr * 1000) ' 10.800.000 Milisegundos
        Tmr = (Tmr + Lottery.Ended) ' Soma o tempo que a loteria acabou com o tempo dela abrir novamente.
        'Debug.Print Tmr
        If Tmr <= getTime Then
            Call StartLottery
        End If
    End If
End Sub

Private Sub CloseBets()
    Lottery.BetEnabled = False
    Lottery.BetTmr = 0
End Sub

Public Sub StartLottery()
    Dim Accumulated As Long
    
    Lottery.Enabled = True
    Lottery.BetEnabled = True
    
    Lottery.Started = getTime
    Lottery.BetTmr = getTime

    Call GlobalMsg("Lottery: Betting is on, place your bets in (" & SecondsToHMS(LOTTERY_TIME_BET) & ")", BrightGreen)
    
    Accumulated = GetBetsAccumulated
    If Accumulated > 0 Then
        Call GlobalMsg("Lottery: The prize is accumulated in " & Accumulated, BrightGreen)
    End If
End Sub

Public Sub HandleBet(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, BetValue As Long, BetID As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    BetID = Buffer.ReadByte
    BetValue = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing

    Call AddBet(Index, BetID, BetValue)
End Sub


Public Sub SendLotteryWindow(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLotteryWindow
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Function GetBetsAccumulated() As Long
    Dim I As Byte

    For I = 1 To MAX_BETS
        If Lottery.Bet(I).Value > 0 Then
            GetBetsAccumulated = GetBetsAccumulated + Lottery.Bet(I).Value
        End If
    Next I
End Function

Private Function CheckBetSlot(ByRef BetID As Byte) As Boolean
' Prevent Subscript out of range
    If BetID > MAX_BETS Or BetID <= 0 Then Exit Function
    
    CheckBetSlot = True

    If LenB(Trim$(Lottery.Bet(BetID).Owner)) > 0 Then
        CheckBetSlot = False
    End If

End Function

Private Function ChooseLoteryNumber() As Byte
    ChooseLoteryNumber = Random(1, CLng(MAX_BETS))
End Function

Public Function VerifyLotteryStatus() As Boolean
    VerifyLotteryStatus = Lottery.Enabled
End Function

Private Function VerifyBetStatus() As Boolean
    VerifyBetStatus = Lottery.BetEnabled
End Function

Public Sub ClearBets()
    Dim I As Byte
    For I = 1 To MAX_BETS
        ClearBetSlot I
    Next I
End Sub

Private Sub ClearBetSlot(ByRef BetID As Byte)
    Call ZeroMemory(ByVal VarPtr(Lottery.Bet(BetID)), LenB(Lottery.Bet(BetID)))
    Lottery.Bet(BetID).Owner = vbNullString
End Sub

Private Sub ClearBetsOwners()
    Dim I As Byte
    For I = 1 To MAX_BETS
        Lottery.Bet(I).Owner = vbNullString
    Next I
End Sub

Public Sub ClearLottery()
    Dim I As Byte
    
    Lottery.Enabled = False
    Lottery.BetEnabled = False
    
    Lottery.Ended = getTime
    Lottery.Started = 0
    Lottery.BetTmr = 0
    
    For I = 1 To MAX_AVISOS
        Lottery.Aviso(I) = False
    Next I
End Sub
