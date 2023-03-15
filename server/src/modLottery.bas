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
    Acumulado As Long
    LastBetNum As Byte ' Last Bet number
    LastBetWinner As String * ACCOUNT_LENGTH ' Last Winner Name
End Type

Private Sub AddBetValue(ByVal Index As Long, ByVal BetID As Byte, ByVal BetValue As Long)

    'Lottery On?
    If Not VerifyBetStatus Then
        Call AlertMsg(Index, DIALOGUE_LOTTERY_CLOSED, , False)
        Exit Sub
    End If

    ' Verify the number
    If BetID < 1 Or BetID > MAX_BETS Then
        Call AlertMsg(Index, DIALOGUE_LOTTERY_NUMBERS, , False)
        Exit Sub
    End If

    ' Verify Bet Slot Null
    If Not CheckBetSlot(BetID) Then
        Call AlertMsg(Index, DIALOGUE_LOTTERY_NUMBERALREADY, , False)
        Exit Sub
    End If

    ' Verify if Bet Value have minium value
    If BetValue < MIN_BETS_VALUE Then
        Call AlertMsg(Index, DIALOGUE_LOTTERY_MINBID, , False)
        Exit Sub
    End If

    ' Verify if Bet Value reached the max value
    If BetValue > MAX_BETS_VALUE Then
        Call AlertMsg(Index, DIALOGUE_LOTTERY_MAXBID, , False)
        Exit Sub
    End If

    ' Verify if player have the value and Remove the value
    If GetPlayerGold(Index) >= BetValue Then
        Call SetPlayerGold(Index, GetPlayerGold(Index) - BetValue)
        Call SendGoldUpdate(Index)
    Else
        Call AlertMsg(Index, DIALOGUE_LOTTERY_GOLD, , False)
        Exit Sub
    End If

    Call AlertMsg(Index, DIALOGUE_LOTTERY_SUCCESS, , False)
    Call SendEventMsgAll("[Lottery]", GetPlayerName(Index) & " bet on the number " & BetID & " value " & BetValue, White)

    Lottery.Bet(BetID).Owner = GetPlayerName(Index)
    Lottery.Bet(BetID).Value = Lottery.Bet(BetID).Value + BetValue
End Sub

Public Sub SendLotteryInfosTo(ByVal Index As Long)
    Dim Buffer As clsBuffer, Tmr As Currency

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLotteryInfo

    Buffer.WriteString Trim$(Lottery.LastBetWinner)
    Buffer.WriteByte Lottery.LastBetNum
    Buffer.WriteByte ConvertBooleanToByte(Lottery.Enabled)
    Buffer.WriteByte ConvertBooleanToByte(Lottery.BetEnabled)

    If Lottery.BetEnabled Or Lottery.Enabled Then
        Buffer.WriteLong 0
    Else
        Tmr = LOTTERY_START_HOURS    ' 3 hrs
        Tmr = (Tmr * 60)    ' 180 min
        Tmr = (Tmr * 60)    ' 10.800 segs
        Tmr = (Tmr * 1000)    ' 10.800.000 Milisegundos
        Debug.Print getTime
        Tmr = (Tmr + Lottery.Ended)    ' Soma o tempo que a loteria acabou com o tempo dela abrir novamente.
        Tmr = (Tmr - getTime)
        Tmr = (Tmr / 1000)
        Buffer.WriteLong CLng(Tmr)
    End If
    
    Buffer.WriteLong Lottery.Acumulado

    Buffer.WriteLong MIN_BETS_VALUE
    Buffer.WriteLong MAX_BETS_VALUE

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendLotteryInfosAll()
    Dim Buffer As clsBuffer, Tmr As Currency

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLotteryInfo
    
    Buffer.WriteString Trim$(Lottery.LastBetWinner)
    Buffer.WriteByte Lottery.LastBetNum
    
    Debug.Print ConvertBooleanToByte(Lottery.Enabled)
    Buffer.WriteByte ConvertBooleanToByte(Lottery.Enabled)
    Buffer.WriteByte ConvertBooleanToByte(Lottery.BetEnabled)

    If Lottery.BetEnabled Or Lottery.Enabled Then
        Buffer.WriteLong 0
    Else
        Tmr = LOTTERY_START_HOURS    ' 3 hrs
        Tmr = (Tmr * 60)    ' 180 min
        Tmr = (Tmr * 60)    ' 10.800 segs
        Tmr = (Tmr * 1000)    ' 10.800.000 Milisegundos
        Tmr = (Tmr + Lottery.Ended)    ' Soma o tempo que a loteria acabou com o tempo dela abrir novamente.
        Tmr = (Tmr - getTime)
        Buffer.WriteLong CLng(Tmr / 1000)
    End If

    Buffer.WriteLong MIN_BETS_VALUE
    Buffer.WriteLong MAX_BETS_VALUE
    
    SendDataToAll Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub CheckBetLoop()
    Dim Number As Byte
    Dim PlayerID As Integer
    Dim i As Byte
    Dim Accumulated As Long, Tmr As Currency

    If VerifyLotteryStatus Then    'On?
        ' Avisos - 1 (last is diferent message!)
        For i = 1 To MAX_AVISOS - 1
            If Not Lottery.Aviso(i) Then
                'Debug.Print "BetTmr " & Lottery.BetTmr + ((LOTTERY_TIME_BET / MAX_AVISOS) * 1000) & " <= " & getTime & " GetTime"
                If Lottery.BetTmr + ((LOTTERY_TIME_BET / MAX_AVISOS) * 1000) <= getTime Then
                    Lottery.Aviso(i) = True
                    Lottery.BetTmr = getTime
                    Call SendEventMsgAll("[Lottery]", "bets close in " & SecondsToHMS(LOTTERY_TIME_BET - ((getTime - Lottery.Started) / 1000)), Yellow)
                End If
            End If
        Next i

        ' Last Aviso
        If Not Lottery.Aviso(MAX_AVISOS) Then
            If Lottery.Started + (LOTTERY_TIME_BET * 1000) <= getTime Then
                Call SendEventMsgAll("[Lottery]", " Bets closed, Good Luck!!!", Green)
                Call CloseBets
                Lottery.Aviso(MAX_AVISOS) = True
                Call SendLotteryInfosAll
            End If
        End If

        If ((getTime - Lottery.Started) / 1000) > LOTTERY_SECS_DURATION Then    ' Time End?
            Number = ChooseLoteryNumber
            Accumulated = GetBetsAccumulated

            Call SendEventMsgAll("[Lottery]", "Drawn Number is " & Number, Yellow)

            If LenB(Trim$(Lottery.Bet(Number).Owner)) > 0 Then
                PlayerID = FindPlayer(Trim$(Lottery.Bet(Number).Owner))
                If PlayerID > 0 Then
                    Call SetPlayerGold(PlayerID, Accumulated)
                    Call SendGoldUpdate(PlayerID)
                    Call SendEventMsgAll("[Lottery]", "The winner is " & Trim$(Lottery.Bet(Number).Owner) & " Congratulations!!!", Green)

                    Lottery.Acumulado = 0
                    Call ClearBets    ' Remove all apostas e all owners
                    Call ClearLottery
                Else
                    Call SendEventMsgAll("[Lottery]", "Player " & Trim$(Lottery.Bet(Number).Owner) & " OFFLINE, jackpot in " & Accumulated, Green)
                    Call ClearBets    ' Remove all apostas e all owners
                    Call ClearLottery
                End If
                
                Lottery.LastBetWinner = Trim$(Lottery.Bet(Number).Owner)
                Lottery.LastBetNum = Number
            Else
                Call SendEventMsgAll("[Lottery]", "There were no winners, jackpot in " & Accumulated, Green)
                Call ClearBets    ' Remove all apostas e all owners
                Call ClearLottery
            End If

            Call SendEventMsgAll("[Lottery]", "The lottery has ended, next lottery starts in " & LOTTERY_START_HOURS & " hours", BrightRed)
            
            Call SendLotteryInfosAll
        End If
    Else    ' Get Off?? take on
        Tmr = LOTTERY_START_HOURS    ' 3 hrs
        Tmr = (Tmr * 60)    ' 180 min
        Tmr = (Tmr * 60)    ' 10.800 segs
        Tmr = (Tmr * 1000)    ' 10.800.000 Milisegundos
        Tmr = (Tmr + Lottery.Ended)    ' Soma o tempo que a loteria acabou com o tempo dela abrir novamente.
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

    Call SendEventMsgAll("[Lottery]", "Betting is on, place your bets in (" & SecondsToHMS(LOTTERY_TIME_BET) & ")", BrightGreen)
    
    Accumulated = Lottery.Acumulado
    If Accumulated > 0 Then
        Call SendEventMsgAll("[Lottery]", "The prize is accumulated in " & Accumulated, BrightGreen)
    End If
    
    Call SendLotteryInfosAll
End Sub

Public Sub HandleBet(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, BetValue As Long, BetID As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    BetID = Buffer.ReadByte
    BetValue = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing

    Call AddBetValue(Index, BetID, BetValue)
End Sub


Public Sub SendLotteryWindow(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLotteryWindow
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Function GetBetsAccumulated() As Long
    Dim i As Byte

    For i = 1 To MAX_BETS
        If Lottery.Bet(i).Value > 0 Then
            GetBetsAccumulated = GetBetsAccumulated + Lottery.Bet(i).Value
        End If
    Next i
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
    Dim i As Byte
    For i = 1 To MAX_BETS
        ClearBetSlot i
    Next i
End Sub

Private Sub ClearBetSlot(ByRef BetID As Byte)
    Call ZeroMemory(ByVal VarPtr(Lottery.Bet(BetID)), LenB(Lottery.Bet(BetID)))
    Lottery.Bet(BetID).Owner = vbNullString
End Sub

Public Sub ClearLottery()
    Dim i As Byte
    
    Lottery.Enabled = False
    Lottery.BetEnabled = False
    
    Lottery.Ended = getTime
    Lottery.Started = 0
    Lottery.BetTmr = 0
    Lottery.LastBetWinner = vbNullString
    Lottery.LastBetNum = 0
    
    For i = 1 To MAX_AVISOS
        Lottery.Aviso(i) = False
    Next i
End Sub
