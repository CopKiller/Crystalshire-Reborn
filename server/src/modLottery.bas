Attribute VB_Name = "modLottery"
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Const MAX_BETS As Byte = 100

Private Const MIN_BETS_VALUE As Long = 20    ' min bet value
Private Const MAX_BETS_VALUE As Long = 100000    ' max bet value
Public Const LOTTERY_START_HOURS As Byte = 3    ' start in..
Private Const LOTTERY_SECS_DURATION As Long = 300    ' 30 mins duration
Public Const LOTTERY_TIME_BET As Long = 180    'starting time of bets / 3 mins

Private Const MAX_AVISOS As Byte = 3

Public lottery As LotteryStruct

Private Type BetStruct
    Owner As String * ACCOUNT_LENGTH
    Value As Long
End Type

Public Type LotteryStruct
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

    lottery.Bet(BetID).Owner = GetPlayerName(Index)
    lottery.Bet(BetID).Value = BetValue
    
    SendLotterySaves Save
End Sub

Public Sub SendLotteryInfosTo(ByVal Index As Long)
    Dim Buffer As clsBuffer, Tmr As Currency

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLotteryInfo

    Buffer.WriteString Trim$(lottery.LastBetWinner)
    Buffer.WriteByte lottery.LastBetNum
    Buffer.WriteByte ConvertBooleanToByte(lottery.Enabled)
    Buffer.WriteByte ConvertBooleanToByte(lottery.BetEnabled)

    If lottery.BetEnabled Or lottery.Enabled Then
        Buffer.WriteLong 0
    Else
        Tmr = LOTTERY_START_HOURS    ' 3 hrs
        Tmr = (Tmr * 60)    ' 180 min
        Tmr = (Tmr * 60)    ' 10.800 segs
        Tmr = (Tmr * 1000)    ' 10.800.000 Milisegundos
        Debug.Print getTime
        Tmr = (Tmr + lottery.Ended)    ' Soma o tempo que a loteria acabou com o tempo dela abrir novamente.
        Tmr = (Tmr - getTime)
        Tmr = (Tmr / 1000)
        Buffer.WriteLong CLng(Tmr)
    End If
    
    Buffer.WriteLong GetBetsAccumulated + lottery.Acumulado

    Buffer.WriteLong MIN_BETS_VALUE
    Buffer.WriteLong MAX_BETS_VALUE

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendLotteryInfosAll()
    Dim Buffer As clsBuffer, Tmr As Currency

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLotteryInfo
    
    Buffer.WriteString Trim$(lottery.LastBetWinner)
    Buffer.WriteByte lottery.LastBetNum
    
    Debug.Print ConvertBooleanToByte(lottery.Enabled)
    Buffer.WriteByte ConvertBooleanToByte(lottery.Enabled)
    Buffer.WriteByte ConvertBooleanToByte(lottery.BetEnabled)

    If lottery.BetEnabled Or lottery.Enabled Then
        Buffer.WriteLong 0
    Else
        Tmr = LOTTERY_START_HOURS    ' 3 hrs
        Tmr = (Tmr * 60)    ' 180 min
        Tmr = (Tmr * 60)    ' 10.800 segs
        Tmr = (Tmr * 1000)    ' 10.800.000 Milisegundos
        Tmr = (Tmr + lottery.Ended)    ' Soma o tempo que a loteria acabou com o tempo dela abrir novamente.
        Tmr = (Tmr - getTime)
        Buffer.WriteLong CLng(Tmr / 1000)
    End If
    
    Buffer.WriteLong GetBetsAccumulated + lottery.Acumulado

    Buffer.WriteLong MIN_BETS_VALUE
    Buffer.WriteLong MAX_BETS_VALUE
    
    SendDataToAll Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub CheckBetLoop()
    Dim Number As Byte
    Dim PlayerID As Integer
    Dim I As Byte
    Dim Accumulated As Long, Tmr As Currency, BackupLastWinner As String

    If VerifyLotteryStatus Then    'Lottery On?

        If VerifyBetStatus Then ' Bets On?
            ' Avisos - 1 (last is diferent message!)
            For I = 1 To MAX_AVISOS - 1
                If Not lottery.Aviso(I) Then
                    If lottery.BetTmr + ((LOTTERY_TIME_BET / MAX_AVISOS) * 1000) <= getTime Then
                        lottery.Aviso(I) = True
                        lottery.BetTmr = getTime
                        Call SendEventMsgAll("[Lottery]", "bets close in " & SecondsToHMS(LOTTERY_TIME_BET - ((getTime - lottery.Started) / 1000)), Yellow)
                    End If
                End If
            Next I

            ' Last Aviso
            If Not lottery.Aviso(MAX_AVISOS) Then
                If lottery.Started + (LOTTERY_TIME_BET * 1000) <= getTime Then
                    Call SendEventMsgAll("[Lottery]", " Bets closed, Good Luck!!!", Green)
                    Call CloseBets
                    lottery.Aviso(MAX_AVISOS) = True
                    Call SendLotteryInfosAll

                    SendLotterySaves Save
                End If
            End If
        End If

        If ((getTime - lottery.Started) / 1000) > LOTTERY_SECS_DURATION Then    ' Time End?
            Number = ChooseLoteryNumber
            Accumulated = GetBetsAccumulated + lottery.Acumulado

            Call SendEventMsgAll("[Lottery]", "Drawn Number is " & Number, Yellow)

            If LenB(Trim$(lottery.Bet(Number).Owner)) > 0 Then
                PlayerID = FindPlayer(Trim$(lottery.Bet(Number).Owner))
                If PlayerID > 0 Then
                    Call SetPlayerGold(PlayerID, Accumulated)
                    Call SendGoldUpdate(PlayerID)
                    Call SendEventMsgAll("[Lottery]", "The winner is " & Trim$(lottery.Bet(Number).Owner) & " Congratulations!!!", Green)

                    lottery.Acumulado = 0
                    BackupLastWinner = Trim$(lottery.Bet(Number).Owner)
                Else
                    Call SendEventMsgAll("[Lottery]", "Player " & Trim$(lottery.Bet(Number).Owner) & " OFFLINE, jackpot in " & Accumulated, Green)
                    lottery.Acumulado = Accumulated
                End If

                Call ClearBets    ' Remove all apostas e all owners
                Call ClearLottery

                lottery.LastBetWinner = BackupLastWinner
            Else
                lottery.Acumulado = Accumulated
                Call SendEventMsgAll("[Lottery]", "There were no winners, jackpot in " & lottery.Acumulado, Green)
                Call ClearBets    ' Remove all apostas e all owners
                Call ClearLottery
            End If

            lottery.LastBetNum = Number
            
            SendLotterySaves Save    ' Faz a limpeza no servidor de eventos

            Call SendEventMsgAll("[Lottery]", "The lottery has ended, next lottery starts in " & LOTTERY_START_HOURS & " hours", BrightRed)

            Call SendLotteryInfosAll
        End If
    Else    ' Get Off?? take on
        Tmr = LOTTERY_START_HOURS    ' 3 hrs
        Tmr = (Tmr * 60)    ' 180 min
        Tmr = (Tmr * 60)    ' 10.800 segs
        Tmr = (Tmr * 1000)    ' 10.800.000 Milisegundos
        Tmr = (Tmr + lottery.Ended)    ' Soma o tempo que a loteria acabou com o tempo dela abrir novamente.
        'Debug.Print Tmr
        If Tmr <= getTime Then
            Call StartLottery
        End If
    End If
End Sub

Public Sub CloseBets()
    lottery.BetEnabled = False
    lottery.BetTmr = 0
End Sub

Public Sub StartLottery()
    Dim Accumulated As Long
    
    lottery.Enabled = True
    lottery.BetEnabled = True
    
    lottery.Started = getTime
    lottery.BetTmr = getTime

    Call SendEventMsgAll("[Lottery]", "Betting is on, place your bets in (" & SecondsToHMS(LOTTERY_TIME_BET) & ")", BrightGreen)
    
    Accumulated = lottery.Acumulado
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
    Dim I As Byte

    For I = 1 To MAX_BETS
        If lottery.Bet(I).Value > 0 Then
            GetBetsAccumulated = GetBetsAccumulated + lottery.Bet(I).Value
        End If
    Next I
End Function

Private Function CheckBetSlot(ByRef BetID As Byte) As Boolean
' Prevent Subscript out of range
    If BetID > MAX_BETS Or BetID <= 0 Then Exit Function
    
    CheckBetSlot = True

    If LenB(Trim$(lottery.Bet(BetID).Owner)) > 0 Then
        CheckBetSlot = False
    End If

End Function

Private Function ChooseLoteryNumber() As Byte
    ChooseLoteryNumber = Random(1, CLng(MAX_BETS))
End Function

Public Function VerifyLotteryStatus() As Boolean
    VerifyLotteryStatus = lottery.Enabled
End Function

Private Function VerifyBetStatus() As Boolean
    VerifyBetStatus = lottery.BetEnabled
End Function

Public Sub ClearBets()
    Dim I As Byte
    For I = 1 To MAX_BETS
        ClearBetSlot I
    Next I
End Sub

Private Sub ClearBetSlot(ByRef BetID As Byte)
    Call ZeroMemory(ByVal VarPtr(lottery.Bet(BetID)), LenB(lottery.Bet(BetID)))
    lottery.Bet(BetID).Owner = vbNullString
End Sub

Public Sub LoadLottery()
    Dim I As Byte, Diretorio As String, SString As String, Filter() As String
    
    Diretorio = App.Path & "/data/EventsData.ini"

    If FileExist(Diretorio, True) Then
        lottery.Enabled = ConvertByteToBool(CByte(GetVar(Diretorio, "LOTTERY", "Status")))
        lottery.BetEnabled = ConvertByteToBool(CByte(GetVar(Diretorio, "LOTTERY", "BetStatus")))
        lottery.Acumulado = CLng(GetVar(Diretorio, "LOTTERY", "Accumulated"))
        lottery.LastBetNum = CByte(GetVar(Diretorio, "LOTTERY", "LastBetNum"))
        lottery.LastBetWinner = CStr(Trim$(GetVar(Diretorio, "LOTTERY", "LastBetWinner")))
        
        SString = CStr(Trim$(GetVar(Diretorio, "LOTTERY", "CountStr")))
        
        Filter = Split(SString, ",")
        
        For I = LBound(Filter) To UBound(Filter)
            lottery.Bet(CByte(Filter(I))).Owner = Trim$(CStr(Trim$(GetVar(Diretorio, "LOTTERY", "BetOwner" & CByte(Filter(I))))))
            lottery.Bet(CByte(Filter(I))).Value = Trim$(CLng(Trim$(GetVar(Diretorio, "LOTTERY", "BetValue" & CByte(Filter(I))))))
        Next I
    Else
        Call RequestLotteryData
    End If
End Sub

Public Sub ClearLottery()
    Dim I As Byte
    
    lottery.Enabled = False
    lottery.BetEnabled = False
    
    lottery.Ended = getTime
    lottery.Started = 0
    lottery.BetTmr = 0
    lottery.LastBetWinner = "Ninguem"
    lottery.LastBetNum = 0
    
    For I = 1 To MAX_AVISOS
        lottery.Aviso(I) = False
    Next I
End Sub

Public Sub SendLotterySaves(ByVal Save As EventOptions)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim I As Byte
    Dim CountStr As String, Filter() As String

    If Save = Save Then
    If Options.EVENTSV = YES And IsEventServerConnected Then
        Buffer.WriteLong SeLotteryData
        Buffer.WriteByte ConvertBooleanToByte(lottery.Enabled)
        Buffer.WriteByte ConvertBooleanToByte(lottery.BetEnabled)
        Buffer.WriteLong lottery.Acumulado
        Buffer.WriteByte lottery.LastBetNum
        Buffer.WriteString lottery.LastBetWinner

        For I = 1 To MAX_BETS
            If LenB(Trim$(lottery.Bet(I).Owner)) > 0 Then
                Buffer.WriteByte I
                Buffer.WriteLong lottery.Bet(I).Value
                Buffer.WriteString Trim$(lottery.Bet(I).Owner)
            End If
        Next I

        SendToEventServer Buffer.ToArray

    ElseIf Options.EVENTSV = NO Or Not IsEventServerConnected Then
        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "Status", ConvertBooleanToByte(lottery.Enabled))
        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "BetStatus", ConvertBooleanToByte(lottery.BetEnabled))
        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "Accumulated", CStr(lottery.Acumulado))
        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "LastBetNum", CStr(lottery.LastBetNum))
        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "LastBetWinner", CStr(lottery.LastBetWinner))

        For I = 1 To MAX_BETS
            If Trim$(lottery.Bet(I).Owner) <> vbNullString Then
                If CountStr <> vbNullString Then CountStr = CountStr & "," & I
                If CountStr = vbNullString Then CountStr = I
            End If
        Next I

        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "CountStr", CountStr)

        Filter = Split(CountStr, ",")

        For I = 0 To UBound(Filter)
            If LenB(Trim$(Filter(I))) > 0 Then
                Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "BetOwner" & Trim$(Filter(I)), Trim$(lottery.Bet(Filter(I)).Owner))
                Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "BetValue" & Trim$(Filter(I)), Trim$(lottery.Bet(Filter(I)).Value))
            End If
        Next I
    End If
    
    Else
        Buffer.WriteLong SeLotteryData
        SendToEventServer Buffer.ToArray
    End If

Set Buffer = Nothing
End Sub

Public Sub RequestLotteryData()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SeReqLotteryInfo

    SendToEventServer Buffer.ToArray
    Set Buffer = Nothing
End Sub

Public Sub HandleLotteryData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim Buffer As clsBuffer
    Dim I As Byte, MAX_INDICE As Byte, Indice As Byte, lot As LotteryStruct

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data


    lot.Enabled = ConvertByteToBool(Buffer.ReadByte)
    lot.BetEnabled = ConvertByteToBool(Buffer.ReadByte)
    lot.Acumulado = Buffer.ReadLong
    lot.LastBetNum = Buffer.ReadByte
    lot.LastBetWinner = Buffer.ReadString
    
    lottery.Acumulado = lot.Acumulado
    lottery.LastBetNum = lot.LastBetNum
    lottery.LastBetWinner = lot.LastBetWinner

    MAX_INDICE = Buffer.ReadLong

    If MAX_INDICE > 0 Then
        For I = 1 To MAX_INDICE
            Indice = Buffer.ReadByte
            If Indice > 0 Then
                lot.Bet(Indice).Owner = Buffer.ReadString
                lot.Bet(Indice).Value = Buffer.ReadLong
                
                lottery.Bet(Indice).Owner = lot.Bet(Indice).Owner
                lottery.Bet(Indice).Value = lot.Bet(Indice).Value
            End If
        Next I
    End If
    Set Buffer = Nothing

    If lot.Enabled Then
        If lot.BetEnabled Then
            Call StartLottery
        Else
            lottery.Enabled = True
            Call CloseBets
            For I = 1 To 3
                lottery.Aviso(I) = True
            Next I
        End If
    End If

    'lottery = lot

    Call TextEventAdd("Lottery Data Received!")
End Sub
