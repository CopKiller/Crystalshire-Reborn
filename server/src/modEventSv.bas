Attribute VB_Name = "modEventSv"
Option Explicit

Private Event_Buffer As New clsBuffer
Private Event_DataTimer As Long
Private Event_DataBytes As Long
Private Event_DataPackets As Long

' Handle
Public Enum HEventPackets
    HaLotteryData = 1
    HaPing
    
    HEMSG_COUNT
End Enum

' Send
Public Enum SEventPackets
    SeLotteryData = 1
    SeReqLotteryInfo
    
    SEMSG_COUNT
End Enum

' Utilidade
Public Enum EventOptions
    Save = 0
    Clear
End Enum

Public Event_HandleDataSub(HEMSG_COUNT) As Long

Private Function Event_GetAddress(FunAddr As Long) As Long
    Event_GetAddress = FunAddr
End Function

Public Sub Event_InitMessages()
    Event_HandleDataSub(HaLotteryData) = Event_GetAddress(AddressOf HandleLotteryData)
End Sub

Sub Event_HandleData(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim MsgType As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    MsgType = Buffer.ReadLong

    If MsgType < 0 Then
        Exit Sub
    End If

    If MsgType >= HEMSG_COUNT Then
        Exit Sub
    End If

    CallWindowProc Event_HandleDataSub(MsgType), 0, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Sub Event_IncomingData(ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long

    ' Check if elapsed time has passed
    Event_DataBytes = Event_DataBytes + DataLength
    If getTime >= Event_DataTimer Then
        Event_DataTimer = getTime + 1000
        Event_DataBytes = 0
        Event_DataPackets = 0
    End If

    ' Get the data from the socket now
    frmServer.EventSocket.GetData Buffer(), vbUnicode, DataLength
    Event_Buffer.WriteBytes Buffer()

    If Event_Buffer.Length >= 4 Then
        pLength = Event_Buffer.ReadLong(False)

        If pLength < 0 Then
            Exit Sub
        End If
    End If

    Do While pLength > 0 And pLength <= Event_Buffer.Length - 4
        If pLength <= Event_Buffer.Length - 4 Then
            Event_DataPackets = Event_DataPackets + 1
            Event_Buffer.ReadLong
            Event_HandleData Event_Buffer.ReadBytes(pLength)
        End If

        pLength = 0
        If Event_Buffer.Length >= 4 Then
            pLength = Event_Buffer.ReadLong(False)

            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop

    Event_Buffer.Trim
End Sub

Public Sub Event_AcceptConnection(ByVal SocketId As Long)
    frmServer.EventSocket.Close
    frmServer.EventSocket.Accept SocketId

    Call TextAdd("Event Server Connected")
End Sub

Function IsEventServerConnected() As Boolean

    If frmServer.EventSocket.State = sckConnected Then
        IsEventServerConnected = True
    End If

End Function

Public Sub SendLotterySaves(ByVal Save As EventOptions)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Byte
    Dim CountStr As String, Filter() As String

    If Save = Save Then
    If Options.EVENTSV = YES And IsEventServerConnected Then
        Buffer.WriteLong SeLotteryData
        Buffer.WriteByte ConvertBooleanToByte(lottery.Enabled)
        Buffer.WriteByte ConvertBooleanToByte(lottery.BetEnabled)
        Buffer.WriteLong lottery.Acumulado
        Buffer.WriteByte lottery.LastBetNum
        Buffer.WriteString lottery.LastBetWinner

        For i = 1 To MAX_BETS
            If LenB(Trim$(lottery.Bet(i).Owner)) > 0 Then
                Buffer.WriteByte i
                Buffer.WriteLong lottery.Bet(i).Value
                Buffer.WriteString Trim$(lottery.Bet(i).Owner)
            End If
        Next i

        SendToEventServer Buffer.ToArray

    ElseIf Options.EVENTSV = NO Or Not IsEventServerConnected Then
        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "Status", ConvertBooleanToByte(lottery.Enabled))
        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "BetStatus", ConvertBooleanToByte(lottery.BetEnabled))
        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "Accumulated", CStr(lottery.Acumulado))
        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "LastBetNum", CStr(lottery.LastBetNum))
        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "LastBetWinner", CStr(lottery.LastBetWinner))

        For i = 1 To MAX_BETS
            If Trim$(lottery.Bet(i).Owner) <> vbNullString Then
                If CountStr <> vbNullString Then CountStr = CountStr & "," & i
                If CountStr = vbNullString Then CountStr = i
            End If
        Next i

        Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "CountStr", CountStr)

        Filter = Split(CountStr, ",")

        For i = 0 To UBound(Filter)
            If LenB(Trim$(Filter(i))) > 0 Then
                Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "BetOwner" & Trim$(Filter(i)), Trim$(lottery.Bet(Filter(i)).Owner))
                Call PutVar(App.Path & "/data/EventsData.ini", "LOTTERY", "BetValue" & Trim$(Filter(i)), Trim$(lottery.Bet(Filter(i)).Value))
            End If
        Next i
    End If
    
    Else
        Buffer.WriteLong SeLotteryData
        SendToEventServer Buffer.ToArray
    End If

Set Buffer = Nothing
End Sub

Sub SendToEventServer(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim tempData() As Byte

    Set Buffer = New clsBuffer
    tempData = Data

    Buffer.PreAllocate 4 + (UBound(tempData) - LBound(tempData)) + 1
    Buffer.WriteLong (UBound(tempData) - LBound(tempData)) + 1
    Buffer.WriteBytes tempData()

    If IsEventServerConnected Then
        frmServer.EventSocket.SendData Buffer.ToArray()
    Else
        Call TextEventAdd("Erro ao enviar dados dos eventos, Event Server Offline!")
    End If

End Sub

Public Sub RequestLotteryData()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SeReqLotteryInfo

    SendToEventServer Buffer.ToArray
    Set Buffer = Nothing
End Sub

Private Sub HandleLotteryData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim Buffer As clsBuffer
    Dim i As Byte, MAX_INDICE As Byte, Indice As Byte, lot As LotteryStruct

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
        For i = 1 To MAX_INDICE
            Indice = Buffer.ReadByte
            If Indice > 0 Then
                lot.Bet(Indice).Owner = Buffer.ReadString
                lot.Bet(Indice).Value = Buffer.ReadLong
                
                lottery.Bet(Indice).Owner = lot.Bet(Indice).Owner
                lottery.Bet(Indice).Value = lot.Bet(Indice).Value
            End If
        Next i
    End If
    Set Buffer = Nothing

    If lot.Enabled Then
        If lot.BetEnabled Then
            Call StartLottery
        Else
            lottery.Enabled = True
            Call CloseBets
            For i = 1 To 3
                lottery.Aviso(i) = True
            Next i
        End If
    End If

    'lottery = lot

    Call TextEventAdd("Lottery Data Received!")
End Sub

Function ConnectToEventServer() As Boolean

    If Options.EVENTSV = NO Then
        ConnectToEventServer = False
        frmServer.EventSocket.Close
        Exit Function
    End If

' Check to see if we are already connected, if so just exit
    If IsEventServerConnected Then
        ConnectToEventServer = True
        Exit Function
    End If

    frmServer.EventSocket.Close
    frmServer.EventSocket.RemoteHost = EVENT_SERVER_IP
    frmServer.EventSocket.RemotePort = EVENT_SERVER_PORT
    frmServer.EventSocket.Connect

    ConnectToEventServer = IsEventServerConnected

End Function
