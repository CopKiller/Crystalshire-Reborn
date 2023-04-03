Attribute VB_Name = "modEventSv"
Option Explicit

Private Event_Buffer As New clsBuffer
Private Event_DataTimer As Long
Private Event_DataBytes As Long
Private Event_DataPackets As Long

' Handle
Public Enum HEventPackets
    HaLotteryData = 1
    
    HEMSG_COUNT
End Enum

' Send
Public Enum SEventPackets
    SeLotteryData = 1
    SeLotteryInfo

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

'Public Lottery As LotteryStruct

'Private Type BetStruct
'    Owner As String * ACCOUNT_LENGTH
'    Value As Long
'End Type

'Private Type LotteryStruct
'    Enabled As Boolean
'    Started As Currency
'    Aviso(1 To MAX_AVISOS) As Boolean
'    Ended As Currency
'    BetEnabled As Boolean ' Are bets open?
'    BetTmr As Currency    ' Comparison time to send the notices
'    Bet(1 To MAX_BETS) As BetStruct
'    Acumulado As Long
'    LastBetNum As Byte ' Last Bet number
'    LastBetWinner As String * ACCOUNT_LENGTH ' Last Winner Name
'End Type

Public Sub SendLotterySaves(ByVal Save As EventOptions)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SeLotteryInfo
    Buffer.WriteByte ConvertBooleanToByte(Lottery.Enabled)
    Buffer.WriteByte ConvertBooleanToByte(Lottery.BetEnabled)
    Buffer.WriteLong Lottery.Acumulado
    Buffer.WriteByte Lottery.LastBetNum
    Buffer.WriteString Lottery.LastBetWinner
    
    For i = 1 To MAX_BETS
        If Trim$(Lottery.Bet(i).Owner) <> vbNullString Then
            Buffer.WriteByte i
            Buffer.WriteLong Lottery.Bet(i).Value
            Buffer.WriteString Trim$(Lottery.Bet(i).Owner)
        End If
    Next i

    SendToEventServer Buffer.ToArray
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
    End If

End Sub

Public Sub RequestLotteryData()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SeLotteryData
    Buffer.WriteString "Packet Recebida nessa porra Status do Socket de onde eu recebi: " & frmServer.EventSocket.State

    SendToEventServer Buffer.ToArray
    Set Buffer = Nothing
End Sub

Private Sub HandleLotteryData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim Buffer As clsBuffer
    Dim num1 As Byte
    Dim num2 As Integer
    Dim num3 As Long
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data
    
    Name = Buffer.ReadString
    num1 = Buffer.ReadByte
    num3 = Buffer.ReadLong
    

    Call TextAdd(Name)
    Call TextAdd(CStr(num1))
    Call TextAdd(CStr(num3))
    'Call TextAdd(num3)

    Set Buffer = Nothing

End Sub

Function ConnectToEventServer() As Boolean

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
