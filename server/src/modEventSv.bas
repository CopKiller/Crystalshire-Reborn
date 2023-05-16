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
    HaItemsPendentes
    
    HEMSG_COUNT
End Enum

' Send
Public Enum SEventPackets
    SeLotteryData = 1
    SeReqLotteryInfo
    SeItemsPendentes
    SeDiscordMsg
    
    SEMSG_COUNT
End Enum

' Utilidade
Public Enum EventOptions
    Saves = 0
    Clear
End Enum

Public Event_HandleDataSub(HEMSG_COUNT) As Long

Private Function Event_GetAddress(FunAddr As Long) As Long
    Event_GetAddress = FunAddr
End Function

Public Sub Event_InitMessages()
    Event_HandleDataSub(HaLotteryData) = Event_GetAddress(AddressOf HandleLotteryData)
    Event_HandleDataSub(HaLotteryData) = Event_GetAddress(AddressOf HandleItemsPendentes)
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

Sub SendToEventServer(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim tempData() As Byte
    
    If Options.EVENTSV = NO Then Exit Sub

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

