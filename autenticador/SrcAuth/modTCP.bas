Attribute VB_Name = "modTCP"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function Current_IP(ByVal Index As Long) As String
    Current_IP = frmMain.Socket(Index).RemoteHostIP
End Function

' Verifica o token foi aceito pelo sistema.
' Se a conexao permanecer aberta por mais que MAX_CONNECTED_TIME sera encerrada.
Sub CheckConnectionTime()
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsConnected(i) Then

            ' Se o token ainda nao foi ativado, verifica o tempo da conexao.
            If Not TempPlayer(i).TokenAccepted Then

                If getTime >= TempPlayer(i).ConnectedTime + MAX_CONNECTED_TIME Then
                    Call HackingAttempt(i)
                End If

            End If

        End If
    Next

End Sub

Function ConnectToGameServer() As Boolean

' Check to see if we are already connected, if so just exit
    If IsConnectedGameServer Then
        ConnectToGameServer = True
        Exit Function
    End If

    frmMain.ServerSocket.Close
    frmMain.ServerSocket.RemoteHost = GAME_SERVER_IP
    frmMain.ServerSocket.RemotePort = SERVER_AUTH_PORT
    frmMain.ServerSocket.Connect

    ConnectToGameServer = IsConnectedGameServer
    
End Function

Function IsConnectedGameServer() As Boolean
     If frmMain.ServerSocket.State = sckConnected Then
        IsConnectedGameServer = True
    End If
End Function

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long
    If isBanned_IP(Current_IP(Index)) Then Exit Sub

    If Index = 0 Then
        i = FindOpenPlayerSlot

        If i <> 0 And Current_IP(Index) <> vbNullString Then
            SetStatus "Received connection from " & Current_IP(Index) & "."
            ' Whoho, we can connect them
            frmMain.Socket(i).Close
            frmMain.Socket(i).Accept SocketId
            SocketConnected i
        End If
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    'Obtem o tempo em que o usuario foi conectado.
    TempPlayer(Index).ConnectedTime = getTime
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long

    ' Check for data flooding
    If TempPlayer(Index).DataBytes > 1000 Then
        Exit Sub
    End If

    ' Check for packet flooding
    If TempPlayer(Index).DataPackets > 25 Then
        Exit Sub
    End If

    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    If getTime >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = getTime + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If

    ' Get the data from the socket now
    frmMain.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(Index).Buffer.WriteBytes Buffer()

    If TempPlayer(Index).Buffer.Length >= 4 Then
        pLength = TempPlayer(Index).Buffer.ReadLong(False)

        If pLength < 0 Then
            Exit Sub
        End If
    End If

    Do While pLength > 0 And pLength <= TempPlayer(Index).Buffer.Length - 4
        If pLength <= TempPlayer(Index).Buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).Buffer.ReadLong
            HandleData Index, TempPlayer(Index).Buffer.ReadBytes(pLength)
        End If

        pLength = 0
        If TempPlayer(Index).Buffer.Length >= 4 Then
            pLength = TempPlayer(Index).Buffer.ReadLong(False)

            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop

    TempPlayer(Index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal Index As Long)
    ClearPlayer Index
    SetStatus "Connection from " & Current_IP(Index) & " has been terminated."
    frmMain.Socket(Index).Close
End Sub

Function FindOpenPlayerSlot() As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next
End Function

Function IsConnected(ByVal Index As Long) As Boolean
    If frmMain.Socket(Index).State = sckConnected Then IsConnected = True
End Function

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim TempData() As Byte

    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        TempData = EncryptPacket(Data, (UBound(Data) - LBound(Data)) + 1)

        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()

        frmMain.Socket(Index).SendData Buffer.ToArray()

        Set Buffer = Nothing
    End If
End Sub

Sub SendDataToGameServer(ByRef Data() As Byte)
    Dim TempBuffer() As Byte
    
    If Not ConnectToGameServer Then
        Call SetStatus("Data Transfer To Game Server Failed!")
        Exit Sub
    End If

    Dim Length As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    TempBuffer = EncryptPacket(Data, (UBound(Data) - LBound(Data)) + 1)
    Length = (UBound(TempBuffer) - LBound(TempBuffer)) + 1

    Buffer.PreAllocate 4 + Length
    Buffer.WriteLong Length
    Buffer.WriteBytes TempBuffer()

    frmMain.ServerSocket.SendData Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HackingAttempt(ByVal Index As Long)
    SendAlertMsg Index, DIALOGUE_MSG_CONNECTION
End Sub

Sub SendAlertMsg(ByVal Index As Long, ByVal Msg As Long, Optional ByVal menuReset As Long = 0, Optional ByVal kick As Boolean = True)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteLong Msg
    Buffer.WriteLong menuReset
    If kick Then Buffer.WriteLong 1 Else Buffer.WriteLong 0

    SendDataTo Index, Buffer.ToArray()

    Set Buffer = Nothing

    If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents

    CloseSocket Index
End Sub

Public Sub SendLoginTokenToPlayer(ByVal Index As Long, ByVal loginToken As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSetPlayerLoginToken
    Buffer.WriteString loginToken

    SendDataTo Index, Buffer.ToArray()

    Set Buffer = Nothing
End Sub

Public Sub SendLoginTokenToGameServer(ByVal Index As Long, ByVal Username As String, ByVal loginToken As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Dim DataSize As Long
    Dim TempData() As Byte

    DataSize = LenB(Player(Index))
    ReDim TempData(DataSize - 1)
    CopyMemory TempData(0), ByVal VarPtr(Player(Index)), DataSize

    Buffer.WriteLong ASetPlayerLoginToken
    Buffer.WriteString Username
    Buffer.WriteString loginToken
    Buffer.WriteBytes TempData
    SendDataToGameServer Buffer.ToArray()

    Set Buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong Class(i).MaxHP
        Buffer.WriteLong Class(i).MaxMP
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next

        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next

        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function
