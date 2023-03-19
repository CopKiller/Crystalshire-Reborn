Attribute VB_Name = "modAuthentication"
Option Explicit

Private Auth_Buffer As New clsBuffer
Private Auth_DataTimer As Currency
Private Auth_DataBytes As Long
Private Auth_DataPackets As Long

Private Type LoginTokenRec
    User As String
    token As String
    TimeCreated As Currency
    Active As Boolean
    LoadPlayer As PlayerRec
End Type

' Packets sent by authentication server to game server
Public Enum AuthPackets
    ASetPlayerLoginToken
    ' Make sure AMSG_COUNT is below everything else
    AMSG_COUNT
End Enum

' Packets sent by Server to Auth Server
Public Enum GSPackets
    GShutDown
    GSavePlayer
    GClassesData
    GBanPlayer
    ' Make sure AMSG_COUNT is below everything else
    GMSG_COUNT
End Enum

' Indica se o usuario foi aceito ou nao no sistema.
Public Auth_HandleDataSub(AMSG_COUNT) As Long
Public LoginTokenAccepted(1 To MAX_PLAYERS) As Boolean
Public LoginToken(1 To MAX_PLAYERS) As LoginTokenRec
Public Const LoginTimer As Long = 60000    ' 60 seconds
Private Const MAX_CONNECTED_TIME As Long = 3000

' Connection details
Public Const GAME_SERVER_IP As String = "127.0.0.1"    ' "46.23.70.66"
Public Const AUTH_SERVER_IP As String = "127.0.0.1"    ' "46.23.70.66"
Public Const EVENT_SERVER_IP As String = "127.0.0.1"    ' "46.23.70.66"
Public Const GAME_SERVER_PORT As Long = 7001    ' the port used by the main game server
Public Const AUTH_SERVER_PORT As Long = 7002    ' the port used for people to connect to auth server
Public Const SERVER_AUTH_PORT As Long = 7003    ' the portal used for server to talk to auth server
Public Const EVENT_SERVER_PORT As Long = 7004    ' the portal used for server to talk to auth server

' Version constants
Public Const CLIENT_MAJOR As Byte = 1
Public Const CLIENT_MINOR As Byte = 8
Public Const CLIENT_REVISION As Byte = 0

' Verifica o token foi aceito pelo sistema.
' Se a conexao permanecer aberta por mais que MAX_CONNECTED_TIME sera encerrada.
Public Sub CheckConnectionTime()
    Dim I As Long
    For I = 1 To Player_HighIndex
        If IsConnected(I) Then
            ' Se o token ainda no foi ativado, verifica o tempo da conexao.
            If Not LoginTokenAccepted(I) Then
                If getTime >= TempPlayer(I).ConnectedTime + MAX_CONNECTED_TIME Then
                    Call AlertMsg(I, DIALOGUE_MSG_CONNECTION, 0, True)
                End If
            End If
        End If
    Next
End Sub

Private Function Auth_GetAddress(FunAddr As Long) As Long
    Auth_GetAddress = FunAddr
End Function

Public Sub Auth_InitMessages()
    Auth_HandleDataSub(ASetPlayerLoginToken) = Auth_GetAddress(AddressOf HandleSetPlayerLoginToken)
End Sub

Sub HandleSetPlayerLoginToken(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, User As String, tLoginToken As String, I As Long, tempData() As Byte, DataSize As Long, JogadorBytes As PlayerRec


    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    User = Buffer.ReadString
    tLoginToken = Buffer.ReadString

    ' find an inactive slot
    For I = 1 To MAX_PLAYERS
        If Not LoginToken(I).Active Then
            ' timed out
            LoginToken(I).User = User
            LoginToken(I).token = tLoginToken
            LoginToken(I).TimeCreated = getTime
            LoginToken(I).Active = True

            DataSize = LenB(JogadorBytes)
            ReDim tempData(DataSize - 1)
            tempData = Buffer.ReadBytes(DataSize)
            
            CopyMemory ByVal VarPtr(LoginToken(I).LoadPlayer), ByVal VarPtr(tempData(0)), DataSize

            Debug.Print "Token LoadPlayer Carregado: " & LoginToken(I).LoadPlayer.Login
            Exit Sub
        End If
    Next

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub Auth_HandleData(ByRef Data() As Byte)
    Dim TempBuffer() As Byte

    TempBuffer = DecryptPacket(Data, (UBound(Data) - LBound(Data)) + 1)

    Dim Buffer As clsBuffer
    Dim MsgType As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes TempBuffer

    MsgType = Buffer.ReadLong

    If MsgType < 0 Then
        Exit Sub
    End If

    If MsgType >= AMSG_COUNT Then
        Exit Sub
    End If

    CallWindowProc Auth_HandleDataSub(MsgType), 0, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Sub Auth_IncomingData(ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long

    ' Check if elapsed time has passed
    Auth_DataBytes = Auth_DataBytes + DataLength
    If getTime >= Auth_DataTimer Then
        Auth_DataTimer = getTime + 1000
        Auth_DataBytes = 0
        Auth_DataPackets = 0
    End If

    ' Get the data from the socket now
    frmServer.AuthSocket.GetData Buffer(), vbUnicode, DataLength
    Auth_Buffer.WriteBytes Buffer()

    If Auth_Buffer.Length >= 4 Then
        pLength = Auth_Buffer.ReadLong(False)

        If pLength < 0 Then
            Exit Sub
        End If
    End If

    Do While pLength > 0 And pLength <= Auth_Buffer.Length - 4
        If pLength <= Auth_Buffer.Length - 4 Then
            Auth_DataPackets = Auth_DataPackets + 1
            Auth_Buffer.ReadLong
            Auth_HandleData Auth_Buffer.ReadBytes(pLength)
        End If

        pLength = 0
        If Auth_Buffer.Length >= 4 Then
            pLength = Auth_Buffer.ReadLong(False)

            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop

    Auth_Buffer.Trim
End Sub



'*************************ENVIANDO DADOS PRO SERVIDOR DE AUTENTICAÇÃO**************************************
Function ConnectToAuthServer() As Boolean

' Check to see if we are already connected, if so just exit
    If IsConnectedAuthServer Then
        ConnectToAuthServer = True
        Exit Function
    End If

    frmServer.AuthSocket.Close
    frmServer.AuthSocket.RemoteHost = AUTH_SERVER_IP
    frmServer.AuthSocket.RemotePort = SERVER_AUTH_PORT
    frmServer.AuthSocket.Connect

    ConnectToAuthServer = IsConnectedAuthServer

End Function

Function IsConnectedAuthServer() As Boolean
    If frmServer.AuthSocket.State = sckConnected Then
        IsConnectedAuthServer = True
    End If
End Function

Sub Auth_SendDataTo(ByRef Data() As Byte)

    If Not IsConnectedAuthServer Then
        Call TextAdd("Dados não foram enviados, servidor de autenticação desconectado!")
        Exit Sub
    End If
    
    Dim TempBuffer() As Byte

    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    TempBuffer = EncryptPacket(Data, (UBound(Data) - LBound(Data)) + 1)

    Buffer.WriteLong (UBound(TempBuffer) - LBound(TempBuffer)) + 1
    Buffer.WriteBytes TempBuffer()
    
    frmServer.AuthSocket.SendData Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub Auth_SendShutdown()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong GShutDown
    Buffer.WriteByte ConvertBooleanToByte(isShuttingDown)

    Auth_SendDataTo Buffer.ToArray
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub Auth_SavePlayer(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim DataSize As Long
    Dim tempData() As Byte

    DataSize = LenB(Player(Index))
    ReDim tempData(DataSize - 1)
    CopyMemory tempData(0), ByVal VarPtr(Player(Index)), DataSize

    Set Buffer = New clsBuffer
    Buffer.WriteLong GSavePlayer
    Buffer.WriteString GetPlayerLogin(Index)
    Buffer.WriteBytes tempData

    Auth_SendDataTo Buffer.ToArray

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Function ConvertBooleanToByte(Variavel As Boolean) As Byte
    If Variavel = True Then
        ConvertBooleanToByte = YES
    Else
        ConvertBooleanToByte = NO
    End If
End Function

Public Function ConvertByteToBool(Variavel As Byte) As Boolean
    If Variavel = True Then
        ConvertBooleanToByte = True
    Else
        ConvertBooleanToByte = False
    End If
End Function

Sub Auth_ClassesData()
    Dim packet As String
    Dim I As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong GClassesData
    Buffer.WriteLong Max_Classes

    For I = 1 To Max_Classes
        Buffer.WriteString GetClassName(I)
        Buffer.WriteLong Class(I).MaxHP
        Buffer.WriteLong Class(I).MaxMP
        
        Buffer.WriteInteger Class(I).START_MAP
        Buffer.WriteInteger Class(I).START_X
        Buffer.WriteInteger Class(I).START_Y

        ' set sprite array size
        n = UBound(Class(I).MaleSprite)

        ' send array size
        Buffer.WriteLong n

        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(I).MaleSprite(q)
        Next

        ' set sprite array size
        n = UBound(Class(I).FemaleSprite)

        ' send array size
        Buffer.WriteLong n

        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(I).FemaleSprite(q)
        Next

        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(I).Stat(q)
        Next
    Next

    Auth_SendDataTo Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
    
    Call SetStatus("## Dados de classes enviados... ##")
End Sub

Sub Auth_BanPlayerIP(ByVal IP As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong GBanPlayer
    Buffer.WriteString IP
    Auth_SendDataTo Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub
