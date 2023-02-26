Attribute VB_Name = "modServerToAuth"
Option Explicit

' Packets sent by Server to Auth Server
Public Enum GSPackets
    GShutDown
    GSavePlayer
    GClassesData
    GBanPlayer
    ' Make sure GMSG_COUNT is below everything else
    GMSG_COUNT
End Enum

Public GS_HandleDataSub(GMSG_COUNT) As Long

Private GS_Buffer As New clsBuffer
Private GS_DataTimer As Currency
Private GS_DataBytes As Long
Private GS_DataPackets As Long

Public Sub GS_InitMessages()
    GS_HandleDataSub(GShutDown) = GetAddress(AddressOf HandleShutDown)
    GS_HandleDataSub(GSavePlayer) = GetAddress(AddressOf HandleSavePlayer)
    GS_HandleDataSub(GClassesData) = GetAddress(AddressOf HandleClassesData)
    GS_HandleDataSub(GBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
End Sub

Sub GS_IncomingData(ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long

    ' Check if elapsed time has passed
    GS_DataBytes = GS_DataBytes + DataLength
    If getTime >= GS_DataTimer Then
        GS_DataTimer = getTime + 1000
        GS_DataBytes = 0
        GS_DataPackets = 0
    End If

    ' Get the data from the socket now
    frmMain.ServerSocket.GetData Buffer(), vbUnicode, DataLength
    GS_Buffer.WriteBytes Buffer()

    If GS_Buffer.Length >= 4 Then
        pLength = GS_Buffer.ReadLong(False)

        If pLength < 0 Then
            Exit Sub
        End If
    End If

    Do While pLength > 0 And pLength <= GS_Buffer.Length - 4
        If pLength <= GS_Buffer.Length - 4 Then
            GS_DataPackets = GS_DataPackets + 1
            GS_Buffer.ReadLong
            GS_HandleData GS_Buffer.ReadBytes(pLength)
        End If

        pLength = 0
        If GS_Buffer.Length >= 4 Then
            pLength = GS_Buffer.ReadLong(False)

            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop

    GS_Buffer.Trim
End Sub

Private Sub GS_HandleData(ByRef Data() As Byte)
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

    If MsgType >= GMSG_COUNT Then
        Exit Sub
    End If

    CallWindowProc GS_HandleDataSub(MsgType), 0, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Private Sub HandleShutDown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    isShuttingDown = ConvertByteToBoolean(Buffer.ReadByte)

    Buffer.Flush: Set Buffer = Nothing

    Exit Sub

errorhandler:
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim filename As String, IP As String, F As Long

    ' IP banning
    filename = App.Path & "\banlist_ip.txt"

    ' Make sure the file exists
    If Not FileExist(filename) Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    IP = Buffer.ReadString
    
    Buffer.Flush: Set Buffer = Nothing

    ' Print the IP in the ip ban list
    F = FreeFile
    Open filename For Append As #F
    Print #F, IP
    Close #F
End Sub

Private Sub HandleSavePlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, DataSize As Long, TempData() As Byte
    Dim filename As String, F As Long
    Dim AccountName As String, Save As PlayerRec

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    AccountName = Buffer.ReadString
    
    DataSize = LenB(Save)
    ReDim TempData(DataSize - 1)
    TempData = Buffer.ReadBytes(DataSize)
    CopyMemory ByVal VarPtr(Save), ByVal VarPtr(TempData(0)), DataSize

    Buffer.Flush: Set Buffer = Nothing
    
    Call SavePlayer_ByGameServer(AccountName, Save)
End Sub

Private Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, x As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong    'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString    'Trim$(Parse(n))
            .MaxHP = Buffer.ReadLong    'CLng(Parse(n + 1))
            .MaxMP = Buffer.ReadLong    'CLng(Parse(n + 2))
            
            .START_MAP = Buffer.ReadInteger
            .START_X = Buffer.ReadInteger
            .START_Y = Buffer.ReadInteger
            
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For x = 0 To z
                .MaleSprite(x) = Buffer.ReadLong
            Next

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For x = 0 To z
                .FemaleSprite(x) = Buffer.ReadLong
            Next

            For x = 1 To Stats.Stat_Count - 1
                .Stat(x) = Buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    Set Buffer = Nothing
End Sub

Public Function SavePlayer_ByGameServer(ByVal AccountName As String, ByRef Jogador As PlayerRec)
    Dim filename As String
    Dim F As Long
    

    filename = App.Path & "\accounts\" & SanitiseString(AccountName) & ".bin"

    F = FreeFile

    Open filename For Binary As #F
    Put #F, , Jogador
    Close #F
    
    Call SetStatus("## Dados de jogador salvo! Conta: " & SanitiseString(AccountName) & " Jogador: " & Trim$(Jogador.Name) & " ##")
End Function

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long

    If Trim$(Name) = vbNullString Then Exit Sub

    Call ClearPlayer(Index)
    filename = App.Path & "\accounts\" & SanitiseString(Trim$(Name)) & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Player(Index)
    Close #F
    
    Call SetStatus("## Dados de jogador carregado! Conta: " & SanitiseString(Trim$(Name)) & " Jogador: " & Trim$(Player(Index).Name) & " ##")
End Sub

Private Function ConvertByteToBoolean(Variavel As Byte) As Boolean
    If Variavel = YES Then
        ConvertByteToBoolean = True
    Else
        ConvertByteToBoolean = False
    End If
End Function
