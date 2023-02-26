Attribute VB_Name = "modGameToAuth"
Option Explicit

' Packets sent by Server to Auth Server
Public Enum GSPackets
    GShutDown
    GSavePlayer
    GSaveInv
    GSaveBank
    ' Make sure GMSG_COUNT is below everything else
    GMSG_COUNT
End Enum

Public GS_HandleDataSub(GMSG_COUNT) As Long

Private GS_Buffer As New clsBuffer
Private GS_DataTimer As Long
Private GS_DataBytes As Long
Private GS_DataPackets As Long

Public Sub GS_InitMessages()
    GS_HandleDataSub(GShutDown) = GetAddress(AddressOf HandleShutDown)
    GS_HandleDataSub(GSavePlayer) = GetAddress(AddressOf HandleSavePlayer)
    GS_HandleDataSub(GSaveInv) = GetAddress(AddressOf HandleSaveInv)
    GS_HandleDataSub(GSaveBank) = GetAddress(AddressOf HandleSaveBank)
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

Private Sub HandleSavePlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, DataSize As Long, TempData() As Byte
    Dim FileName As String, F As Long
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

Private Sub HandleSaveInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, DataSize As Long, TempData() As Byte
    Dim FileName As String, F As Long
    Dim AccountName As String, Save As InvRec

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    AccountName = Buffer.ReadString
    
    DataSize = LenB(Save)
    ReDim TempData(DataSize - 1)
    TempData = Buffer.ReadBytes(DataSize)
    CopyMemory ByVal VarPtr(Save), ByVal VarPtr(TempData(0)), DataSize

    Buffer.Flush: Set Buffer = Nothing
    
    Call SaveInv_ByGameServer(AccountName, Save)
End Sub

Private Sub HandleSaveBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, DataSize As Long, TempData() As Byte
    Dim FileName As String, F As Long
    Dim AccountName As String, Save As BankRec

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    AccountName = Buffer.ReadString
    
    DataSize = LenB(Save)
    ReDim TempData(DataSize - 1)
    TempData = Buffer.ReadBytes(DataSize)
    CopyMemory ByVal VarPtr(Save), ByVal VarPtr(TempData(0)), DataSize

    Buffer.Flush: Set Buffer = Nothing
    
    Call SaveBank_ByGameServer(AccountName, Save)
End Sub

Private Function SaveBank_ByGameServer(ByVal AccountName As String, ByRef Jogador As BankRec)
    Dim FileName As String
    Dim F As Long
    

    FileName = App.Path & "\bank\" & AccountName & ".bin"

    F = FreeFile

    Open FileName For Binary As #F
    Put #F, , Jogador
    Close #F
End Function

Private Function SaveInv_ByGameServer(ByVal AccountName As String, ByRef Jogador As InvRec)
    Dim FileName As String
    Dim F As Long
    

    FileName = App.Path & "\inv\" & AccountName & ".bin"

    F = FreeFile

    Open FileName For Binary As #F
    Put #F, , Jogador
    Close #F
End Function

Private Function SavePlayer_ByGameServer(ByVal AccountName As String, ByRef Jogador As PlayerRec)
    Dim FileName As String
    Dim F As Long
    

    FileName = App.Path & "\accounts\" & AccountName & ".bin"

    F = FreeFile

    Open FileName For Binary As #F
    Put #F, , Jogador
    Close #F
    
    Call SetStatus("## Dados de jogador salvo! Conta: " & AccountName & " Jogador: " & Trim$(Jogador.Name) & " ##")
End Function

Sub LoadInv(ByVal Index As Long, ByVal Name As String)
    Dim FileName As String
    Dim F As Long

    If Trim$(Name) = vbNullString Then Exit Sub
    
    Call ClearInv(Index)

    FileName = App.Path & "\inv\" & SanitiseString(Trim$(Name)) & ".bin"
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , Inv(Index)
    Close #F
End Sub

Sub ClearInv(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Inv(Index)), LenB(Inv(Index)))
    Inv(Index).Login = vbNullString
End Sub

Sub LoadBank(ByVal Index As Long, ByVal Name As String)
    Dim FileName As String
    Dim F As Long

    If Trim$(Name) = vbNullString Then Exit Sub

    Call ClearBank(Index)
    
    FileName = App.Path & "\bank\" & SanitiseString(Trim$(Name)) & ".bin"
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , Bank(Index)
    Close #F
End Sub

Sub ClearBank(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(Index)), LenB(Bank(Index)))
    Bank(Index).Login = vbNullString
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim FileName As String
    Dim F As Long

    If Trim$(Name) = vbNullString Then Exit Sub

    Call ClearPlayer(Index)
    FileName = App.Path & "\accounts\" & SanitiseString(Trim$(Name)) & ".bin"
    F = FreeFile
    Open FileName For Binary As #F
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
