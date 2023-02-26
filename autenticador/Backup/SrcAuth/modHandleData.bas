Attribute VB_Name = "modHandleData"
Option Explicit

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CAuthLogin) = GetAddress(AddressOf HandleLogin)
End Sub

Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer, MsgType As Long, packetCallback As Long
    
    Dim TempBuffer() As Byte
    
    TempBuffer = DecryptPacket(Data, (UBound(Data) - LBound(Data)) + 1)

    Set Buffer = New clsBuffer
    Buffer.WriteBytes TempBuffer

    MsgType = Buffer.ReadLong

    If (MsgType < 0) Or (MsgType >= CMSG_COUNT) Then
        HackingAttempt Index
        Exit Sub
    End If

    packetCallback = HandleDataSub(MsgType)
    
    If packetCallback <> 0 Then
        CallWindowProc HandleDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
    End If
    
    Buffer.Flush
    Set Buffer = Nothing
End Sub

Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Username As String, Password As String
    Dim loginToken As String
    Dim vMAJOR As Long, vMINOR As Long, vREVISION As Long
    Dim i As Long, n As Long, FileName As String

    If isShuttingDown Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_REBOOTING)
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    Username = Buffer.ReadString

    If Username = vbNullString Then
        SendAlertMsg Index, DIALOGUE_MSG_CONNECTION
        Buffer.Flush: Set Buffer = Nothing
        Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Username)
        n = AscW(Mid$(Username, i, 1))
        If Not isNameLegal(n) Then
            Call SendAlertMsg(Index, DIALOGUE_MSG_USERILLEGAL, MenuCount.menuLogin)
            Buffer.Flush: Set Buffer = Nothing
            Exit Sub
        End If
    Next

    Password = Buffer.ReadString

    If Password = vbNullString Then
        SendAlertMsg Index, DIALOGUE_MSG_CONNECTION
        Buffer.Flush: Set Buffer = Nothing
        Exit Sub
    End If

    vMAJOR = Buffer.ReadLong
    vMINOR = Buffer.ReadLong
    vREVISION = Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing

    ' right version
    If vMAJOR <> CLIENT_MAJOR Or vMINOR <> CLIENT_MINOR Or vREVISION <> CLIENT_REVISION Then
        SendAlertMsg Index, DIALOGUE_MSG_OUTDATED
        Exit Sub
    End If

    If Len(Username) < 3 Or Len(Password) < 3 Then
        SendAlertMsg Index, DIALOGUE_MSG_USERLENGTH, MenuCount.menuLogin
        Exit Sub
    End If

    ' username found
    If Not AccountExist(Username) Then
        SendAlertMsg Index, DIALOGUE_MSG_WRONGPASS, MenuCount.menuLogin
        Exit Sub
    End If

    ' check password
    If Not PasswordOK(Username, Password) Then
        SendAlertMsg Index, DIALOGUE_MSG_WRONGPASS, MenuCount.menuLogin
        Exit Sub
    End If

    TempPlayer(Index).TokenAccepted = True

    ' Everything passed, create the token and send it off
    loginToken = RandomString("AN-##AA-ANHHAN-H")

    SendLoginTokenToGameServer Index, Username, loginToken

    SendLoginTokenToPlayer Index, loginToken

    If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
    CloseSocket Index
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String, Pass As String, mail As String, i As Long, n As Long, Major As Long, Minor As Long, Revision As Long

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    ' Get the data
    Name = Buffer.ReadString
    Pass = Buffer.ReadString
    mail = Buffer.ReadString

    Major = Buffer.ReadLong
    Minor = Buffer.ReadLong
    Revision = Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing

    If isShuttingDown Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_REBOOTING)
        Exit Sub
    End If

    ' right version
    If Major <> CLIENT_MAJOR Or Minor <> CLIENT_MINOR Or Revision <> CLIENT_REVISION Then
        SendAlertMsg Index, DIALOGUE_MSG_OUTDATED
        Exit Sub
    End If

    If Len(Trim$(Name)) < 3 Or Len(Trim$(Pass)) < 3 Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_NAMELENGTH, MenuCount.menuRegister)
        Exit Sub
    End If

    If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Pass)) > NAME_LENGTH Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_NAMELENGTH, MenuCount.menuRegister)
        Exit Sub
    End If

    If InStr(1, mail, "@") = 0 Or Len(Trim$(mail)) < 4 Or Len(Trim$(mail)) > EMAIL_LENGTH Then
        Call SendAlertMsg(Index, DIALOGUE_ACCOUNT_EMAILINVALID, MenuCount.menuRegister)
        Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Name)
        n = AscW(Mid$(Name, i, 1))
        If Not isNameLegal(n) Then
            Call SendAlertMsg(Index, DIALOGUE_MSG_USERILLEGAL, MenuCount.menuRegister)
            Exit Sub
        End If
    Next

    If AccountExist(Name) Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_NAMETAKEN, MenuCount.menuRegister)
        Exit Sub
    Else
        Call AddAccount(Index, Name, Pass, mail)
        Call SendAlertMsg(Index, DIALOGUE_ACCOUNT_CREATED, MenuCount.menuLogin)
    End If
    Exit Sub
End Sub

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String, ByVal Code As String)
    Dim i As Long
    Dim F As Long, FileName As String
    
    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
    
    ClearPlayer Index
    ClearInv Index
    ClearBank Index

    Player(Index).Login = Name
    Player(Index).Password = Password
    Player(Index).mail = Code
    
    Inv(Index).Login = Name
    Bank(Index).Login = Name

    ' Append name to file
    FileName = App.Path & "\emailList.txt"
    F = FreeFile
    Open FileName For Append As #F
    Print #F, Code & ":" & Password
    Close #F
    
    ' Save Player archive
    FileName = App.Path & "\accounts\" & SanitiseString(Trim$(Player(Index).Login)) & ".bin"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Player(Index)
    Close #F
    
    ' Save Inv archive
    FileName = App.Path & "\inv\" & SanitiseString(Trim$(Inv(Index).Login)) & ".bin"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Inv(Index)
    Close #F
    
    ' Save Bank archive
    FileName = App.Path & "\bank\" & SanitiseString(Trim$(Bank(Index).Login)) & ".bin"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Bank(Index)
    Close #F
    
    

End Sub
