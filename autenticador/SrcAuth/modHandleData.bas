Attribute VB_Name = "modHandleData"
Option Explicit

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CAuthLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAuthAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CAuthAccountRecovery) = GetAddress(AddressOf HandleAccountRecovery)
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

' ::::::::::::::::::::::::::
' :: Account Recovery packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAccountRecovery(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim sEmail As String
    Dim vMAJOR As Long, vMINOR As Long, vREVISION As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    vMAJOR = Buffer.ReadLong
    vMINOR = Buffer.ReadLong
    vREVISION = Buffer.ReadLong
    
    sEmail = Buffer.ReadString
    
    Buffer.Flush: Set Buffer = Nothing
    
    ' right version
    If vMAJOR <> CLIENT_MAJOR Or vMINOR <> CLIENT_MINOR Or vREVISION <> CLIENT_REVISION Then
        SendAlertMsg Index, DIALOGUE_MSG_OUTDATED, MenuCount.menuLogin
        Exit Sub
    End If

   If FindEmail(sEmail) = False Then
        SendAlertMsg Index, DIALOGUE_ACCOUNT_EMAILINVALID, MenuCount.menuLogin
        Exit Sub
    End If

    Call SendEmail(Index, sEmail)
    
End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim Username As String
    Dim Password As String
    Dim vMAJOR As Long, vMINOR As Long, vREVISION As Long
    Dim i As Long
    Dim n As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Username = Buffer.ReadString
    Password = Buffer.ReadString
    vMAJOR = Buffer.ReadLong
    vMINOR = Buffer.ReadLong
    vREVISION = Buffer.ReadLong
    Name = Buffer.ReadString
    Sex = Buffer.ReadLong
    Class = Buffer.ReadLong
    Sprite = Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing

    If Username = vbNullString Then
        SendAlertMsg Index, DIALOGUE_MSG_CONNECTION
        Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Username)
        n = AscW(Mid$(Username, i, 1))
        If Not isNameLegal(n) Then
            Call SendAlertMsg(Index, DIALOGUE_MSG_USERILLEGAL, MenuCount.menuLogin)
            Exit Sub
        End If
    Next

    If Password = vbNullString Then
        SendAlertMsg Index, DIALOGUE_MSG_CONNECTION, MenuCount.menuLogin
        Exit Sub
    End If

    ' right version
    If vMAJOR <> CLIENT_MAJOR Or vMINOR <> CLIENT_MINOR Or vREVISION <> CLIENT_REVISION Then
        SendAlertMsg Index, DIALOGUE_MSG_OUTDATED, MenuCount.menuLogin
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

    Call LoadPlayer(Index, Username)
    
    '*********************VERIFY NEW CHAR**********************
    ' Prevent hacking
    If Len(Trim$(Name)) < 3 Or Len(Trim$(Name)) > NAME_LENGTH Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_NAMELENGTH, menuNewChar, False)
        Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Name)
        n = AscW(Mid$(Name, i, 1))
        If Not isNameLegal(n) Then
            Call SendAlertMsg(Index, DIALOGUE_MSG_NAMEILLEGAL, menuNewChar, False)
            Exit Sub
        End If
    Next

    ' Prevent hacking
    If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_CONNECTION)
        Exit Sub
    End If

    ' Prevent hacking
    If Class < 1 Or Class > Max_Classes Then
        Exit Sub
    End If

    ' Check if char already exists in slot
    If CharExist(Index) Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_CONNECTION)
        Exit Sub
    End If

    ' Check if name is already in use
    If FindChar(Name) Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_NAMETAKEN, menuNewChar, False)
        Exit Sub
    End If
    '**********************************************************

    ' Everything went ok, add the character
    Call AddChar(Index, Name, Sex, Class, Sprite)
    Call SetStatus("Character " & Name & " added to " & Trim$(Player(Index).Login) & "'s account.")
    
    Call Login(Index)
End Sub

Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Username As String, Password As String
    Dim loginToken As String
    Dim vMAJOR As Long, vMINOR As Long, vREVISION As Long
    Dim i As Long, n As Long, filename As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    Username = Buffer.ReadString

    Password = Buffer.ReadString

    vMAJOR = Buffer.ReadLong
    vMINOR = Buffer.ReadLong
    vREVISION = Buffer.ReadLong

    Buffer.Flush: Set Buffer = Nothing

    '****************VERIFY LOGIN*********************
    If Username = vbNullString Then
        SendAlertMsg Index, DIALOGUE_MSG_CONNECTION
        Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Username)
        n = AscW(Mid$(Username, i, 1))
        If Not isNameLegal(n) Then
            Call SendAlertMsg(Index, DIALOGUE_MSG_USERILLEGAL, MenuCount.menuLogin)
            Exit Sub
        End If
    Next

    If Password = vbNullString Then
        SendAlertMsg Index, DIALOGUE_MSG_CONNECTION
        Exit Sub
    End If

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

    Call LoadPlayer(Index, Username)

    ' make sure they're not banned
    If isBanned_Account(Index) Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_BANNED, MenuCount.menuLogin)
        Exit Sub
    End If

    Call Login(Index)
End Sub

Sub Login(ByVal Index As Long)
    Dim i As Long
    Dim loginToken As String
    
    If isShuttingDown Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_REBOOTING)
        Exit Sub
    End If

    If LenB(Trim$(Player(Index).Name)) > 0 Then

        TempPlayer(Index).TokenAccepted = True
        ' Everything passed, create the token and send it off
        loginToken = RandomString("AN-##AA-ANHHAN-H")
        SendLoginTokenToGameServer Index, Trim$(Player(Index).Login), loginToken
        SendLoginTokenToPlayer Index, loginToken
    Else
        Call SendNewCharClasses(Index)
    End If

    If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents

    CloseSocket Index
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String, Pass As String, mail As String, i As Long, n As Long, Major As Long, Minor As Long, Revision As Long, BirthDay As Date

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    ' Get the data
    Name = Buffer.ReadString
    Pass = Buffer.ReadString
    mail = Buffer.ReadString
    BirthDay = CDate(Buffer.ReadString)

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

    If Not IsDate(BirthDay) Then
        Call SendAlertMsg(Index, DIALOGUE_BIRTHDAY_INCORRECT, MenuCount.menuRegister)
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

    ' Check if email is already in use
    If FindEmail(mail) Then
        Call SendAlertMsg(Index, DIALOGUE_ACCOUNT_EMAILTAKEN, MenuCount.menuRegister)
        Exit Sub
    End If

    If AccountExist(Name) Then
        Call SendAlertMsg(Index, DIALOGUE_MSG_NAMETAKEN, MenuCount.menuRegister)
        Exit Sub
    Else
        Call AddAccount(Index, Name, Pass, mail, BirthDay)
        Call SetStatus("-> Nova Conta: " & Name)
        Call SendAlertMsg(Index, DIALOGUE_ACCOUNT_CREATED, MenuCount.menuLogin)
    End If
    Exit Sub
End Sub
