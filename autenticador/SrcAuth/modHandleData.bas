Attribute VB_Name = "modHandleData"
Option Explicit

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CAuthLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAuthAddChar) = GetAddress(AddressOf HandleAddChar)
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

' Convert date string to long value
Private Function ConvertBirthDayToLng(ByVal BirthDay As String) As Long
    Dim tmpString() As String, tmpChar As String
    Dim i As Integer, CountCharacters As Integer

    ' Return 0 if have a problem in format
    ConvertBirthDayToLng = 0

    '00/00/0000 <- Length = 10, Verify if is a valid Date Format!   '
    If Len(BirthDay) < 10 Or Len(BirthDay) > 10 Then Exit Function  '

    ' Verify if have 8 numerics characters! \
    For i = 1 To Len(BirthDay)                 '\
        If IsNumeric(Mid$(BirthDay, i, 1)) Then    '\
            CountCharacters = CountCharacters + 1       '\
        End If                                              '\
    Next i                                                      '\
    If CountCharacters < 8 Or CountCharacters > 8 Then Exit Function    '\

    ' Verify if have 3 "/" or "\" or "-"
    If InStr(1, BirthDay, "/") = 3 Then
        tmpChar = "/"
    ElseIf InStr(1, BirthDay, "\") = 3 Then
        tmpChar = "\"
    ElseIf InStr(1, BirthDay, "-") = 3 Then
        tmpChar = "-"
    Else: Exit Function
    End If
    
    ' All OK, Go Split the string, to convert to long!
    tmpString = Split(BirthDay, tmpChar)

    tmpChar = vbNullString
    For i = 0 To UBound(tmpString)
        tmpChar = tmpChar + tmpString(i)
    Next i

    ConvertBirthDayToLng = CLng(tmpChar)

End Function

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim s As String
    F = FreeFile
    Open App.Path & "\_charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

Function CharExist(ByVal Index As Long) As Boolean
    If LenB(Trim$(Player(Index).Name)) > 0 Then
        CharExist = True
    End If
End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String, ByVal Code As String, ByVal BirthDay As Date)
    Dim i As Long
    Dim F As Long, filename As String
    
    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
    
    ClearPlayer Index

    Player(Index).Login = Name
    Player(Index).Password = Password
    Player(Index).mail = Code
    Player(Index).BirthDay = BirthDay

    ' Append name to file
    filename = App.Path & "\emailList.txt"
    F = FreeFile
    Open filename For Append As #F
    Print #F, Code & ":" & Password
    Close #F
    
    ' Save Player archive
    filename = App.Path & "\accounts\" & SanitiseString(Trim$(Player(Index).Login)) & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Player(Index)
    Close #F

End Sub

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long)
    Dim F As Long
    Dim n As Long

    If LenB(Trim$(Player(Index).Name)) = 0 Then

        Player(Index).Name = Name
        Player(Index).Sex = Sex
        Player(Index).Class = ClassNum
        Player(Index).Premium = NO
        Player(Index).StartPremium = vbNullString
        Player(Index).DaysPremium = 0

        If Player(Index).Sex = SEX_MALE Then
            Player(Index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(Index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(Index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(Index).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        Player(Index).Map = Class(ClassNum).START_MAP
        Player(Index).x = Class(ClassNum).START_X
        Player(Index).Y = Class(ClassNum).START_Y
        Player(Index).dir = 0
        Player(Index).Vital(Vitals.HP) = Class(ClassNum).MaxHP
        Player(Index).Vital(Vitals.MP) = Class(ClassNum).MaxMP

        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    Player(Index).Inv(n).Num = Class(ClassNum).StartItem(n)
                    Player(Index).Inv(n).value = Class(ClassNum).StartValue(n)
                End If
            Next
        End If

        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    Player(Index).Spell(n).Spell = Class(ClassNum).StartSpell(n)
                    Player(Index).Hotbar(n).Slot = Class(ClassNum).StartSpell(n)
                    Player(Index).Hotbar(n).sType = 2    ' spells
                End If
            Next
        End If

        ' Append name to file
        F = FreeFile
        Open App.Path & "\_charlist.txt" For Append As #F
        Print #F, Name
        Close #F
        Call SavePlayer_ByGameServer(Trim$(Player(Index).Login), Player(Index))
        Exit Sub
    End If

End Sub
