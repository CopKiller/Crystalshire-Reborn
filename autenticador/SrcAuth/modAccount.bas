Attribute VB_Name = "modLogic"
' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    filename = App.Path & "\accounts\" & SanitiseString(Trim$(Name)) & ".bin"

    If FileExist(filename) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    If AccountExist(Name) Then
        filename = App.Path & "\accounts\" & SanitiseString(Trim$(Name)) & ".bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Function SanitiseString(ByVal theString As String) As String
    Dim I As Long, tmpString As String
    tmpString = vbNullString
    If Len(theString) <= 0 Then Exit Function
    For I = 1 To Len(theString)
        Select Case Mid$(theString, I, 1)
        Case "*"
            tmpString = tmpString + "[s]"
        Case ":"
            tmpString = tmpString + "[c]"
        Case Else
            tmpString = tmpString + Mid$(theString, I, 1)
        End Select
    Next
    SanitiseString = tmpString
End Function

Public Function FindEmail(ByVal Email As String) As Boolean
    Dim F As Long
    Dim s As String
    Dim g() As String

    F = FreeFile
    Open App.Path & "\emailList.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        g = Split(s, ":")

        If Trim$(LCase(g(0))) = Trim$(LCase$(Email)) Then
            FindEmail = True
            Close #F
            Exit Function
        End If
    Loop

    Close #F
End Function

' Convert date string to long value
Public Function ConvertBirthDayToLng(ByVal BirthDay As String) As Long
    Dim tmpString() As String, tmpChar As String
    Dim I As Integer, CountCharacters As Integer

    ' Return 0 if have a problem in format
    ConvertBirthDayToLng = 0

    '00/00/0000 <- Length = 10, Verify if is a valid Date Format!   '
    If Len(BirthDay) < 10 Or Len(BirthDay) > 10 Then Exit Function  '

    ' Verify if have 8 numerics characters! \
    For I = 1 To Len(BirthDay)                 '\
        If IsNumeric(Mid$(BirthDay, I, 1)) Then    '\
            CountCharacters = CountCharacters + 1       '\
        End If                                              '\
    Next I                                                      '\
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
    For I = 0 To UBound(tmpString)
        tmpChar = tmpChar + tmpString(I)
    Next I

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
    Dim I As Long
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
        Player(Index).X = Class(ClassNum).START_X
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

