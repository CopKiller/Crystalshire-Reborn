Attribute VB_Name = "modCheckIn"
Option Explicit

Private Const MONTHS As Byte = 12

Public MonthReward(1 To MONTHS) As MonthRewardRec

Private Type DayRewardRec
    ItemNum As Integer
    ItemQuant As Long
End Type

Private Type MonthRewardRec
    DayReward() As DayRewardRec
End Type

Public Sub DayRewardInit()
    Dim FileName As String, I As Byte, n As Byte, GetDaysInMonth As Byte, GetYear As Integer
    Dim sPath As String, sFile As String

    ChkDir App.Path & "\data\", "checkin"

    FileName = App.Path & "\data\checkin\Year.ini"
    If Not FileExist(FileName, True) Then
        PutVar FileName, "ActualYear", "Year", Year(Date)
    Else
        GetYear = CInt(Trim$(GetVar(FileName, "ActualYear", "Year")))

        If GetYear <> Year(Date) Then
            PutVar FileName, "ActualYear", "Year", Year(Date)

            For I = 1 To MONTHS
                If FileExist(App.Path & "\data\checkin\Month" & I & "\DaysReward.ini", True) Then
                    Kill App.Path & "\data\checkin\Month" & I & "\DaysReward.ini"
                End If
            Next I
        End If
    End If

    For I = 1 To MONTHS

        ChkDir App.Path & "\data\checkin\", "month" & I
        GetDaysInMonth = DaysInMonth(Year(Date), I)

        ReDim MonthReward(I).DayReward(1 To GetDaysInMonth)

        FileName = App.Path & "\data\checkin\month" & I & "\daysreward.ini"
        If Not FileExist(FileName, True) Then
            For n = 1 To GetDaysInMonth
                'ItemNum
                PutVar FileName, "Day" & n, "ItemNum", 0
                'Quant
                PutVar FileName, "Day" & n, "ItemQuant", 0
            Next n
        End If
    Next I

    Call LoadDayReward
End Sub

Public Sub LoadDayReward()
    Dim I As Byte, n As Byte, FileName As String

    For I = 1 To MONTHS

        For n = 1 To UBound(MonthReward(I).DayReward)
            FileName = App.Path & "\data\checkin\month" & I & "\daysReward.ini"
            MonthReward(I).DayReward(n).ItemNum = CInt(Trim$(GetVar(FileName, "Day" & n, "ItemNum")))
            MonthReward(I).DayReward(n).ItemQuant = CLng(Trim$(GetVar(FileName, "Day" & n, "ItemQuant")))
        Next n

    Next I

End Sub

Public Sub SendDayReward(ByVal Index As Long)
    Dim I As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SSendDayReward

    Buffer.WriteByte Player(Index).CheckIn

    Buffer.WriteByte UBound(MonthReward(Month(Date)).DayReward)

    For I = 1 To UBound(MonthReward(Month(Date)).DayReward)
        Buffer.WriteInteger MonthReward(Month(Date)).DayReward(I).ItemNum
        Buffer.WriteLong MonthReward(Month(Date)).DayReward(I).ItemQuant
    Next I

    SendDataTo Index, Buffer.ToArray

    Buffer.Flush: Set Buffer = Nothing

End Sub

Public Sub HandleCheckIn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ItemNum As Integer, ItemQuant As Long
    
    Debug.Print UBound(MonthReward(Month(Date)).DayReward)

    ' Prevent errors
    If Player(Index).CheckIn = 0 Or Player(Index).CheckIn > UBound(MonthReward(Month(Date)).DayReward) Then
        Exit Sub
    End If
    
    If GetPlayerLastCheckIn(Index) = Date Then
        Call PlayerMsg(Index, "Você já realizou o CheckIn hoje, volte amanhã!", Yellow)
        Exit Sub
    End If

    ItemNum = MonthReward(Month(Date)).DayReward(Player(Index).CheckIn).ItemNum
    ItemQuant = MonthReward(Month(Date)).DayReward(Player(Index).CheckIn).ItemQuant

    If ItemNum > 0 Then
        If GiveInvItem(Index, ItemNum, ItemQuant, 0) Then
            Call PlayerMsg(Index, "Voce recebeu " & ItemQuant & Space(1) & Trim$(Item(ItemNum).Name), Yellow)
        Else
            Call PlayerMsg(Index, "Voce nao conseguiu realizar o CheckIn por falta de espaco na mochila!", Yellow)
        End If

    End If

    Player(Index).LastCheckIn = Date
    Player(Index).CheckIn = Player(Index).CheckIn + 1
End Sub

Public Function GetPlayerLastCheckIn(ByVal Index As Long) As Date
    
    If Not IsDate(Player(Index).LastCheckIn) Then Exit Function

    GetPlayerLastCheckIn = Player(Index).LastCheckIn
    
End Function

' Convert date string to long value
Public Function ConvertDateToLng(ByVal BirthDay As String) As Long
    Dim tmpString() As String, tmpChar As String
    Dim I As Integer, CountCharacters As Integer

    ' Return 0 if have a problem in format
    ConvertDateToLng = 0

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

    ConvertDateToLng = CLng(tmpChar)

End Function

Public Function DaysInMonth(ByVal Yr As Long, ByVal Mnth As Long) As Byte
' Return the number of days in the specified month.
    DaysInMonth = Day(DateSerial(Yr, Mnth + 1, 1) - 1)
End Function
