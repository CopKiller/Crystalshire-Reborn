Attribute VB_Name = "modCheckIn"
Option Explicit

Public Const MONTHS As Byte = 12

Public MonthReward(1 To MONTHS) As MonthRewardRec

Public CopyDayReward As DayRewardRec
Public CopyMonthReward As MonthRewardRec

Private Type DayRewardRec
    ItemNum As Integer
    ItemQuant As Long
End Type

Private Type MonthRewardRec
    DayReward() As DayRewardRec
End Type

Public Sub DayRewardInit()
    Dim Filename As String, i As Byte, n As Byte, GetDaysInMonth As Byte, GetYear As Integer

    ChkDir App.Path & "\data\", "checkin"
    
    Filename = App.Path & "\data\checkin\Year.ini"
    If Not FileExist(Filename, True) Then
        PutVar Filename, "ActualYear", "Year", Year(Date)
    Else
        GetYear = CInt(Trim$(GetVar(Filename, "ActualYear", "Year")))
        
        If GetYear <> Year(Date) Then
            PutVar Filename, "ActualYear", "Year", Year(Date)
            
            For i = 1 To MONTHS
                If FileExist(App.Path & "\data\checkin\Month" & i & "\DaysReward.ini", True) Then
                    Kill App.Path & "\data\checkin\Month" & i & "\DaysReward.ini"
                End If
            Next i
        End If
    End If

    For i = 1 To MONTHS

        ChkDir App.Path & "\data\checkin\", "month" & i
        GetDaysInMonth = DaysInMonth(Year(Date), i)

        ReDim MonthReward(i).DayReward(1 To GetDaysInMonth)

        Filename = App.Path & "\data\checkin\month" & i & "\daysreward.ini"
        If Not FileExist(Filename, True) Then
            For n = 1 To GetDaysInMonth
                'ItemNum
                PutVar Filename, "Day" & n, "ItemNum", 0
                'Quant
                PutVar Filename, "Day" & n, "ItemQuant", 0
            Next n
        End If
    Next i

    Call LoadDayReward
End Sub

Public Sub LoadDayReward()
    Dim i As Byte, n As Byte, Filename As String

    For i = 1 To MONTHS
        
        For n = 1 To UBound(MonthReward(i).DayReward)
            Filename = App.Path & "\data\CheckIn\Month" & i & "\DaysReward.ini"
            MonthReward(i).DayReward(n).ItemNum = CInt(Trim$(GetVar(Filename, "Day" & n, "ItemNum")))
            MonthReward(i).DayReward(n).ItemQuant = CLng(Trim$(GetVar(Filename, "Day" & n, "ItemQuant")))
        Next n

    Next i
    
End Sub

Public Function DaysInMonth(ByVal Yr As Long, ByVal Mnth As Long) As Byte
' Return the number of days in the specified month.
    DaysInMonth = Day(DateSerial(Yr, Mnth + 1, 1) - 1)
End Function

Public Sub SaveCheckIn()
    Dim Filename As String
    Dim i As Byte, n As Byte

    For i = 1 To UBound(MonthReward)

        Filename = App.Path & "\data\checkin\month" & i & "\DaysReward.ini"

        For n = 1 To UBound(MonthReward(i).DayReward)
            'ItemNum
            PutVar Filename, "Day" & n, "ItemNum", CStr(MonthReward(i).DayReward(n).ItemNum)
            'Quant
            PutVar Filename, "Day" & n, "ItemQuant", CStr(MonthReward(i).DayReward(n).ItemQuant)
        Next n
    Next i

End Sub

