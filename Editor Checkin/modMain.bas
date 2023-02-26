Attribute VB_Name = "modMain"
Option Explicit

Public MAX_ITEMS As Long

Public Const ITEM_PATH As String = "\data\items\"

Public Sub Main()
    Dim i As Byte
    
    Call LoadItems

    Call DayRewardInit

    With frmCheckIn

        For i = 1 To MONTHS
            .optMonth(i - 1).Caption = MonthName(i)
        Next i
        
        For i = 1 To 31
            .optDay(i - 1).Caption = i
        Next i
        
        ' Set data
        .optMonth(Month(Date) - 1).value = True
        
        .Show
    End With

End Sub

Sub LoadItems()
    Dim Filename As String
    Dim i As Long
    Dim F As Long
    
    MAX_ITEMS = CountFiles(App.Path & ITEM_PATH)
    ReDim Item(1 To MAX_ITEMS)

    For i = 1 To MAX_ITEMS
        Filename = App.Path & ITEM_PATH & "item" & i & ".dat"
        F = FreeFile
        Open Filename For Binary As #F
        Get #F, , Item(i)
        Close #F
        
        Item(i).Name = Trim$(Item(i).Name)
    Next

End Sub
