Attribute VB_Name = "modConvs"
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Conv(1 To MAX_CONVS) As ConvWrapperRec

Private Type ConvRec
    Conv As String * DESC_LENGTH
    rText(1 To 4) As String * DESC_LENGTH
    rTarget(1 To 4) As Long
    Event As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type

Private Type ConvWrapperRec
    Name As String * NAME_LENGTH
    chatCount As Long
    Conv() As ConvRec
End Type

' ***********
' ** Convs **
' ***********
Sub SaveConvs()
    Dim i As Long

    For i = 1 To MAX_CONVS
        Call SaveConv(i)
    Next
End Sub

Sub SaveConv(ByVal convNum As Long)
    Dim filename As String
    Dim i As Long, X As Long, F As Long

    filename = App.Path & "\data\convs\conv" & convNum & ".dat"
    F = FreeFile

    Open filename For Binary As #F
    With Conv(convNum)

        Put #F, , .Name
        Put #F, , .chatCount

        For i = 1 To .chatCount
            Put #F, , .Conv(i).Conv

            For X = 1 To 4
                Put #F, , .Conv(i).rText(X)
                Put #F, , .Conv(i).rTarget(X)
            Next

            Put #F, , .Conv(i).Event
            Put #F, , .Conv(i).Data1
            Put #F, , .Conv(i).Data2
            Put #F, , .Conv(i).Data3
        Next
    End With
    Close #F
End Sub
Sub LoadConvs()
    Dim filename As String
    Dim i As Long, n As Long, X As Long, F As Long
    Dim sLen As Long

    Call CheckConvs

    For i = 1 To MAX_CONVS
        filename = App.Path & "\data\convs\conv" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        With Conv(i)
            Get #F, , .Name
            Get #F, , .chatCount

            If .chatCount > 0 Then ReDim .Conv(1 To .chatCount)

            For n = 1 To .chatCount
                Get #F, , .Conv(n).Conv

                For X = 1 To 4
                    Get #F, , .Conv(n).rText(X)
                    Get #F, , .Conv(n).rTarget(X)
                Next

                Get #F, , .Conv(n).Event
                Get #F, , .Conv(n).Data1
                Get #F, , .Conv(n).Data2
                Get #F, , .Conv(n).Data3
            Next
        End With
        Close #F
    Next
End Sub

Sub CheckConvs()
    Dim i As Long

    For i = 1 To MAX_CONVS
        If Not FileExist("\data\convs\conv" & i & ".dat") Then
            Call SaveConv(i)
        End If
    Next
End Sub

Sub ClearConv(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Conv(Index)), LenB(Conv(Index)))
    Conv(Index).Name = vbNullString
    ReDim Conv(Index).Conv(1)
End Sub

Sub ClearConvs()
    Dim i As Long

    For i = 1 To MAX_CONVS
        Call ClearConv(i)
    Next

End Sub

Sub SendConvs(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_CONVS
        If LenB(Trim$(Conv(i).Name)) > 0 Then
            Call SendConvTo(Index, i)
        End If
    Next
End Sub
