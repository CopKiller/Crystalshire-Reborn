Attribute VB_Name = "modAccount"
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
