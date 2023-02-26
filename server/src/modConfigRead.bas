Attribute VB_Name = "modConfig"
' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

' Gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' Writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, value As String)
    Call WritePrivateProfileString$(Header, Var, value, File)
End Sub

' Load mysql config.
Public Sub ReadSqlConfig()
Dim SqlConfig As String

    SqlConfig = App.Path & "\SqlConfig.ini"

    strServer = GetVar(SqlConfig, "SQL", "Server")
    strUsername = GetVar(SqlConfig, "SQL", "Username")
    strPassword = GetVar(SqlConfig, "SQL", "Password")
    strPort = GetVar(SqlConfig, "SQL", "Port")
    strDatabase = GetVar(SqlConfig, "SQL", "Database")
    
    AuthCode = GetVar(SqlConfig, "SERVER", "AuthCode")

End Sub

Function SanitiseString(ByVal theString As String) As String
Dim i As Long, tmpString As String
    tmpString = vbNullString
    If Len(theString) <= 0 Then Exit Function
    For i = 1 To Len(theString)
        Select Case Mid$(theString, i, 1)
            Case "*"
                tmpString = tmpString + "[s]"
            Case ":"
                tmpString = tmpString + "[c]"
            Case Else
                tmpString = tmpString + Mid$(theString, i, 1)
        End Select
    Next
    SanitiseString = tmpString
End Function
