Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, value As String)
    Call WritePrivateProfileString$(Header, Var, value, File)
End Sub

Public Function FileExist(ByVal Filename As String, Optional RAW As Boolean = False) As Boolean
    If Not RAW Then
        If LenB(Dir(App.Path & "\" & Filename)) > 0 Then
            FileExist = True
        End If
    Else
        If LenB(Dir(Filename)) > 0 Then
            FileExist = True
        End If
    End If
End Function

Public Function CountFiles(ByVal Folder As String) As Long
    Dim s As String
    Dim z As Long
    
    s = Dir(Folder & "\*.*")
    Do
        If Len(s) = 0 Then
            Exit Do
        End If
        If (GetAttr(Folder & "\" & s) And vbDirectory) = 0 Then
            z = z + 1
        End If
        s = Dir()
    Loop
    
    CountFiles = z

End Function

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

