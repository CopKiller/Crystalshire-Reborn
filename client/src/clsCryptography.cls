VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCryptography_Sem_Uso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Function Crypto(ByVal Text As String, ByVal KeyPair As Long, ByVal KeyOdd As Long) As String
    Dim intCont, nCar, cCar

    On Error GoTo errLabel

    If KeyPair < 1 Or KeyPair > 255 Then
        Crypto = vbNullString
        Exit Function
    End If

    If KeyOdd < 1 Or KeyOdd > 255 Then
        Crypto = vbNullString
        Exit Function
    End If

    cCar = ""

    For intCont = 1 To Len(Text)

        If intCont / 2 = Int(intCont / 2) Then
            nCar = Asc(Mid(Text, intCont, 1)) Xor KeyPair
        Else
            nCar = Asc(Mid(Text, intCont, 1)) Xor KeyOdd
        End If

        cCar = cCar + Chr(nCar)
    Next

    Crypto = cCar
    Exit Function

errLabel:
    Crypto = vbNullString
End Function

Public Function GetArray(ByVal Text As String, ByVal KeyPair As Long, ByVal KeyOdd As Long) As Long()
    Dim i As Long, Values() As Long, Lenght As Long
    Dim result As String

    result = Crypto(Text, KeyPair, KeyOdd)
    Lenght = Len(result)

    ReDim Values(1 To Lenght) As Long

    For i = 1 To Lenght
        Values(i) = AscW(Mid(result, i, 1))
    Next

    GetArray = Values
End Function

Public Function GetString(ByRef Values() As Long, ByVal KeyPair As Long, ByVal KeyOdd As Long) As String
    Dim i As Long, Lenght As Long, result As String

    '   On Error GoTo errLabel

    Lenght = UBound(Values)

    For i = 1 To Lenght
        result = result & ChrW(Values(i))
    Next

    GetString = Crypto(result, KeyPair, KeyOdd)
    Exit Function

    'errLabel:
    '  GetString = vbNullString
End Function

Private Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Sub WriteText(ByRef Text As String, ByRef Buffer As clsBuffer)
    Dim arr() As Long, Lenght As Long, i As Long
    Dim KeyPair As Long, KeyOdd As Long

    KeyPair = Rand(1, 255)
    KeyOdd = Rand(1, 255)

    arr = GetArray(Text, KeyPair, KeyOdd)
    Lenght = UBound(arr)

    Buffer.WriteLong Lenght
    Buffer.WriteLong KeyPair

    For i = 1 To Lenght
        Buffer.WriteLong arr(i)
    Next

    Buffer.WriteLong KeyOdd
End Sub

Public Function ReadText(ByRef Buffer As clsBuffer) As String
    Dim arr() As Long, Lenght As Long, i As Long
    Dim KeyPair As Long, KeyOdd As Long

    Lenght = Buffer.ReadLong

    If Lenght <= 0 Or Lenght >= MAX_LONG Then
        ReadText = vbNullString
        Exit Function
    End If

    KeyPair = Buffer.ReadLong

    ReDim arr(1 To Lenght) As Long

    For i = 1 To Lenght
        arr(i) = Buffer.ReadLong
    Next

    KeyOdd = Buffer.ReadLong

    ReadText = GetString(arr, KeyPair, KeyOdd)
End Function



