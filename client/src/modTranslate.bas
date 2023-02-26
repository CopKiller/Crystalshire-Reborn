Attribute VB_Name = "modTranslate"
Option Explicit

'Save the collection every 50 translations
Private Const saveEvery = 50
'Max of 6 translations per second. Good luck hitting that very often!
Private Const intTransPerSec As Integer = 10
'Sleep the server every 1ms until it's able to translate again.
Public Const blWaitToTranslate As Boolean = True

'these could be const, but I didn't want them to be.
Public LangTo As String
Public LangFrom As String
Public strTransPath As String
Public strOrigPath As String

Public Enum UnTrimType

    UnTrimFront = 1
    UnTrimBack = 2
    UnTrimBoth = 3

End Enum

'last number of total translations saved at
Private lastSave As Long
'timer for the translations
Private tmrTrans As Currency
'number of translations done in the last second
Private TransCount As Integer
'the collection for the current language.
'You can make multiples of these and pass them to each of the functions/subs that use them.
Public langCol As Collection
Public origCol As Collection

Public isLocked As Boolean

Private currentTranslation As String

Private T As GTranslate.DLL

'I try to use byref wherever possible here to prevent VB from having to
'make copies of strings continuously just to process them.
'However, I realized that I was trimming the text, which might not always
'be what the server wants to happen!
'it should actually lead to a small speed increase,
'but you have to be careful to not modify the original string.

Public Function GetTranslation(ByRef Text As String, Optional transLock As Boolean, Optional UnTrim As UnTrimType) As String
'weed out easy stuff
Dim txtTemp As String
txtTemp = Trim$(Text) 'if there are spaces on the ends, trim them
If LenB(txtTemp) <= 1 Then GetTranslation = txtTemp: Exit Function 'if the length of the string is <=1 then we aren't translating it.
If IsNumeric(txtTemp) = True Then GetTranslation = Text: Exit Function 'if it's a number.. we aren't translating it.


'I didn't feel like creating a translation queue or something of the sort,
'so this is the simple method to prevent translating the same thing multiple times while
'a translation is pending. Especially an issue when showing item descriptions/names.
If isLocked = True Then
    AddLog "Trying to translate when transLock is active!"
    Exit Function
End If

If transLock = True Then isLocked = True

    If txtTemp = currentTranslation Then
    
        AddLog "Trying to translate the same thing multiple times!"
        Exit Function
    End If
    
currentTranslation = txtTemp

'some loaded data will contain nullchars, which can waste time translating.
If InStr(1, txtTemp, vbNullChar, vbBinaryCompare) <> 0 Then
    txtTemp = Replace(txtTemp, vbNullChar, "") 'replace the nullchars.
    currentTranslation = vbNullString
    isLocked = False
    If LenB(txtTemp) <= 1 Then GetTranslation = txtTemp: Exit Function 'if the new length is too short, we're not translating it.
End If

'get the actual translation (either from cache, or from a translation service)
GetTranslation = Translate(txtTemp)

'currentTranslation = vbNullString
'isLocked = False

'checking for a new line in the text first is faster than
'simply running the replace on every translation
If InStr(1, GetTranslation, "\r\n", vbBinaryCompare) <> 0 Then
'Some lines in this game require this to look right!
    GetTranslation = Replace(GetTranslation, "\r\n", vbNewLine)
End If

Select Case UnTrim

    Case UnTrimType.UnTrimFront
        GetTranslation = " " & GetTranslation
    
    Case UnTrimType.UnTrimBack
        GetTranslation = GetTranslation & " "
    
    Case UnTrimType.UnTrimBoth
        GetTranslation = " " & GetTranslation & " "

End Select

'release the lock and reset current translation
currentTranslation = vbNullString
isLocked = False

End Function

Private Sub AddLog(Text As String)
With frmTransLog.txtLog

If frmTransLog.txtLog.visible = False Then Exit Sub

    .SelText = vbCrLf & Time & ": " & Text
    '.Text = .Text & vbCrLf & Time & ": " & Text
End With
End Sub

Private Function Translate(ByRef Text As String) As String
Dim strTranslation As String
Dim strHash As String
Dim SleepTime As Long
Dim i As Long

'init our translation/md5 dll
If T Is Nothing Then
AddLog "Init GTranslate.dll.. (If you get an activeX error at this point, it's probably not registered properly!)"
Set T = New GTranslate.DLL
AddLog "Init complete. Your GTranslate.dll is registered properly!"
End If
'init the collection for the lang if it doesn't exist (and load our current language)
If langCol Is Nothing Then Set langCol = New Collection: loadLang strTransPath, langCol
If origCol Is Nothing Then Set origCol = New Collection: loadLang strOrigPath, origCol

'get the md5 of our current string
strHash = T.GetMD5Hash(Text)
'read from the file to see if it's already been translated
strTranslation = ReadFromCache(strHash, langCol)
'If the length of it is 0, translate it.
If LenB(strTranslation$) = 0 Then

StartOver:
    'Check to see if we're over the timer, if we are, reset it.
    If getTime > tmrTrans Then
        tmrTrans = getTime + 1000
        TransCount = 0
    End If
    
    'check to see if we can do more translations this second
        If TransCount + 1 > intTransPerSec Then
            'too many translations. bail. or wait.
            If blWaitToTranslate = True Then
                AddLog "Over translate quota. Sleeping.. "
                'calculate the sleeptime from now to when we can do another translation
                SleepTime = (tmrTrans) - (getTime + 1)
                'sleep 1ms every iteration between 0 and sleeptime
                For i = 0 To SleepTime
                    'this should prevent problems for players from small lag maybe?
                    Sleep 1
                    DoEvents
                Next i
                'go back up a bit and try again.
                GoTo StartOver
            Else
                'blWaitToTranslate is false. Return untranslated text.
                'This is the best setting for a populated server, as otherwise
                'it would slow down a bit and lag for people.
                AddLog "Skipping translation; over quota.."
                Translate = Text
                Exit Function
            End If
        Else
        
        'russian roulette! WHOEVER WINS, GETS TO TRANSLATE FOR US! (not truly random :o)
        Select Case Rand(0, 2)

        Case Is = 0
            Translate = T.BingTranslate(LangTo, LangFrom, Text, "myBTranslate", "zgQQfksRpj8H60LVHq4afeHtmVTldKrE7PQxRnqxOy4=")
            AddLog "Translated(Bing): [" & Text & "] to [" & Translate & "]"
        Case Is = 1
            Translate = T.GoogleTranslate(LangTo, LangFrom, Text)
            AddLog "Translated(Google): [" & Text & "] to [" & Translate & "]"
        Case Is = 2
            Translate = T.YandexTranslate(LangTo, LangFrom, Text, "trnsl.1.1.20141229T202549Z.5f61901044d9ab3e.4d5c2d268897918f1adbfa15eb58b66d970ecbef")
            AddLog "Translated(Yandex): [" & Text & "] to [" & Translate & "]"
        End Select
        DoEvents
        
        'for now, if it's blank, just return the original text. However, this means that a translation error happened most likely.
        If LenB(Translate) <= 1 Then
        Translate = Text
        Exit Function
        End If
        
        'check that an error didn't occur.. log it to server log?
            AddToCache strHash, Translate, langCol ', Text
            AddToCache strHash, Text, origCol
        'uncomment this to save the collection every time a new translation is made, but be careful as it could get slow..
            'saveLang strTransPath, langCol
            'saveLang strOrigPath, origCol
        'increase the number of translations currently
            TransCount = TransCount + 1
        
        End If

Else
Translate = strTranslation
If Exists(origCol, strHash) = False Then
'if it's not in the cache, let's add it.
'This shouldn't happen, but I didn't have a separate collection
'for the original untranslated text.
    AddToCache strHash, Text, origCol
End If

AddLog "Cached: [" & strTranslation & "] original: [" & Text & "]"

End If

If (lastSave + saveEvery) < (langCol.Count) And (lastSave + saveEvery) < (origCol.Count) Then
    saveLang strTransPath, langCol
    saveLang strOrigPath, origCol
End If
'Set T = Nothing

End Function

Public Function ReadFromCache(ByRef strHash As String, ByRef col As Collection) As String
Dim Temp() As String

If col.Count = 1 Then ReadFromCache = "": Exit Function

If Exists(col, strHash) = True Then ReadFromCache = col.Item(strHash)(1)

End Function

Private Sub AddToCache(ByRef strHash As String, ByRef Translate As String, ByRef col As Collection)
On Error Resume Next
Dim Temp(1) As String

buildArray strHash, Translate, Temp
col.Add Temp, Temp(0)

End Sub

Public Sub loadLang(path As String, ByRef col As Collection)
'Dim strTemp As String
Dim Temp() As Byte
Dim tempArray(1) As String
Dim Buffer As New clsBuffer
Dim lngBufferCount As Long
Dim NF As Integer
Dim NotNull As Boolean
Dim bfFail As Boolean
NF = FreeFile

AddLog "Loading Lang file: " & path

' check exists
    Open path For Binary As NF
    Close NF

Temp = ReadFile(path, NotNull)

If NotNull = False Then Exit Sub

    Temp = Decompress(Temp, bfFail)
    If bfFail = True Then GoTo skip
    Buffer.WriteBytes Temp
    
lngBufferCount = Buffer.ReadLong
lastSave = lngBufferCount
Dim i As Long
For i = 1 To lngBufferCount
    buildArray Buffer.ReadString, Buffer.ReadString, tempArray
    col.Add tempArray, tempArray(0)
Next

'set our lastSave variable so we can save again in 50 translations
lastSave = col.Count

skip:
Set Buffer = Nothing
End Sub

Public Sub saveLang(path As String, ByRef col As Collection, Optional blForceSave As Boolean = False)
Dim NF As Integer
Dim tempOut() As Byte
Dim Buffer As New clsBuffer
Dim i As Long
NF = FreeFile

If col Is Nothing Then Exit Sub

If blForceSave = False Then If (lastSave) = (langCol.Count) Then Exit Sub
AddLog "Saving lang to: " & path
Buffer.WriteLong col.Count

For i = 1 To col.Count
'write the key first
    Buffer.WriteString (col.Item(i)(0))
'write the actual translation
    Buffer.WriteString (col.Item(i)(1))
Next

'write buffer to temp out
tempOut = Compress(Buffer.ReadBytes(Buffer.length))

    Open path For Binary As #NF
    Put #NF, , tempOut
    Close #NF
    
lastSave = langCol.Count
End Sub

Private Sub buildArray(ByRef Key As String, ByRef Text As String, ByRef myArr() As String)

myArr(0) = Key
myArr(1) = Text

End Sub

'use a buffer class to write the
'key length, then the key. use long - string.
'the variable length, then the variable. use long - string
'write them into a collection. How long can the keys be?

Public Function ReadFile(sFile As String, Optional ByRef NotNull As Boolean) As Byte()
    Dim nFile As Integer

    nFile = FreeFile
    Open sFile For Binary Access Read As #nFile
    If LOF(nFile) > 0 Then
        ReDim ReadFile(0 To LOF(nFile) - 1)
        Get nFile, , ReadFile
        NotNull = True
    Else
        NotNull = False
    End If
    Close #nFile
End Function

Public Sub debugLangFile(path As String)

Dim NF As Integer
Dim tempOut() As Byte
Dim Buffer As New clsBuffer
NF = FreeFile

Dim col As Collection
Set col = New Collection

Dim myArr(1) As String
col.Add Array("0a-6d-d0-dd-a2-ee-52-6b-57-55-6b-68-97-33-4a-b1", "Level 40-50"), "0a-6d-d0-dd-a2-ee-52-6b-57-55-6b-68-97-33-4a-b1"
col.Add Array("da-f8-9b-9c-2d-b8-51-d6-91-84-f0-95-6a-44-a0-d3", "Level 5-10"), "da-f8-9b-9c-2d-b8-51-d6-91-84-f0-95-6a-44-a0-d3"
col.Add Array("7e-54-5b-21-c5-a5-49-23-4b-ca-43-23-32-cf-54-82", "Level 10-20"), "7e-54-5b-21-c5-a5-49-23-4b-ca-43-23-32-cf-54-82"
Buffer.WriteLong col.Count

Dim i As Long


For i = 1 To col.Count
'write the key first
    Buffer.WriteString (col.Item(i)(0))
'write the actual translation
    Buffer.WriteString (col.Item(i)(1))
Next i

'write buffer to temp out
tempOut = Compress(Buffer.ReadBytes(Buffer.length))

    Open path For Binary As #NF
    Put #NF, , tempOut
    Close #NF

Set col = Nothing

End Sub

Public Function Exists(ByRef col As Collection, ByRef Index) As Boolean
On Error GoTo ExistsTryNonObject
    Dim o As Object

    Set o = col(Index)
    Exists = True
    Exit Function

ExistsTryNonObject:
    Exists = ExistsNonObject(col, Index)
End Function

Private Function ExistsNonObject(ByRef col As Collection, ByRef Index) As Boolean
On Error GoTo ExistsNonObjectErrorHandler
    Dim v As Variant

    v = col(Index)
    ExistsNonObject = True
    Exit Function

ExistsNonObjectErrorHandler:
    ExistsNonObject = False
End Function

