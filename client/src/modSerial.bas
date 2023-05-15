Attribute VB_Name = "modSerial_TPC"
Option Explicit

'serial
Public Const MAX_SERIAL_NUMBER As Byte = 100

Public Const MAX_SERIAL_ITEMS As Byte = 10

Public Serial_Changed(1 To MAX_SERIAL_NUMBER) As Boolean

'map Serial
Public SerialEditorNum As Long

Public Serial(1 To MAX_SERIAL_NUMBER) As SerialRec

Public Type SerialRec
    ' INIT
    Name As String * NAME_LENGTH
    Serial As String * NAME_LENGTH
    ' CONFIG
    NamePlayer As String * NAME_LENGTH
    GiveOne As Byte
    Blocked As Byte
    BirthDay As Byte
    ' ITEMS
    Item(1 To MAX_SERIAL_ITEMS) As Integer
    ItemValue(1 To MAX_SERIAL_ITEMS) As Long
    ' ADICIONAIS
    VipDays As Integer
    GiveSpell As Integer
    GiveGuildSlot As Byte
    ' MSG
    Msg As String * DESC_LENGTH
End Type

Public Sub HandleSerialWindow(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ShowWindow GetWindowIndex("winSerial")
End Sub

Public Sub HandleUpdateSerial(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim SerialSize As Long
    Dim SerialData() As Byte
    Dim DecompData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    DecompData = buffer.UnCompressData
    Set buffer = Nothing

    Set buffer = New clsBuffer
    buffer.WriteBytes DecompData

    n = buffer.ReadLong
    ' Update the item
    SerialSize = LenB(Serial(n))
    ReDim SerialData(SerialSize - 1)
    SerialData = buffer.ReadBytes(SerialSize)
    CopyMemory ByVal VarPtr(Serial(n)), ByVal VarPtr(SerialData(0)), SerialSize
    Set buffer = Nothing
End Sub

Public Sub HandleSerialEditor()
    Dim i As Long

    With frmEditor_Serial
        Editor = EDITOR_SERIAL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SERIAL_NUMBER
            .lstIndex.AddItem i & ": " & Trim$(Serial(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SerialEditorInit
    End With

End Sub

'////////////////
'// Client TCP //
'////////////////

Public Sub SendSaveSerial(ByVal serialNum As Long)
    Dim buffer As clsBuffer
    Dim SerialSize As Long
    Dim SerialData() As Byte

    Set buffer = New clsBuffer
    SerialSize = LenB(Serial(serialNum))
    ReDim SerialData(SerialSize - 1)
    CopyMemory SerialData(0), ByVal VarPtr(Serial(serialNum)), SerialSize
    buffer.WriteLong CSaveSerial
    buffer.WriteLong serialNum
    buffer.WriteBytes SerialData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditSerial()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditSerial
    SendData buffer.ToArray()
    Set buffer = Nothing

End Sub

'////////////////
'// Database   //
'////////////////
Sub ClearSerial(ByVal Index As Long)

    Call ZeroMemory(ByVal VarPtr(Serial(Index)), LenB(Serial(Index)))
    Serial(Index).Name = vbNullString
    Serial(Index).Serial = vbNullString
    Serial(Index).Msg = vbNullString
    Serial(Index).NamePlayer = vbNullString

End Sub

Sub ClearSerials()
    Dim i As Long

    For i = 1 To MAX_SERIAL_NUMBER
        Call ClearSerial(i)
    Next

End Sub

'///////////////////
'// Game Editor   //
'///////////////////

Public Sub SerialEditorInit()
    Dim i As Integer
    If frmEditor_Serial.visible = False Then Exit Sub
    EditorIndex = frmEditor_Serial.lstIndex.ListIndex + 1

    With frmEditor_Serial

        .txtName.text = Trim$(Serial(EditorIndex).Name)
        .txtSerial = Trim$(Serial(EditorIndex).Serial)
        .txtPName = Trim$(Serial(EditorIndex).NamePlayer)
        .chkObtain = Serial(EditorIndex).GiveOne
        .chkBlocked = Serial(EditorIndex).Blocked
        .chkBirthday = Serial(EditorIndex).BirthDay
        .txtDias = Serial(EditorIndex).VipDays
        .scrlTechnique = Serial(EditorIndex).GiveSpell
        .scrlGuildSlot = Serial(EditorIndex).GiveGuildSlot
        .txtMsg = Trim$(Serial(EditorIndex).Msg)

        ' Items
        .cmbItems.Clear
        .cmbItems.AddItem "No Items"
        .cmbItems.ListIndex = 0
        If .cmbItems.ListCount >= 0 Then
            For i = 1 To MAX_ITEMS
                .cmbItems.AddItem (Trim$(Item(i).Name))
            Next
        End If
        ' re-load the list
        .lstItems.Clear
        For i = 1 To MAX_SERIAL_ITEMS
            If Serial(EditorIndex).Item(i) > 0 Then
                .lstItems.AddItem i & ": " & Serial(EditorIndex).ItemValue(i) & "x " & Trim$(Item(Serial(EditorIndex).Item(i)).Name)
            Else
                .lstItems.AddItem i & ": No Items"
            End If
        Next
        .lstItems.ListIndex = 0


    End With
    Serial_Changed(EditorIndex) = True
End Sub

Public Sub SerialEditorOk()
    Dim i As Long

    For i = 1 To MAX_SERIAL_NUMBER
        If Serial_Changed(i) Then
            Call SendSaveSerial(i)
        End If
    Next

    Unload frmEditor_Serial
    Editor = 0
    ClearChanged_Serial
End Sub

Public Sub SerialEditorCancel()
    Editor = 0
    Unload frmEditor_Serial
    ClearChanged_Serial
    ClearSerials
    SendRequestSerial
End Sub

Public Sub ClearChanged_Serial()
    ZeroMemory Serial_Changed(1), MAX_SERIAL_NUMBER * 2    ' 2 = boolean length
End Sub

Public Sub SendRequestSerial()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestSerial
    SendData buffer.ToArray()
    Set buffer = Nothing

End Sub
