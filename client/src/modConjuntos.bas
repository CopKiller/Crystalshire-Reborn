Attribute VB_Name = "modConjuntos"
Option Explicit

Public Const MAX_CONJUNTOS As Byte = 50
Public Const ACTIONMSG_LENGTH As Byte = 100
Private Conjunto_Changed(1 To MAX_CONJUNTOS) As Boolean

Public Conjunto(1 To MAX_CONJUNTOS) As ConjuntoRec

Public UsingSet As Integer

' Type recs
Private Type BonusRec ' 8 Bonus Totais
    'Atributos
    Add_Stat(1 To Stats.Stat_Count - 1) As Long
    Add_Stat_Percent(1 To Stats.Stat_Count - 1) As Byte
    'Dano
    Dano As Long
    DanoPercent As Byte
    'Defesa
    Defesa As Long
    DefesaPercent As Byte
    'Exp
    EXP As Integer
    Drop As Byte
End Type

Private Type ActionsRec
    Msg As String * ACTIONMSG_LENGTH
    Animation As Integer
End Type

Private Type ConjuntoRec
    Name As String * NAME_LENGTH
    Item(1 To Equipment.Equipment_Count - 1) As Integer
    Bonus As BonusRec
    Actions As ActionsRec
End Type

Public Sub ConjuntoEditorCancel()
    Editor = 0
    Unload frmEditor_Conjuntos
    ClearChanged_Conjunto
    ClearConjuntos
    SendRequestConjuntos
End Sub

Private Sub ClearChanged_Conjunto()
    ZeroMemory Conjunto_Changed(1), MAX_CONJUNTOS * 2    ' 2 = boolean length
End Sub

Public Sub ClearConjunto(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Conjunto(Index)), LenB(Conjunto(Index)))
    Conjunto(Index).Name = vbNullString
End Sub

Private Sub ClearConjuntos()
    Dim i As Long

    For i = 1 To MAX_CONJUNTOS
        Call ClearConjunto(i)
    Next

End Sub

Private Sub SendRequestConjuntos()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestConjuntos
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub ConjuntoEditorOk()
    Dim i As Long

    For i = 1 To MAX_CONJUNTOS

        If Conjunto_Changed(i) Then
            Call SendSaveConjunto(i)
        End If

    Next

    Unload frmEditor_Conjuntos
    Editor = 0
    ClearChanged_Conjunto
End Sub

Public Sub SendRequestEditConjunto()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditConjunto
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Private Sub SendSaveConjunto(ByVal Conjuntonum As Long)
    Dim Buffer As clsBuffer
    Dim ConjuntoSize As Long
    Dim ConjuntoData() As Byte
    Set Buffer = New clsBuffer
    ConjuntoSize = LenB(Conjunto(Conjuntonum))
    ReDim ConjuntoData(ConjuntoSize - 1)
    CopyMemory ConjuntoData(0), ByVal VarPtr(Conjunto(Conjuntonum)), ConjuntoSize
    Buffer.WriteLong CSaveConjunto
    Buffer.WriteLong Conjuntonum
    Buffer.WriteBytes ConjuntoData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' /////////////////
' // Conjunto Editor //
' /////////////////
Public Sub ConjuntoEditorInit()
    Dim i As Long

    If frmEditor_Conjuntos.visible = False Then Exit Sub
    EditorIndex = frmEditor_Conjuntos.lstIndex.ListIndex + 1

    With Conjunto(EditorIndex)
        frmEditor_Conjuntos.txtName.Text = Trim$(.Name)

        ' Drop Items
        frmEditor_Conjuntos.cmbItems.Clear
        frmEditor_Conjuntos.cmbItems.AddItem "No Items"
        frmEditor_Conjuntos.cmbItems.ListIndex = 0
        If frmEditor_Conjuntos.cmbItems.ListCount >= 0 Then
            For i = 1 To MAX_ITEMS
                frmEditor_Conjuntos.cmbItems.AddItem (Trim$(Item(i).Name))
            Next
        End If
        ' re-load the list
        frmEditor_Conjuntos.lstItems.Clear
        For i = 1 To (Equipment_Count - 1)
            If .Item(i) > 0 Then
                frmEditor_Conjuntos.lstItems.AddItem i & ": " & Trim$(Item(.Item(i)).Name)
            Else
                frmEditor_Conjuntos.lstItems.AddItem i & ": No Items"
            End If
        Next

        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Conjuntos.txtStatBonus(i) = CLng(.Bonus.Add_Stat(i))
            frmEditor_Conjuntos.chkPercentStats(i) = CByte(.Bonus.Add_Stat_Percent(i))
        Next i
        frmEditor_Conjuntos.txtDamage = CLng(.Bonus.Dano)
        frmEditor_Conjuntos.chkPercentDamage = CByte(.Bonus.DanoPercent)
        frmEditor_Conjuntos.txtDefense = CLng(.Bonus.Defesa)
        frmEditor_Conjuntos.chkPercentDefense = CByte(.Bonus.DefesaPercent)
        frmEditor_Conjuntos.txtExp = CLng(.Bonus.EXP)
        frmEditor_Conjuntos.txtDrop = CByte(.Bonus.Drop)
        frmEditor_Conjuntos.scrlAnim = CInt(.Actions.Animation)
        frmEditor_Conjuntos.txtMsg = Trim$(.Actions.Msg)

        frmEditor_Conjuntos.lstItems.ListIndex = 0

        EditorIndex = frmEditor_Conjuntos.lstIndex.ListIndex + 1
    End With

    Conjunto_Changed(EditorIndex) = True
End Sub


' ////////////////////
' //////HANDLE////////
' ////////////////////
Public Sub HandleConjuntoEditor()
    Dim i As Long

    With frmEditor_Conjuntos
        Editor = EDITOR_CONJUNTO
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_CONJUNTOS
            .lstIndex.AddItem i & ": " & Trim$(Conjunto(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ConjuntoEditorInit
    End With

End Sub

Public Sub HandleUpdateConjunto(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ConjuntoSize As Long
    Dim ConjuntoData() As Byte
    Dim DecompData()   As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    DecompData = Buffer.UnCompressData
    Set Buffer = Nothing
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes DecompData
    
    n = Buffer.ReadLong
    ' Update the Conjunto
    ConjuntoSize = LenB(Conjunto(n))
    ReDim ConjuntoData(ConjuntoSize - 1)
    ConjuntoData = Buffer.ReadBytes(ConjuntoSize)
    CopyMemory ByVal VarPtr(Conjunto(n)), ByVal VarPtr(ConjuntoData(0)), ConjuntoSize
    Set Buffer = Nothing
End Sub

Public Sub HandleUpdateConjuntoWindow(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    UsingSet = Buffer.ReadInteger
    
    Set Buffer = Nothing
    
    Call UpdateConjuntoWindow(UsingSet)
End Sub

Public Sub UpdateConjuntoWindow(ByRef ConjNum As Integer)
    Dim i As Byte, z As Byte, SString As String

    z = 1

    With Windows(GetWindowIndex("winCharacter"))

        ' for Clear
        If ConjNum = NO Then
            For i = 1 To 8
                .Controls(GetControlIndex("winCharacter", "lblBonus" & i)).Text = vbNullString
                .Controls(GetControlIndex("winCharacter", "lblBonus" & i)).visible = False
            Next i
            Exit Sub
        End If

        For i = 1 To Stats.Stat_Count - 1
            If Conjunto(ConjNum).Bonus.Add_Stat(i) > 0 Then
                If .Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = YES Then .Controls(GetControlIndex("winCharacter", "lblBonus" & z)).visible = True

                SString = Replace$(SString, ColourChar, vbNullString)
                If Conjunto(ConjNum).Bonus.Add_Stat_Percent(i) = YES Then
                    SString = ColourChar & GetColStr(Green) & GetAtributeName(i) & "+ " & ColourChar & GetColStr(White) & Conjunto(ConjNum).Bonus.Add_Stat(i) & "%"
                Else
                    SString = ColourChar & GetColStr(Green) & GetAtributeName(i) & "+ " & ColourChar & GetColStr(White) & Conjunto(ConjNum).Bonus.Add_Stat(i)
                End If
                .Controls(GetControlIndex("winCharacter", "lblBonus" & z)).Text = SString
                z = z + 1
            End If
        Next i

        ' Dano
        If Conjunto(ConjNum).Bonus.Dano > 0 Then
            If .Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = YES Then .Controls(GetControlIndex("winCharacter", "lblBonus" & z)).visible = True
            SString = Replace$(SString, ColourChar, vbNullString)
            If Conjunto(ConjNum).Bonus.DanoPercent = YES Then
                SString = ColourChar & GetColStr(Green) & "DMG+ " & ColourChar & GetColStr(White) & Conjunto(ConjNum).Bonus.Dano & "%"
            Else
                SString = ColourChar & GetColStr(Green) & "DMG+ " & ColourChar & GetColStr(White) & Conjunto(ConjNum).Bonus.Dano
            End If
            .Controls(GetControlIndex("winCharacter", "lblBonus" & z)).Text = SString
            z = z + 1
        End If

        ' Defence
        If Conjunto(ConjNum).Bonus.Defesa > 0 Then
            If .Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = YES Then .Controls(GetControlIndex("winCharacter", "lblBonus" & z)).visible = True
            SString = Replace$(SString, ColourChar, vbNullString)
            If Conjunto(ConjNum).Bonus.DanoPercent = YES Then
                SString = ColourChar & GetColStr(Green) & "DEF+ " & ColourChar & GetColStr(White) & Conjunto(ConjNum).Bonus.Defesa & "%"
            Else
                SString = ColourChar & GetColStr(Green) & "DEF+ " & ColourChar & GetColStr(White) & Conjunto(ConjNum).Bonus.Defesa
            End If
            .Controls(GetControlIndex("winCharacter", "lblBonus" & z)).Text = SString
            z = z + 1
        End If

        ' Exp
        If Conjunto(ConjNum).Bonus.EXP > 0 Then
            If .Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = YES Then .Controls(GetControlIndex("winCharacter", "lblBonus" & z)).visible = True
            SString = Replace$(SString, ColourChar, vbNullString)
            SString = ColourChar & GetColStr(Green) & "EXP+ " & ColourChar & GetColStr(White) & Conjunto(ConjNum).Bonus.EXP & "%"
            .Controls(GetControlIndex("winCharacter", "lblBonus" & z)).Text = SString
            z = z + 1
        End If
        
        ' Drop
        If Conjunto(ConjNum).Bonus.Drop > 0 Then
            If .Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = YES Then .Controls(GetControlIndex("winCharacter", "lblBonus" & z)).visible = True
            SString = Replace$(SString, ColourChar, vbNullString)
            SString = ColourChar & GetColStr(Green) & "DROP+ " & ColourChar & GetColStr(White) & Conjunto(ConjNum).Bonus.Drop & "%"
            .Controls(GetControlIndex("winCharacter", "lblBonus" & z)).Text = SString
            z = z + 1
        End If
    End With
End Sub

Public Function GetAtributeName(ByVal Atribute As Stats) As String
    Select Case Atribute
        Case Stats.strength
            GetAtributeName = "STR"
        Case Stats.Endurance
            GetAtributeName = "END"
        Case Stats.Agility
            GetAtributeName = "AGI"
        Case Stats.Intelligence
            GetAtributeName = "INT"
        Case Stats.Willpower
            GetAtributeName = "WILL"
    End Select
End Function
















