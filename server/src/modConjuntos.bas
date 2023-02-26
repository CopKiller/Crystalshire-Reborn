Attribute VB_Name = "modConjuntos"
Option Explicit

'For Clear functions
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Const MAX_CONJUNTOS As Byte = 50
Private Const ACTIONMSG_LENGTH As Byte = 100

Public Conjunto(1 To MAX_CONJUNTOS) As ConjuntoRec

' Type recs
Public Type BonusRec                                    '\\Contagem dos Bonus//
    'Atributos
    Add_Stat(1 To Stats.Stat_Count - 1) As Long         ' 1,2,3,4,5
    Add_Stat_Percent(1 To Stats.Stat_Count - 1) As Byte
    'Dano
    Dano As Long                                        ' 6
    DanoPercent As Byte
    'Defesa
    Defesa As Long                                      ' 7
    DefesaPercent As Byte
    'Exp
    EXP As Integer                                      ' 8
    Drop As Byte                                        ' 9
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

'////////////////////////
'////////HANDLE//////////
'////////////////////////
Sub HandleRequestEditConjunto(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SConjuntoEditor
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleSaveConjunto(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ConjuntoSize As Long
    Dim ConjuntoData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong    'CLng(Parse(1))

    If n < 0 Or n > MAX_CONJUNTOS Then
        Exit Sub
    End If

    ' Update the Conjunto
    ConjuntoSize = LenB(Conjunto(n))
    ReDim ConjuntoData(ConjuntoSize - 1)
    ConjuntoData = Buffer.ReadBytes(ConjuntoSize)
    CopyMemory ByVal VarPtr(Conjunto(n)), ByVal VarPtr(ConjuntoData(0)), ConjuntoSize
    Buffer.Flush: Set Buffer = Nothing

    ' Save it
    Call ConjuntoCache_Create(n)
    Call SendConjuntosAll(n)
    Call SaveConjunto(n)
    Call AddLog(GetPlayerName(Index) & " saved Conjunto #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestConjuntos(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendConjuntos Index
End Sub

'////////////////////
'/////TCP////////////
'////////////////////
Private Sub SendConjuntosTo(ByVal Index As Long, ByVal ConjuntoNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateConjunto
    Buffer.WriteBytes ConjuntoCache(ConjuntoNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Private Sub SendConjuntosAll(ByVal ConjuntoNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateConjunto
    Buffer.WriteBytes ConjuntoCache(ConjuntoNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendConjuntos(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_CONJUNTOS

        If LenB(Trim$(Conjunto(I).Name)) > 0 Then
            Call SendConjuntosTo(Index, I)
        End If

    Next

End Sub

Private Sub UpdateConjuntoWindow(ByVal Index As Long, ByVal ConjuntoNum As Integer)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateConjuntoWindow
    Buffer.WriteInteger ConjuntoNum
    
    SendDataTo Index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

'///////////////////
'/////DATA BASE/////
'///////////////////
' **********
' ** Conjuntos **
' **********
Sub SaveConjuntos()
    Dim I As Long

    For I = 1 To MAX_CONJUNTOS
        Call SaveConjunto(I)
    Next

End Sub

Sub SaveConjunto(ByVal ConjuntoNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\Conjuntos\Conjunto" & ConjuntoNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Conjunto(ConjuntoNum)
    Close #F
End Sub

Public Sub LoadConjuntos()
    Dim FileName As String
    Dim I As Long
    Dim F As Long

    Call CheckConjuntos

    For I = 1 To MAX_CONJUNTOS
        FileName = App.Path & "\data\Conjuntos\Conjunto" & I & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Conjunto(I)
        Close #F
    Next

End Sub

Private Sub CheckConjuntos()
    Dim I As Long

    For I = 1 To MAX_CONJUNTOS
        If Not FileExist("\Data\Conjuntos\Conjunto" & I & ".dat") Then
            Call SaveConjunto(I)
        End If
    Next

End Sub

Private Sub ClearConjunto(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Conjunto(Index)), LenB(Conjunto(Index)))
    Conjunto(Index).Name = vbNullString
    Conjunto(Index).Actions.Msg = vbNullString
End Sub

Sub ClearConjuntos()
    Dim I As Long

    For I = 1 To MAX_CONJUNTOS
        Call ClearConjunto(I)
    Next
End Sub

'//////////////////////////
'////////LOGICA////////////
'//////////////////////////
Public Sub CheckConjunto(ByVal Index As Long)
    Dim I As Integer, z As Byte

    ' Faz a limpeza antes e recalcula os bonus se tiver um conjunto a ser ativado ou atualizado!
    Call ClearConjuntoBonus(Index)

    For I = 1 To MAX_CONJUNTOS
        If LenB(Trim$(Conjunto(I).Name)) > 0 Then
            For z = 1 To Equipment.Equipment_Count - 1
                If Conjunto(I).Item(z) > 0 And Conjunto(I).Item(z) <= MAX_ITEMS Then
                    If GetEquipment(Index, Conjunto(I).Item(z)) Then
                        If HaveAllConjuntoItems(Index, I) Then

                            TempPlayer(Index).ConjuntoID = I
                            Call AllocateConjuntoBonus(Index)
                            
                            If LenB(Trim$(Conjunto(I).Actions.Msg)) > 0 Then
                                Call PlayerMsg(Index, Trim$(Conjunto(I).Actions.Msg), BrightGreen)
                            End If
                            If Conjunto(I).Actions.Animation > 0 Then
                                Call SendAnimation(GetPlayerMap(Index), Conjunto(I).Actions.Animation, GetPlayerX(Index), GetPlayerY(Index), TARGET_TYPE_PLAYER, Index)
                            End If

                            Exit Sub
                        End If
                    End If
                End If
            Next z
        End If
    Next I
End Sub

Private Sub ClearConjuntoBonus(ByVal Index As Long)
    Dim I As Integer
    
    TempPlayer(Index).ConjuntoID = 0

    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index).Bonus), LenB(TempPlayer(Index).Bonus))
    
    Call UpdateConjuntoWindow(Index, NO)
    Call SendStats(Index)
End Sub

Public Sub AllocateConjuntoBonus(ByVal Index As Long)
    Dim I As Integer
    Dim ConjuntoID As Integer

    ConjuntoID = TempPlayer(Index).ConjuntoID

    If ConjuntoID <= 0 Or ConjuntoID > MAX_CONJUNTOS Then Exit Sub
    If LenB(Trim$(Conjunto(ConjuntoID).Name)) = 0 Then Exit Sub

    ' Atributos
    For I = 1 To Stats.Stat_Count - 1
        If Conjunto(ConjuntoID).Bonus.Add_Stat(I) > 0 Then
            If Conjunto(ConjuntoID).Bonus.Add_Stat_Percent(I) = YES Then
                TempPlayer(Index).Bonus.Add_Stat(I) = (GetPlayerRawStat(Index, I) / 100) * Conjunto(ConjuntoID).Bonus.Add_Stat(I)
                TempPlayer(Index).Bonus.Add_Stat_Percent(I) = YES
            ElseIf Conjunto(ConjuntoID).Bonus.Add_Stat_Percent(I) = NO Then
                TempPlayer(Index).Bonus.Add_Stat(I) = Conjunto(ConjuntoID).Bonus.Add_Stat(I)
                TempPlayer(Index).Bonus.Add_Stat_Percent(I) = NO
            End If
        End If
    Next I

    ' Damage
    If Conjunto(ConjuntoID).Bonus.Dano > 0 Then
        If Conjunto(ConjuntoID).Bonus.DanoPercent = YES Then
            TempPlayer(Index).Bonus.Dano = (GetPlayerDamage(Index) / 100) * Conjunto(ConjuntoID).Bonus.Dano
            TempPlayer(Index).Bonus.DanoPercent = YES
        ElseIf Conjunto(ConjuntoID).Bonus.DanoPercent = NO Then
            TempPlayer(Index).Bonus.Dano = Conjunto(ConjuntoID).Bonus.Dano
            TempPlayer(Index).Bonus.DanoPercent = NO
        End If
    End If

    ' Defence
    If Conjunto(ConjuntoID).Bonus.Defesa > 0 Then
        If Conjunto(ConjuntoID).Bonus.DefesaPercent = YES Then
            TempPlayer(Index).Bonus.Defesa = (GetPlayerDefence(Index) / 100) * Conjunto(ConjuntoID).Bonus.Defesa
            TempPlayer(Index).Bonus.DefesaPercent = YES
        ElseIf Conjunto(ConjuntoID).Bonus.DefesaPercent = NO Then
            TempPlayer(Index).Bonus.Defesa = Conjunto(ConjuntoID).Bonus.Defesa
            TempPlayer(Index).Bonus.DefesaPercent = NO
        End If
    End If

    ' Exp
    If Conjunto(ConjuntoID).Bonus.EXP > 0 Then
        TempPlayer(Index).Bonus.EXP = Conjunto(ConjuntoID).Bonus.EXP
    End If
    
    ' Drop
    If Conjunto(ConjuntoID).Bonus.Drop > 0 Then
        TempPlayer(Index).Bonus.Drop = Conjunto(ConjuntoID).Bonus.Drop
    End If
    
    Call UpdateConjuntoWindow(Index, ConjuntoID)
    Call SendStats(Index)
End Sub

Private Function GetEquipment(ByVal Index As Long, ByVal ID As Integer) As Boolean
    Dim I As Byte
    
    GetEquipment = False
    
    For I = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipmentNum(Index, I) > 0 And GetPlayerEquipmentNum(Index, I) <= MAX_ITEMS Then
            If GetPlayerEquipmentNum(Index, I) = ID Then
                GetEquipment = True
                Exit Function
            End If
        End If
    Next I
End Function

Public Function HaveAllConjuntoItems(ByVal Index As Long, ByVal ConjuntoID As Integer) As Boolean
    Dim I As Integer
    
    HaveAllConjuntoItems = True

    For I = 1 To Equipment.Equipment_Count - 1
        If Conjunto(ConjuntoID).Item(I) > 0 And Conjunto(ConjuntoID).Item(I) <= MAX_ITEMS Then
            If Not GetEquipment(Index, Conjunto(ConjuntoID).Item(I)) Then
                HaveAllConjuntoItems = False
            End If
        End If
    Next I

End Function
