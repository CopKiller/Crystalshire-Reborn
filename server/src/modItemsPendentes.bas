Attribute VB_Name = "modItemsPendentes"
Option Explicit

Public Enum Operacao
    None = 0
    Save
    Delete
    Request
End Enum

Public Sub PendingItem(ByVal Name As String, ByVal Operacao As Operacao, Optional ByVal ItemID As Long, Optional ByVal ItemValue As Long, _
                                                                                                        Optional ByVal Mensagem As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    ' Envia uma solicitação conforme a operação pro servidor de eventos!
    Buffer.WriteLong ItemsPendentes

    Buffer.WriteByte Operacao
    Select Case Operacao
    
    Case 1, 2 'Save Or Delete
        Buffer.WriteString Name
        Buffer.WriteLong ItemID
        Buffer.WriteLong ItemValue
        Buffer.WriteString Mensagem
    Case 3 ' Request
        Buffer.WriteString Name
    End Select
    
    SendToEventServer Buffer.ToArray

    Set Buffer = Nothing
End Sub

Public Sub HandleItemsPendentes(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String, ItemID As Long, ItemValue As Long, Indice As Long, CountItems As Long, I As Integer, Msg As String
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data

    ' recebe os dados do servidor de eventos
    CountItems = Buffer.ReadLong

    For I = 1 To CountItems
        Name = Trim$(Buffer.ReadString)
        Indice = FindPlayer(Name)

        ItemID = Buffer.ReadLong
        ItemValue = Buffer.ReadLong
        
        Msg = Buffer.ReadString

        ' Jogador Online?
        If Indice > 0 Then
            If GiveInvItem(Indice, ItemID, ItemValue, 0, , True) Then   'Da o item ao jogador, caso aconteça envia um pedido pra deletar da fila
                Call PlayerMsg(Indice, Msg, Yellow)
                Call PendingItem(Name, Delete, ItemID, ItemValue)
                
                If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
            End If
        End If
    Next I

    Set Buffer = Nothing

    Call TextEventAdd(Name & ": Pending items Received!")
End Sub
