Attribute VB_Name = "modItemsPendentes"
Option Explicit

Private Enum Operacao
    None = 0
    Save
    Delete
    Request
End Enum

Public Sub PendingItem(ByVal Index As Long, ByVal Operacao As Operacao, Optional ByVal ItemID As Long, Optional ByVal ItemValue As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    ' Envia uma solicitação conforme a operação pro servidor de eventos!
    Buffer.WriteLong SeItemsPendentes

    Buffer.WriteByte Operacao
    Select Case Operacao
    
    Case 1, 2 'Save Or Delete
        Buffer.WriteString GetPlayerName(Index)
        Buffer.WriteLong ItemID
        Buffer.WriteLong ItemValue
    Case 3 ' Request
        Buffer.WriteString GetPlayerName(Index)
    End Select
    
    SendToEventServer Buffer.ToArray

    Set Buffer = Nothing
End Sub

Public Sub HandleItemsPendentes(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String, ItemID As Long, ItemValue As Long, Indice As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data
    
    ' recebe os dados do servidor de eventos
    Name = Buffer.ReadString
    Indice = FindPlayer(Name)
    ItemID = Buffer.ReadLong
    ItemValue = Buffer.ReadLong
    
    ' Jogador Online?
    If Indice > 0 Then
        If GiveInvItem(Indice, ItemID, ItemValue, 0) Then 'Da o item ao jogador, caso aconteça envia um pedido pra deletar da fila
            Call PlayerMsg(Indice, "You received an item that was pending!", Green)
            Call PendingItem(Indice, Delete, ItemID, ItemValue)
        End If
    End If
    Set Buffer = Nothing
End Sub
