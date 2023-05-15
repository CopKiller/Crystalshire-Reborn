Attribute VB_Name = "modPremium"
Option Explicit

' Premium
Public PPremium As String
Public RPremium As String

Public Sub HandlePlayerDPremium(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim a As String
    Dim B As Long, c As Long, i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Catch Data
    i = buffer.ReadLong
    a = buffer.ReadByte
    B = buffer.ReadLong
    c = buffer.ReadLong

    ' Changing global variables
    Player(i).Premium = a

    ' Exclusivo do client do próprio jogador!
    If i = MyIndex Then
        If a = YES Then
            PPremium = "Sim"
            RPremium = c - B & " Dias"
        Else
            PPremium = "Nao"
            RPremium = "0 Dias"
        End If
        Windows(GetWindowIndex("winCharacter")).Controls(GetControlIndex("winCharacter", "lblVip")).text = "Vip: " & PPremium
        Windows(GetWindowIndex("winCharacter")).Controls(GetControlIndex("winCharacter", "lblVipD")).text = "Days: " & RPremium
    End If
End Sub

Public Sub HandlePremiumEditor()
    Dim i As Long

    ' Check Access
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    ' If you have everything right, up the Editor.
    With frmEditor_Premium
        .visible = True
    End With
End Sub

Sub SendRequestEditPremium()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditPremium
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendChangePremium(ByVal Name As String, ByVal Start As String, ByVal Days As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CChangePremium
    buffer.WriteString Name
    buffer.WriteString Start
    buffer.WriteLong Days
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendRemovePremium(ByVal Name As String)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRemovePremium
    buffer.WriteString Name
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub
