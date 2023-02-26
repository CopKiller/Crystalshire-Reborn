Attribute VB_Name = "modPremium"
Option Explicit

' Premium
Public PPremium As String
Public RPremium As String

Public Sub HandlePlayerDPremium(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim a As String
    Dim B As Long, c As Long, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Catch Data
    i = Buffer.ReadLong
    a = Buffer.ReadByte
    B = Buffer.ReadLong
    c = Buffer.ReadLong

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
        Windows(GetWindowIndex("winCharacter")).Controls(GetControlIndex("winCharacter", "lblVip")).Text = "Vip: " & PPremium
        Windows(GetWindowIndex("winCharacter")).Controls(GetControlIndex("winCharacter", "lblVipD")).Text = "Days: " & RPremium
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
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditPremium
    SendData Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendChangePremium(ByVal Name As String, ByVal Start As String, ByVal Days As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CChangePremium
    Buffer.WriteString Name
    Buffer.WriteString Start
    Buffer.WriteLong Days
    SendData Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendRemovePremium(ByVal Name As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRemovePremium
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
