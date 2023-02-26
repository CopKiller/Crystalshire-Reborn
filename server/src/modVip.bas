Attribute VB_Name = "modVip"
Option Explicit

Public Sub CheckPremium(ByVal Index As Long)
' Check Premium
    If GetPlayerPremium(Index) = YES Then
        If GetPlayerStartPremium(Index) <> vbNullString Then
            If DateDiff("d", GetPlayerStartPremium(Index), Date) < GetPlayerDaysPremium(Index) Then
                If GetPlayerPremium(Index) = YES Then
                    Call PlayerMsg(Index, "Obrigado por adquirir seu Premium, Bom Jogo!", White)
                End If
            ElseIf DateDiff("d", GetPlayerStartPremium(Index), Date) >= GetPlayerDaysPremium(Index) Then
                If GetPlayerPremium(Index) = YES Then
                    Call SetPlayerPremium(Index, NO)
                    Call PlayerMsg(Index, "Seus dias de Premium acabaram... Bom Jogo!", White)
                End If
            End If
        Else
            Call SetPlayerStartPremium(Index, Date)
            Call CheckPremium(Index)
        End If
    End If
End Sub

Public Sub CheckPremiumLoop()
    Dim I As Integer
    For I = 1 To Player_HighIndex
        ' Check Premium
        If GetPlayerPremium(I) = YES Then
            If DateDiff("d", GetPlayerStartPremium(I), Date) >= GetPlayerDaysPremium(I) Then
                Call SetPlayerPremium(I, NO)
                Call PlayerMsg(I, "Seus dias de Premium acabaram... bom jogo!", White)
                Call SavePlayer(I)
                Call SendPlayerData(I)
            End If
        End If
    Next I
End Sub

Public Sub HandleRequestEditPremium(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
' Check Access
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call PlayerMsg(Index, "You do not have access to complete this action!", White)
        Exit Sub
    End If
    Call SendPremiumEditor(Index)
End Sub

Public Sub HandleChangePremium(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim a As String
    Dim b As String
    Dim c As Long
    Dim d As String
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    a = Buffer.ReadString
    b = Buffer.ReadString
    c = Buffer.ReadLong
    d = FindPlayer(a)
    Buffer.Flush: Set Buffer = Nothing
    ChangePremium Index, a, b, c, d
End Sub

Public Sub HandleRemovePremium(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim a As String
    Dim b As String
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    a = Buffer.ReadString
    b = FindPlayer(a)
    Buffer.Flush: Set Buffer = Nothing
    RemovePremium Index, a, b
End Sub

Private Sub ChangePremium(ByVal Index As Long, ByVal a As String, ByVal b As String, ByVal c As Long, ByVal d As String)
    If IsPlaying(d) Then
        If d < 0 Then Exit Sub
        ' Check access if everything is right, change Premium
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call PlayerMsg(Index, "Você não tem acesso pra completar essa ação.", White)
            Exit Sub
        Else
            Call SetPlayerPremium(d, YES)
            Call SetPlayerStartPremium(d, b)
            Call SetPlayerDaysPremium(d, c)
            GlobalMsg "O jogador " & GetPlayerName(d) & " se tornou VIP. Parabéns!", BrightCyan
        End If
        SendDataPremium d
        
    Else: Call PlayerMsg(Index, "Error: Jogador está offline ou nome inválido!.", BrightRed)
    End If
End Sub

Private Sub RemovePremium(ByVal Index As Long, ByVal a As String, ByVal b As String)
    If IsPlaying(b) Then
        ' Check access if everything is right, change Premium
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call PlayerMsg(Index, "You do not have access to complete this action!", White)
            Exit Sub
        Else
            Call SetPlayerPremium(b, NO)
            Call SetPlayerStartPremium(b, vbNullString)
            Call SetPlayerDaysPremium(b, 0)
            PlayerMsg b, "Seus dias de VIP acabaram.", BrightCyan
        End If
        SendPlayerData b
        SendDataPremium b
    End If
End Sub

Public Sub SendDataPremium(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim a As Long
    If GetPlayerPremium(Index) = YES Then
        a = DateDiff("d", GetPlayerStartPremium(Index), Now)
    Else
        a = 0
    End If
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDPremium
    Buffer.WriteLong Index
    Buffer.WriteByte GetPlayerPremium(Index)
    Buffer.WriteLong a
    Buffer.WriteLong GetPlayerDaysPremium(Index)
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub SendPremiumEditor(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPremiumEditor
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub AddPremiumTime(ByVal PlayerIndex As Long, ByVal Days As Long)
    Dim MStr
    Dim I As Integer

    If IsPlaying(PlayerIndex) Then
        If GetPlayerPremium(PlayerIndex) = YES Then
            Call SetPlayerDaysPremium(PlayerIndex, GetPlayerDaysPremium(PlayerIndex) + Days)
            Call PlayerMsg(PlayerIndex, "Foi adicionado +" & Days & " dias de Premium!", BrightGreen)
        Else
            Call SetPlayerPremium(PlayerIndex, YES)
            Call SetPlayerStartPremium(PlayerIndex, Date)
            Call SetPlayerDaysPremium(PlayerIndex, Days)
            GlobalMsg "O jogador " & GetPlayerName(PlayerIndex) & " se tornou Premium. Parabéns!", BrightCyan
        End If
        SavePlayer PlayerIndex
        SendPlayerData PlayerIndex
        SendDataPremium PlayerIndex
    End If
End Sub

' Premium
Function GetPlayerPremium(ByVal Index As Long) As Byte
    GetPlayerPremium = Trim$(Player(Index).Premium)
End Function

Sub SetPlayerPremium(ByVal Index As Long, ByVal Premium As Byte)
    Player(Index).Premium = Premium
End Sub

' Start Premium
Private Function GetPlayerStartPremium(ByVal Index As Long) As String
    GetPlayerStartPremium = Trim$(Player(Index).StartPremium)
End Function

Sub SetPlayerStartPremium(ByVal Index As Long, ByVal StartPremium As String)
    Player(Index).StartPremium = StartPremium
End Sub

' Days Premium
Private Function GetPlayerDaysPremium(ByVal Index As Long) As Long
    GetPlayerDaysPremium = Player(Index).DaysPremium
End Function

Sub SetPlayerDaysPremium(ByVal Index As Long, ByVal DaysPremium As Long)
    Player(Index).DaysPremium = DaysPremium
End Sub
