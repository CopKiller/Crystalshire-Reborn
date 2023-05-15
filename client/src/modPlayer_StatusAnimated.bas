Attribute VB_Name = "modPlayer_StatusAnimated"
Option Explicit

Public Type StatusRec
    Ativo As Byte
End Type

Public Sub StatusPlayer(ByVal Index As Long, ByVal StatusNum As Byte, ByVal OnOff As Byte)
    Player(Index).StatusNum(StatusNum).Ativo = OnOff
End Sub

Public Sub DrawPlayerStatus(ByVal Index As Long)
    Dim X As Long, Y As Long, rec As RECT, i As Long, SString As String

    With Player(Index)
        X = (.X * PIC_X) + .xOffset
        Y = (.Y * PIC_Y) - 38 + .yOffset
        X = ConvertMapX(X)
        Y = ConvertMapY(Y)

        For i = 1 To (status_count - 1)
            If .StatusNum(i).Ativo > 0 Then
                rec.top = 0
                rec.Left = .StatusFrame * PIC_X
                RenderTexture Tex_Status(i), X, Y, rec.Left, rec.top, 25, 25, 32, 32

                Select Case i
                Case Status.typing
                    SString = "Jogador Digitando..."
                Case Status.Afk
                    SString = "Jogador Ausente..."
                Case Status.Confused
                    SString = "Jogador Confuso..."
                Case Else
                    SString = "Status de Jogador..."
                End Select

                If GlobalX >= X And GlobalX <= X + PIC_X Then
                    If GlobalY >= Y And GlobalY <= Y + PIC_Y Then
                        Call RenderEntity_Square(Tex_Design(6), GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), SString) / 2)) - 5, GlobalY - 35, TextWidth(font(Fonts.georgiaBold_16), SString) + 10, 20, 5, 200)
                        Call RenderText(font(Fonts.georgiaBold_16), SString, GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), SString) / 2)), GlobalY - 32, White)
                    End If
                End If
            End If
        Next i
    End With

End Sub

Public Sub SendStatusDigitando(ByVal Ativar As Byte)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CStatus
    buffer.WriteByte Status.typing
    buffer.WriteByte Ativar
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleStatusPlayer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim StatusNum As Byte
    Dim OnOff As Byte
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    i = buffer.ReadLong
    StatusNum = buffer.ReadByte
    OnOff = buffer.ReadByte

    buffer.Flush: Set buffer = Nothing

    StatusPlayer i, StatusNum, OnOff
End Sub
