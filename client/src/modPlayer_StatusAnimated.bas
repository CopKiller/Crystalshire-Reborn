Attribute VB_Name = "modPlayer_StatusAnimated"
Option Explicit

Public Type StatusRec
    Ativo As Byte
End Type

Public Sub StatusPlayer(ByVal Index As Long, ByVal StatusNum As Byte, ByVal OnOff As Byte)
    Player(Index).StatusNum(StatusNum).Ativo = OnOff
End Sub

Public Sub DrawPlayerStatus(ByVal Index As Long)
    Dim x As Long, Y As Long, rec As RECT, i As Long, sString As String

    With Player(Index)
        x = (.x * PIC_X) + .xOffset
        Y = (.Y * PIC_Y) - 38 + .yOffset
        x = ConvertMapX(x)
        Y = ConvertMapY(Y)

        For i = 1 To (status_count - 1)
            If .StatusNum(i).Ativo > 0 Then
                rec.top = 0
                rec.Left = .StatusFrame * PIC_X
                RenderTexture Tex_Status(i), x, Y, rec.Left, rec.top, 25, 25, 32, 32

                Select Case i
                Case Status.typing
                    sString = "Jogador Digitando..."
                Case Status.Afk
                    sString = "Jogador Ausente..."
                Case Status.Confused
                    sString = "Jogador Confuso..."
                Case Else
                    sString = "Status de Jogador..."
                End Select

                If GlobalX >= x And GlobalX <= x + PIC_X Then
                    If GlobalY >= Y And GlobalY <= Y + PIC_Y Then
                        Call RenderEntity_Square(Tex_Design(6), GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)) - 5, GlobalY - 35, TextWidth(font(Fonts.georgiaBold_16), sString) + 10, 20, 5, 200)
                        Call RenderText(font(Fonts.georgiaBold_16), sString, GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)), GlobalY - 32, White)
                    End If
                End If
            End If
        Next i
    End With

End Sub

Public Sub SendStatusDigitando(ByVal Ativar As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CStatus
    Buffer.WriteByte Status.typing
    Buffer.WriteByte Ativar
    SendData Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleStatusPlayer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim StatusNum As Byte
    Dim OnOff As Byte
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    i = Buffer.ReadLong
    StatusNum = Buffer.ReadByte
    OnOff = Buffer.ReadByte

    Buffer.Flush: Set Buffer = Nothing

    StatusPlayer i, StatusNum, OnOff
End Sub
