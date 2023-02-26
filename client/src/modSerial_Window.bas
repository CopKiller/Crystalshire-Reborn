Attribute VB_Name = "modSerial_Window"
Option Explicit

Private Const MAX_SERIAL_LENGTH As Byte = 14

Public Sub CreateWindow_Serial()
' Create the window
    CreateWindow "winSerial", "Reivindicar Pacote!", zOrder_Win, 0, 0, 200, 170, Tex_Item(1), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf CloseWindowSerial)

    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 187, 136, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment

    CreateLabel WindowCount, "lblSerial", 18, 45, 190, , "Serial:", rockwellDec_15, , Alignment.alignLeft

    CreateTextbox WindowCount, "txtSerial", 18, 70, 170, 18, , Fonts.rockwellDec_15, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , , 1, , GetAddress(AddressOf SendSerial), MAX_SERIAL_LENGTH
    ' Button
    CreateButton WindowCount, "btnCancel", (Windows(WindowCount).Window.Width - 68 - 16), (Windows(WindowCount).Window.Height - 40), 68, 24, "Cancelar", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf CloseWindowSerial)
    CreateButton WindowCount, "btnReivindicar", 16, (Windows(WindowCount).Window.Height - 40), 68, 24, "Reivindicar", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf SendSerial)
End Sub

'Public Function lblSerialFormat() As String
'Dim i As Byte, X As Byte
'lblSerialFormat = "("
'For X = 1 To SERIAL_NUCLEOS
'    For i = 1 To SERIAL_CHARACTERES_NUCLEO
'        lblSerialFormat = lblSerialFormat & "#"
'    Next i
'    If X < SERIAL_NUCLEOS Then
'        lblSerialFormat = lblSerialFormat & SERIAL_SEPARATOR
'    End If
'Next X
'lblSerialFormat = lblSerialFormat & ")"
'End Function

Public Sub CloseWindowSerial()

If Windows(GetWindowIndex("winSerial")).Window.visible = False Then
    ShowWindow GetWindowIndex("winSerial")
Else
    HideWindow GetWindowIndex("winSerial")
End If

End Sub

'Public Function ValidateSerialFormat(ByVal Str As String, _
                                     ByVal Separador As String, _
                                     ByVal QuantNucleos As Byte, _
                                     ByVal QntChrInNucleo) As Boolean
'Dim Format() As String
'Dim i As Byte
'ValidateSerialFormat = False
'Format = Split(Str, Separador)
'If (UBound(Format) + 1) = QuantNucleos Then
'For i = 0 To UBound(Format)
'    If Len(Format(i)) <> QntChrInNucleo Then
'        Exit Function
'    End If
'Next i
'    ValidateSerialFormat = True
'End If
'End Function

Private Sub SendSerial()
Dim Buffer As clsBuffer
Dim SerialNumber As String

With Windows(GetWindowIndex("winSerial"))

SerialNumber = .Controls(GetControlIndex("winSerial", "txtSerial")).Text

If Len(SerialNumber) > MAX_SERIAL_LENGTH Then
    DialogueAlert DialogueMsg.MsgSERIALINCORRECT
    Exit Sub
End If

End With

Set Buffer = New clsBuffer
Buffer.WriteLong CSendSerial
Buffer.WriteString SerialNumber
SendData Buffer.ToArray()

Set Buffer = Nothing

End Sub
