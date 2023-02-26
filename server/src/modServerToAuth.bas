Attribute VB_Name = "modServerToAuth"
Option Explicit


Sub SendDataToGameServer(ByRef Data() As Byte)
    Dim TempBuffer() As Byte
    
    If Not ConnectToGameServer Then
        Exit Sub
    End If

    Dim Length As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    TempBuffer = EncryptPacket(Data, (UBound(Data) - LBound(Data)) + 1)
    Length = (UBound(TempBuffer) - LBound(TempBuffer)) + 1

    Buffer.PreAllocate 4 + Length
    Buffer.WriteLong Length
    Buffer.WriteBytes TempBuffer()

    frmMain.AuthSocket.SendData Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub
