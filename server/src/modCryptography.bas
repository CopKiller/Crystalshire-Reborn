Attribute VB_Name = "modCryptography"
Option Explicit

Public Const CRYPTO_KEY_LENGTH As Long = 16

Private Const padding As Long = 64

Public Declare Function Encrypt Lib "Encryptor.dll" (ByVal SourceLngPtr As Long, ByVal SourceLength As Long, ByRef DestLngPtr As Long, ByRef DestLengthLngPtr As Long, ByVal KeyRef As Long, ByVal IvRef As Long) As Long
Public Declare Function Decrypt Lib "Encryptor.dll" (ByVal SourceLngPtr As Long, ByVal SourceLength As Long, ByRef DestLngPtr As Long, ByVal KeyRef As Long, ByVal IvRef As Long) As Long

Public Key(0 To CRYPTO_KEY_LENGTH - 1) As Byte
Public IV(0 To CRYPTO_KEY_LENGTH - 1) As Byte

Public Sub InitCryptographyKey()
    Dim I As Long

    For I = 0 To CRYPTO_KEY_LENGTH - 1
        Key(I) = I * 5
        IV(I) = I * 3
    Next

End Sub

Public Function EncryptPacket(ByRef Data() As Byte, ByVal DataLength As Long) As Byte()
    Dim EncryptedLength As Long
    Dim Encrypted() As Byte

    ReDim Encrypted(0 To DataLength + padding)

    EncryptedLength = Encrypt(ByVal VarPtr(Data(0)), DataLength, ByVal VarPtr(Encrypted(0)), ByVal VarPtr(EncryptedLength), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    ReDim Preserve Encrypted(0 To EncryptedLength - 1)

    EncryptPacket = Encrypted

End Function

Public Function DecryptPacket(ByRef Data() As Byte, ByVal DataLengh As Long) As Byte()
    Dim Decrypted() As Byte
    Dim Count As Long

    ReDim Decrypted(0 To DataLengh - 1)

    Count = Decrypt(ByVal VarPtr(Data(0)), DataLengh, ByVal VarPtr(Decrypted(0)), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    DecryptPacket = Decrypted
End Function
