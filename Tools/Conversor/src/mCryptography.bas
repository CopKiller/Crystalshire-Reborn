Attribute VB_Name = "mCryptography"
Option Explicit

Public Const CRYPTO_KEY_LENGTH As Long = 16

Private Const PADDING As Long = 64

Public Declare Function encrypt Lib "Encryptor.dll" Alias "Encrypt" (ByVal SourceLngPtr As Long, ByVal SourceLength As Long, ByRef DestLngPtr As Long, ByRef DestLengthLngPtr As Long, ByVal KeyRef As Long, ByVal IvRef As Long) As Long
Public Declare Function decrypt Lib "Encryptor.dll" Alias "Decrypt" (ByVal SourceLngPtr As Long, ByVal SourceLength As Long, ByRef DestLngPtr As Long, ByVal KeyRef As Long, ByVal IvRef As Long) As Long

Public Key(0 To CRYPTO_KEY_LENGTH - 1) As Byte
Public IV(0 To CRYPTO_KEY_LENGTH - 1) As Byte

Public Sub InitCryptographyKey()
    Dim i As Long

    For i = 0 To CRYPTO_KEY_LENGTH - 1
        Key(i) = i * 5
        IV(i) = i * 3
    Next

End Sub

Public Function EncryptPacket(ByRef data() As Byte, ByVal dataLength As Long) As Byte()
    Dim EncryptedLength As Long
    Dim Encrypted() As Byte

    ReDim Encrypted(0 To dataLength + PADDING)

    EncryptedLength = encrypt(ByVal VarPtr(data(0)), dataLength, ByVal VarPtr(Encrypted(0)), ByVal VarPtr(EncryptedLength), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    ReDim Preserve Encrypted(0 To EncryptedLength - 1)

    EncryptPacket = Encrypted

End Function

Public Function DecryptPacket(ByRef data() As Byte, ByVal DataLengh As Long) As Byte()
    Dim Decrypted() As Byte
    Dim Count As Long

    ReDim Decrypted(0 To DataLengh - 1)

    Count = decrypt(ByVal VarPtr(data(0)), DataLengh, ByVal VarPtr(Decrypted(0)), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    DecryptPacket = Decrypted
End Function

Public Function EncryptFile(ByRef data() As Byte, ByVal dataLength As Long) As Byte()
    Dim EncryptedLength As Long
    Dim Encrypted() As Byte

    ReDim Encrypted(0 To dataLength + PADDING)

    EncryptedLength = encrypt(ByVal VarPtr(data(0)), dataLength, ByVal VarPtr(Encrypted(0)), ByVal VarPtr(EncryptedLength), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    ReDim Preserve Encrypted(0 To EncryptedLength - 1)

    EncryptFile = Encrypted

End Function

Public Function DecryptFile(ByRef data() As Byte, ByVal DataLengh As Long) As Byte()
    Dim Decrypted() As Byte
    Dim Count As Long

    ReDim Decrypted(0 To DataLengh - 1)

    Count = decrypt(ByVal VarPtr(data(0)), DataLengh, ByVal VarPtr(Decrypted(0)), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    DecryptFile = Decrypted
End Function

