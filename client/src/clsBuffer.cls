VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private buffer() As Byte
Private BufferSize As Long
Private WriteHead As Long
Private ReadHead As Long

Private Declare Function ZCompress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function ZUncompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private Sub Class_Initialize()
    Flush
End Sub

Public Sub PreAllocate(ByVal nLength As Long)
    WriteHead = 0
    ReadHead = 0
    BufferSize = nLength - 1
    ReDim buffer(0 To BufferSize)
End Sub

Public Sub Allocate(ByVal nLength As Long)

    If BufferSize = 0 And nLength > 1 Then nLength = nLength - 1
    BufferSize = BufferSize + nLength
    ReDim Preserve buffer(0 To BufferSize)
End Sub

Public Sub Flush()
    WriteHead = 0
    ReadHead = 0
    BufferSize = 0
    ReDim buffer(0)
End Sub

Public Sub Trim()
    Dim c As Long

    ' If the readhead is past the buffersize, this means everything has been read in the array, flush it
    If ReadHead >= Count Then
        Flush
    End If

End Sub

Public Sub WriteByte(ByVal nByte As Byte)

    If WriteHead > BufferSize Then Allocate 1
    buffer(WriteHead) = nByte
    WriteHead = WriteHead + 1
End Sub

Public Sub WriteBytes(ByRef nByte() As Byte)
    Dim nLength As Long
    nLength = (UBound(nByte) - LBound(nByte)) + 1

    If WriteHead + nLength - 1 > BufferSize Then Allocate nLength
    CopyMemory buffer(WriteHead), nByte(0), nLength
    WriteHead = WriteHead + nLength
End Sub

Public Sub WriteInteger(ByVal nInteger As Integer)

    If WriteHead + 1 > BufferSize Then Allocate 2
    CopyMemory buffer(WriteHead), nInteger, 2
    WriteHead = WriteHead + 2
End Sub

Public Sub WriteLong(ByVal nLong As Long)

    If WriteHead + 3 > BufferSize Then Allocate 4
    CopyMemory buffer(WriteHead), nLong, 4
    WriteHead = WriteHead + 4
End Sub

Public Sub WriteString(ByRef nString As String)
    Dim sBytes() As Byte
    Dim sLength As Long
    sLength = Len(nString)
    sBytes = StrConv(nString, vbFromUnicode)
    WriteLong sLength

    If sLength <= 0 Then Exit Sub
    If WriteHead + sLength - 1 > BufferSize Then Allocate sLength
    CopyMemory buffer(WriteHead), sBytes(0), sLength
    WriteHead = WriteHead + sLength
End Sub

Public Function ReadByte(Optional MoveReadHead As Boolean = True) As Byte

    If ReadHead > BufferSize Then Exit Function
    ReadByte = buffer(ReadHead)

    If MoveReadHead Then ReadHead = ReadHead + 1
End Function

Public Function ReadBytes(ByVal nLength As Long, Optional MoveReadHead As Boolean = True) As Byte()
    Dim data() As Byte

    If nLength = 0 Then Exit Function
    If ReadHead + nLength - 1 > BufferSize Then Exit Function
    ReDim data(nLength - 1)
    CopyMemory data(0), buffer(ReadHead), nLength

    If MoveReadHead Then ReadHead = ReadHead + nLength
    ReadBytes = data
End Function

Public Function ReadInteger(Optional MoveReadHead As Boolean = True) As Integer

    If ReadHead + 1 > BufferSize Then Exit Function
    CopyMemory ReadInteger, buffer(ReadHead), 2

    If MoveReadHead Then ReadHead = ReadHead + 2
End Function

Public Function ReadLong(Optional MoveReadHead As Boolean = True) As Long

    If ReadHead + 3 > BufferSize Then Exit Function
    CopyMemory ReadLong, buffer(ReadHead), 4

    If MoveReadHead Then ReadHead = ReadHead + 4
End Function

Public Function ReadString(Optional MoveReadHead As Boolean = True) As String
    Dim sLength As Long
    Dim sBytes() As Byte
    sLength = ReadLong(False)

    If sLength <= 0 Then
        If MoveReadHead Then ReadHead = ReadHead + 4
        Exit Function
    End If

    ReDim sBytes(sLength - 1)
    CopyMemory sBytes(0), buffer(ReadHead + 4), sLength
    ReadString = StrConv(sBytes, vbUnicode)

    If MoveReadHead Then ReadHead = ReadHead + sLength + 4
End Function

Public Function Count() As Long
    Count = (UBound(buffer) - LBound(buffer)) + 1
End Function

Public Function length() As Long
    length = Count - ReadHead
End Function

Public Function ToArray() As Byte()
    ToArray = buffer()
End Function

Public Function ToString() As String
    ToString = StrConv(buffer, vbUnicode)
End Function

Public Function CompressData(Optional ByVal WriteToBuffer As Boolean = True) As Byte()
' If WriteToBuffer is true then all data in the buffer is comopressed and overwritten along with the compressed data header

    Dim CompressedSize As Long, UnCompressedSize As Long, i As Long, data() As Byte

    UnCompressedSize = Count
    CompressedSize = 1.01 * UnCompressedSize + 12    ' This value is an assumption, compressed size may be lower
    ReDim data(CompressedSize - 1)    ' Set the dest buffer
    i = ZCompress(data(0), CompressedSize, buffer(0), UnCompressedSize)
    ReDim Preserve data(CompressedSize - 1)    ' In case we were able to compress more than what we assumed

    If WriteToBuffer Then
        Flush    ' Clean the buffer
        ' Write some info about the compression and the actual compressed data
        Call WriteLong(UnCompressedSize)
        Call WriteLong(CompressedSize)
        Call WriteBytes(data)
    Else
        CompressData = data
    End If
End Function

Public Function UnCompressData() As Byte()
' Reads the data and tries to uncompress (Decompression starts from the ReadHead)

    Dim CompressedSize As Long, UnCompressedSize As Long, i As Long, cData() As Byte
    Dim data() As Byte
    UnCompressedSize = ReadLong()
    CompressedSize = ReadLong()
    cData = ReadBytes(CompressedSize)

    ' Uncompress the data
    On Error Resume Next
    ReDim data(UnCompressedSize - 1)
    ZUncompress data(0), UnCompressedSize, cData(0), CompressedSize

    UnCompressData = data
End Function
