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

Private Buffer() As Byte
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
    ReDim Buffer(0 To BufferSize)
End Sub

Public Sub Allocate(ByVal nLength As Long)
    If BufferSize = 0 And nLength > 1 Then nLength = nLength - 1
    BufferSize = BufferSize + nLength
    ReDim Preserve Buffer(0 To BufferSize)
End Sub

Public Sub Flush()
    WriteHead = 0
    ReadHead = 0
    BufferSize = 0
    ReDim Buffer(0)
End Sub

Public Sub Trim()
' If the readhead is past the buffersize, this means everything has been read in the array, flush it
    If ReadHead >= Count Then Flush
End Sub

Public Sub WriteByte(ByVal nByte As Byte)

    If WriteHead > BufferSize Then Allocate 1

    Buffer(WriteHead) = nByte
    WriteHead = WriteHead + 1
End Sub

Public Sub WriteBytes(ByRef nByte() As Byte)
    Dim nLength As Long

    On Error GoTo errorhandler

    nLength = (UBound(nByte) - LBound(nByte)) + 1

    If WriteHead + nLength - 1 > BufferSize Then Allocate nLength

    CopyMemory Buffer(WriteHead), nByte(0), nLength
    WriteHead = WriteHead + nLength

errorhandler:
    Exit Sub
End Sub

Public Sub WriteInteger(ByVal nInteger As Integer)

    If WriteHead + 1 > BufferSize Then Allocate 2

    CopyMemory Buffer(WriteHead), nInteger, 2
    WriteHead = WriteHead + 2
End Sub

Public Sub WriteLong(ByVal nLong As Long)

    If WriteHead + 3 > BufferSize Then Allocate 4

    CopyMemory Buffer(WriteHead), nLong, 4
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

    CopyMemory Buffer(WriteHead), sBytes(0), sLength
    WriteHead = WriteHead + sLength
End Sub

Public Function ReadByte(Optional MoveReadHead As Boolean = True) As Byte

    If ReadHead > BufferSize Then Exit Function

    ReadByte = Buffer(ReadHead)
    If MoveReadHead Then ReadHead = ReadHead + 1
End Function

Public Function ReadBytes(ByVal nLength As Long, Optional MoveReadHead As Boolean = True) As Byte()
    Dim Data() As Byte

    If nLength = 0 Then Exit Function
    If ReadHead + nLength - 1 > BufferSize Then Exit Function

    ReDim Data(nLength - 1)

    CopyMemory Data(0), Buffer(ReadHead), nLength
    If MoveReadHead Then ReadHead = ReadHead + nLength

    ReadBytes = Data
End Function

Public Function ReadInteger(Optional MoveReadHead As Boolean = True) As Integer

    If ReadHead + 1 > BufferSize Then Exit Function

    CopyMemory ReadInteger, Buffer(ReadHead), 2
    If MoveReadHead Then ReadHead = ReadHead + 2
End Function

Public Function ReadLong(Optional MoveReadHead As Boolean = True) As Long

    If ReadHead + 3 > BufferSize Then Exit Function

    CopyMemory ReadLong, Buffer(ReadHead), 4
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

    'If sLength > 255 Then Exit Function
    ReDim sBytes(sLength - 1)

    CopyMemory sBytes(0), Buffer(ReadHead + 4), sLength

    ReadString = StrConv(sBytes, vbUnicode)
    If MoveReadHead Then ReadHead = ReadHead + sLength + 4
End Function

Public Function Count() As Long
    Count = (UBound(Buffer) - LBound(Buffer)) + 1
End Function

Public Function Length() As Long
    Length = Count - ReadHead
End Function

Public Function ToArray() As Byte()
    ToArray = Buffer()
End Function

Public Function ToString() As String
    ToString = StrConv(Buffer, vbUnicode)
End Function

Public Function CompressData(Optional ByVal WriteToBuffer As Boolean = True) As Byte()

Dim h As Integer
    ' If WriteToBuffer is true then all data in the buffer is comopressed and overwritten along with the compressed data header
    
    Dim CompressedSize As Long, UnCompressedSize As Long, I As Long, Data() As Byte
     
    UnCompressedSize = Count
    CompressedSize = 1.01 * UnCompressedSize + 12 ' This value is an assumption, compressed size may be lower
    ReDim Data(CompressedSize - 1) ' Set the dest buffer
    I = ZCompress(Data(0), CompressedSize, Buffer(0), UnCompressedSize)
    ReDim Preserve Data(CompressedSize - 1) ' In case we were able to compress more than what we assumed
    
    If WriteToBuffer Then
        Flush ' Clean the buffer
        ' Write some info about the compression and the actual compressed data
        h = h + 1
        Call WriteLong(UnCompressedSize)
        Call WriteLong(CompressedSize)
        Call WriteBytes(Data)
    Else
        CompressData = Data
    End If
End Function

Public Function UnCompressData() As Byte()
    ' Reads the data and tries to uncompress (Decompression starts from the ReadHead)
    
    Dim CompressedSize As Long, UnCompressedSize As Long, I As Long, cData() As Byte
    Dim Data() As Byte
    UnCompressedSize = ReadLong()
    CompressedSize = ReadLong()
    cData = ReadBytes(CompressedSize)
    
    ' Uncompress the data
    ReDim Data(UnCompressedSize - 1)
    ZUncompress Data(0), UnCompressedSize, cData(0), CompressedSize
    
    UnCompressData = Data
End Function
