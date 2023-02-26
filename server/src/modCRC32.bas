Attribute VB_Name = "modCRC32"
Option Explicit

Public crcTable(0 To 255) As Long

Public MapCRC32(1 To MAX_MAPS) As MapCRCStruct
Public ItemCRC32(1 To MAX_ITEMS) As ItemCRCStruct
Public NpcCRC32(1 To MAX_NPCS) As NpcCRCStruct

Public Type MapCRCStruct
    MapDataCRC As Long
    MapTileCRC As Long
End Type

Public Type ItemCRCStruct
    ItemDataCRC As Long
End Type

Public Type NpcCRCStruct
    NpcDataCRC As Long
End Type

Public Sub InitCRC32()
    Dim I As Long, n As Long, CRC As Long

    For I = 0 To 255
        CRC = I
        For n = 0 To 7
            If CRC And 1 Then
                CRC = (((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor &HEDB88320
            Else
                CRC = ((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next
        crcTable(I) = CRC
    Next
End Sub

Public Function CRC32(ByRef Data() As Byte) As Long
    Dim lCurPos As Long
    Dim lLen As Long

    lLen = AryCount(Data) - 1
    CRC32 = &HFFFFFFFF

    For lCurPos = 0 To lLen
        CRC32 = (((CRC32 And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (crcTable((CRC32 And 255) Xor Data(lCurPos)))
    Next

    CRC32 = CRC32 Xor &HFFFFFFFF
End Function

Sub GetMapCRC32(MapNum As Long)
    Dim Data() As Byte, FileName As String, F As Long
    ' map data
    FileName = App.Path & "\data\maps\map" & MapNum & ".ini"
    If FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Binary As #F
        Data = Space$(LOF(F))
        Get #F, , Data
        Close #F
        MapCRC32(MapNum).MapDataCRC = CRC32(Data)
    Else
        MapCRC32(MapNum).MapDataCRC = 0
    End If
    ' clear
    Erase Data
    ' tile data
    FileName = App.Path & "\data\maps\map" & MapNum & ".dat"
    If FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Binary As #F
        Data = Space$(LOF(F))
        Get #F, , Data
        Close #F
        MapCRC32(MapNum).MapTileCRC = CRC32(Data)
    Else
        MapCRC32(MapNum).MapTileCRC = 0
    End If
End Sub

Sub GetItemCRC32(ItemNum As Long)
    Dim Data() As Byte, FileName As String, F As Long
    ' Item data
    FileName = App.Path & "\data\items\item" & ItemNum & ".dat"
    If FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Binary As #F
        Data = Space$(LOF(F))
        Get #F, , Data
        Close #F
        ItemCRC32(ItemNum).ItemDataCRC = CRC32(Data)
    Else
        ItemCRC32(ItemNum).ItemDataCRC = 0
    End If
End Sub

Sub GetNpcCRC32(NpcNum As Long)
    Dim Data() As Byte, FileName As String, F As Long
    ' Npc data
    FileName = App.Path & "\data\npcs\npc" & NpcNum & ".dat"
    If FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Binary As #F
        Data = Space$(LOF(F))
        Get #F, , Data
        Close #F
        NpcCRC32(NpcNum).NpcDataCRC = CRC32(Data)
    Else
        NpcCRC32(NpcNum).NpcDataCRC = 0
    End If
End Sub

Public Sub SendItemsCRC32(ByVal Index As Long)
    Dim I As Integer
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SCheckItemCRC
    
    For I = 1 To MAX_ITEMS
        If LenB(Trim$(Item(I).Name)) > 0 Then
            Buffer.WriteInteger I
            Buffer.WriteLong ItemCRC32(I).ItemDataCRC
        End If
    Next I
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendNpcsCRC32(ByVal Index As Long)
    Dim I As Integer
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SCheckNpcCRC
    
    For I = 1 To MAX_NPCS
        If LenB(Trim$(NPC(I).Name)) > 0 Then
            Buffer.WriteInteger I
            Buffer.WriteLong NpcCRC32(I).NpcDataCRC
        End If
    Next I
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

