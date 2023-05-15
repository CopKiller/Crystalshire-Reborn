Attribute VB_Name = "modCRC32"
Option Explicit

Private crcTable(0 To 255) As Long

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
    Dim i As Long, n As Long, CRC As Long

    For i = 0 To 255
        CRC = i
        For n = 0 To 7
            If CRC And 1 Then
                CRC = (((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor &HEDB88320
            Else
                CRC = ((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next
        crcTable(i) = CRC
    Next
End Sub

Public Function CRC32(ByRef data() As Byte) As Long
    Dim lCurPos As Long
    Dim lLen As Long

    lLen = AryCount(data) - 1
    CRC32 = &HFFFFFFFF

    For lCurPos = 0 To lLen
        CRC32 = (((CRC32 And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (crcTable((CRC32 And 255) Xor data(lCurPos)))
    Next

    CRC32 = CRC32 Xor &HFFFFFFFF
End Function

Public Sub GetMapCRC32(MapNum As Long)
    Dim data() As Byte, filename As String, f As Long
    ' map data
    filename = App.path & MAP_PATH & MapNum & "_.dat"
    If FileExist(filename) Then
        f = FreeFile
        Open filename For Binary As #f
        data = Space$(LOF(f))
        Get #f, , data
        Close #f
        MapCRC32(MapNum).MapDataCRC = CRC32(data)
    Else
        MapCRC32(MapNum).MapDataCRC = 0
    End If
    ' clear
    Erase data
    ' tile data
    filename = App.path & MAP_PATH & MapNum & ".dat"
    If FileExist(filename) Then
        f = FreeFile
        Open filename For Binary As #f
        data = Space$(LOF(f))
        Get #f, , data
        Close #f
        MapCRC32(MapNum).MapTileCRC = CRC32(data)
    Else
        MapCRC32(MapNum).MapTileCRC = 0
    End If
End Sub

Public Sub GetItemCRC32(itemNum As Long)
    Dim data() As Byte, filename As String, f As Long

    ' item data
    filename = App.path & ITEM_PATH & "item" & itemNum & ".dat"
    If FileExist(filename) Then
        f = FreeFile
        Open filename For Binary As #f
        data = Space$(LOF(f))
        Get #f, , data
        Close #f
        ItemCRC32(itemNum).ItemDataCRC = CRC32(data)
    Else
        ItemCRC32(itemNum).ItemDataCRC = 0
    End If
End Sub

Public Sub GetNpcCRC32(NpcNum As Long)
    Dim data() As Byte, filename As String, f As Long

    ' item data
    filename = App.path & NPC_PATH & "npc" & NpcNum & ".dat"
    If FileExist(filename) Then
        f = FreeFile
        Open filename For Binary As #f
        data = Space$(LOF(f))
        Get #f, , data
        Close #f
        NpcCRC32(NpcNum).NpcDataCRC = CRC32(data)
    Else
        NpcCRC32(NpcNum).NpcDataCRC = 0
    End If
End Sub

Public Sub HandleItemsCRC(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Integer, ItemCRC As Long, itemNum As Integer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_ITEMS
        itemNum = buffer.ReadInteger

        If itemNum > 0 Then
            ItemCRC = buffer.ReadLong

            If ItemCRC32(i).ItemDataCRC <> ItemCRC Then
                Call SendRequestItems
                Set buffer = Nothing
                Exit Sub
            End If
        End If
    Next i


    Set buffer = Nothing
End Sub

Public Sub HandleNpcsCRC(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Integer, NpcCRC As Long, NpcNum As Integer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_NPCS
        NpcNum = buffer.ReadInteger

        If NpcNum > 0 Then
            NpcCRC = buffer.ReadLong

            If NpcCRC32(NpcNum).NpcDataCRC <> NpcCRC Then
                Call SendRequestNPCS
                Set buffer = Nothing
                Exit Sub
            Else

            End If
        End If
    Next i




    Set buffer = Nothing
End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        If Not FileExist(App.path & "\data files\items\item" & i & ".dat") Then
            Call SaveItem(i)
        End If
    Next

End Sub

Sub SaveItem(ByVal itemNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.path & "\data files\items\item" & itemNum & ".dat"

    ' if it exists then kill the dat
    If FileExist(filename) Then
        Kill filename
    End If

    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Item(itemNum)
    Close #f
End Sub

Public Sub LoadItems()
    Dim filename As String, i As Integer

    Call CheckItems

    For i = 1 To MAX_ITEMS
        Call LoadItem(i)
    Next

End Sub

Public Sub LoadItem(ByVal itemNum As Long)
    Dim filename As String, f As Long

    filename = App.path & "\data files\items\item" & itemNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Item(itemNum)
    Close #f

    GetItemCRC32 itemNum
End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        If Not FileExist(App.path & "\data files\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If
    Next

End Sub

Sub SaveNpc(ByVal NpcNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.path & "\data files\npcs\npc" & NpcNum & ".dat"

    ' if it exists then kill the dat
    If FileExist(filename) Then
        Kill filename
    End If

    f = FreeFile
    Open filename For Binary As #f
    Put #f, , NPC(NpcNum)
    Close #f
End Sub

Public Sub LoadNpcs()
    Dim filename As String, i As Integer

    Call CheckNpcs

    For i = 1 To MAX_NPCS
        Call LoadNpc(i)
    Next

End Sub

Public Sub LoadNpc(ByVal NpcNum As Long)
    Dim filename As String, f As Long

    filename = App.path & "\data files\npcs\npc" & NpcNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , NPC(NpcNum)
    Close #f

    GetNpcCRC32 NpcNum

End Sub
