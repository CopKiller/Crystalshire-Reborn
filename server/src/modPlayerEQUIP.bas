Attribute VB_Name = "modPlayerEQUIPMENT"
Option Explicit

Function GetPlayerEquipmentNum(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function

    GetPlayerEquipmentNum = Player(Index).Equipment(EquipmentSlot).Num
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal ItemNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot).Num = ItemNum
    Player(Index).Equipment(EquipmentSlot).Level = ItemLevel
End Sub

Sub SetPlayerEquipmentBound(ByVal Index As Long, ByVal Bound As Byte, ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot).Bound = Bound
End Sub

Function GetPlayerEquipmentBound(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte
    GetPlayerEquipmentBound = Player(Index).Equipment(EquipmentSlot).Bound
End Function
