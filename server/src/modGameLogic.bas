Attribute VB_Name = "modGameLogic"
Option Explicit

'For Clear functions
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Function FindOpenPlayerSlot() As Long
    Dim i As Long

    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long, Optional ByVal PlayerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, ItemVal, MapNum, x, Y, PlayerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long, Optional ByVal PlayerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True, Optional ByVal isSB As Boolean = False)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
            MapItem(MapNum, i).PlayerName = PlayerName
            MapItem(MapNum, i).playerTimer = getTime + ITEM_SPAWN_TIME
            MapItem(MapNum, i).canDespawn = canDespawn
            MapItem(MapNum, i).despawnTimer = getTime + ITEM_DESPAWN_TIME
            MapItem(MapNum, i).Num = ItemNum
            MapItem(MapNum, i).Value = ItemVal
            MapItem(MapNum, i).x = x
            MapItem(MapNum, i).Y = Y
            MapItem(MapNum, i).Bound = isSB
            ' send to map
            SendSpawnItemToMap MapNum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim x As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(MapNum).MapData.MaxX
        For Y = 0 To Map(MapNum).MapData.MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).TileData.Tile(x, Y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).TileData.Tile(x, Y).Data1).Stackable > 0 And Map(MapNum).TileData.Tile(x, Y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).TileData.Tile(x, Y).Data1, 1, MapNum, x, Y)
                Else
                    Call SpawnItem(Map(MapNum).TileData.Tile(x, Y).Data1, Map(MapNum).TileData.Tile(x, Y).Data2, MapNum, x, Y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal mapnpcnum As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Dim NpcNum As Long
    Dim i As Long
    Dim x As Long
    Dim Y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    NpcNum = Map(MapNum).MapData.NPC(mapnpcnum)

    If NpcNum > 0 Then

        With MapNpc(MapNum).NPC(mapnpcnum)
            .Num = NpcNum
            .target = 0
            .TargetType = 0    ' clear
            .Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
            .Vital(Vitals.MP) = GetNpcMaxVital(NpcNum, Vitals.MP)
            .dir = Int(Rnd * 4)
            .spellBuffer.Spell = 0
            .spellBuffer.Timer = 0
            .spellBuffer.target = 0
            .spellBuffer.tType = 0
            
            .SecondsToSpawn = 0
            .ActionMsgSpawn = 0
            .Dead = NO

            'Check if theres a spawn tile for the specific npc
            For x = 0 To Map(MapNum).MapData.MaxX
                For Y = 0 To Map(MapNum).MapData.MaxY
                    If Map(MapNum).TileData.Tile(x, Y).Type = TILE_TYPE_NPCSPAWN Then
                        If Map(MapNum).TileData.Tile(x, Y).Data1 = mapnpcnum Then
                            .x = x
                            .Y = Y
                            .dir = Map(MapNum).TileData.Tile(x, Y).Data2
                            Spawned = True
                            Exit For
                        End If
                    End If
                Next Y
            Next x

            If Not Spawned Then

                ' Well try 100 times to randomly place the sprite
                For i = 1 To 100
                    x = Random(0, Map(MapNum).MapData.MaxX)
                    Y = Random(0, Map(MapNum).MapData.MaxY)

                    If x > Map(MapNum).MapData.MaxX Then x = Map(MapNum).MapData.MaxX
                    If Y > Map(MapNum).MapData.MaxY Then Y = Map(MapNum).MapData.MaxY

                    ' Check if the tile is walkable
                    If NpcTileIsOpen(MapNum, x, Y) Then
                        .x = x
                        .Y = Y
                        Spawned = True
                        Exit For
                    End If

                Next

            End If

            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then

                For x = 0 To Map(MapNum).MapData.MaxX
                    For Y = 0 To Map(MapNum).MapData.MaxY

                        If NpcTileIsOpen(MapNum, x, Y) Then
                            .x = x
                            .Y = Y
                            Spawned = True
                        End If

                    Next
                Next

            End If

            ' If we suceeded in spawning then send it to everyone
            If Spawned Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong SSpawnNpc
                Buffer.WriteLong mapnpcnum
                Buffer.WriteLong .Num
                Buffer.WriteLong .x
                Buffer.WriteLong .Y
                Buffer.WriteLong .dir
                SendDataToMap MapNum, Buffer.ToArray()
                Buffer.Flush: Set Buffer = Nothing
            End If

            SendMapNpcVitals MapNum, mapnpcnum
        End With
    End If
End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = Y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum).NPC(LoopI).Num > 0 Then
            If MapNpc(MapNum).NPC(LoopI).x = x Then
                If MapNpc(MapNum).NPC(LoopI).Y = Y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(MapNum).TileData.Tile(x, Y).Type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).TileData.Tile(x, Y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).TileData.Tile(x, Y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(MapNum).NPC(mapnpcnum).x
    Y = MapNpc(MapNum).NPC(mapnpcnum).Y
    CanNpcMove = True

    Select Case dir
    Case DIR_UP

        ' Check to make sure not outside of boundries
        If Y > 0 Then
            n = Map(MapNum).TileData.Tile(x, Y - 1).Type

            ' Check to make sure that the tile is walkable
            If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                CanNpcMove = False
                Exit Function
            End If

            ' Check to make sure that there is not a player in the way
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).NPC(mapnpcnum).x) And (GetPlayerY(i) = MapNpc(MapNum).NPC(mapnpcnum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                End If
            Next

            ' Check to make sure that there is not another npc in the way
            For i = 1 To MAX_MAP_NPCS
                If (i <> mapnpcnum) And (MapNpc(MapNum).NPC(i).Num > 0) And (MapNpc(MapNum).NPC(i).x = MapNpc(MapNum).NPC(mapnpcnum).x) And (MapNpc(MapNum).NPC(i).Y = MapNpc(MapNum).NPC(mapnpcnum).Y - 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Next

            ' Directional blocking
            If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).NPC(mapnpcnum).x, MapNpc(MapNum).NPC(mapnpcnum).Y).DirBlock, DIR_UP + 1) Then
                CanNpcMove = False
                Exit Function
            End If
        Else
            CanNpcMove = False
        End If

    Case DIR_DOWN

        ' Check to make sure not outside of boundries
        If Y < Map(MapNum).MapData.MaxY Then
            n = Map(MapNum).TileData.Tile(x, Y + 1).Type

            ' Check to make sure that the tile is walkable
            If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                CanNpcMove = False
                Exit Function
            End If

            ' Check to make sure that there is not a player in the way
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).NPC(mapnpcnum).x) And (GetPlayerY(i) = MapNpc(MapNum).NPC(mapnpcnum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                End If
            Next

            ' Check to make sure that there is not another npc in the way
            For i = 1 To MAX_MAP_NPCS
                If (i <> mapnpcnum) And (MapNpc(MapNum).NPC(i).Num > 0) And (MapNpc(MapNum).NPC(i).x = MapNpc(MapNum).NPC(mapnpcnum).x) And (MapNpc(MapNum).NPC(i).Y = MapNpc(MapNum).NPC(mapnpcnum).Y + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Next

            ' Directional blocking
            If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).NPC(mapnpcnum).x, MapNpc(MapNum).NPC(mapnpcnum).Y).DirBlock, DIR_DOWN + 1) Then
                CanNpcMove = False
                Exit Function
            End If
        Else
            CanNpcMove = False
        End If

    Case DIR_LEFT

        ' Check to make sure not outside of boundries
        If x > 0 Then
            n = Map(MapNum).TileData.Tile(x - 1, Y).Type

            ' Check to make sure that the tile is walkable
            If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                CanNpcMove = False
                Exit Function
            End If

            ' Check to make sure that there is not a player in the way
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).NPC(mapnpcnum).x - 1) And (GetPlayerY(i) = MapNpc(MapNum).NPC(mapnpcnum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                End If
            Next

            ' Check to make sure that there is not another npc in the way
            For i = 1 To MAX_MAP_NPCS
                If (i <> mapnpcnum) And (MapNpc(MapNum).NPC(i).Num > 0) And (MapNpc(MapNum).NPC(i).x = MapNpc(MapNum).NPC(mapnpcnum).x - 1) And (MapNpc(MapNum).NPC(i).Y = MapNpc(MapNum).NPC(mapnpcnum).Y) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Next

            ' Directional blocking
            If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).NPC(mapnpcnum).x, MapNpc(MapNum).NPC(mapnpcnum).Y).DirBlock, DIR_LEFT + 1) Then
                CanNpcMove = False
                Exit Function
            End If
        Else
            CanNpcMove = False
        End If

    Case DIR_RIGHT

        ' Check to make sure not outside of boundries
        If x < Map(MapNum).MapData.MaxX Then
            n = Map(MapNum).TileData.Tile(x + 1, Y).Type

            ' Check to make sure that the tile is walkable
            If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                CanNpcMove = False
                Exit Function
            End If

            ' Check to make sure that there is not a player in the way
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).NPC(mapnpcnum).x + 1) And (GetPlayerY(i) = MapNpc(MapNum).NPC(mapnpcnum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                End If
            Next

            ' Check to make sure that there is not another npc in the way
            For i = 1 To MAX_MAP_NPCS
                If (i <> mapnpcnum) And (MapNpc(MapNum).NPC(i).Num > 0) And (MapNpc(MapNum).NPC(i).x = MapNpc(MapNum).NPC(mapnpcnum).x + 1) And (MapNpc(MapNum).NPC(i).Y = MapNpc(MapNum).NPC(mapnpcnum).Y) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Next

            ' Directional blocking
            If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).NPC(mapnpcnum).x, MapNpc(MapNum).NPC(mapnpcnum).Y).DirBlock, DIR_RIGHT + 1) Then
                CanNpcMove = False
                Exit Function
            End If
        Else
            CanNpcMove = False
        End If

    End Select

End Function

Sub NpcMove(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(MapNum).NPC(mapnpcnum).dir = dir

    Select Case dir
    Case DIR_UP
        MapNpc(MapNum).NPC(mapnpcnum).Y = MapNpc(MapNum).NPC(mapnpcnum).Y - 1
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcMove
        Buffer.WriteLong mapnpcnum
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).x
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).Y
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).dir
        Buffer.WriteLong movement
        SendDataToMap MapNum, Buffer.ToArray()
        Buffer.Flush: Set Buffer = Nothing
    Case DIR_DOWN
        MapNpc(MapNum).NPC(mapnpcnum).Y = MapNpc(MapNum).NPC(mapnpcnum).Y + 1
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcMove
        Buffer.WriteLong mapnpcnum
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).x
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).Y
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).dir
        Buffer.WriteLong movement
        SendDataToMap MapNum, Buffer.ToArray()
        Buffer.Flush: Set Buffer = Nothing
    Case DIR_LEFT
        MapNpc(MapNum).NPC(mapnpcnum).x = MapNpc(MapNum).NPC(mapnpcnum).x - 1
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcMove
        Buffer.WriteLong mapnpcnum
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).x
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).Y
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).dir
        Buffer.WriteLong movement
        SendDataToMap MapNum, Buffer.ToArray()
        Buffer.Flush: Set Buffer = Nothing
    Case DIR_RIGHT
        MapNpc(MapNum).NPC(mapnpcnum).x = MapNpc(MapNum).NPC(mapnpcnum).x + 1
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcMove
        Buffer.WriteLong mapnpcnum
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).x
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).Y
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).dir
        Buffer.WriteLong movement
        SendDataToMap MapNum, Buffer.ToArray()
        Buffer.Flush: Set Buffer = Nothing
    End Select

End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal mapnpcnum As Long, ByVal dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(MapNum).NPC(mapnpcnum).dir = dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapnpcnum
    Buffer.WriteLong dir
    SendDataToMap MapNum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal MapNum As Long)
    Dim Y As Long
    Dim x As Long
    TempTile(MapNum).DoorTimer = 0
    ReDim TempTile(MapNum).DoorOpen(0 To Map(MapNum).MapData.MaxX, 0 To Map(MapNum).MapData.MaxY)

    For x = 0 To Map(MapNum).MapData.MaxX
        For Y = 0 To Map(MapNum).MapData.MaxY
            TempTile(MapNum).DoorOpen(x, Y) = NO
        Next
    Next

End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim OldNum As Long, NewNum As Long, OldUses As Long, NewUses As Long

    If OldSlot = 0 Or NewSlot = 0 Then
        Exit Sub
    End If

    OldNum = Player(Index).Spell(OldSlot).Spell
    NewNum = Player(Index).Spell(NewSlot).Spell
    OldUses = Player(Index).Spell(OldSlot).Uses
    NewUses = Player(Index).Spell(NewSlot).Uses

    Player(Index).Spell(OldSlot).Spell = NewNum
    Player(Index).Spell(OldSlot).Uses = NewUses
    Player(Index).Spell(NewSlot).Spell = OldNum
    Player(Index).Spell(NewSlot).Uses = OldUses
    SendPlayerSpells Index
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long)
    Dim ItemNum As Long, ItemBound As Long

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub    ' exit out early if error'd

    If FindOpenInvSlot(Index, GetPlayerEquipmentNum(Index, EqSlot)) > 0 Then

        ItemNum = GetPlayerEquipmentNum(Index, EqSlot)
        ItemBound = GetPlayerEquipmentBound(Index, EqSlot)

        GiveInvItem Index, ItemNum, 1, ItemBound

        If Item(ItemNum).GiveSpellNum > 0 Then
            Call RemovePlayerSpell(Index, Item(ItemNum).GiveSpellNum)
        End If

        PlayerMsg Index, "Você desequipou " & Trim$(Item(GetPlayerEquipmentNum(Index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipmentNum(Index, EqSlot)

        ' remove equipment
        SetPlayerEquipment Index, 0, EqSlot
        SetPlayerEquipmentBound Index, 0, EqSlot
        
        Call CheckConjunto(Index)

        SendWornEquipment Index
        SendMapEquipment Index

        ' send vitals
        Call SendVital(Index, Vitals.HP)
        Call SendVital(Index, Vitals.MP)
        Call SendMapVitals(Index)
        Call SendStats(Index)

        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
    Else
        PlayerMsg Index, "O inventário está cheio.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
    Dim FirstLetter As String * 1

    FirstLetter = LCase$(left$(Word, 1))

    If FirstLetter = "$" Then
        CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
        Exit Function
    End If

    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean
    If Not blockvar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal Index As Long)
    Dim partynum As Long, i As Long

    partynum = TempPlayer(Index).inParty
    If partynum > 0 Then
        ' find out how many members we have
        Party_CountMembers partynum
        ' make sure there's more than 2 people
        If Party(partynum).MemberCount > 2 Then
            ' check if leader
            If Party(partynum).Leader = Index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) > 0 And Party(partynum).Member(i) <> Index Then
                        Party(partynum).Leader = Party(partynum).Member(i)
                        PartyMsg partynum, GetPlayerName(i) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partynum, GetPlayerName(Index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) = Index Then
                        Party(partynum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partynum
                ' set update to all
                SendPartyUpdate partynum
                ' send clear to player
                SendPartyUpdateTo Index
            Else
                ' not the leader, just leave
                PartyMsg partynum, GetPlayerName(Index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) = Index Then
                        Party(partynum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partynum
                ' set update to all
                SendPartyUpdate partynum
                ' send clear to player
                SendPartyUpdateTo Index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partynum
            ' only 2 people, disband
            PartyMsg partynum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                Index = Party(partynum).Member(i)
                ' player exist?
                If Index > 0 Then
                    ' remove them
                    TempPlayer(Index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo Index
                End If
            Next
            ' clear out the party itself
            ClearParty partynum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal Index As Long, ByVal targetPlayer As Long)
    Dim partynum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub

    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg Index, "This player has an outstanding party invitation already.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg Index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If

    ' check if we're in a party
    If TempPlayer(Index).inParty > 0 Then
        partynum = TempPlayer(Index).inParty
        ' make sure we're the leader
        If Party(partynum).Leader = Index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partynum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, Index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = Index
                    ' let them know
                    PlayerMsg Index, "Invitation sent.", Green
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg Index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg Index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, Index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = Index
        ' let them know
        PlayerMsg Index, "Invitation sent.", Green
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal Index As Long, ByVal targetPlayer As Long)
    Dim partynum As Long, i As Long, x As Long

    If Index = 0 Then Exit Sub

    If Not IsConnected(Index) Or Not IsPlaying(Index) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If

    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If

    If TempPlayer(targetPlayer).inParty > 0 Then
        PlayerMsg Index, Trim$(GetPlayerName(targetPlayer)) & " is already in a party.", BrightRed
        PlayerMsg targetPlayer, "You're already in a party.", BrightRed
        Exit Sub
    End If

    ' check if already in a party
    If TempPlayer(Index).inParty > 0 Then
        ' get the partynumber
        partynum = TempPlayer(Index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partynum).Member(i) = 0 Then
                'add to the party
                Party(partynum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partynum
                ' send everyone's data to everyone
                SendPlayerData_Party partynum
                ' send update to all - including new player
                SendPartyUpdate partynum
                ' Send party vitals to everyone again
                For x = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(x) > 0 Then
                        SendPartyVitals partynum, Party(partynum).Member(x)
                    End If
                Next
                ' let everyone know they've joined
                PartyMsg partynum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partynum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg Index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partynum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partynum).MemberCount = 2
        Party(partynum).Leader = Index
        Party(partynum).Member(1) = Index
        Party(partynum).Member(2) = targetPlayer
        SendPlayerData_Party partynum
        SendPartyUpdate partynum
        SendPartyVitals partynum, Index
        SendPartyVitals partynum, targetPlayer
        ' let them know it's created
        PartyMsg partynum, "Party created.", BrightGreen
        PartyMsg partynum, GetPlayerName(Index) & " has joined the party.", Pink
        PartyMsg partynum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(Index).inParty = partynum
        TempPlayer(targetPlayer).inParty = partynum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal Index As Long, ByVal targetPlayer As Long)
    If Not IsConnected(Index) Or Not IsPlaying(Index) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If

    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If

    PlayerMsg Index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partynum As Long)
    Dim i As Long, highIndex As Long, x As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partynum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partynum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partynum).Member(x) = Party(partynum).Member(x + 1)
                    Party(partynum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partynum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partynum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partynum
End Sub

Public Sub Party_ShareExp(ByVal partynum As Long, ByVal EXP As Long, ByVal Index As Long, Optional ByVal enemyLevel As Long = 0)
    Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long

    If Party(partynum).MemberCount <= 0 Then Exit Sub

    ' check if it's worth sharing
    If Not EXP >= Party(partynum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP Index, EXP, enemyLevel
        Exit Sub
    End If

    ' find out the equal share
    expShare = EXP \ Party(partynum).MemberCount
    leftOver = EXP Mod Party(partynum).MemberCount

    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partynum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' give them their share
                GivePlayerEXP tmpIndex, expShare, enemyLevel
            End If
        End If
    Next

    ' give the remainder to a random member
    tmpIndex = Party(partynum).Member(Rand(1, Party(partynum).MemberCount))
    ' give the exp
    If leftOver > 0 Then GivePlayerEXP tmpIndex, leftOver, enemyLevel
End Sub

Public Sub GivePlayerEXP(ByVal Index As Long, ByVal EXP As Long, Optional ByVal enemyLevel As Long = 0)
    Dim multiplier As Long, partynum As Long, expBonus As Long
    ' no exp
    If EXP = 0 Then Exit Sub
    ' rte9
    If Index <= 0 Or Index > Player_HighIndex Then Exit Sub
    ' make sure we're not max level
    If Not GetPlayerLevel(Index) >= MAX_LEVELS Then
        ' check for exp deduction
        If enemyLevel > 0 Then
            ' exp deduction
            If enemyLevel <= GetPlayerLevel(Index) - 3 Then
                ' 3 levels lower, exit out
                Exit Sub
            ElseIf enemyLevel <= GetPlayerLevel(Index) - 2 Then
                ' half exp if enemy is 2 levels lower
                EXP = EXP / 2
            End If
        End If
        ' check if in party
        partynum = TempPlayer(Index).inParty
        If partynum > 0 Then
            If Party(partynum).MemberCount > 1 Then
                multiplier = Party(partynum).MemberCount - 1
                ' multiply the exp
                expBonus = (EXP / 100) * (multiplier * Options.PartyBonus)    ' Edit in Server Window
                ' Modify the exp
                EXP = EXP + expBonus
            End If
        End If
        ' give the exp
        Call SetPlayerExp(Index, GetPlayerExp(Index) + EXP)
        SendEXP Index
        SendActionMsg GetPlayerMap(Index), "+" & EXP & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        
        If GetPlayerPremium(Index) = YES Then
            SendActionMsg GetPlayerMap(Index), "Exp Bonus + " & (Options.PREMIUMEXP + TempPlayer(Index).Bonus.EXP) & "%", Yellow, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32 - 32
        End If
        
        
        ' check if we've leveled
        CheckPlayerLevelUp Index
    Else
        Call SetPlayerExp(Index, 0)
        SendEXP Index
    End If
End Sub

Public Function loginTokenOk(ByVal Index As Long, ByVal User As String, ByVal tLoginToken As String) As Boolean
    Dim i As Long
    loginTokenOk = False
    For i = 1 To MAX_PLAYERS
        If LoginToken(i).Active Then
            If LoginToken(i).User = User And LoginToken(i).token = tLoginToken Then
                ' return true
                loginTokenOk = True
                'load player
                ClearPlayer Index
                Player(Index) = LoginToken(i).LoadPlayer
                ' clear the token
                ClearLoginToken i
                ' exit out
                Exit Function
            End If
        End If
    Next
End Function

Public Sub ClearLoginToken(ByVal i As Integer)
    Debug.Print "Limpando Token:" & LoginToken(i).LoadPlayer.Login
    
    Call ZeroMemory(ByVal VarPtr(LoginToken(i)), LenB(LoginToken(i)))
    LoginToken(i).User = vbNullString
    LoginToken(i).token = vbNullString
    LoginToken(i).LoadPlayer.Login = vbNullString
    LoginToken(i).LoadPlayer.Password = vbNullString
    LoginToken(i).LoadPlayer.Name = vbNullString
    LoginToken(i).LoadPlayer.StartPremium = vbNullString
    LoginToken(i).LoadPlayer.Class = 1
    
    Debug.Print "Limpando Token:" & LoginToken(i).LoadPlayer.Login
End Sub

Function ActiveEventPage(ByVal Index As Long, ByVal eventNum As Long) As Long
    Dim x As Long, MapNum As Long, process As Boolean
    MapNum = GetPlayerMap(Index)
    For x = Map(MapNum).TileData.Events(eventNum).PageCount To 1 Step -1
        ' check if we match
        With Map(MapNum).TileData.Events(eventNum).EventPage(x)
            process = True
            ' player var check
            If .chkPlayerVar Then
                If .PlayerVarNum > 0 Then
                    If Player(Index).Variable(.PlayerVarNum) < .PlayerVariable Then
                        process = False
                    End If
                End If
            End If
            ' has item check
            If .chkHasItem Then
                If .HasItemNum > 0 Then
                    If HasItem(Index, .HasItemNum) = 0 Then
                        process = False
                    End If
                End If
            End If
            ' this page
            If process = True Then
                ActiveEventPage = x
                Exit Function
            End If
        End With
    Next
End Function

Public Function KeepTwoDigit(Num As Byte)
    If (Num < 10) Then
        KeepTwoDigit = "0" & Num
    Else
        KeepTwoDigit = Num
    End If
End Function

Public Sub SetPlayerGold(ByVal Index As Long, ByVal Value As Long)
    
    If Not IsPlaying(Index) Then Exit Sub
    
    Player(Index).Gold = Value
End Sub

Public Function GetPlayerGold(ByVal Index As Long) As Long
    If Not IsPlaying(Index) Then Exit Function
    
    GetPlayerGold = Player(Index).Gold
End Function







