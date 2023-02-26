Attribute VB_Name = "modCACHE"
Option Explicit

Public MapCache(1 To MAX_MAPS) As CACHE
Private ItemCache(1 To MAX_ITEMS) As CACHE
Private AnimationCache(1 To MAX_ANIMATIONS) As CACHE
Private NpcCache(1 To MAX_NPCS) As CACHE
Private ShopCache(1 To MAX_SHOPS) As CACHE
Private SpellCache(1 To MAX_SPELLS) As CACHE
Private ResourceCache(1 To MAX_RESOURCES) As CACHE
Private SerialCache(1 To MAX_SERIAL_NUMBER) As CACHE
Private ConvCache(1 To MAX_CONVS) As CACHE
Private QuestCache(1 To MAX_QUESTS) As CACHE
Private GuildCache(1 To MAX_GUILDS) As CACHE
Public ConjuntoCache(1 To MAX_CONJUNTOS) As CACHE

Public MapResourceCache(1 To MAX_MAPS) As ResourceCacheRec

Private Type CACHE
    Data() As Byte
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Currency
    x As Long
    Y As Long
    cur_health As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Sub CreateFullCache()
    Dim n As Long

    For n = 1 To MAX_MAPS
        Call MapCache_Create(n)
    Next
    
    For n = 1 To MAX_NPCS
        Call NpcCache_Create(n)
    Next

    For n = 1 To MAX_ITEMS
        Call ItemCache_Create(n)
    Next

    For n = 1 To MAX_SPELLS
        Call SpellCache_Create(n)
    Next

    For n = 1 To MAX_SHOPS
        Call ShopCache_Create(n)
    Next

    For n = 1 To MAX_ANIMATIONS
        Call AnimationCache_Create(n)
    Next

    For n = 1 To MAX_RESOURCES
        Call ResourceCache_Create(n)
    Next
    
    For n = 1 To MAX_SERIAL_NUMBER
        Call SerialCache_Create(n)
    Next
    
    For n = 1 To MAX_CONVS
        Call ConvCache_Create(n)
    Next
    
    For n = 1 To MAX_QUESTS
        Call QuestCache_Create(n)
    Next
    
    For n = 1 To MAX_GUILDS
        Call GuildCache_Create(n)
    Next
    
    For n = 1 To MAX_CONJUNTOS
        Call ConjuntoCache_Create(n)
    Next n

    If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents

End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim x As Long
    Dim Y As Long
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong MapNum
    Buffer.WriteString Trim$(Map(MapNum).MapData.Name)
    Buffer.WriteString Trim$(Map(MapNum).MapData.Music)
    Buffer.WriteByte Map(MapNum).MapData.Moral
    Buffer.WriteLong Map(MapNum).MapData.Up
    Buffer.WriteLong Map(MapNum).MapData.Down
    Buffer.WriteLong Map(MapNum).MapData.left
    Buffer.WriteLong Map(MapNum).MapData.Right
    Buffer.WriteLong Map(MapNum).MapData.BootMap
    Buffer.WriteByte Map(MapNum).MapData.BootX
    Buffer.WriteByte Map(MapNum).MapData.BootY
    Buffer.WriteByte Map(MapNum).MapData.MaxX
    Buffer.WriteByte Map(MapNum).MapData.MaxY
    Buffer.WriteLong Map(MapNum).MapData.BossNpc
    Buffer.WriteByte Map(MapNum).MapData.Panorama
    
    Buffer.WriteByte Map(MapNum).MapData.Weather
    Buffer.WriteByte Map(MapNum).MapData.WeatherIntensity
    
    Buffer.WriteByte Map(MapNum).MapData.Fog
    Buffer.WriteByte Map(MapNum).MapData.FogSpeed
    Buffer.WriteByte Map(MapNum).MapData.FogOpacity
    
    Buffer.WriteByte Map(MapNum).MapData.Red
    Buffer.WriteByte Map(MapNum).MapData.Green
    Buffer.WriteByte Map(MapNum).MapData.Blue
    Buffer.WriteByte Map(MapNum).MapData.Alpha
    
    Buffer.WriteByte Map(MapNum).MapData.Sun
    
    Buffer.WriteByte Map(MapNum).MapData.DAYNIGHT
    
    
    For I = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(MapNum).MapData.NPC(I)
    Next I

    Buffer.WriteLong Map(MapNum).TileData.EventCount
    If Map(MapNum).TileData.EventCount > 0 Then
        For I = 1 To Map(MapNum).TileData.EventCount
            With Map(MapNum).TileData.Events(I)
                Buffer.WriteString .Name
                Buffer.WriteLong .x
                Buffer.WriteLong .Y
                Buffer.WriteLong .PageCount
            End With
            If Map(MapNum).TileData.Events(I).PageCount > 0 Then
                For x = 1 To Map(MapNum).TileData.Events(I).PageCount
                    With Map(MapNum).TileData.Events(I).EventPage(x)
                        Buffer.WriteByte .chkPlayerVar
                        Buffer.WriteByte .chkSelfSwitch
                        Buffer.WriteByte .chkHasItem
                        Buffer.WriteLong .PlayerVarNum
                        Buffer.WriteLong .SelfSwitchNum
                        Buffer.WriteLong .HasItemNum
                        Buffer.WriteLong .PlayerVariable
                        Buffer.WriteByte .GraphicType
                        Buffer.WriteLong .Graphic
                        Buffer.WriteLong .GraphicX
                        Buffer.WriteLong .GraphicY
                        Buffer.WriteByte .MoveType
                        Buffer.WriteByte .MoveSpeed
                        Buffer.WriteByte .MoveFreq
                        Buffer.WriteByte .WalkAnim
                        Buffer.WriteByte .StepAnim
                        Buffer.WriteByte .DirFix
                        Buffer.WriteByte .WalkThrough
                        Buffer.WriteByte .Priority
                        Buffer.WriteByte .Trigger
                        Buffer.WriteLong .CommandCount
                    End With
                    If Map(MapNum).TileData.Events(I).EventPage(x).CommandCount > 0 Then
                        For Y = 1 To Map(MapNum).TileData.Events(I).EventPage(x).CommandCount
                            With Map(MapNum).TileData.Events(I).EventPage(x).Commands(Y)
                                Buffer.WriteByte .Type
                                Buffer.WriteString .Text
                                Buffer.WriteLong .colour
                                Buffer.WriteByte .Channel
                                Buffer.WriteByte .TargetType
                                Buffer.WriteLong .target
                            End With
                        Next Y
                    End If
                Next x
            End If
        Next I
    End If

    For x = 0 To Map(MapNum).MapData.MaxX
        For Y = 0 To Map(MapNum).MapData.MaxY
            With Map(MapNum).TileData.Tile(x, Y)
                For I = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(I).x
                    Buffer.WriteLong .Layer(I).Y
                    Buffer.WriteLong .Layer(I).Tileset
                    Buffer.WriteByte .Autotile(I)
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteLong .Data4
                Buffer.WriteLong .Data5
                Buffer.WriteByte .DirBlock
            End With
        Next Y
    Next x
    
    Buffer.CompressData
    
    MapCache(MapNum).Data = Buffer.ToArray

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub GuildCache_Create(ByVal GuildNum As Long)
    Dim Buffer As clsBuffer, I As Byte
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong GuildNum
    Buffer.WriteString Guild(GuildNum).Name
    Buffer.WriteString Guild(GuildNum).MOTD
    Buffer.WriteByte Guild(GuildNum).Color
    Buffer.WriteLong Guild(GuildNum).Honra
    Buffer.WriteByte Guild(GuildNum).GuildID
    Buffer.WriteByte Guild(GuildNum).GuildDisponivel
    Buffer.WriteByte Guild(GuildNum).Capacidade
    Buffer.WriteByte Guild(GuildNum).Boost
    Buffer.WriteLong Guild(GuildNum).Kills
    Buffer.WriteLong Guild(GuildNum).Victory
    Buffer.WriteLong Guild(GuildNum).Lose
    Buffer.WriteByte Guild(GuildNum).Icon

    For I = 1 To Guild(GuildNum).Capacidade
        With GuildMembers(GuildNum).Membro(I)
            Buffer.WriteString .Login
            Buffer.WriteString .Name
            Buffer.WriteLong .Level
            Buffer.WriteByte .Online
            Buffer.WriteByte .Dono
            Buffer.WriteByte .Admin
            Buffer.WriteLong .MembroID
            Buffer.WriteByte .MembroDisponivel
        End With
    Next
    
    Buffer.CompressData
    
    GuildCache(GuildNum).Data = Buffer.ToArray()
End Sub

Public Sub QuestCache_Create(ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    Buffer.CompressData
    QuestCache(QuestNum).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub ItemCache_Create(ByVal ItemNum As Long)
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteLong ItemNum
    Buffer.WriteLong ItemCRC32(ItemNum).ItemDataCRC
    Buffer.WriteBytes ItemData
    Buffer.CompressData
    ItemCache(ItemNum).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub NpcCache_Create(ByVal NpcNum As Long)
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Set Buffer = New clsBuffer
    NpcSize = LenB(NPC(NpcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(NPC(NpcNum)), NpcSize
    Buffer.WriteLong NpcNum
    Buffer.WriteLong NpcCRC32(NpcNum).NpcDataCRC
    Buffer.WriteBytes NpcData
    Buffer.CompressData
    NpcCache(NpcNum).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SpellCache_Create(ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    Set Buffer = New clsBuffer
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    Buffer.CompressData
    SpellCache(SpellNum).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub ShopCache_Create(ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    Buffer.WriteLong ShopNum
    Buffer.WriteBytes ShopData
    Buffer.CompressData
    ShopCache(ShopNum).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub AnimationCache_Create(ByVal AnimationNum As Long)
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    Buffer.CompressData
    AnimationCache(AnimationNum).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub ResourceCache_Create(ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    Set Buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    Buffer.CompressData
    ResourceCache(ResourceNum).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SerialCache_Create(ByVal Num As Long)
    Dim Buffer As clsBuffer
    Dim SerialSize As Long
    Dim SerialData() As Byte
    Set Buffer = New clsBuffer
    SerialSize = LenB(Serial(Num))
    ReDim SerialData(SerialSize - 1)
    CopyMemory SerialData(0), ByVal VarPtr(Serial(Num)), SerialSize
    Buffer.WriteLong Num
    Buffer.WriteBytes SerialData
    Buffer.CompressData
    SerialCache(Num).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub ConvCache_Create(ByVal convNum As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Dim x As Long
    
    Set Buffer = New clsBuffer

    Buffer.WriteLong convNum

    With Conv(convNum)
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteLong .chatCount

        For I = 1 To .chatCount
            Buffer.WriteString Trim$(.Conv(I).Conv)

            For x = 1 To 4
                Buffer.WriteString Trim$(.Conv(I).rText(x))
                Buffer.WriteLong .Conv(I).rTarget(x)
            Next

            Buffer.WriteLong .Conv(I).Event
            Buffer.WriteLong .Conv(I).Data1
            Buffer.WriteLong .Conv(I).Data2
            Buffer.WriteLong .Conv(I).Data3
        Next
    End With

    Buffer.CompressData
    
    ConvCache(convNum).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub ConjuntoCache_Create(ByVal ConjuntoNum As Long)
    Dim Buffer As clsBuffer
    Dim ConjuntoSize As Long
    Dim ConjuntoData() As Byte
    Set Buffer = New clsBuffer
    ConjuntoSize = LenB(Conjunto(ConjuntoNum))
    ReDim ConjuntoData(ConjuntoSize - 1)
    CopyMemory ConjuntoData(0), ByVal VarPtr(Conjunto(ConjuntoNum)), ConjuntoSize
    Buffer.WriteLong ConjuntoNum
    Buffer.WriteBytes ConjuntoData
    Buffer.CompressData
    ConjuntoCache(ConjuntoNum).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendGuildAll(ByVal GuildNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateGuild
    Buffer.WriteBytes GuildCache(GuildNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendUpdateGuildTo(ByVal Index As Long, ByVal GuildNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateGuild
    Buffer.WriteBytes GuildCache(GuildNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendQuestAll(ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteBytes QuestCache(QuestNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteBytes QuestCache(QuestNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendSerialsAll(ByVal SerialNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateSerial
    Buffer.WriteBytes SerialCache(SerialNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAdmins Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendUpdateSerialsTo(ByVal Index As Long, ByVal SerialNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateSerial
    Buffer.WriteBytes SerialCache(SerialNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendItemsTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateItem
    Buffer.WriteBytes ItemCache(ItemNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendItemsAll(ByVal ItemNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateItem
    Buffer.WriteBytes ItemCache(ItemNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendAnimationsTo(ByVal Index As Long, ByVal AnimationNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteBytes AnimationCache(AnimationNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendAnimationsAll(ByVal AnimationNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteBytes AnimationCache(AnimationNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendNpcsTo(ByVal Index As Long, ByVal NpcNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteBytes NpcCache(NpcNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendNpcsAll(ByVal NpcNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteBytes NpcCache(NpcNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendResourcesTo(ByVal Index As Long, ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateResource
    Buffer.WriteBytes ResourceCache(ResourceNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendResourcesAll(ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateResource
    Buffer.WriteBytes ResourceCache(ResourceNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendSpellsTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteBytes SpellCache(SpellNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendSpellsAll(ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteBytes SpellCache(SpellNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendShopsTo(ByVal Index As Long, ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateShop
    Buffer.WriteBytes ShopCache(ShopNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendShopsAll(ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateShop
    Buffer.WriteBytes ShopCache(ShopNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendConvAll(ByVal convNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateConv
    Buffer.WriteBytes ConvCache(convNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
Public Sub SendConvTo(ByVal Index As Long, ByVal convNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateConv
    Buffer.WriteBytes ConvCache(convNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub


Sub SendResourceCacheTo(ByVal Index As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong MapResourceCache(GetPlayerMap(Index)).Resource_Count

    If MapResourceCache(GetPlayerMap(Index)).Resource_Count > 0 Then

        For I = 0 To MapResourceCache(GetPlayerMap(Index)).Resource_Count
            Buffer.WriteByte MapResourceCache(GetPlayerMap(Index)).ResourceData(I).ResourceState
            Buffer.WriteLong MapResourceCache(GetPlayerMap(Index)).ResourceData(I).x
            Buffer.WriteLong MapResourceCache(GetPlayerMap(Index)).ResourceData(I).Y
        Next

    End If

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal MapNum As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong MapResourceCache(MapNum).Resource_Count

    If MapResourceCache(MapNum).Resource_Count > 0 Then

        For I = 0 To MapResourceCache(MapNum).Resource_Count
            Buffer.WriteByte MapResourceCache(MapNum).ResourceData(I).ResourceState
            Buffer.WriteLong MapResourceCache(MapNum).ResourceData(I).x
            Buffer.WriteLong MapResourceCache(MapNum).ResourceData(I).Y
        Next

    End If

    SendDataToMap MapNum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub CacheResources(ByVal MapNum As Long)
    Dim x As Long, Y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(MapNum).MapData.MaxX
        For Y = 0 To Map(MapNum).MapData.MaxY

            If Map(MapNum).TileData.Tile(x, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResourceCache(MapNum).ResourceData(0 To Resource_Count)
                MapResourceCache(MapNum).ResourceData(Resource_Count).x = x
                MapResourceCache(MapNum).ResourceData(Resource_Count).Y = Y
                MapResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(Map(MapNum).TileData.Tile(x, Y).Data1).health
            End If

        Next
    Next

    MapResourceCache(MapNum).Resource_Count = Resource_Count
End Sub

Sub SendItems(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_ITEMS

        If LenB(Trim$(Item(I).Name)) > 0 Then
            Call SendItemsTo(Index, I)
        End If

    Next

End Sub

Sub SendAnimations(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(I).Name)) > 0 Then
            Call SendAnimationsTo(Index, I)
        End If

    Next

End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_NPCS

        If LenB(Trim$(NPC(I).Name)) > 0 Then
            Call SendNpcsTo(Index, I)
        End If

    Next

End Sub

Sub SendResources(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(I).Name)) > 0 Then
            Call SendResourcesTo(Index, I)
        End If

    Next

End Sub

Sub SendShops(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(I).Name)) > 0 Then
            Call SendShopsTo(Index, I)
        End If

    Next

End Sub

Sub SendSpells(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(I).Name)) > 0 Then
            Call SendSpellsTo(Index, I)
        End If

    Next

End Sub

Sub SendSerial(ByVal Index As Long)
    Dim I As Long

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    For I = 1 To MAX_SERIAL_NUMBER
        If LenB(Trim$(Serial(I).Name)) > 0 Then
            Call SendUpdateSerialsTo(Index, I)
        End If
    Next I

End Sub
