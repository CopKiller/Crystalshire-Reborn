Attribute VB_Name = "modClientTCP"
Option Explicit
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer

Sub TcpInit(ByVal IP As String, ByVal Port As Long)
    Set PlayerBuffer = Nothing
    Set PlayerBuffer = New clsBuffer
    ' connect
    frmMain.Socket.Close
    frmMain.Socket.RemoteHost = IP
    frmMain.Socket.RemotePort = Port
End Sub

Sub DestroyTCP()
    frmMain.Socket.Close
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long
    frmMain.Socket.GetData Buffer, vbUnicode, DataLength
    PlayerBuffer.WriteBytes Buffer()

    If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)

    Do While pLength > 0 And pLength <= PlayerBuffer.length - 4

        If pLength <= PlayerBuffer.length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0

        If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop

    PlayerBuffer.Trim

    If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
End Sub

Public Function ConnectToServer() As Boolean
    Dim Wait As Currency

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If

    Wait = getTime
    frmMain.Socket.Close
    frmMain.Socket.Connect
    SetStatus "Connecting to server."

    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (getTime <= Wait + 3000)
        If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
    Loop

    ConnectToServer = IsConnected
    SetStatus vbNullString
End Function

Function IsConnected() As Boolean

    If frmMain.Socket.state = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If

End Function

Sub SendData(ByRef data() As Byte)

    If Not IsConnected Then
        Exit Sub
    End If

    Dim TempBuffer() As Byte

    Dim length As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    TempBuffer = EncryptPacket(data, (UBound(data) - LBound(data)) + 1)
    length = (UBound(TempBuffer) - LBound(TempBuffer)) + 1

    Buffer.PreAllocate 4 + length
    Buffer.WriteLong length
    Buffer.WriteBytes TempBuffer()

    frmMain.Socket.SendData Buffer.ToArray()

    Buffer.Flush
    Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************
Public Sub SendNewAccount(ByVal AName As String, ByVal APass As String, ByVal ACode As String, BirthDay As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteLong CNewAccount
    Buffer.WriteString AName
    Buffer.WriteString APass
    Buffer.WriteString ACode
    Buffer.WriteString BirthDay
    
    Buffer.WriteLong CLIENT_MAJOR
    Buffer.WriteLong CLIENT_MINOR
    Buffer.WriteLong CLIENT_REVISION

    SendData Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendLogin(ByVal Name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong CLogin

    Buffer.WriteString Name
    Buffer.WriteString loginToken

    Buffer.WriteLong CLIENT_MAJOR
    Buffer.WriteLong CLIENT_MINOR
    Buffer.WriteLong CLIENT_REVISION

    SendData Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendAuthLogin(ByVal Name As String, ByVal Password As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteLong CAuthLogin

    Buffer.WriteString Name
    Buffer.WriteString Password

    Buffer.WriteLong CLIENT_MAJOR
    Buffer.WriteLong CLIENT_MINOR
    Buffer.WriteLong CLIENT_REVISION
    SendData Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal sex As Long, ByVal ClassNum As Long, ByVal Sprite As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAuthAddChar
    Buffer.WriteString Options.TmpLogin
    Buffer.WriteString Options.TmpPassword
    Buffer.WriteLong CLIENT_MAJOR
    Buffer.WriteLong CLIENT_MINOR
    Buffer.WriteLong CLIENT_REVISION
    Buffer.WriteString Name
    Buffer.WriteLong sex
    Buffer.WriteLong ClassNum
    Buffer.WriteLong Sprite
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SayMsg(ByVal Text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub BroadcastMsg(ByVal Text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBroadcastMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub GuildMsg(ByVal Text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CGuildMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal Text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub EmoteMsg(ByVal Text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CEmoteMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
    Buffer.WriteString MsgTo
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerMove()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong GetPlayerDir(MyIndex)
    Buffer.WriteLong Player(MyIndex).Moving
    Buffer.WriteLong Player(MyIndex).X
    Buffer.WriteLong Player(MyIndex).Y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerDir()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerDir
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerRequestNewMap()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNewMap
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMap()
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    CanMoveNow = False

    Buffer.WriteLong CMapData

    Buffer.WriteString Trim$(Map.MapData.Name)
    Buffer.WriteString Trim$(Map.MapData.Music)
    Buffer.WriteByte Map.MapData.Moral
    Buffer.WriteLong Map.MapData.Up
    Buffer.WriteLong Map.MapData.Down
    Buffer.WriteLong Map.MapData.Left
    Buffer.WriteLong Map.MapData.Right
    Buffer.WriteLong Map.MapData.BootMap
    Buffer.WriteByte Map.MapData.BootX
    Buffer.WriteByte Map.MapData.BootY
    Buffer.WriteByte Map.MapData.MaxX
    Buffer.WriteByte Map.MapData.MaxY
    Buffer.WriteLong Map.MapData.BossNpc
    Buffer.WriteByte Map.MapData.Panorama

    Buffer.WriteByte Map.MapData.Weather
    Buffer.WriteByte Map.MapData.WeatherIntensity

    Buffer.WriteByte Map.MapData.Fog
    Buffer.WriteByte Map.MapData.FogSpeed
    Buffer.WriteByte Map.MapData.FogOpacity

    Buffer.WriteByte Map.MapData.Red
    Buffer.WriteByte Map.MapData.Green
    Buffer.WriteByte Map.MapData.Blue
    Buffer.WriteByte Map.MapData.Alpha
    
    Buffer.WriteByte Map.MapData.Sun
    
    Buffer.WriteByte Map.MapData.DayNight
    
    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map.MapData.NPC(i)
    Next

    Buffer.WriteLong Map.TileData.EventCount
    If Map.TileData.EventCount > 0 Then
        For i = 1 To Map.TileData.EventCount
            With Map.TileData.Events(i)
                Buffer.WriteString .Name
                Buffer.WriteLong .X
                Buffer.WriteLong .Y
                Buffer.WriteLong .pageCount
            End With
            If Map.TileData.Events(i).pageCount > 0 Then
                For X = 1 To Map.TileData.Events(i).pageCount
                    With Map.TileData.Events(i).EventPage(X)
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
                    If Map.TileData.Events(i).EventPage(X).CommandCount > 0 Then
                        For Y = 1 To Map.TileData.Events(i).EventPage(X).CommandCount
                            With Map.TileData.Events(i).EventPage(X).Commands(Y)
                                Buffer.WriteByte .Type
                                Buffer.WriteString .Text
                                Buffer.WriteLong .Colour
                                Buffer.WriteByte .channel
                                Buffer.WriteByte .TargetType
                                Buffer.WriteLong .Target
                                Buffer.WriteLong .X
                                Buffer.WriteLong .Y
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            With Map.TileData.Tile(X, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).X
                    Buffer.WriteLong .Layer(i).Y
                    Buffer.WriteLong .Layer(i).tileSet
                    Buffer.WriteByte .Autotile(i)
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteLong .Data4
                Buffer.WriteLong .Data5
                Buffer.WriteByte .DirBlock
            End With
        Next
    Next

    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpMeTo(ByVal Name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpMeTo
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpToMe(ByVal Name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpToMe
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpTo
    Buffer.WriteLong MapNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetAccess
    Buffer.WriteString Name
    Buffer.WriteLong Access
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetSprite
    Buffer.WriteLong SpriteNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendKick(ByVal Name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CKickPlayer
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBan(ByVal Name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanPlayer
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditItem()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditItem
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveItem(ByVal itemNum As Long)
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemNum)), ItemSize
    Buffer.WriteLong CSaveItem
    Buffer.WriteLong itemNum
    Buffer.WriteBytes ItemData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditAnimation()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditAnimation
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    Buffer.WriteLong CSaveAnimation
    Buffer.WriteLong Animationnum
    Buffer.WriteBytes AnimationData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditNpc()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditNpc
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveNpc(ByVal NpcNum As Long)
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Set Buffer = New clsBuffer
    NpcSize = LenB(NPC(NpcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(NPC(NpcNum)), NpcSize
    Buffer.WriteLong CSaveNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteBytes NpcData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditResource()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditResource
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    Set Buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    Buffer.WriteLong CSaveResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapRespawn()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapRespawn
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendUseItem(ByVal invNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseItem
    Buffer.WriteLong invNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendDropItem(ByVal invNum As Long, ByVal Amount As Long)
    Dim Buffer As clsBuffer

    If InBank Or InShop Then Exit Sub

    ' do basic checks
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    If PlayerInv(invNum).num < 1 Or PlayerInv(invNum).num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
        If Amount < 1 Or Amount > PlayerInv(invNum).Value Then Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapDropItem
    Buffer.WriteLong invNum
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendWhosOnline()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWhosOnline
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMOTDChange(ByVal Motd As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetMotd
    Buffer.WriteString Motd
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditShop()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveShop(ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    Buffer.WriteLong CSaveShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditSpell()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditSpell
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    Set Buffer = New clsBuffer
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    Buffer.WriteLong CSaveSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditMap()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapInvSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    ' buffer it
    PlayerSwitchInvSlots OldSlot, NewSlot
End Sub

Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)

    ' Check for subscript out of range
    If OldSlot < 1 Or OldSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    ElseIf NewSlot < 1 Or NewSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

   ' dont let them swap a spell which is in CD
    If SpellCD(OldSlot) > 0 Then
        AddText "Cannot swap a spell which is cooling down!", BrightRed
        Exit Sub
    ElseIf SpellCD(NewSlot) > 0 Then
        AddText "Cannot swap a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If SpellBuffer = OldSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    ElseIf SpellBuffer = NewSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapSpellSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    ' buffer it
    PlayerSwitchSpellSlots OldSlot, NewSlot
End Sub

Sub GetPing()
    Dim Buffer As clsBuffer
    PingStart = getTime
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCheckPing
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUnequip(ByVal eqNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUnequip
    Buffer.WriteLong eqNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestPlayerData()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPlayerData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestItems()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestItems
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestAnimations()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestAnimations
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestNPCS()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNPCS
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestResources()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestResources
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestSpells()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestSpells
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestShops()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestShops
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSpawnItem
    Buffer.WriteLong tmpItem
    Buffer.WriteLong tmpAmount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTrainStat(ByVal statNum As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseStatPoint
    Buffer.WriteByte statNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestLevelUp()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestLevelUp
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub BuyItem(ByVal shopSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBuyItem
    Buffer.WriteLong shopSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SellItem(ByVal InvSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSellItem
    Buffer.WriteLong InvSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub AdminWarp(ByVal X As Long, ByVal Y As Long)
    If X < 0 Or Y < 0 Or X > Map.MapData.MaxX Or Y > Map.MapData.MaxY Then Exit Sub
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAdminWarp
    Buffer.WriteLong X
    Buffer.WriteLong Y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub AcceptTrade()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub DeclineTrade()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub TradeItem(ByVal InvSlot As Long, ByVal Amount As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeItem
    Buffer.WriteLong InvSlot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub TradeGold(ByVal Amount As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeGold
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub UntradeItem(ByVal InvSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUntradeItem
    Buffer.WriteLong InvSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
    Dim Buffer As clsBuffer
    
    'Clear the hotbarnum if droped
    If sType = 0 And Slot = 0 Then
        Hotbar(hotbarNum).sType = 0
        Hotbar(hotbarNum).Slot = 0
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarChange
    Buffer.WriteLong sType
    Buffer.WriteLong Slot
    Buffer.WriteLong hotbarNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerSwitchHotbarSlots(ByVal OldSlot As Byte, ByVal NewSlot As Byte)
Dim OldSlotNum As Integer, NewSlotNum As Integer, OldSlotType As Byte, NewSlotType As Byte
    
    OldSlotNum = Hotbar(OldSlot).Slot
    NewSlotNum = Hotbar(NewSlot).Slot
    OldSlotType = Hotbar(OldSlot).sType
    NewSlotType = Hotbar(NewSlot).sType
    
    Hotbar(OldSlot).Slot = NewSlotNum
    Hotbar(NewSlot).Slot = OldSlotNum
    Hotbar(OldSlot).sType = NewSlotType
    Hotbar(NewSlot).sType = OldSlotType
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
    Dim Buffer As clsBuffer, X As Long

    ' check if spell
    If Hotbar(Slot).sType = 2 Then    ' spell

        For X = 1 To MAX_PLAYER_SPELLS

            ' is the spell matching the hotbar?
            If PlayerSpells(X).Spell = Hotbar(Slot).Slot Then
                ' found it, cast it
                CastSpell X
                Exit Sub
            End If

        Next

        ' can't find the spell, exit out
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarUse
    Buffer.WriteLong Slot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapReport()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapReport
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerTarget(ByVal Target As Long, ByVal TargetType As Long)
    Dim Buffer As clsBuffer

    If myTargetType = TargetType And myTarget = Target Then
        myTargetType = 0
        myTarget = 0
    Else
        myTarget = Target
        myTargetType = TargetType
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTarget
    Buffer.WriteLong Target
    Buffer.WriteLong TargetType
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest(playerIndex As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeRequest
    Buffer.WriteLong playerIndex
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAcceptTradeRequest()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDeclineTradeRequest()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyLeave()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyLeave
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyRequest(Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyRequest
    Buffer.WriteLong Index
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAcceptParty()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDeclineParty()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditConv()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditConv
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveConv(ByVal Convnum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim X As Long

    Set Buffer = New clsBuffer

    Buffer.WriteLong CSaveConv
    Buffer.WriteLong Convnum
    With Conv(Convnum)
        Buffer.WriteString .Name
        Buffer.WriteLong .chatCount
        For i = 1 To .chatCount
            Buffer.WriteString .Conv(i).Conv
            For X = 1 To 4
                Buffer.WriteString .Conv(i).rText(X)
                Buffer.WriteLong .Conv(i).rTarget(X)
            Next
            Buffer.WriteLong .Conv(i).Event
            Buffer.WriteLong .Conv(i).Data1
            Buffer.WriteLong .Conv(i).Data2
            Buffer.WriteLong .Conv(i).Data3
        Next
    End With

    SendData Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendRequestConvs()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestConvs
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendChatOption(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CChatOption
    Buffer.WriteLong Index
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendFinishTutorial()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CFinishTutorial
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendCloseShop()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
