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

Public Sub IncomingData(ByVal dataLength As Long)
    Dim buffer() As Byte
    Dim pLength As Long
    frmMain.Socket.GetData buffer, vbUnicode, dataLength
    PlayerBuffer.WriteBytes buffer()

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
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    TempBuffer = EncryptPacket(data, (UBound(data) - LBound(data)) + 1)
    length = (UBound(TempBuffer) - LBound(TempBuffer)) + 1

    buffer.PreAllocate 4 + length
    buffer.WriteLong length
    buffer.WriteBytes TempBuffer()

    frmMain.Socket.SendData buffer.ToArray()

    buffer.Flush
    Set buffer = Nothing
End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************
Public Sub SendNewAccount(ByVal AName As String, ByVal APass As String, ByVal ACode As String, BirthDay As String)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer

    buffer.WriteLong CNewAccount
    buffer.WriteString AName
    buffer.WriteString APass
    buffer.WriteString ACode
    buffer.WriteString BirthDay

    buffer.WriteLong CLIENT_MAJOR
    buffer.WriteLong CLIENT_MINOR
    buffer.WriteLong CLIENT_REVISION

    SendData buffer.ToArray()

    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendLogin(ByVal Name As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong CLogin

    buffer.WriteString Name
    buffer.WriteString loginToken

    buffer.WriteLong CLIENT_MAJOR
    buffer.WriteLong CLIENT_MINOR
    buffer.WriteLong CLIENT_REVISION

    SendData buffer.ToArray()

    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendAuthLogin(ByVal Name As String, ByVal Password As String)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer

    buffer.WriteLong CAuthLogin

    buffer.WriteString Name
    buffer.WriteString Password

    buffer.WriteLong CLIENT_MAJOR
    buffer.WriteLong CLIENT_MINOR
    buffer.WriteLong CLIENT_REVISION
    SendData buffer.ToArray()

    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal sex As Long, ByVal ClassNum As Long, ByVal Sprite As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAuthAddChar
    buffer.WriteString Options.TmpLogin
    buffer.WriteString Options.TmpPassword
    buffer.WriteLong CLIENT_MAJOR
    buffer.WriteLong CLIENT_MINOR
    buffer.WriteLong CLIENT_REVISION
    buffer.WriteString Name
    buffer.WriteLong sex
    buffer.WriteLong ClassNum
    buffer.WriteLong Sprite
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SayMsg(ByVal text As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub BroadcastMsg(ByVal text As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CBroadcastMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub GuildMsg(ByVal text As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal text As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub EmoteMsg(ByVal text As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CEmoteMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal text As String, ByVal MsgTo As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString MsgTo
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerMove()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMove
    buffer.WriteLong GetPlayerDir(MyIndex)
    buffer.WriteLong Player(MyIndex).Moving
    buffer.WriteLong Player(MyIndex).X
    buffer.WriteLong Player(MyIndex).Y
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerDir()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerDir
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerRequestNewMap()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNewMap
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMap()
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    CanMoveNow = False

    buffer.WriteLong CMapData

    buffer.WriteString Trim$(Map.MapData.Name)
    buffer.WriteString Trim$(Map.MapData.Music)
    buffer.WriteByte Map.MapData.Moral
    buffer.WriteLong Map.MapData.Up
    buffer.WriteLong Map.MapData.Down
    buffer.WriteLong Map.MapData.Left
    buffer.WriteLong Map.MapData.Right
    buffer.WriteLong Map.MapData.BootMap
    buffer.WriteByte Map.MapData.BootX
    buffer.WriteByte Map.MapData.BootY
    buffer.WriteByte Map.MapData.MaxX
    buffer.WriteByte Map.MapData.MaxY
    buffer.WriteLong Map.MapData.BossNpc
    buffer.WriteByte Map.MapData.Panorama

    buffer.WriteByte Map.MapData.Weather
    buffer.WriteByte Map.MapData.WeatherIntensity

    buffer.WriteByte Map.MapData.Fog
    buffer.WriteByte Map.MapData.FogSpeed
    buffer.WriteByte Map.MapData.FogOpacity

    buffer.WriteByte Map.MapData.Red
    buffer.WriteByte Map.MapData.Green
    buffer.WriteByte Map.MapData.Blue
    buffer.WriteByte Map.MapData.Alpha

    buffer.WriteByte Map.MapData.Sun

    buffer.WriteByte Map.MapData.DayNight

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong Map.MapData.NPC(i)
    Next

    buffer.WriteLong Map.TileData.EventCount
    If Map.TileData.EventCount > 0 Then
        For i = 1 To Map.TileData.EventCount
            With Map.TileData.Events(i)
                buffer.WriteString .Name
                buffer.WriteLong .X
                buffer.WriteLong .Y
                buffer.WriteLong .pageCount
            End With
            If Map.TileData.Events(i).pageCount > 0 Then
                For X = 1 To Map.TileData.Events(i).pageCount
                    With Map.TileData.Events(i).EventPage(X)
                        buffer.WriteByte .chkPlayerVar
                        buffer.WriteByte .chkSelfSwitch
                        buffer.WriteByte .chkHasItem
                        buffer.WriteLong .PlayerVarNum
                        buffer.WriteLong .SelfSwitchNum
                        buffer.WriteLong .HasItemNum
                        buffer.WriteLong .PlayerVariable
                        buffer.WriteByte .GraphicType
                        buffer.WriteLong .Graphic
                        buffer.WriteLong .GraphicX
                        buffer.WriteLong .GraphicY
                        buffer.WriteByte .MoveType
                        buffer.WriteByte .MoveSpeed
                        buffer.WriteByte .MoveFreq
                        buffer.WriteByte .WalkAnim
                        buffer.WriteByte .StepAnim
                        buffer.WriteByte .DirFix
                        buffer.WriteByte .WalkThrough
                        buffer.WriteByte .Priority
                        buffer.WriteByte .Trigger
                        buffer.WriteLong .CommandCount
                    End With
                    If Map.TileData.Events(i).EventPage(X).CommandCount > 0 Then
                        For Y = 1 To Map.TileData.Events(i).EventPage(X).CommandCount
                            With Map.TileData.Events(i).EventPage(X).Commands(Y)
                                buffer.WriteByte .Type
                                buffer.WriteString .text
                                buffer.WriteLong .Colour
                                buffer.WriteByte .channel
                                buffer.WriteByte .TargetType
                                buffer.WriteLong .Target
                                buffer.WriteLong .X
                                buffer.WriteLong .Y
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
                    buffer.WriteLong .Layer(i).X
                    buffer.WriteLong .Layer(i).Y
                    buffer.WriteLong .Layer(i).tileSet
                    buffer.WriteByte .Autotile(i)
                Next
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteLong .Data4
                buffer.WriteLong .Data5
                buffer.WriteByte .DirBlock
            End With
        Next
    Next

    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub WarpMeTo(ByVal Name As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpMeTo
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub WarpToMe(ByVal Name As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpToMe
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpTo
    buffer.WriteLong MapNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSetAccess
    buffer.WriteString Name
    buffer.WriteLong Access
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSetSprite
    buffer.WriteLong SpriteNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendKick(ByVal Name As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CKickPlayer
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendBan(ByVal Name As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CBanPlayer
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditItem()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditItem
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveItem(ByVal itemNum As Long)
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(itemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemNum)), ItemSize
    buffer.WriteLong CSaveItem
    buffer.WriteLong itemNum
    buffer.WriteBytes ItemData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditAnimation()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditAnimation
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    buffer.WriteLong CSaveAnimation
    buffer.WriteLong Animationnum
    buffer.WriteBytes AnimationData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditNpc()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditNpc
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveNpc(ByVal NpcNum As Long)
    Dim buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Set buffer = New clsBuffer
    NpcSize = LenB(NPC(NpcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(NPC(NpcNum)), NpcSize
    buffer.WriteLong CSaveNpc
    buffer.WriteLong NpcNum
    buffer.WriteBytes NpcData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditResource()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditResource
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    Set buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    buffer.WriteLong CSaveResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMapRespawn()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CMapRespawn
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUseItem(ByVal invNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CUseItem
    buffer.WriteLong invNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendDropItem(ByVal invNum As Long, ByVal Amount As Long)
    Dim buffer As clsBuffer

    If InBank Or InShop Then Exit Sub

    ' do basic checks
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    If PlayerInv(invNum).Num < 1 Or PlayerInv(invNum).Num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
        If Amount < 1 Or Amount > PlayerInv(invNum).Value Then Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CMapDropItem
    buffer.WriteLong invNum
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendWhosOnline()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CWhosOnline
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMOTDChange(ByVal Motd As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSetMotd
    buffer.WriteString Motd
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditShop()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditShop
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveShop(ByVal shopNum As Long)
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set buffer = New clsBuffer
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    buffer.WriteLong CSaveShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditSpell()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditSpell
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    Set buffer = New clsBuffer
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    buffer.WriteLong CSaveSpell
    buffer.WriteLong spellnum
    buffer.WriteBytes SpellData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditMap()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditMap
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSwapInvSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
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

    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSwapSpellSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    ' buffer it
    PlayerSwitchSpellSlots OldSlot, NewSlot
End Sub

Sub GetPing()
    Dim buffer As clsBuffer
    PingStart = getTime
    Set buffer = New clsBuffer
    buffer.WriteLong CCheckPing
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUnequip(ByVal eqNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CUnequip
    buffer.WriteLong eqNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestPlayerData()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestPlayerData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestItems()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestItems
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestAnimations()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestAnimations
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestNPCS()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNPCS
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestResources()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestResources
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestSpells()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestSpells
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendRequestShops()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestShops
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSpawnItem
    buffer.WriteLong tmpItem
    buffer.WriteLong tmpAmount
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTrainStat(ByVal statNum As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CUseStatPoint
    buffer.WriteByte statNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestLevelUp()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestLevelUp
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub BuyItem(ByVal shopSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CBuyItem
    buffer.WriteLong shopSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SellItem(ByVal InvSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSellItem
    buffer.WriteLong InvSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub AdminWarp(ByVal X As Long, ByVal Y As Long)
    If X < 0 Or Y < 0 Or X > Map.MapData.MaxX Or Y > Map.MapData.MaxY Then Exit Sub
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAdminWarp
    buffer.WriteLong X
    buffer.WriteLong Y
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub AcceptTrade()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTrade
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub DeclineTrade()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTrade
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub TradeItem(ByVal InvSlot As Long, ByVal Amount As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeItem
    buffer.WriteLong InvSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub TradeGold(ByVal Amount As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeGold
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub UntradeItem(ByVal InvSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CUntradeItem
    buffer.WriteLong InvSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
    Dim buffer As clsBuffer

    'Clear the hotbarnum if droped
    If sType = 0 And Slot = 0 Then
        Hotbar(hotbarNum).sType = 0
        Hotbar(hotbarNum).Slot = 0
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarChange
    buffer.WriteLong sType
    buffer.WriteLong Slot
    buffer.WriteLong hotbarNum
    SendData buffer.ToArray()
    Set buffer = Nothing
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
    Dim buffer As clsBuffer, X As Long

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

    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarUse
    buffer.WriteLong Slot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMapReport()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CMapReport
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub PlayerTarget(ByVal Target As Long, ByVal TargetType As Long)
    Dim buffer As clsBuffer

    If myTargetType = TargetType And myTarget = Target Then
        myTargetType = 0
        myTarget = 0
    Else
        myTarget = Target
        myTargetType = TargetType
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CTarget
    buffer.WriteLong Target
    buffer.WriteLong TargetType
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeRequest(playerIndex As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeRequest
    buffer.WriteLong playerIndex
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendAcceptTradeRequest()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendDeclineTradeRequest()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyLeave()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyLeave
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyRequest(Index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyRequest
    buffer.WriteLong Index
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendAcceptParty()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptParty
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendDeclineParty()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineParty
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditConv()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditConv
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveConv(ByVal Convnum As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim X As Long

    Set buffer = New clsBuffer

    buffer.WriteLong CSaveConv
    buffer.WriteLong Convnum
    With Conv(Convnum)
        buffer.WriteString .Name
        buffer.WriteLong .chatCount
        For i = 1 To .chatCount
            buffer.WriteString .Conv(i).Conv
            For X = 1 To 4
                buffer.WriteString .Conv(i).rText(X)
                buffer.WriteLong .Conv(i).rTarget(X)
            Next
            buffer.WriteLong .Conv(i).Event
            buffer.WriteLong .Conv(i).Data1
            buffer.WriteLong .Conv(i).Data2
            buffer.WriteLong .Conv(i).Data3
        Next
    End With

    SendData buffer.ToArray()

    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendRequestConvs()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestConvs
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendChatOption(ByVal Index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CChatOption
    buffer.WriteLong Index
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendFinishTutorial()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CFinishTutorial
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCloseShop()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CCloseShop
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
