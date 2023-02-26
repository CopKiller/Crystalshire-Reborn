Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = Options.GAME_NAME & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Function IsConnected(ByVal Index As Long) As Boolean

    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean

    If IsConnected(Index) Then
        If LenB(Trim$(Player(Index).Name)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal Index As Integer, ByVal Login As String)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsConnected(I) Then
            If LCase$(Trim$(Player(I).Login)) = SanitiseString(LCase$(Login)) Then
                If I <> Index Then
                    Call AlertMsg(I, DIALOGUE_MSG_KICKED, NO, True)
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim I As Long
    Dim n As Long

    For I = 1 To Player_HighIndex

        If IsConnected(I) Then
            If Trim$(GetPlayerIP(I)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
    On Error GoTo ErrorHandle

    If Not IsConnected(Index) Then
        Exit Sub
    End If

    Dim TempBuffer() As Byte

    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    TempBuffer = EncryptPacket(Data, (UBound(Data) - LBound(Data)) + 1)

    Buffer.WriteLong (UBound(TempBuffer) - LBound(TempBuffer)) + 1
    Buffer.WriteBytes TempBuffer()

    frmServer.Socket(Index).SendData Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing

    Exit Sub

ErrorHandle:
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            Call SendDataTo(I, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            If I <> Index Then
                Call SendDataTo(I, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                If I <> Index Then
                    Call SendDataTo(I, Data)
                End If
            End If
        End If

    Next

End Sub

Sub SendDataToParty(ByVal partynum As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Party(partynum).MemberCount
        If Party(partynum).Member(I) > 0 Then
            Call SendDataTo(Party(partynum).Member(I), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToAll Buffer.ToArray

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer

    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color

    For I = 1 To Player_HighIndex
        If IsPlaying(I) And GetPlayerAccess(I) > 0 Then
            SendDataTo I, Buffer.ToArray
        End If
    Next

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataTo Index, Buffer.ToArray

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToMap MapNum, Buffer.ToArray

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal MessageNo As Long, Optional ByVal MenuReset As Long = 0, Optional ByVal kick As Boolean = True)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteLong MessageNo
    Buffer.WriteLong MenuReset
    If kick Then Buffer.WriteLong 1 Else Buffer.WriteLong 0
    SendDataTo Index, Buffer.ToArray

    Buffer.Flush: Set Buffer = Nothing

    If kick Then
        If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
        Call CloseSocket(Index)
    End If

End Sub

Public Sub PartyMsg(ByVal partynum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim I As Long
    ' send message to all people
    For I = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partynum).Member(I) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partynum).Member(I)) And IsPlaying(Party(partynum).Member(I)) Then
                SayMsg_Party Party(partynum).Member(I), Msg, Color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal Index As Long)
    Call AlertMsg(Index, DIALOGUE_MSG_CONNECTION)
End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim I As Long

    If (Index = 0) Then
        I = FindOpenPlayerSlot
        If I <> 0 And GetPlayerIP(Index) <> vbNullString Then
            ' we can connect them
            frmServer.Socket(I).Close
            frmServer.Socket(I).Accept SocketId
            Call SocketConnected(I)
        End If
    End If

End Sub

Sub SocketConnected(ByVal Index As Long)
    Dim I As Long

    If Index <> 0 Then
        Call TextAdd("Received connection from " & GetPlayerIP(Index) & ".")

        TempPlayer(Index).ConnectedTime = getTime
        SetHighIndex

        LoginTokenAccepted(Index) = False
        ' re-set the high index
        SendHighIndex
    End If
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long

    If GetPlayerAccess(Index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 1000 Then
            Exit Sub
        End If

        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 25 Then
            Exit Sub
        End If
    End If

    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    If getTime >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = getTime + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If

    ' Get the data from the socket now
    frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(Index).Buffer.WriteBytes Buffer()

    If TempPlayer(Index).Buffer.Length >= 4 Then
        pLength = TempPlayer(Index).Buffer.ReadLong(False)

        If pLength < 0 Then
            Exit Sub
        End If
    End If

    Do While pLength > 0 And pLength <= TempPlayer(Index).Buffer.Length - 4
        If pLength <= TempPlayer(Index).Buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).Buffer.ReadLong
            HandleData Index, TempPlayer(Index).Buffer.ReadBytes(pLength)
        End If

        pLength = 0
        If TempPlayer(Index).Buffer.Length >= 4 Then
            pLength = TempPlayer(Index).Buffer.ReadLong(False)

            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop

    TempPlayer(Index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal Index As Long)

    If Index > 0 And GetPlayerIP(Index) <> vbNullString Then

        LoginTokenAccepted(Index) = False

        Call LeftGame(Index)

        Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
        frmServer.Socket(Index).Close
        Call UpdateCaption

        Call ClearPlayer(Index)

        ' Set The High Index
        Call SetHighIndex
        SendHighIndex
    Else
        LoginTokenAccepted(Index) = False
        frmServer.Socket(Index).Close
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If

End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim n As Long
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            If I <> Index Then
                s = s & GetPlayerName(I) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(Index, s, WhoColor)
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
    Dim Buffer As clsBuffer, I As Long

    If Index > Player_HighIndex Then Exit Function
    Set Buffer = New clsBuffer

    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerLevel(Index)
    Buffer.WriteLong GetPlayerPOINTS(Index)
    Buffer.WriteLong GetPlayerSprite(Index)
    Buffer.WriteLong GetPlayerMap(Index)
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteLong GetPlayerClass(Index)
    Buffer.WriteInteger Player(Index).Guild_ID
    Buffer.WriteByte Player(Index).Guild_MembroID
    Buffer.WriteByte GetPlayerPremium(Index)
    Buffer.WriteLong GetPlayerGold(Index)

    For I = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, I)
    Next I

    For I = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(Index, I)
        Buffer.WriteLong GetPlayerVital(Index, I)
    Next I

    For I = 1 To (Status_Count - 1)
        If TempPlayer(Index).StatusNum(I).Ativo > 0 Then
            Buffer.WriteByte I
        End If
    Next

    PlayerData = Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Function

Sub SendJoinMap(ByVal Index As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If I <> Index Then
                If GetPlayerMap(I) = GetPlayerMap(Index) Then
                    SendDataTo Index, PlayerData(I)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SLeft
    Buffer.WriteLong Index
    SendDataToMapBut Index, MapNum, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
End Sub

Sub SendPlayerData_Party(partynum As Long)
    Dim I As Long, X As Long
    ' loop through all the party members
    For I = 1 To Party(partynum).MemberCount
        For X = 1 To Party(partynum).MemberCount
            SendDataTo Party(partynum).Member(X), PlayerData(Party(partynum).Member(I))
        Next
    Next
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    'Buffer.PreAllocate (UBound(MapCache(mapNum).Data) - LBound(MapCache(mapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapItemData

    For I = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(MapNum, I).PlayerName
        Buffer.WriteLong MapItem(MapNum, I).Num
        Buffer.WriteLong MapItem(MapNum, I).Value
        Buffer.WriteLong MapItem(MapNum, I).X
        Buffer.WriteLong MapItem(MapNum, I).Y
        Buffer.WriteByte MapItem(MapNum, I).Bound
    Next

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapItemData

    For I = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(MapNum, I).PlayerName
        Buffer.WriteLong MapItem(MapNum, I).Num
        Buffer.WriteLong MapItem(MapNum, I).Value
        Buffer.WriteLong MapItem(MapNum, I).X
        Buffer.WriteLong MapItem(MapNum, I).Y
        Buffer.WriteByte MapItem(MapNum, I).Bound
    Next

    SendDataToMap MapNum, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal MapNum As Long, ByVal mapnpcnum As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapNpcVitals
    Buffer.WriteByte mapnpcnum

    For I = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNpc(MapNum).NPC(mapnpcnum).Vital(I)
    Next

    SendDataToMap MapNum, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapNpcData

    For I = 1 To MAX_MAP_NPCS
        Buffer.WriteInteger MapNpc(MapNum).NPC(I).Num
        Buffer.WriteByte MapNpc(MapNum).NPC(I).X
        Buffer.WriteByte MapNpc(MapNum).NPC(I).Y
        Buffer.WriteByte MapNpc(MapNum).NPC(I).dir
        Buffer.WriteLong MapNpc(MapNum).NPC(I).Vital(Vitals.HP)
        Buffer.WriteLong MapNpc(MapNum).NPC(I).Vital(Vitals.MP)
        Buffer.WriteLong MapNpc(MapNum).NPC(I).StunDuration
        
        Buffer.WriteByte MapNpc(MapNum).NPC(I).Dead
        Buffer.WriteInteger MapNpc(MapNum).NPC(I).tmpNum
    Next

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapNpcData

    For I = 1 To MAX_MAP_NPCS
        Buffer.WriteInteger MapNpc(MapNum).NPC(I).Num
        Buffer.WriteByte MapNpc(MapNum).NPC(I).X
        Buffer.WriteByte MapNpc(MapNum).NPC(I).Y
        Buffer.WriteByte MapNpc(MapNum).NPC(I).dir
        Buffer.WriteLong MapNpc(MapNum).NPC(I).Vital(Vitals.HP)
        Buffer.WriteLong MapNpc(MapNum).NPC(I).Vital(Vitals.MP)
        Buffer.WriteLong MapNpc(MapNum).NPC(I).StunDuration
        
        Buffer.WriteByte MapNpc(MapNum).NPC(I).Dead
        Buffer.WriteInteger MapNpc(MapNum).NPC(I).tmpNum
    Next

    SendDataToMap MapNum, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SPlayerInv

    For I = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(Index, I)
        Buffer.WriteLong GetPlayerInvItemValue(Index, I)
        Buffer.WriteByte GetPlayerInvItemBound(Index, I)
    Next

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteByte InvSlot
    Buffer.WriteInteger GetPlayerInvItemNum(Index, InvSlot)
    Buffer.WriteLong GetPlayerInvItemValue(Index, InvSlot)
    Buffer.WriteByte GetPlayerInvItemBound(Index, InvSlot)
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SPlayerWornEq
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Armor)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Weapon)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Helmet)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Shield)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Legs)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Boots)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Amulet)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, RingLeft)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, RingRight)
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong Index
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Armor)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Weapon)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Helmet)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Shield)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Legs)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Boots)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, Amulet)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, RingLeft)
    Buffer.WriteInteger GetPlayerEquipmentNum(Index, RingRight)

    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    Buffer.WriteInteger GetPlayerEquipmentNum(PlayerNum, Armor)
    Buffer.WriteInteger GetPlayerEquipmentNum(PlayerNum, Weapon)
    Buffer.WriteInteger GetPlayerEquipmentNum(PlayerNum, Helmet)
    Buffer.WriteInteger GetPlayerEquipmentNum(PlayerNum, Shield)
    Buffer.WriteInteger GetPlayerEquipmentNum(PlayerNum, Legs)
    Buffer.WriteInteger GetPlayerEquipmentNum(PlayerNum, Boots)
    Buffer.WriteInteger GetPlayerEquipmentNum(PlayerNum, Amulet)
    Buffer.WriteInteger GetPlayerEquipmentNum(PlayerNum, RingLeft)
    Buffer.WriteInteger GetPlayerEquipmentNum(PlayerNum, RingRight)

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
    Case HP
        Buffer.WriteLong SPlayerHp
        Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
        Buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
    Case MP
        Buffer.WriteLong SPlayerMp
        Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
        Buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
    End Select

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing

    ' check if they're in a party
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
End Sub

Sub SendMapVitals(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim I As Byte

    Buffer.WriteLong SSendMapHpMp
    Buffer.WriteByte Index

    For I = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
        Buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
    Next I

    SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing

    ' check if they're in a party
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
End Sub

Sub SendEXP(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(Index)
    Buffer.WriteLong GetPlayerNextLevel(Index)

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendStats(ByVal Index As Long)
    Dim I As Long
    Dim packet As String
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    For I = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, I)
    Next
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendWelcome(ByVal Index As Long)

' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(Index, Options.MOTD, BrightCyan)
    End If

    ' Guild MOTD
    If Player(Index).Guild_ID > 0 Then
        If LenB(Guild(Player(Index).Guild_ID).MOTD) > 0 Then
            Call PlayerMsg(Index, Guild(Player(Index).Guild_ID).MOTD, Yellow)
        End If
    End If

    ' Send whos online
    Call SendWhosOnline(Index)
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim packet As String
    Dim I As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClassesData
    Buffer.WriteLong Max_Classes

    For I = 1 To Max_Classes
        Buffer.WriteString GetClassName(I)
        Buffer.WriteLong Class(I).MaxHP
        Buffer.WriteLong Class(I).MaxMP

        ' set sprite array size
        n = UBound(Class(I).MaleSprite)

        ' send array size
        Buffer.WriteLong n

        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(I).MaleSprite(q)
        Next

        ' set sprite array size
        n = UBound(Class(I).FemaleSprite)

        ' send array size
        Buffer.WriteLong n

        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(I).FemaleSprite(q)
        Next

        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(I).Stat(q)
        Next
    Next

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendLeftGame(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    SendDataToAllBut Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For I = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong Player(Index).Spell(I).Spell
        Buffer.WriteLong Player(Index).Spell(I).Uses
    Next

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPlayerSpellsCD(ByVal Index As Long, ByVal Slot As Byte)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong Slot
    Buffer.WriteInteger Player(Index).SpellCD(Slot)

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SDoorAnimation
    Buffer.WriteLong X
    Buffer.WriteLong Y

    SendDataToMap MapNum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendActionMsg(ByVal MapNum As Long, ByVal message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal X As Long, ByVal Y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString message
    Buffer.WriteLong Color
    Buffer.WriteLong MsgType
    Buffer.WriteLong X
    Buffer.WriteLong Y

    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendBlood(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong X
    Buffer.WriteLong Y

    SendDataToMap MapNum, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendAnimation(ByVal MapNum As Long, ByVal Anim As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0, Optional isCasting As Byte = 0)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    Buffer.WriteByte isCasting

    SendDataToMap MapNum, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendCooldown(ByVal Index As Long, ByVal Slot As Long, ByVal Seconds As Integer)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong Slot
    Buffer.WriteInteger Seconds

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal MapNum As Long, ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour

    SendDataToMap MapNum, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour

    SendDataToAll Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SayMsg_Party(ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Party] "
    Buffer.WriteLong saycolour

    SendDataToAll Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SayMsg_Guild(ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Guild] "
    Buffer.WriteLong saycolour

    SendDataToAll Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub ResetShopAction(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction

    SendDataToAll Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendStunned(ByVal MapNum As Long, ByVal Index As Long, ByVal TargetType As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong MapNum
    Buffer.WriteLong Index
    Buffer.WriteByte TargetType

    If TargetType = TARGET_TYPE_PLAYER Then
        Buffer.WriteLong TempPlayer(Index).StunDuration
    Else
        Buffer.WriteLong MapNpc(MapNum).NPC(Index).StunDuration
    End If

    SendDataToMap MapNum, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapKey(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte Value
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte Value
    SendDataToMap MapNum, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal Index As Long, ByVal ShopNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong ShopNum
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal Index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong movement

    If Not sendToSelf Then
        SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    End If

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendTrade(ByVal Index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal Index As Long, ByVal dataType As Byte)
    Dim Buffer As clsBuffer
    Dim I As Long
    Dim tradeTarget As Long
    Dim totalWorth As Long, multiplier As Long

    tradeTarget = TempPlayer(Index).InTrade
    
    Call SendGoldUpdate(Index)

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType

    If dataType = 0 Then    ' own inventory
        For I = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(Index).TradeOffer(I).Num
            Buffer.WriteLong TempPlayer(Index).TradeOffer(I).Value
        Next
        totalWorth = TempPlayer(Index).TradeGold
    ElseIf dataType = 1 Then    ' other inventory
        For I = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(I).Value
        Next
        totalWorth = TempPlayer(tradeTarget).TradeGold
    End If

    ' send total worth of trade
    Buffer.WriteLong totalWorth

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendTarget(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    Buffer.WriteLong TempPlayer(Index).target
    Buffer.WriteLong TempPlayer(Index).TargetType
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal Index As Long)
    Dim I As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For I = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(Index).Hotbar(I).Slot
        Buffer.WriteByte Player(Index).Hotbar(I).sType
    Next
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendLoginOk(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong Index
    Buffer.WriteLong Player_HighIndex
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendInGame(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendHighIndex()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighIndex
    Buffer.WriteLong Player_HighIndex
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal Index As Long, ByVal targetPlayer As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(targetPlayer).Name)
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partynum As Long)
    Dim Buffer As clsBuffer, I As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partynum).Leader
    For I = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partynum).Member(I)
    Next
    Buffer.WriteLong Party(partynum).MemberCount
    SendDataToParty partynum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal Index As Long)
    Dim Buffer As clsBuffer, I As Long, partynum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate

    ' check if we're in a party
    partynum = TempPlayer(Index).inParty
    If partynum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partynum).Leader
        For I = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partynum).Member(I)
        Next
        Buffer.WriteLong Party(partynum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partynum As Long, ByVal Index As Long)
    Dim Buffer As clsBuffer, I As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong Index
    Buffer.WriteInteger GetPlayerLevel(Index)
    For I = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(Index, I)
        Buffer.WriteLong Player(Index).Vital(I)
    Next
    SendDataToParty partynum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal MapNum As Long, ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong Index
    Buffer.WriteString MapItem(MapNum, Index).PlayerName
    Buffer.WriteLong MapItem(MapNum, Index).Num
    Buffer.WriteLong MapItem(MapNum, Index).Value
    Buffer.WriteLong MapItem(MapNum, Index).X
    Buffer.WriteLong MapItem(MapNum, Index).Y
    Buffer.WriteByte MapItem(MapNum, Index).Bound

    SendDataToMap MapNum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendChatUpdate(ByVal Index As Long, ByVal NpcNum As Long, ByVal mT As String, ByVal o1 As String, ByVal o2 As String, ByVal o3 As String, ByVal o4 As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatUpdate
    Buffer.WriteLong NpcNum
    Buffer.WriteString mT
    Buffer.WriteString o1
    Buffer.WriteString o2
    Buffer.WriteString o3
    Buffer.WriteString o4
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendStartTutorial(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SStartTutorial
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendNpcDeath(ByVal MapNum As Long, ByVal mapnpcnum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDead
    Buffer.WriteLong mapnpcnum
    SendDataToMap MapNum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendChatBubble(ByVal MapNum As Long, ByVal target As Long, ByVal TargetType As Long, ByVal message As String, ByVal colour As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatBubble
    Buffer.WriteLong target
    Buffer.WriteLong TargetType
    Buffer.WriteString message
    Buffer.WriteLong colour
    SendDataToMap MapNum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendAttack(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAttack
    Buffer.WriteLong Index
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Function SanitiseString(ByVal theString As String) As String
    Dim I As Long, tmpString As String
    tmpString = vbNullString
    If Len(theString) <= 0 Then Exit Function
    For I = 1 To Len(theString)
        Select Case Mid$(theString, I, 1)
        Case "*"
            tmpString = tmpString + "[s]"
        Case ":"
            tmpString = tmpString + "[c]"
        Case Else
            tmpString = tmpString + Mid$(theString, I, 1)
        End Select
    Next
    SanitiseString = tmpString
End Function

Sub SendCancelAnimation(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCancelAnimation
    Buffer.WriteLong Index
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPlayerVariables(ByVal Index As Long)
    Dim Buffer As clsBuffer, I As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerVariables
    For I = 1 To MAX_BYTE
        Buffer.WriteLong Player(Index).Variable(I)
    Next
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendCheckForMap(Index As Long, MapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong MapNum
    Buffer.WriteLong MapCRC32(MapNum).MapDataCRC
    Buffer.WriteLong MapCRC32(MapNum).MapTileCRC

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendEvent(Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SEvent
    If TempPlayer(Index).inEvent Then
        Buffer.WriteLong 1
    Else
        Buffer.WriteLong 0
    End If
    Buffer.WriteLong TempPlayer(Index).eventNum
    Buffer.WriteLong TempPlayer(Index).pageNum
    Buffer.WriteLong TempPlayer(Index).commandNum

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendClientTime()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SClientTime
    Buffer.WriteByte GameSecondsPerSecond
    Buffer.WriteByte GameMinutesPerMinute
    Buffer.WriteByte GameSeconds
    Buffer.WriteByte GameMinutes
    Buffer.WriteByte GameHours

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

End Sub
Sub SendClientTimeTo(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SClientTime
    Buffer.WriteByte GameSecondsPerSecond
    Buffer.WriteByte GameMinutesPerMinute
    Buffer.WriteByte GameSeconds
    Buffer.WriteByte GameMinutes
    Buffer.WriteByte GameHours

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendMessageTo(ByVal Index As Long, ByVal WindowName As String, ByVal message As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMessage
    Buffer.WriteString WindowName
    Buffer.WriteString message

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMessageToAll(ByVal WindowName As String, ByVal message As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMessage
    Buffer.WriteString WindowName
    Buffer.WriteString message

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendGoldUpdate(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SGoldUpdate
    Buffer.WriteLong GetPlayerGold(Index)
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub
