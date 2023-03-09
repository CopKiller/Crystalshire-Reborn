Attribute VB_Name = "modGameEditors"
Option Explicit

' Temp event storage
Public tmpEvent As EventRec
Public tmpItem As ItemRec
Public tmpSpell As SpellRec
Public tmpNPC As NpcRec

Public curPageNum As Long
Public curCommand As Long
Public GraphicSelX As Long
Public GraphicSelY As Long

' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
    Dim i As Long
    ' set the width
    frmEditor_Map.Width = 9585
    ' we're in the map editor
    InMapEditor = True
    ' show the form
    frmEditor_Map.visible = True
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.max = Count_Tileset
    frmEditor_Map.fraTileSet.caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.Value = 1
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.max = (frmEditor_Map.picBackSelect.Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (frmEditor_Map.picBackSelect.Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    shpSelectedWidth = 32
    shpSelectedHeight = 32
    MapEditorTileScroll
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"

    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).Name
    Next

    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0
End Sub

Public Sub MapEditorProperties()
    Dim X As Long, i As Long, tmpNum As Long

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    tmpNum = UBound(musicCache)

    For i = 1 To tmpNum
        frmEditor_MapProperties.lstMusic.AddItem musicCache(i)
    Next

    ' finished populating
    With frmEditor_MapProperties
        .scrlBoss.max = MAX_MAP_NPCS
        .txtName.Text = Trim$(Map.MapData.Name)

        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            tmpNum = .lstMusic.ListCount

            For i = 0 To tmpNum - 1

                If .lstMusic.list(i) = Trim$(Map.MapData.Music) Then
                    .lstMusic.ListIndex = i
                End If

            Next

        End If

        ' rest of it
        .txtUp.Text = CStr(Map.MapData.Up)
        .txtDown.Text = CStr(Map.MapData.Down)
        .txtLeft.Text = CStr(Map.MapData.Left)
        .txtRight.Text = CStr(Map.MapData.Right)
        .cmbMoral.ListIndex = Map.MapData.Moral
        .txtBootMap.Text = CStr(Map.MapData.BootMap)
        .txtBootX.Text = CStr(Map.MapData.BootX)
        .txtBootY.Text = CStr(Map.MapData.BootY)
        .scrlBoss = Map.MapData.BossNpc
        .scrlPanorama = Map.MapData.Panorama
        
        .CmbWeather.ListIndex = Map.MapData.Weather
        .scrlWeatherIntensity.Value = Map.MapData.WeatherIntensity
        
        .ScrlFog.max = Count_Fog
        .ScrlFog.Value = Map.MapData.Fog
        .ScrlFogSpeed.Value = Map.MapData.FogSpeed
        .scrlFogOpacity.Value = Map.MapData.FogOpacity
        
        .scrlR.Value = Map.MapData.Red
        .scrlG.Value = Map.MapData.Green
        .scrlB.Value = Map.MapData.Blue
        .scrlA.Value = Map.MapData.Alpha
        
        .scrlSun.max = Count_Sun
        .scrlSun.Value = Map.MapData.Sun
        
        .cmbDayNight.ListIndex = Map.MapData.DayNight
        
        ' show the map npcs
        .lstNpcs.Clear

        For X = 1 To MAX_MAP_NPCS

            If Map.MapData.NPC(X) > 0 Then
                .lstNpcs.AddItem X & ": " & Trim$(NPC(Map.MapData.NPC(X)).Name)
            Else
                .lstNpcs.AddItem X & ": No NPC"
            End If

        Next

        .lstNpcs.ListIndex = 0
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"

        For X = 1 To MAX_NPCS
            .cmbNpc.AddItem X & ": " & Trim$(NPC(X).Name)
        Next

        ' set the combo box properly
        Dim tmpString() As String
        Dim NpcNum As Long
        tmpString = Split(.lstNpcs.list(.lstNpcs.ListIndex))
        NpcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = Map.MapData.NPC(NpcNum)
        ' show the current map
        .lblMap.caption = "Current map: " & GetPlayerMap(MyIndex)
        .txtMaxX.Text = Map.MapData.MaxX
        .txtMaxY.Text = Map.MapData.MaxY
    End With

End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False, Optional ByVal theAutotile As Byte = 0)
    Dim X2 As Long, Y2 As Long

    If theAutotile > 0 Then

        With Map.TileData.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).tileSet = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurLayer) = theAutotile
            cacheRenderState X, Y, CurLayer
        End With

        ' do a re-init so we can see our changes
        initAutotiles
        Exit Sub
    End If

    If Not multitile Then    ' single

        With Map.TileData.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).tileSet = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurLayer) = 0
            cacheRenderState X, Y, CurLayer
        End With

    Else    ' multitile
        Y2 = 0    ' starting tile for y axis

        For Y = CurY To CurY + EditorTileHeight - 1
            X2 = 0    ' re-set x count every y loop

            For X = CurX To CurX + EditorTileWidth - 1

                If X >= 0 And X <= Map.MapData.MaxX Then
                    If Y >= 0 And Y <= Map.MapData.MaxY Then

                        With Map.TileData.Tile(X, Y)
                            .Layer(CurLayer).X = EditorTileX + X2
                            .Layer(CurLayer).Y = EditorTileY + Y2
                            .Layer(CurLayer).tileSet = frmEditor_Map.scrlTileSet.Value
                            .Autotile(CurLayer) = 0
                            cacheRenderState X, Y, CurLayer
                        End With

                    End If
                End If

                X2 = X2 + 1
            Next

            Y2 = Y2 + 1
        Next

    End If

End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
    Dim i As Long
    Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1

        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If

    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.Value Then

            ' no autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then    'single tile
                MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
            Else    ' multi tile!

                If frmEditor_Map.scrlAutotile.Value = 0 Then
                    MapEditorSetTile CurX, CurY, CurLayer, True
                Else
                    MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
                End If
            End If

        ElseIf frmEditor_Map.optAttribs.Value Then

            With Map.TileData.Tile(CurX, CurY)

                ' blocked tile
                If frmEditor_Map.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED

                ' warp tile
                If frmEditor_Map.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = EditorWarpFall
                    .Data5 = 0
                End If

                ' item spawn
                If frmEditor_Map.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' npc avoid
                If frmEditor_Map.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' key
                If frmEditor_Map.optKey.Value Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = KeyEditorTime
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' key open
                If frmEditor_Map.optKeyOpen.Value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' resource
                If frmEditor_Map.optResource.Value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' door
                If frmEditor_Map.optDoor.Value Then
                    .Type = TILE_TYPE_DOOR
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' npc spawn
                If frmEditor_Map.optNpcSpawn.Value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' shop
                If frmEditor_Map.optShop.Value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' bank
                If frmEditor_Map.optBank.Value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' heal
                If frmEditor_Map.optHeal.Value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' trap
                If frmEditor_Map.optTrap.Value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' slide
                If frmEditor_Map.optSlide.Value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If
                
                ' Light
                If frmEditor_Map.optLight.Value Then
                    .Type = TILE_TYPE_LIGHT
                    .Data1 = MapEditorLightA
                    .Data2 = MapEditorLightR
                    .Data3 = MapEditorLightG
                    .Data4 = MapEditorLightB
                    .Data5 = MapEditorLightSize
                End If

                ' chat
                If frmEditor_Map.optChat.Value Then
                    .Type = TILE_TYPE_CHAT
                    .Data1 = MapEditorChatNpc
                    .Data2 = MapEditorChatDir
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If

                ' appear
                If frmEditor_Map.optAppear.Value Then
                    .Type = TILE_TYPE_APPEAR
                    .Data1 = EditorAppearRange
                    .Data2 = EditorAppearBottom
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
                End If
            End With

        ElseIf frmEditor_Map.optBlock.Value Then

            If movedMouse Then Exit Sub
            ' find what tile it is
            X = X - ((X \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)

            ' see if it hits an arrow
            For i = 1 To 4
                If X >= DirArrowX(i) And X <= DirArrowX(i) + 8 Then
                    If Y >= DirArrowY(i) And Y <= DirArrowY(i) + 8 Then
                        ' flip the value.
                        setDirBlock Map.TileData.Tile(CurX, CurY).DirBlock, CByte(i), Not isDirBlocked(Map.TileData.Tile(CurX, CurY).DirBlock, CByte(i))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.Value Then

            With Map.TileData.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).X = 0
                .Layer(CurLayer).Y = 0
                .Layer(CurLayer).tileSet = 0

                If .Autotile(CurLayer) > 0 Then
                    .Autotile(CurLayer) = 0
                    ' do a re-init so we can see our changes
                    initAutotiles
                End If

                cacheRenderState X, Y, CurLayer
            End With

        ElseIf frmEditor_Map.optAttribs.Value Then

            With Map.TileData.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
                .Data4 = 0
                .Data5 = 0
            End With

        End If
    End If

    CacheResources
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
        shpSelectedTop = EditorTileY * PIC_Y
        shpSelectedLeft = EditorTileX * PIC_X
        shpSelectedWidth = PIC_X
        shpSelectedHeight = PIC_Y
    End If

End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1

        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > frmEditor_Map.picBackSelect.Width / PIC_X Then X = frmEditor_Map.picBackSelect.Width / PIC_X
        If Y < 0 Then Y = 0
        If Y > frmEditor_Map.picBackSelect.Height / PIC_Y Then Y = frmEditor_Map.picBackSelect.Height / PIC_Y

        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then    ' drag right
            EditorTileWidth = X - EditorTileX
        Else    ' drag left
            ' TO DO
        End If

        If Y > EditorTileY Then    ' drag down
            EditorTileHeight = Y - EditorTileY
        Else    ' drag up
            ' TO DO
        End If

        shpSelectedWidth = EditorTileWidth * PIC_X
        shpSelectedHeight = EditorTileHeight * PIC_Y
    End If

End Sub

Public Sub NudgeMap(ByVal theDir As Byte)
    Dim X As Long, Y As Long, i As Long

    ' if left or right
    If theDir = DIR_UP Or theDir = DIR_LEFT Then
        For Y = 0 To Map.MapData.MaxY
            For X = 0 To Map.MapData.MaxX
                Select Case theDir
                Case DIR_UP
                    ' move up all one
                    If Y > 0 Then CopyTile Map.TileData.Tile(X, Y), Map.TileData.Tile(X, Y - 1)
                Case DIR_LEFT
                    ' move left all one
                    If X > 0 Then CopyTile Map.TileData.Tile(X, Y), Map.TileData.Tile(X - 1, Y)
                End Select
            Next
        Next
    Else
        For Y = Map.MapData.MaxY To 0 Step -1
            For X = Map.MapData.MaxX To 0 Step -1
                Select Case theDir
                Case DIR_DOWN
                    ' move down all one
                    If Y < Map.MapData.MaxY Then CopyTile Map.TileData.Tile(X, Y), Map.TileData.Tile(X, Y + 1)
                Case DIR_RIGHT
                    ' move right all one
                    If X < Map.MapData.MaxX Then CopyTile Map.TileData.Tile(X, Y), Map.TileData.Tile(X + 1, Y)
                End Select
            Next
        Next
    End If

    ' do events
    If Map.TileData.EventCount > 0 Then
        For i = 1 To Map.TileData.EventCount
            Select Case theDir
            Case DIR_UP
                Map.TileData.Events(i).Y = Map.TileData.Events(i).Y - 1
            Case DIR_LEFT
                Map.TileData.Events(i).X = Map.TileData.Events(i).X - 1
            Case DIR_RIGHT
                Map.TileData.Events(i).X = Map.TileData.Events(i).X + 1
            Case DIR_DOWN
                Map.TileData.Events(i).Y = Map.TileData.Events(i).Y + 1
            End Select
        Next
    End If

    initAutotiles
End Sub

Public Sub CopyTile(ByRef origTile As TileRec, ByRef newTile As TileRec)
    Dim tilesize As Long
    tilesize = LenB(origTile)
    CopyMemory ByVal VarPtr(newTile), ByVal VarPtr(origTile), tilesize
    ZeroMemory ByVal VarPtr(origTile), tilesize
End Sub

Public Sub MapEditorTileScroll()

' horizontal scrolling
    If frmEditor_Map.picBackSelect.Width < frmEditor_Map.picBack.Width Then
        frmEditor_Map.scrlPictureX.enabled = False
    Else
        frmEditor_Map.scrlPictureX.enabled = True
        frmEditor_Map.picBackSelect.Left = (frmEditor_Map.scrlPictureX.Value * PIC_X) * -1
    End If

    ' vertical scrolling
    If frmEditor_Map.picBackSelect.Height < frmEditor_Map.picBack.Height Then
        frmEditor_Map.scrlPictureY.enabled = False
    Else
        frmEditor_Map.scrlPictureY.enabled = True
        frmEditor_Map.picBackSelect.top = (frmEditor_Map.scrlPictureY.Value * PIC_Y) * -1
    End If

End Sub

Public Sub MapEditorSend()
    Call SendMap
    InMapEditor = False
    'Unload frmEditor_Map
    frmEditor_Map.Hide
End Sub

Public Sub MapEditorCancel()
    InMapEditor = False
    LoadMap GetPlayerMap(MyIndex)
    initAutotiles
    'Unload frmEditor_Map
    frmEditor_Map.Hide
End Sub

Public Sub MapEditorClearLayer()
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1

        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If

    Next

    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo, GAME_NAME) = vbYes Then

        For X = 0 To Map.MapData.MaxX
            For Y = 0 To Map.MapData.MaxY
                Map.TileData.Tile(X, Y).Layer(CurLayer).X = 0
                Map.TileData.Tile(X, Y).Layer(CurLayer).Y = 0
                Map.TileData.Tile(X, Y).Layer(CurLayer).tileSet = 0
                cacheRenderState X, Y, CurLayer
            Next
        Next

        ' re-cache autos
        initAutotiles
    End If

End Sub

Public Sub MapEditorFillLayer()
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1

        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If

    Next

    ' Ground layer
    If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, GAME_NAME) = vbYes Then

        For X = 0 To Map.MapData.MaxX
            For Y = 0 To Map.MapData.MaxY
                Map.TileData.Tile(X, Y).Layer(CurLayer).X = EditorTileX
                Map.TileData.Tile(X, Y).Layer(CurLayer).Y = EditorTileY
                Map.TileData.Tile(X, Y).Layer(CurLayer).tileSet = frmEditor_Map.scrlTileSet.Value
                Map.TileData.Tile(X, Y).Autotile(CurLayer) = frmEditor_Map.scrlAutotile.Value
                cacheRenderState X, Y, CurLayer
            Next
        Next

        ' now cache the positions
        initAutotiles
    End If

End Sub

Public Sub MapEditorClearAttribs()
    Dim X As Long
    Dim Y As Long

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME) = vbYes Then

        For X = 0 To Map.MapData.MaxX
            For Y = 0 To Map.MapData.MaxY
                Map.TileData.Tile(X, Y).Type = 0
            Next
        Next

    End If

End Sub

Public Sub MapEditorLeaveMap()

    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If

End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
    Dim i As Long, SoundSet As Boolean, tmpNum As Long
    
    On Error Resume Next

    If frmEditor_Item.visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    tmpNum = UBound(soundCache)

    For i = 1 To tmpNum
        frmEditor_Item.cmbSound.AddItem soundCache(i)
    Next

    ' finished populating
    With Item(EditorIndex)
        frmEditor_Item.txtName.Text = Trim$(.Name)

        If .Pic > frmEditor_Item.scrlPic.max Then .Pic = 0
        frmEditor_Item.scrlPic.Value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.Value = .Animation
        frmEditor_Item.txtDesc.Text = Trim$(.Desc)
        frmEditor_Item.chkStackable.Value = Item(EditorIndex).Stackable
        frmEditor_Item.scrlGiveSpell.Value = Item(EditorIndex).GiveSpellNum
        frmEditor_Item.chkDropDead.Value = Item(EditorIndex).DropDead
        frmEditor_Item.scrlChance.Value = Item(EditorIndex).DropDeadChance

        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            tmpNum = frmEditor_Item.cmbSound.ListCount

            For i = 0 To tmpNum

                If frmEditor_Item.cmbSound.list(i) = Trim$(.sound) Then
                    frmEditor_Item.cmbSound.ListIndex = i
                    SoundSet = True
                End If

            Next

            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_RINGRIGHT) Then
            frmEditor_Item.fraEquipment.visible = True
            frmEditor_Item.txtDamage.Text = .Data2
            frmEditor_Item.chkPercentDamage.Value = .Data2_Percent
            frmEditor_Item.cmbTool.ListIndex = .Data3

            If .Speed < 100 Then .Speed = 100
            frmEditor_Item.scrlSpeed.Value = .Speed

            ' loop for stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.txtStatBonus(i).Text = .Add_Stat(i)
                frmEditor_Item.chkPercentStats(i).Value = .Stat_Percent(i)
                
                'Base Atribute
                frmEditor_Item.optBase(.AtributeBase).Value = True
            Next

            If Not .Paperdoll > Count_Paperdoll Then frmEditor_Item.scrlPaperdoll = .Paperdoll
            frmEditor_Item.scrlProf.Value = .proficiency
        Else
            frmEditor_Item.fraEquipment.visible = False
        End If

        ' Block Chance
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SHIELD) Then
            frmEditor_Item.scrlBlockChance.Value = .BlockChance
            frmEditor_Item.lblBlockChance.caption = "Block Chance(Shield): " & .BlockChance & " %"
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.visible = True
            frmEditor_Item.scrlAddHp.Value = .AddHP
            frmEditor_Item.scrlAddMP.Value = .AddMP
            frmEditor_Item.scrlAddExp.Value = .AddEXP
            frmEditor_Item.scrlCastSpell.Value = .CastSpell
            frmEditor_Item.chkInstant.Value = .instaCast
        Else
            frmEditor_Item.fraVitals.visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.visible = True
            frmEditor_Item.scrlSpell.Value = .Data1
        Else
            frmEditor_Item.fraSpell.visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_FOOD Then
            If .HPorSP = 2 Then
                frmEditor_Item.optSP.Value = True
            Else
                frmEditor_Item.optHP.Value = True
            End If

            frmEditor_Item.scrlFoodHeal = .FoodPerTick
            frmEditor_Item.scrlFoodTick = .FoodTickCount
            frmEditor_Item.scrlFoodInterval = .FoodInterval
            frmEditor_Item.fraFood.visible = True
        Else
            frmEditor_Item.fraFood.visible = False
        End If

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.Value = .AccessReq
        frmEditor_Item.scrlLevelReq.Value = .LevelReq

        ' loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Item.txtStatReq(i).Text = .Stat_Req(i)
        Next

        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(i).Name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.txtPrice.Text = .price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.Value = .Rarity
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With

    Item_Changed(EditorIndex) = True
End Sub

Public Sub ItemEditorOk()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If

    Next

    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
End Sub

Sub ItemEditorCopy()
    CopyMemory ByVal VarPtr(tmpItem), ByVal VarPtr(Item(EditorIndex)), LenB(Item(EditorIndex))
End Sub

Sub ItemEditorPaste()
    CopyMemory ByVal VarPtr(Item(EditorIndex)), ByVal VarPtr(tmpItem), LenB(tmpItem)
    ItemEditorInit
    frmEditor_Item.txtName_Validate False
End Sub

Public Sub ItemEditorCancel()
    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
End Sub

Public Sub ClearChanged_Item()
    ZeroMemory Item_Changed(1), MAX_ITEMS * 2    ' 2 = boolean length
End Sub

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
    Dim i As Long
    Dim SoundSet As Boolean, tmpNum As Long

    If frmEditor_Animation.visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    tmpNum = UBound(soundCache)

    For i = 1 To tmpNum
        frmEditor_Animation.cmbSound.AddItem soundCache(i)
    Next

    ' finished populating
    With Animation(EditorIndex)
        frmEditor_Animation.txtName.Text = Trim$(.Name)

        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            tmpNum = frmEditor_Animation.cmbSound.ListCount

            For i = 0 To tmpNum

                If frmEditor_Animation.cmbSound.list(i) = Trim$(.sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = i
                    SoundSet = True
                End If

            Next

            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If

        For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).Value = .Sprite(i)
            frmEditor_Animation.scrlFrameCount(i).Value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).Value = .LoopCount(i)

            If .looptime(i) > 0 Then
                frmEditor_Animation.scrlLoopTime(i).Value = .looptime(i)
            Else
                frmEditor_Animation.scrlLoopTime(i).Value = 45
            End If

        Next

        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With

    Animation_Changed(EditorIndex) = True
End Sub

Public Sub AnimationEditorOk()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If

    Next

    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
End Sub

Public Sub AnimationEditorCancel()
    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
End Sub

Public Sub ClearChanged_Animation()
    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2    ' 2 = boolean length
End Sub

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
    Dim i As Long
    Dim SoundSet As Boolean

    If frmEditor_NPC.visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."

    For i = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(i)
    Next

    ' finished populating
    With frmEditor_NPC
        .scrlSpell.max = MAX_NPC_SPELLS
        .txtName.Text = Trim$(NPC(EditorIndex).Name)
        .txtAttackSay.Text = Trim$(NPC(EditorIndex).AttackSay)

        If NPC(EditorIndex).Sprite < 0 Or NPC(EditorIndex).Sprite > .scrlSprite.max Then NPC(EditorIndex).Sprite = 0
        .scrlSprite.Value = NPC(EditorIndex).Sprite
        .cmbBehaviour.ListIndex = NPC(EditorIndex).Behaviour
        .scrlRange.Value = NPC(EditorIndex).Range
        .txtExp.Text = NPC(EditorIndex).EXP
        .txtLevel.Text = NPC(EditorIndex).Level
        .scrlConv.Value = NPC(EditorIndex).Conv
        .scrlAnimation.Value = NPC(EditorIndex).Animation
        .chkShadow.Value = NPC(EditorIndex).Shadow
        .scrlBalao = NPC(EditorIndex).Balao

        ' spawn variavel
        .chkRndSpawn.Value = NPC(EditorIndex).RndSpawn
        .txtSpawnSecs.Text = CStr(NPC(EditorIndex).SpawnSecs)
        .txtSpawnSecsMin.Text = CStr(NPC(EditorIndex).SpawnSecsMin)

        ' exp variavel
        .chkRndExp.Value = NPC(EditorIndex).RandExp
        .opPercent_5.Value = CBool(NPC(EditorIndex).Percent_5)
        .opPercent_10.Value = CBool(NPC(EditorIndex).Percent_10)
        .opPercent_20.Value = CBool(NPC(EditorIndex).Percent_20)
        
        ' block chance
        .scrlBlockChance.Value = NPC(EditorIndex).BlockChance

        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then

            For i = 0 To .cmbSound.ListCount

                If .cmbSound.list(i) = Trim$(NPC(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If

            Next

            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If

        For i = 1 To Stats.Stat_Count - 1
            .txtStat(i).Text = NPC(EditorIndex).Stat(i)
        Next

        ' Drop Items
        .cmbItems.Clear
        .cmbItems.AddItem "No Items"
        .cmbItems.ListIndex = 0
        If .cmbItems.ListCount >= 0 Then
            For i = 1 To MAX_ITEMS
                .cmbItems.AddItem (Trim$(Item(i).Name))
            Next
        End If
        ' re-load the list
        .lstItems.Clear
        For i = 1 To MAX_NPC_DROPS
            If NPC(EditorIndex).DropItem(i) > 0 Then
                .lstItems.AddItem i & ": " & NPC(EditorIndex).DropItemValue(i) & "x " & Trim$(Item(NPC(EditorIndex).DropItem(i)).Name) & " : 1 em " & NPC(EditorIndex).DropChance(i)
            Else
                .lstItems.AddItem i & ": No Items"
            End If
        Next
        .lstItems.ListIndex = 0

        ' show 1 data
        .scrlSpell.Value = 1
    End With

    NPC_Changed(EditorIndex) = True
End Sub

Public Sub NpcEditorOk()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If

    Next

    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
End Sub

Sub NpcEditorCopy()
    CopyMemory ByVal VarPtr(tmpNPC), ByVal VarPtr(NPC(EditorIndex)), LenB(NPC(EditorIndex))
End Sub

Sub NpcEditorPaste()
    CopyMemory ByVal VarPtr(NPC(EditorIndex)), ByVal VarPtr(tmpNPC), LenB(tmpNPC)
    NpcEditorInit
    frmEditor_NPC.txtName_Validate False
End Sub

Public Sub NpcEditorCancel()
    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
End Sub

Public Sub ClearChanged_NPC()
    ZeroMemory NPC_Changed(1), MAX_NPCS * 2    ' 2 = boolean length
End Sub

' /////////////////
' // Conv Editor //
' /////////////////
Public Sub ConvEditorInit()
    Dim i As Long, n As Long

    If frmEditor_Conv.visible = False Then Exit Sub
    EditorIndex = Val(frmEditor_Conv.lstIndex.list(frmEditor_Conv.lstIndex.ListIndex))
    If EditorIndex <= 0 Then Exit Sub

    With frmEditor_Conv
        ' Indica que está no modo de inicialização.
        ' Não permitindo que o controle event seja alterado.
        .IsIniting = True

        .txtName.Text = Trim$(Conv(EditorIndex).Name)

        If Conv(EditorIndex).chatCount = 0 Then
            .curConv = 1
            Conv(EditorIndex).chatCount = 1
            ReDim Preserve Conv(EditorIndex).Conv(1 To Conv(EditorIndex).chatCount)
        End If

        For n = 1 To 4
            .cmbReply(n).Clear
            .cmbReply(n).AddItem "None"

            For i = 1 To Conv(EditorIndex).chatCount
                .cmbReply(n).AddItem i
            Next
        Next

        .scrlChatCount.Value = Conv(EditorIndex).chatCount
        .scrlConv.max = Conv(EditorIndex).chatCount
        .curConv = 1
        .scrlConv.Value = 1

        .txtConv = Conv(EditorIndex).Conv(.scrlConv.Value).Conv

        For i = 1 To 4
            If Conv(EditorIndex).Conv(.scrlConv.Value).rTarget(i) > Conv(EditorIndex).chatCount Then
                Conv(EditorIndex).Conv(.scrlConv.Value).rTarget(i) = 0
            End If
            .txtReply(i).Text = Conv(EditorIndex).Conv(.scrlConv.Value).rText(i)
            .cmbReply(i).ListIndex = Conv(EditorIndex).Conv(.scrlConv.Value).rTarget(i)
        Next

        .cmbEvent.ListIndex = Conv(EditorIndex).Conv(.scrlConv.Value).Event

        .scrlData1.Value = Conv(EditorIndex).Conv(.scrlConv.Value).Data1
        .scrlData2.Value = Conv(EditorIndex).Conv(.scrlConv.Value).Data2
        .scrlData3.Value = Conv(EditorIndex).Conv(.scrlConv.Value).Data3

        .UpdateText

        ' Indica que está no modo de inicialização.
        ' Não permitindo que o controle event seja alterado.
        .IsIniting = False
    End With

    Conv_Changed(EditorIndex) = True
End Sub

Public Sub ConvEditorOk()
    Dim i As Long

    For i = 1 To MAX_CONVS

        If Conv_Changed(i) Then
            Call SendSaveConv(i)
        End If

    Next

    Unload frmEditor_Conv
    Editor = 0
    ClearChanged_Conv
End Sub

Public Sub ConvEditorCancel()
    Editor = 0
    Unload frmEditor_Conv
    ClearChanged_Conv
    ClearConvs
    SendRequestConvs
End Sub

Public Sub ClearChanged_Conv()
    ZeroMemory Conv_Changed(1), MAX_CONVS * 2    ' 2 = boolean length
End Sub

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
    Dim i As Long
    Dim SoundSet As Boolean

    If frmEditor_Resource.visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."

    For i = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(i)
    Next

    ' finished populating
    With frmEditor_Resource
        .scrlExhaustedPic.max = Count_Resource
        .scrlNormalPic.max = Count_Resource
        .scrlAnimation.max = MAX_ANIMATIONS
        .txtName.Text = Trim$(Resource(EditorIndex).Name)
        .txtMessage.Text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.Text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlTool.Value = Resource(EditorIndex).ToolRequired
        .scrlHealth.Value = Resource(EditorIndex).health
        .scrlRespawn.Value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.Value = Resource(EditorIndex).Animation
        .chkShadow.Value = Resource(EditorIndex).Shadow

        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then

            For i = 0 To .cmbSound.ListCount

                If .cmbSound.list(i) = Trim$(Resource(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If

            Next

            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If

    End With

    Resource_Changed(EditorIndex) = True
End Sub

Public Sub ResourceEditorOk()
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If

    Next

    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
End Sub

Public Sub ResourceEditorCancel()
    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
End Sub

Public Sub ClearChanged_Resource()
    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2    ' 2 = boolean length
End Sub

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
    Dim i As Long

    If frmEditor_Shop.visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    frmEditor_Shop.txtName.Text = Trim$(Shop(EditorIndex).Name)

    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.Value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.Value = 100
    End If

    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(Item(i).Name)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(Item(i).Name)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    UpdateShopTrade
    Shop_Changed(EditorIndex) = True
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
    Dim i As Long
    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES

        With Shop(EditorIndex).TradeItem(i)

            ' if none, show as none
            If .Item = 0 And .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                If .CostItem > 0 Then
                    frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).Name) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).Name)
                Else
                    frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).Name) & " for " & .CostValue & " $"
                End If
            End If

        End With

    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
End Sub

Public Sub ShopEditorOk()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If

    Next

    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
End Sub

Public Sub ShopEditorCancel()
    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
End Sub

Public Sub ClearChanged_Shop()
    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2    ' 2 = boolean length
End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Sub SpellEditorCopy()
    CopyMemory ByVal VarPtr(tmpSpell), ByVal VarPtr(Spell(EditorIndex)), LenB(Spell(EditorIndex))
End Sub

Sub SpellEditorPaste()
    CopyMemory ByVal VarPtr(Spell(EditorIndex)), ByVal VarPtr(tmpSpell), LenB(tmpSpell)
    SpellEditorInit
    frmEditor_Spell.txtName_Validate False
End Sub

Public Sub SpellEditorInit()
    Dim i As Long
    Dim SoundSet As Boolean

    If frmEditor_Spell.visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."

    For i = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(i)
    Next

    ' finished populating
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.max = MAX_ANIMATIONS
        .scrlAnim.max = MAX_ANIMATIONS
        .scrlAOE.max = MAX_BYTE
        .scrlRange.max = MAX_BYTE
        .scrlMap.max = MAX_MAPS
        .scrlNext.max = MAX_SPELLS
        ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"

        For i = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(i).Name)
        Next

        .cmbClass.ListIndex = 0
        ' set values
        .txtName.Text = Trim$(Spell(EditorIndex).Name)
        .txtDesc.Text = Trim$(Spell(EditorIndex).Desc)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.Value = Spell(EditorIndex).MPCost
        .scrlLevel.Value = Spell(EditorIndex).LevelReq
        .scrlAccess.Value = Spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        .scrlCast.Value = Spell(EditorIndex).CastTime
        .scrlCool.Value = Spell(EditorIndex).CDTime
        .scrlIcon.Value = Spell(EditorIndex).Icon
        .scrlMap.Value = Spell(EditorIndex).Map
        .scrlX.Value = Spell(EditorIndex).X
        .scrlY.Value = Spell(EditorIndex).Y
        .scrlDir.Value = Spell(EditorIndex).dir
        .scrlVital.Value = Spell(EditorIndex).Vital
        .scrlDuration.Value = Spell(EditorIndex).Duration
        .scrlInterval.Value = Spell(EditorIndex).Interval
        .scrlRange.Value = Spell(EditorIndex).Range

        If Spell(EditorIndex).IsAoE Then
            .chkAOE.Value = 1
        Else
            .chkAOE.Value = 0
        End If
        
        .chkCanRun.Value = Spell(EditorIndex).CanRun

        .scrlAOE.Value = Spell(EditorIndex).AoE
        .scrlAnimCast.Value = Spell(EditorIndex).CastAnim
        .scrlAnim.Value = Spell(EditorIndex).SpellAnim
        .scrlStun.Value = Spell(EditorIndex).StunDuration
        .scrlNext.Value = Spell(EditorIndex).NextRank
        .scrlIndex.Value = Spell(EditorIndex).UniqueIndex
        .scrlUses.Value = Spell(EditorIndex).NextUses

        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then

            For i = 0 To .cmbSound.ListCount

                If .cmbSound.list(i) = Trim$(Spell(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If

            Next

            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If

    End With

    Spell_Changed(EditorIndex) = True
End Sub

Public Sub SpellEditorOk()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If

    Next

    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
End Sub

Public Sub SpellEditorCancel()
    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
End Sub

Public Sub ClearChanged_Spell()
    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2    ' 2 = boolean length
End Sub

Public Sub ClearAttributeDialogue()
    frmEditor_Map.fraNpcSpawn.visible = False
    frmEditor_Map.fraResource.visible = False
    frmEditor_Map.fraMapItem.visible = False
    frmEditor_Map.fraMapKey.visible = False
    frmEditor_Map.fraKeyOpen.visible = False
    frmEditor_Map.fraMapWarp.visible = False
    frmEditor_Map.fraShop.visible = False
    frmEditor_Map.fraLight.visible = False
End Sub
