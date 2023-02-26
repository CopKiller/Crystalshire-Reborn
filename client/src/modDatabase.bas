Attribute VB_Name = "modDatabase"
Option Explicit
' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)

    If LCase$(dir$(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

Public Function FileExist(ByVal filename As String) As Boolean

    If LenB(dir$(filename)) > 0 Then
        FileExist = True
    End If

End Function

' gets a string from a text file
Public Function GetVar(File As String, header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(header, Var, Value, File)
End Sub

Public Sub SaveOptions()
    Dim filename As String, i As Long

    filename = App.path & "\Data Files\config_v2.ini"

    Call PutVar(filename, "Options", "Username", Options.Username)
    Call PutVar(filename, "Options", "Password", Options.Password)
    Call PutVar(filename, "Options", "Music", Str$(Options.Music))
    Call PutVar(filename, "Options", "Sound", Str$(Options.sound))
    Call PutVar(filename, "Options", "NoAuto", Str$(Options.NoAuto))
    Call PutVar(filename, "Options", "Render", Str$(Options.Render))
    Call PutVar(filename, "Options", "SaveUser", Str$(Options.SaveUser))
    Call PutVar(filename, "Options", "SavePass", Str$(Options.SavePass))
    Call PutVar(filename, "Options", "Resolution", Str$(Options.Resolution))
    Call PutVar(filename, "Options", "Fullscreen", Str$(Options.Fullscreen))
    Call PutVar(filename, "Options", "Reconnect", Str$(Options.Reconnect))
    Call PutVar(filename, "Options", "PlayIntro", Str$(Options.PlayIntro))
    For i = 0 To ChatChannel.Channel_Count - 1
        Call PutVar(filename, "Options", "Channel" & i, Str$(Options.channelState(i)))
    Next

    'Change Controls
    Call PutVar(filename, "Controles", "Correr", Str$(Options.Correr))
    Call PutVar(filename, "Controles", "Atacar", Str$(Options.Atacar))
    Call PutVar(filename, "Controles", "PegarItem", Str$(Options.PegarItem))
    Call PutVar(filename, "Controles", "Chat", Str$(Options.Chat))
    For i = 1 To MAX_HOTBAR
        Call PutVar(filename, "Controles", "Hotbar" & i, Str$(Options.Hotbar(i)))
    Next i
    Call PutVar(filename, "Controles", "Bolsa_Window", Str$(Options.Bolsa))
    Call PutVar(filename, "Controles", "Magias_Window", Str$(Options.Magias))
    Call PutVar(filename, "Controles", "Personagem_Window", Str$(Options.Personagem))
    Call PutVar(filename, "Controles", "Options_Window", Str$(Options.Options))
    Call PutVar(filename, "Controles", "Guild_Window", Str$(Options.Guild))
    Call PutVar(filename, "Controles", "Quest_Window", Str$(Options.Quests))
    ' moves
    Call PutVar(filename, "Controles", "Up", Str$(Options.Up))
    Call PutVar(filename, "Controles", "Down", Str$(Options.Down))
    Call PutVar(filename, "Controles", "Left", Str$(Options.Left))
    Call PutVar(filename, "Controles", "Right", Str$(Options.Right))
    Call PutVar(filename, "Controles", "UsarSetas", Str$(Options.UsarSetas))
    Call PutVar(filename, "Controles", "Target", Str$(Options.Target))
    Call PutVar(filename, "Controles", "ItemName", Str$(Options.ItemName))
    Call PutVar(filename, "Controles", "ItemAnimation", Str$(Options.ItemAnimation))
    Call PutVar(filename, "Controles", "FPSConection", Str$(Options.FPSConection))
End Sub

Public Sub LoadOptions()
    Dim filename As String, i As Long

    On Error GoTo ErrorHandler

    filename = App.path & "\Data Files\config_v2.ini"

    If Not FileExist(filename) Then
        GoTo ErrorHandler
    Else
        Options.Username = GetVar(filename, "Options", "Username")
        Options.Password = GetVar(filename, "Options", "Password")
        Options.Music = GetVar(filename, "Options", "Music")
        Options.sound = Val(GetVar(filename, "Options", "Sound"))
        Options.NoAuto = Val(GetVar(filename, "Options", "NoAuto"))
        Options.Render = Val(GetVar(filename, "Options", "Render"))
        Options.SaveUser = Val(GetVar(filename, "Options", "SaveUser"))
        Options.SavePass = Val(GetVar(filename, "Options", "SavePass"))
        Options.Resolution = Val(GetVar(filename, "Options", "Resolution"))
        Options.Fullscreen = Val(GetVar(filename, "Options", "Fullscreen"))
        Options.Reconnect = Val(GetVar(filename, "Options", "Reconnect"))
        Options.PlayIntro = Val(GetVar(filename, "Options", "PlayIntro"))
        For i = 0 To ChatChannel.Channel_Count - 1
            Options.channelState(i) = Val(GetVar(filename, "Options", "Channel" & i))
        Next

        'Change Controls
        Options.Correr = Val(GetVar(filename, "Controles", "Correr"))
        Options.Atacar = Val(GetVar(filename, "Controles", "Atacar"))
        Options.PegarItem = Val(GetVar(filename, "Controles", "PegarItem"))
        Options.Chat = Val(GetVar(filename, "Controles", "Chat"))
        For i = 1 To MAX_HOTBAR
            Options.Hotbar(i) = Val(GetVar(filename, "Controles", "Hotbar" & i))
        Next i
        Options.Bolsa = Val(GetVar(filename, "Controles", "Bolsa_Window"))
        Options.Magias = Val(GetVar(filename, "Controles", "Magias_Window"))
        Options.Personagem = Val(GetVar(filename, "Controles", "Personagem_Window"))
        Options.Options = Val(GetVar(filename, "Controles", "Options_Window"))
        Options.Guild = Val(GetVar(filename, "Controles", "Guild_Window"))
        Options.Quests = Val(GetVar(filename, "Controles", "Quest_Window"))

        Options.Up = Val(GetVar(filename, "Controles", "Up"))
        Options.Down = Val(GetVar(filename, "Controles", "Down"))
        Options.Left = Val(GetVar(filename, "Controles", "Left"))
        Options.Right = Val(GetVar(filename, "Controles", "Right"))
        Options.UsarSetas = Val(GetVar(filename, "Controles", "UsarSetas"))
        Options.Target = Val(GetVar(filename, "Controles", "Target"))
        Options.ItemName = Val(GetVar(filename, "Controles", "ItemName"))
        Options.ItemAnimation = Val(GetVar(filename, "Controles", "ItemAnimation"))
        Options.FPSConection = Val(GetVar(filename, "Controles", "FPSConection"))
    End If

    Exit Sub
    
ErrorHandler:
    Options.Music = YES
    Options.sound = YES
    Options.NoAuto = NO
    Options.Username = vbNullString
    Options.Password = vbNullString
    Options.Fullscreen = NO
    Options.Render = NO
    Options.SaveUser = NO
    Options.SavePass = NO
    Options.Reconnect = NO
    Options.PlayIntro = YES
    Options.ItemName = YES
    Options.ItemAnimation = YES
    Options.FPSConection = YES
    For i = 0 To ChatChannel.Channel_Count - 1
        Options.channelState(i) = 1
    Next
    ChangeControls_Restore
    SaveOptions
    
    Exit Sub
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
    Dim filename As String, f As Long, X As Long, Y As Long, i As Long

    ' save map data
    filename = App.path & MAP_PATH & MapNum & "_.dat"

    ' if it exists then kill the ini
    If FileExist(filename) Then Kill filename

    ' General
    With Map.MapData
        PutVar filename, "General", "Name", .Name
        PutVar filename, "General", "Music", .Music
        PutVar filename, "General", "Moral", Val(.Moral)
        PutVar filename, "General", "Up", Val(.Up)
        PutVar filename, "General", "Down", Val(.Down)
        PutVar filename, "General", "Left", Val(.Left)
        PutVar filename, "General", "Right", Val(.Right)
        PutVar filename, "General", "BootMap", Val(.BootMap)
        PutVar filename, "General", "BootX", Val(.BootX)
        PutVar filename, "General", "BootY", Val(.BootY)
        PutVar filename, "General", "MaxX", Val(.MaxX)
        PutVar filename, "General", "MaxY", Val(.MaxY)
        PutVar filename, "General", "BossNpc", Val(.BossNpc)
        PutVar filename, "General", "Panorama", Val(.Panorama)
        
        PutVar filename, "General", "Weather", Val(.Weather)
        PutVar filename, "General", "WeatherIntensity", Val(.WeatherIntensity)
        PutVar filename, "General", "Fog", Val(.Fog)
        PutVar filename, "General", "FogSpeed", Val(.FogSpeed)
        PutVar filename, "General", "FogOpacity", Val(.FogOpacity)
        PutVar filename, "General", "Red", Val(.Red)
        PutVar filename, "General", "Green", Val(.Green)
        PutVar filename, "General", "Blue", Val(.Blue)
        PutVar filename, "General", "Alpha", Val(.Alpha)
        PutVar filename, "General", "Sun", Val(.Sun)
        PutVar filename, "General", "DayNight", Val(.DayNight)
        For i = 1 To MAX_MAP_NPCS
            PutVar filename, "General", "Npc" & i, Val(.NPC(i))
        Next
    End With

    ' Events
    PutVar filename, "Events", "EventCount", Val(Map.TileData.EventCount)

    If Map.TileData.EventCount > 0 Then
        For i = 1 To Map.TileData.EventCount
            With Map.TileData.Events(i)
                PutVar filename, "Event" & i, "Name", .Name
                PutVar filename, "Event" & i, "x", Val(.X)
                PutVar filename, "Event" & i, "y", Val(.Y)
                PutVar filename, "Event" & i, "PageCount", Val(.pageCount)
            End With
            If Map.TileData.Events(i).pageCount > 0 Then
                For X = 1 To Map.TileData.Events(i).pageCount
                    With Map.TileData.Events(i).EventPage(X)
                        PutVar filename, "Event" & i & "Page" & X, "chkPlayerVar", Val(.chkPlayerVar)
                        PutVar filename, "Event" & i & "Page" & X, "chkSelfSwitch", Val(.chkSelfSwitch)
                        PutVar filename, "Event" & i & "Page" & X, "chkHasItem", Val(.chkHasItem)
                        PutVar filename, "Event" & i & "Page" & X, "PlayerVarNum", Val(.PlayerVarNum)
                        PutVar filename, "Event" & i & "Page" & X, "SelfSwitchNum", Val(.SelfSwitchNum)
                        PutVar filename, "Event" & i & "Page" & X, "HasItemNum", Val(.HasItemNum)
                        PutVar filename, "Event" & i & "Page" & X, "PlayerVariable", Val(.PlayerVariable)
                        PutVar filename, "Event" & i & "Page" & X, "GraphicType", Val(.GraphicType)
                        PutVar filename, "Event" & i & "Page" & X, "Graphic", Val(.Graphic)
                        PutVar filename, "Event" & i & "Page" & X, "GraphicX", Val(.GraphicX)
                        PutVar filename, "Event" & i & "Page" & X, "GraphicY", Val(.GraphicY)
                        PutVar filename, "Event" & i & "Page" & X, "MoveType", Val(.MoveType)
                        PutVar filename, "Event" & i & "Page" & X, "MoveSpeed", Val(.MoveSpeed)
                        PutVar filename, "Event" & i & "Page" & X, "MoveFreq", Val(.MoveFreq)
                        PutVar filename, "Event" & i & "Page" & X, "WalkAnim", Val(.WalkAnim)
                        PutVar filename, "Event" & i & "Page" & X, "StepAnim", Val(.StepAnim)
                        PutVar filename, "Event" & i & "Page" & X, "DirFix", Val(.DirFix)
                        PutVar filename, "Event" & i & "Page" & X, "WalkThrough", Val(.WalkThrough)
                        PutVar filename, "Event" & i & "Page" & X, "Priority", Val(.Priority)
                        PutVar filename, "Event" & i & "Page" & X, "Trigger", Val(.Trigger)
                        PutVar filename, "Event" & i & "Page" & X, "CommandCount", Val(.CommandCount)
                    End With
                    If Map.TileData.Events(i).EventPage(X).CommandCount > 0 Then
                        For Y = 1 To Map.TileData.Events(i).EventPage(X).CommandCount
                            With Map.TileData.Events(i).EventPage(X).Commands(Y)
                                PutVar filename, "Event" & i & "Page" & X & "Command" & Y, "Type", Val(.Type)
                                PutVar filename, "Event" & i & "Page" & X & "Command" & Y, "Text", .Text
                                PutVar filename, "Event" & i & "Page" & X & "Command" & Y, "Colour", Val(.Colour)
                                PutVar filename, "Event" & i & "Page" & X & "Command" & Y, "Channel", Val(.channel)
                                PutVar filename, "Event" & i & "Page" & X & "Command" & Y, "TargetType", Val(.TargetType)
                                PutVar filename, "Event" & i & "Page" & X & "Command" & Y, "Target", Val(.Target)
                                PutVar filename, "Event" & i & "Page" & X & "Command" & Y, "x", Val(.X)
                                PutVar filename, "Event" & i & "Page" & X & "Command" & Y, "y", Val(.Y)
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If

    ' dump tile data
    filename = App.path & MAP_PATH & MapNum & ".dat"

    ' if it exists then kill the ini
    If FileExist(filename) Then Kill filename

    f = FreeFile
    With Map
        Open filename For Binary As #f
        For X = 0 To .MapData.MaxX
            For Y = 0 To .MapData.MaxY
                Put #f, , .TileData.Tile(X, Y).Type
                Put #f, , .TileData.Tile(X, Y).Data1
                Put #f, , .TileData.Tile(X, Y).Data2
                Put #f, , .TileData.Tile(X, Y).Data3
                Put #f, , .TileData.Tile(X, Y).Data4
                Put #f, , .TileData.Tile(X, Y).Data5
                Put #f, , .TileData.Tile(X, Y).Autotile
                Put #f, , .TileData.Tile(X, Y).DirBlock
                For i = 1 To MapLayer.Layer_Count - 1
                    Put #f, , .TileData.Tile(X, Y).Layer(i).tileSet
                    Put #f, , .TileData.Tile(X, Y).Layer(i).X
                    Put #f, , .TileData.Tile(X, Y).Layer(i).Y
                Next
            Next
        Next
        Close #f
    End With

    Close #f
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
    Dim filename As String, i As Long, f As Long, X As Long, Y As Long

    ' load map data
    filename = App.path & MAP_PATH & MapNum & "_.dat"

    ' General
    With Map.MapData
        .Name = GetVar(filename, "General", "Name")
        .Music = GetVar(filename, "General", "Music")
        .Moral = Val(GetVar(filename, "General", "Moral"))
        .Up = Val(GetVar(filename, "General", "Up"))
        .Down = Val(GetVar(filename, "General", "Down"))
        .Left = Val(GetVar(filename, "General", "Left"))
        .Right = Val(GetVar(filename, "General", "Right"))
        .BootMap = Val(GetVar(filename, "General", "BootMap"))
        .BootX = Val(GetVar(filename, "General", "BootX"))
        .BootY = Val(GetVar(filename, "General", "BootY"))
        .MaxX = Val(GetVar(filename, "General", "MaxX"))
        .MaxY = Val(GetVar(filename, "General", "MaxY"))
        .BossNpc = Val(GetVar(filename, "General", "BossNpc"))
        .Panorama = Val(GetVar(filename, "General", "Panorama"))
        
        .Weather = Val(GetVar(filename, "General", "Weather"))
        .WeatherIntensity = Val(GetVar(filename, "General", "WeatherIntensity"))
        .Fog = Val(GetVar(filename, "General", "Fog"))
        .FogSpeed = Val(GetVar(filename, "General", "FogSpeed"))
        .FogOpacity = Val(GetVar(filename, "General", "FogOpacity"))
        .Red = Val(GetVar(filename, "General", "Red"))
        .Green = Val(GetVar(filename, "General", "Green"))
        .Blue = Val(GetVar(filename, "General", "Blue"))
        .Alpha = Val(GetVar(filename, "General", "Alpha"))
        .Sun = Val(GetVar(filename, "General", "Sun"))
        .DayNight = Val(GetVar(filename, "General", "DayNight"))
        For i = 1 To MAX_MAP_NPCS
            .NPC(i) = Val(GetVar(filename, "General", "Npc" & i))
        Next
    End With

    ' Events
    Map.TileData.EventCount = Val(GetVar(filename, "Events", "EventCount"))

    If Map.TileData.EventCount > 0 Then
        ReDim Preserve Map.TileData.Events(1 To Map.TileData.EventCount)
        For i = 1 To Map.TileData.EventCount
            With Map.TileData.Events(i)
                .Name = GetVar(filename, "Event" & i, "Name")
                .X = Val(GetVar(filename, "Event" & i, "x"))
                .Y = Val(GetVar(filename, "Event" & i, "y"))
                .pageCount = Val(GetVar(filename, "Event" & i, "PageCount"))
            End With
            If Map.TileData.Events(i).pageCount > 0 Then
                ReDim Preserve Map.TileData.Events(i).EventPage(1 To Map.TileData.Events(i).pageCount)
                For X = 1 To Map.TileData.Events(i).pageCount
                    With Map.TileData.Events(i).EventPage(X)
                        .chkPlayerVar = Val(GetVar(filename, "Event" & i & "Page" & X, "chkPlayerVar"))
                        .chkSelfSwitch = Val(GetVar(filename, "Event" & i & "Page" & X, "chkSelfSwitch"))
                        .chkHasItem = Val(GetVar(filename, "Event" & i & "Page" & X, "chkHasItem"))
                        .PlayerVarNum = Val(GetVar(filename, "Event" & i & "Page" & X, "PlayerVarNum"))
                        .SelfSwitchNum = Val(GetVar(filename, "Event" & i & "Page" & X, "SelfSwitchNum"))
                        .HasItemNum = Val(GetVar(filename, "Event" & i & "Page" & X, "HasItemNum"))
                        .PlayerVariable = Val(GetVar(filename, "Event" & i & "Page" & X, "PlayerVariable"))
                        .GraphicType = Val(GetVar(filename, "Event" & i & "Page" & X, "GraphicType"))
                        .Graphic = Val(GetVar(filename, "Event" & i & "Page" & X, "Graphic"))
                        .GraphicX = Val(GetVar(filename, "Event" & i & "Page" & X, "GraphicX"))
                        .GraphicY = Val(GetVar(filename, "Event" & i & "Page" & X, "GraphicY"))
                        .MoveType = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveType"))
                        .MoveSpeed = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveSpeed"))
                        .MoveFreq = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveFreq"))
                        .WalkAnim = Val(GetVar(filename, "Event" & i & "Page" & X, "WalkAnim"))
                        .StepAnim = Val(GetVar(filename, "Event" & i & "Page" & X, "StepAnim"))
                        .DirFix = Val(GetVar(filename, "Event" & i & "Page" & X, "DirFix"))
                        .WalkThrough = Val(GetVar(filename, "Event" & i & "Page" & X, "WalkThrough"))
                        .Priority = Val(GetVar(filename, "Event" & i & "Page" & X, "Priority"))
                        .Trigger = Val(GetVar(filename, "Event" & i & "Page" & X, "Trigger"))
                        .CommandCount = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandCount"))
                    End With
                    If Map.TileData.Events(i).EventPage(X).CommandCount > 0 Then
                        ReDim Preserve Map.TileData.Events(i).EventPage(X).Commands(1 To Map.TileData.Events(i).EventPage(X).CommandCount)
                        For Y = 1 To Map.TileData.Events(i).EventPage(X).CommandCount
                            With Map.TileData.Events(i).EventPage(X).Commands(Y)
                                .Type = GetVar(filename, "Event" & i & "Page" & X & "Command" & Y, "Type")
                                .Text = GetVar(filename, "Event" & i & "Page" & X & "Command" & Y, "Text")
                                .Colour = Val(GetVar(filename, "Event" & i & "Page" & X & "Command" & Y, "Colour"))
                                .channel = Val(GetVar(filename, "Event" & i & "Page" & X & "Command" & Y, "Channel"))
                                .TargetType = Val(GetVar(filename, "Event" & i & "Page" & X & "Command" & Y, "TargetType"))
                                .Target = Val(GetVar(filename, "Event" & i & "Page" & X & "Command" & Y, "Target"))
                                .X = Val(GetVar(filename, "Event" & i & "Page" & X & "Command" & Y, "x"))
                                .Y = Val(GetVar(filename, "Event" & i & "Page" & X & "Command" & Y, "y"))
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If

    ' dump tile data
    filename = App.path & MAP_PATH & MapNum & ".dat"
    f = FreeFile

    ReDim Map.TileData.Tile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY) As TileRec

    With Map
        Open filename For Binary As #f
        For X = 0 To .MapData.MaxX
            For Y = 0 To .MapData.MaxY
                Get #f, , .TileData.Tile(X, Y).Type
                Get #f, , .TileData.Tile(X, Y).Data1
                Get #f, , .TileData.Tile(X, Y).Data2
                Get #f, , .TileData.Tile(X, Y).Data3
                Get #f, , .TileData.Tile(X, Y).Data4
                Get #f, , .TileData.Tile(X, Y).Data5
                Get #f, , .TileData.Tile(X, Y).Autotile
                Get #f, , .TileData.Tile(X, Y).DirBlock
                For i = 1 To MapLayer.Layer_Count - 1
                    Get #f, , .TileData.Tile(X, Y).Layer(i).tileSet
                    Get #f, , .TileData.Tile(X, Y).Layer(i).X
                    Get #f, , .TileData.Tile(X, Y).Layer(i).Y
                Next
            Next
        Next
        Close #f
    End With

    ClearTempTile
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = vbNullString
End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).sound = "None."
    
    ' Clear CRC32
    Call ZeroMemory(ByVal VarPtr(ItemCRC32(Index)), LenB(ItemCRC32(Index)))
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
End Sub

Sub ClearAnimation(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

End Sub

Sub ClearNPC(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).sound = "None."
    
    ' Clear CRC32
    Call ZeroMemory(ByVal VarPtr(NpcCRC32(Index)), LenB(NpcCRC32(Index)))
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

End Sub

Sub ClearMapItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
End Sub

Sub ClearMap()
    Erase Map.TileData.Tile
    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.MapData.Name = vbNullString
    Map.MapData.MaxX = MAX_MAPX
    Map.MapData.MaxY = MAX_MAPY
    ReDim Map.TileData.Tile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)
    initAutotiles
End Sub

Sub ClearMapItems()
    Dim i As Long

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
End Sub

Sub ClearMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String

    If Index > Player_HighIndex Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerExp = Player(Index).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).EXP = EXP
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerMaxVital = Player(Index).MaxVital(Vital)
End Function

Sub SetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)

    If Index > Player_HighIndex Then Exit Sub
    
    Player(Index).MaxVital(Vital) = Value

End Sub

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)

    If Index > Player_HighIndex Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_BYTE Then Value = MAX_BYTE
    Player(Index).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long

    If Index > Player_HighIndex Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerDir = Player(Index).dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal dir As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).dir = dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    If InvSlot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal itemNum As Long)

    If Index > Player_HighIndex Then Exit Sub
    PlayerInv(InvSlot).num = itemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerInvItemValue = PlayerInv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)

    If Index > Player_HighIndex Then Exit Sub
    PlayerInv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)

    If Index < 1 Or Index > Player_HighIndex Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = invNum
End Sub

Sub ClearConv(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Conv(Index)), LenB(Conv(Index)))
    Conv(Index).Name = vbNullString
    ReDim Conv(Index).Conv(1)
End Sub

Sub ClearConvs()
    Dim i As Long

    For i = 1 To MAX_CONVS
        Call ClearConv(i)
    Next

End Sub

Public Function GetNpcMaxVitals(ByVal MapNpcNum As Byte, ByVal Vital As Vitals) As Long
    Dim NpcNum As Integer

    GetNpcMaxVitals = 0

    NpcNum = MapNpc(MapNpcNum).num

    If NpcNum <= 0 Then Exit Function

    Select Case Vital
    Case HP
        GetNpcMaxVitals = ((NPC(NpcNum).Stat(Endurance) / 2)) * 10
    Case MP
        GetNpcMaxVitals = (NPC(NpcNum).Stat(Intelligence) / 2) * 5 + 35
    End Select

End Function

Public Function GetNpcVitals(ByVal MapNpcNum As Byte, ByVal Vital As Vitals) As Long

    GetNpcVitals = 0

    If MapNpcNum <= 0 Then Exit Function

    Select Case Vital
    Case HP
        GetNpcVitals = MapNpc(MapNpcNum).Vital(HP)
    Case MP
        GetNpcVitals = MapNpc(MapNpcNum).Vital(MP)
    End Select

End Function
