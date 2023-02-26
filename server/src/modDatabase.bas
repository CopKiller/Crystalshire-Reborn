Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
    Dim FileName As String
    FileName = App.Path & "\data files\logs\errors.txt"
    Open FileName For Append As #1
    Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
    Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
    Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim FileName As String
    Dim F As Long

    If ServerLog Then
        FileName = App.Path & "\data\logs\" & FN

        If Not FileExist(FileName, True) Then
            F = FreeFile
            Open FileName For Output As #F
            Close #F
        End If

        F = FreeFile
        Open FileName For Append As #F
        Print #F, DateValue(Now) & " " & Time & ": " & Text
        Close #F
    End If

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(dir(App.Path & "\" & FileName)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Trim$(Options.MOTD)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "PARTY BONUS", Str(Options.PartyBonus)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "START_MAP", Str(Options.START_MAP)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "START_X", Str(Options.START_X)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "START_Y", Str(Options.START_Y)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "GAME_NAME", Trim$(Options.GAME_NAME)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "GAME_WEBSITE", Trim$(Options.GAME_WEBSITE)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "DAYNIGHT", Str(Options.DAYNIGHT)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "PREMIUMEXP", Str(Options.PREMIUMEXP)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "PREMIUMDROP", Str(Options.PREMIUMDROP)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "LOTTERYBONUS", Str(Options.LOTTERYBONUS)
End Sub

Public Sub LoadOptions()
    On Error GoTo Conserta
    ' Load Database
    Options.MOTD = Trim$(GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD"))
    Options.PartyBonus = GetVar(App.Path & "\data\options.ini", "OPTIONS", "PARTY BONUS")
    Options.START_MAP = GetVar(App.Path & "\data\options.ini", "OPTIONS", "START_MAP")
    Options.START_X = GetVar(App.Path & "\data\options.ini", "OPTIONS", "START_X")
    Options.START_Y = GetVar(App.Path & "\data\options.ini", "OPTIONS", "START_Y")
    Options.GAME_NAME = Trim$(GetVar(App.Path & "\data\options.ini", "OPTIONS", "GAME_NAME"))
    Options.GAME_WEBSITE = Trim$(GetVar(App.Path & "\data\options.ini", "OPTIONS", "GAME_WEBSITE"))
    Options.DAYNIGHT = GetVar(App.Path & "\data\options.ini", "OPTIONS", "DAYNIGHT")
    Options.PREMIUMEXP = GetVar(App.Path & "\data\options.ini", "OPTIONS", "PREMIUMEXP")
    Options.PREMIUMDROP = GetVar(App.Path & "\data\options.ini", "OPTIONS", "PREMIUMDROP")
    Options.LOTTERYBONUS = GetVar(App.Path & "\data\options.ini", "OPTIONS", "LOTTERYBONUS")

    ' Change Options in Server Window
    With frmServer
        .txtMOTD.Text = Options.MOTD
        .txtMap = Options.START_MAP
        .txtX.Text = Options.START_X
        .txtY.Text = Options.START_Y
        .txtGameName.Text = Options.GAME_NAME
        .txtGameSite.Text = Options.GAME_WEBSITE
    End With

    ' Change Options in Configurations Window
    With frmConfiguration
        .scrlPartyBonus.Value = Options.PartyBonus
        .scrlPremiumExp = Options.PREMIUMEXP
        .scrlPremiumDrop = Options.PREMIUMDROP
        .scrlLottery = Options.LOTTERYBONUS
    End With

    Exit Sub

Conserta:
    SaveOptions
    LoadOptions
    Exit Sub
End Sub

Public Sub ToggleMute(ByVal Index As Long)
' exit out for rte9
    If Index <= 0 Or Index > Player_HighIndex Then Exit Sub

    ' toggle the player's mute
    If Player(Index).isMuted = 1 Then
        Player(Index).isMuted = 0
        ' Let them know
        PlayerMsg Index, "You have been unmuted and can now talk in global.", BrightGreen
        TextAdd GetPlayerName(Index) & " has been unmuted."
    Else
        Player(Index).isMuted = 1
        ' Let them know
        PlayerMsg Index, "You have been muted and can no longer talk in global.", BrightRed
        TextAdd GetPlayerName(Index) & " has been muted."
    End If

    ' save the player
    SavePlayer Index
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long)
    Dim FileName As String, IP As String, I As Long

    ' Print the IP in the ip ban list
    IP = GetPlayerIP(BanPlayerIndex)

    ' Tell them they're banned
    ' Pega o ip do jogador e verifica todos os jogadores com esse ip e da ban em todos os chars!
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If GetPlayerIP(I) = IP Then
                ' Add banned to the player's index
                Player(I).isBanned = 1
                
                ' Add in banlist autenticador
                Call Auth_BanPlayerIP(IP)
                
                Call GlobalMsg(GetPlayerName(I) & " has been banned from " & Options.GAME_NAME & ".", White)
                Call AddLog(GetPlayerName(I) & " has been banned.", ADMIN_LOG)
                Call AlertMsg(I, DIALOGUE_MSG_BANNED)
            End If
        End If
    Next I
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.Path & "\data\_charlist.txt", App.Path & "\data\_chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\_chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\_charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\_chartemp.txt")
End Sub

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            Call SavePlayer(I)
        End If
    Next
End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim FileName As String
    Dim F As Long

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
    
    ' Se não estiver conectado no servidor de autenticação, salva o arquivo no game server mesmo!
    ' pra posteriormente quando reconectar com o autenticador, enviar os dados perdidos!
    If Not IsConnectedAuthServer Then
        ChkDir App.Path & "\data\", "accounts"
        FileName = App.Path & "\data\accounts\" & Trim$(Player(Index).Login) & ".bin"

        F = FreeFile

        Open FileName For Binary As #F
        Put #F, , Player(Index)
        Close #F
    Else
        Auth_SavePlayer Index
    End If
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim I As Long

    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).Buffer = New clsBuffer

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Login = vbNullString
    Player(Index).Name = vbNullString
    Player(Index).Password = vbNullString
    Player(Index).StartPremium = vbNullString
    Player(Index).Class = 1

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim FileName As String
    Dim File As String
    FileName = App.Path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(FileName, True) Then
        File = FreeFile
        Open FileName For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Sub LoadClasses()
    Dim FileName As String
    Dim I As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim X As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        FileName = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For I = 1 To Max_Classes
        Class(I).Name = GetVar(FileName, "CLASS" & I, "Name")
        
        Class(I).START_MAP = GetVar(FileName, "CLASS" & I, "START_MAP")
        Class(I).START_X = GetVar(FileName, "CLASS" & I, "START_X")
        Class(I).START_Y = GetVar(FileName, "CLASS" & I, "START_Y")

        ' read string of sprites
        tmpSprite = GetVar(FileName, "CLASS" & I, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(I).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(I).MaleSprite(n) = Val(tmpArray(n))
        Next

        ' read string of sprites
        tmpSprite = GetVar(FileName, "CLASS" & I, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(I).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(I).FemaleSprite(n) = Val(tmpArray(n))
        Next

        ' continue
        Class(I).Stat(Stats.Strength) = Val(GetVar(FileName, "CLASS" & I, "Strength"))
        Class(I).Stat(Stats.Endurance) = Val(GetVar(FileName, "CLASS" & I, "Endurance"))
        Class(I).Stat(Stats.Intelligence) = Val(GetVar(FileName, "CLASS" & I, "Intelligence"))
        Class(I).Stat(Stats.Agility) = Val(GetVar(FileName, "CLASS" & I, "Agility"))
        Class(I).Stat(Stats.Willpower) = Val(GetVar(FileName, "CLASS" & I, "Willpower"))
        
        ' Get Max Vitals
        Class(I).MaxHP = GetClassMaxVital(I, Vitals.HP)
        Class(I).MaxMP = GetClassMaxVital(I, Vitals.MP)

        ' how many starting items?
        startItemCount = Val(GetVar(FileName, "CLASS" & I, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(I).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(I).StartValue(1 To startItemCount)

        ' loop for items & values
        Class(I).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For X = 1 To startItemCount
                Class(I).StartItem(X) = Val(GetVar(FileName, "CLASS" & I, "StartItem" & X))
                Class(I).StartValue(X) = Val(GetVar(FileName, "CLASS" & I, "StartValue" & X))
            Next
        End If

        ' how many starting spells?
        startSpellCount = Val(GetVar(FileName, "CLASS" & I, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(I).StartSpell(1 To startSpellCount)

        ' loop for spells
        Class(I).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_INV Then
            For X = 1 To startSpellCount
                Class(I).StartSpell(X) = Val(GetVar(FileName, "CLASS" & I, "StartSpell" & X))
            Next
        End If
    Next

End Sub

Sub SaveClasses()
    Dim FileName As String
    Dim I As Long
    Dim X As Long

    FileName = App.Path & "\data\classes.ini"

    For I = 1 To Max_Classes
        Call PutVar(FileName, "CLASS" & I, "Name", Trim$(Class(I).Name))
        Call PutVar(FileName, "CLASS" & I, "Maleprite", "1")
        Call PutVar(FileName, "CLASS" & I, "Femaleprite", "1")
        Call PutVar(FileName, "CLASS" & I, "Strength", Str(Class(I).Stat(Stats.Strength)))
        Call PutVar(FileName, "CLASS" & I, "Endurance", Str(Class(I).Stat(Stats.Endurance)))
        Call PutVar(FileName, "CLASS" & I, "Intelligence", Str(Class(I).Stat(Stats.Intelligence)))
        Call PutVar(FileName, "CLASS" & I, "Agility", Str(Class(I).Stat(Stats.Agility)))
        Call PutVar(FileName, "CLASS" & I, "Willpower", Str(Class(I).Stat(Stats.Willpower)))
        Call PutVar(FileName, "CLASS" & I, "START_MAP", Str(Options.START_MAP))
        Call PutVar(FileName, "CLASS" & I, "START_X", Str(Options.START_X))
        Call PutVar(FileName, "CLASS" & I, "START_Y", Str(Options.START_Y))
        ' loop for items & values
        For X = 1 To UBound(Class(I).StartItem)
            Call PutVar(FileName, "CLASS" & I, "StartItem" & X, Str(Class(I).StartItem(X)))
            Call PutVar(FileName, "CLASS" & I, "StartValue" & X, Str(Class(I).StartValue(X)))
        Next
        ' loop for spells
        For X = 1 To UBound(Class(I).StartSpell)
            Call PutVar(FileName, "CLASS" & I, "StartSpell" & X, Str(Class(I).StartSpell(X)))
        Next
    Next

End Sub

Function CheckClasses() As Boolean
    Dim FileName As String
    FileName = App.Path & "\data\classes.ini"

    If Not FileExist(FileName, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
    Dim I As Long

    For I = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(I)), LenB(Class(I)))
        Class(I).Name = vbNullString
    Next

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim I As Long

    For I = 1 To MAX_ITEMS
        Call SaveItem(I)
    Next

End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim FileName As String
    Dim F As Long
    
    FileName = App.Path & "\data\items\item" & ItemNum & ".dat"
    
    If FileExist(FileName, True) Then Kill FileName
    
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Item(ItemNum)
    Close #F
End Sub

Sub LoadItems()
    Dim FileName As String
    Dim I As Long
    Dim F As Long
    
    Call CheckItems

    For I = 1 To MAX_ITEMS
        FileName = App.Path & "\data\Items\Item" & I & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Item(I)
        Close #F
    Next

End Sub

Sub CheckItems()
    Dim I As Long

    For I = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & I & ".dat") Then
            Call SaveItem(I)
        End If

    Next

End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."
End Sub

Sub ClearItems()
    Dim I As Long

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim I As Long

    For I = 1 To MAX_SHOPS
        Call SaveShop(I)
    Next

End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\shops\shop" & ShopNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Shop(ShopNum)
    Close #F
End Sub

Sub LoadShops()
    Dim FileName As String
    Dim I As Long
    Dim F As Long
    Call CheckShops

    For I = 1 To MAX_SHOPS
        FileName = App.Path & "\data\shops\shop" & I & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Shop(I)
        Close #F
    Next

End Sub

Sub CheckShops()
    Dim I As Long

    For I = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & I & ".dat") Then
            Call SaveShop(I)
        End If

    Next

End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim I As Long

    For I = 1 To MAX_SHOPS
        Call ClearShop(I)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal SpellNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\spells\spells" & SpellNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Spell(SpellNum)
    Close #F
End Sub

Sub SaveSpells()
    Dim I As Long
    Call SetStatus("Saving spells... ")

    For I = 1 To MAX_SPELLS
        Call SaveSpell(I)
    Next

End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim I As Long
    Dim F As Long
    Call CheckSpells

    For I = 1 To MAX_SPELLS
        FileName = App.Path & "\data\spells\spells" & I & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Spell(I)
        Close #F
    Next

End Sub

Sub CheckSpells()
    Dim I As Long

    For I = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & I & ".dat") Then
            Call SaveSpell(I)
        End If

    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).LevelReq = 1    'Needs to be 1 for the spell editor
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim I As Long

    For I = 1 To MAX_SPELLS
        Call ClearSpell(I)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim I As Long

    For I = 1 To MAX_NPCS
        Call SaveNpc(I)
    Next

End Sub

Sub SaveNpc(ByVal NpcNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\npcs\npc" & NpcNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , NPC(NpcNum)
    Close #F
End Sub

Sub LoadNpcs()
    Dim FileName As String
    Dim I As Long
    Dim F As Long
    Call CheckNpcs

    For I = 1 To MAX_NPCS
        FileName = App.Path & "\data\npcs\npc" & I & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , NPC(I)
        Close #F
    Next

End Sub

Sub CheckNpcs()
    Dim I As Long

    For I = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & I & ".dat") Then
            Call SaveNpc(I)
        End If

    Next

End Sub

Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).AttackSay = vbNullString
    NPC(Index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim I As Long

    For I = 1 To MAX_NPCS
        Call ClearNpc(I)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim I As Long

    For I = 1 To MAX_RESOURCES
        Call SaveResource(I)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
    Dim FileName As String
    Dim I As Long
    Dim F As Long
    Dim sLen As Long

    Call CheckResources

    For I = 1 To MAX_RESOURCES
        FileName = App.Path & "\data\resources\resource" & I & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Resource(I)
        Close #F
    Next

End Sub

Sub CheckResources()
    Dim I As Long

    For I = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & I & ".dat") Then
            Call SaveResource(I)
        End If
    Next

End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."
End Sub

Sub ClearResources()
    Dim I As Long

    For I = 1 To MAX_RESOURCES
        Call ClearResource(I)
    Next
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim I As Long

    For I = 1 To MAX_ANIMATIONS
        Call SaveAnimation(I)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Animation(AnimationNum)
    Close #F
End Sub

Sub LoadAnimations()
    Dim FileName As String
    Dim I As Long
    Dim F As Long
    Dim sLen As Long

    Call CheckAnimations

    For I = 1 To MAX_ANIMATIONS
        FileName = App.Path & "\data\animations\animation" & I & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Animation(I)
        Close #F
    Next

End Sub

Sub CheckAnimations()
    Dim I As Long

    For I = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & I & ".dat") Then
            Call SaveAnimation(I)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = "None."
End Sub

Sub ClearAnimations()
    Dim I As Long

    For I = 1 To MAX_ANIMATIONS
        Call ClearAnimation(I)
    Next
End Sub

' **********
' ** Maps **
' **********

Sub SaveMap(ByVal MapNum As Long)
    Dim FileName As String, F As Long, X As Long, Y As Long, I As Long

    ' save map data
    FileName = App.Path & "\data\maps\map" & MapNum & ".ini"

    ' if it exists then kill the ini
    If FileExist(FileName, True) Then Kill FileName

    ' General
    With Map(MapNum).MapData
        PutVar FileName, "General", "Name", .Name
        PutVar FileName, "General", "Music", .Music
        PutVar FileName, "General", "Moral", Val(.Moral)
        PutVar FileName, "General", "Up", Val(.Up)
        PutVar FileName, "General", "Down", Val(.Down)
        PutVar FileName, "General", "Left", Val(.left)
        PutVar FileName, "General", "Right", Val(.Right)
        PutVar FileName, "General", "BootMap", Val(.BootMap)
        PutVar FileName, "General", "BootX", Val(.BootX)
        PutVar FileName, "General", "BootY", Val(.BootY)
        PutVar FileName, "General", "MaxX", Val(.MaxX)
        PutVar FileName, "General", "MaxY", Val(.MaxY)
        PutVar FileName, "General", "BossNpc", Val(.BossNpc)
        PutVar FileName, "General", "Panorama", Val(.Panorama)
        
        PutVar FileName, "General", "Weather", Val(.Weather)
        PutVar FileName, "General", "WeatherIntensity", Val(.WeatherIntensity)
        PutVar FileName, "General", "Fog", Val(.Fog)
        PutVar FileName, "General", "FogSpeed", Val(.FogSpeed)
        PutVar FileName, "General", "FogOpacity", Val(.FogOpacity)
        PutVar FileName, "General", "Red", Val(.Red)
        PutVar FileName, "General", "Green", Val(.Green)
        PutVar FileName, "General", "Blue", Val(.Blue)
        PutVar FileName, "General", "Alpha", Val(.Alpha)
        PutVar FileName, "General", "Sun", Val(.Sun)
        PutVar FileName, "General", "DayNight", Val(.DAYNIGHT)
        For I = 1 To MAX_MAP_NPCS
            PutVar FileName, "General", "Npc" & I, Val(.NPC(I))
        Next
    End With

    ' Events
    PutVar FileName, "Events", "EventCount", Val(Map(MapNum).TileData.EventCount)

    If Map(MapNum).TileData.EventCount > 0 Then
        For I = 1 To Map(MapNum).TileData.EventCount
            With Map(MapNum).TileData.Events(I)
                PutVar FileName, "Event" & I, "Name", .Name
                PutVar FileName, "Event" & I, "x", Val(.X)
                PutVar FileName, "Event" & I, "y", Val(.Y)
                PutVar FileName, "Event" & I, "PageCount", Val(.PageCount)
            End With
            If Map(MapNum).TileData.Events(I).PageCount > 0 Then
                For X = 1 To Map(MapNum).TileData.Events(I).PageCount
                    With Map(MapNum).TileData.Events(I).EventPage(X)
                        PutVar FileName, "Event" & I & "Page" & X, "chkPlayerVar", Val(.chkPlayerVar)
                        PutVar FileName, "Event" & I & "Page" & X, "chkSelfSwitch", Val(.chkSelfSwitch)
                        PutVar FileName, "Event" & I & "Page" & X, "chkHasItem", Val(.chkHasItem)
                        PutVar FileName, "Event" & I & "Page" & X, "PlayerVarNum", Val(.PlayerVarNum)
                        PutVar FileName, "Event" & I & "Page" & X, "SelfSwitchNum", Val(.SelfSwitchNum)
                        PutVar FileName, "Event" & I & "Page" & X, "HasItemNum", Val(.HasItemNum)
                        PutVar FileName, "Event" & I & "Page" & X, "PlayerVariable", Val(.PlayerVariable)
                        PutVar FileName, "Event" & I & "Page" & X, "GraphicType", Val(.GraphicType)
                        PutVar FileName, "Event" & I & "Page" & X, "Graphic", Val(.Graphic)
                        PutVar FileName, "Event" & I & "Page" & X, "GraphicX", Val(.GraphicX)
                        PutVar FileName, "Event" & I & "Page" & X, "GraphicY", Val(.GraphicY)
                        PutVar FileName, "Event" & I & "Page" & X, "MoveType", Val(.MoveType)
                        PutVar FileName, "Event" & I & "Page" & X, "MoveSpeed", Val(.MoveSpeed)
                        PutVar FileName, "Event" & I & "Page" & X, "MoveFreq", Val(.MoveFreq)
                        PutVar FileName, "Event" & I & "Page" & X, "WalkAnim", Val(.WalkAnim)
                        PutVar FileName, "Event" & I & "Page" & X, "StepAnim", Val(.StepAnim)
                        PutVar FileName, "Event" & I & "Page" & X, "DirFix", Val(.DirFix)
                        PutVar FileName, "Event" & I & "Page" & X, "WalkThrough", Val(.WalkThrough)
                        PutVar FileName, "Event" & I & "Page" & X, "Priority", Val(.Priority)
                        PutVar FileName, "Event" & I & "Page" & X, "Trigger", Val(.Trigger)
                        PutVar FileName, "Event" & I & "Page" & X, "CommandCount", Val(.CommandCount)
                    End With
                    If Map(MapNum).TileData.Events(I).EventPage(X).CommandCount > 0 Then
                        For Y = 1 To Map(MapNum).TileData.Events(I).EventPage(X).CommandCount
                            With Map(MapNum).TileData.Events(I).EventPage(X).Commands(Y)
                                PutVar FileName, "Event" & I & "Page" & X & "Command" & Y, "Type", Val(.Type)
                                PutVar FileName, "Event" & I & "Page" & X & "Command" & Y, "Text", .Text
                                PutVar FileName, "Event" & I & "Page" & X & "Command" & Y, "Colour", Val(.colour)
                                PutVar FileName, "Event" & I & "Page" & X & "Command" & Y, "Channel", Val(.Channel)
                                PutVar FileName, "Event" & I & "Page" & X & "Command" & Y, "TargetType", Val(.TargetType)
                                PutVar FileName, "Event" & I & "Page" & X & "Command" & Y, "Target", Val(.target)
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If

    ' dump tile data
    FileName = App.Path & "\data\maps\map" & MapNum & ".dat"
    F = FreeFile

    ' if it exists then kill the ini
    If FileExist(FileName, True) Then Kill FileName

    With Map(MapNum)
        Open FileName For Binary As #F
        For X = 0 To .MapData.MaxX
            For Y = 0 To .MapData.MaxY
                Put #F, , .TileData.Tile(X, Y).Type
                Put #F, , .TileData.Tile(X, Y).Data1
                Put #F, , .TileData.Tile(X, Y).Data2
                Put #F, , .TileData.Tile(X, Y).Data3
                Put #F, , .TileData.Tile(X, Y).Data4
                Put #F, , .TileData.Tile(X, Y).Data5
                Put #F, , .TileData.Tile(X, Y).Autotile
                Put #F, , .TileData.Tile(X, Y).DirBlock
                For I = 1 To MapLayer.Layer_Count - 1
                    Put #F, , .TileData.Tile(X, Y).Layer(I).Tileset
                    Put #F, , .TileData.Tile(X, Y).Layer(I).X
                    Put #F, , .TileData.Tile(X, Y).Layer(I).Y
                Next
            Next
        Next
        Close #F
    End With

    If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
End Sub

Sub SaveMaps()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call SaveMap(I)
    Next

End Sub

Sub LoadMaps()
    Dim FileName As String, MapNum As Long

    Call CheckMaps

    For MapNum = 1 To MAX_MAPS
        LoadMap MapNum
        ClearTempTile MapNum
        CacheResources MapNum
        If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
    Next
End Sub

Sub LoadMap(MapNum As Long)
    Dim FileName As String, I As Long, F As Long, X As Long, Y As Long

    ' load map data
    FileName = App.Path & "\data\maps\map" & MapNum & ".ini"

    ' General
    With Map(MapNum).MapData
        .Name = GetVar(FileName, "General", "Name")
        .Music = GetVar(FileName, "General", "Music")
        .Moral = Val(GetVar(FileName, "General", "Moral"))
        .Up = Val(GetVar(FileName, "General", "Up"))
        .Down = Val(GetVar(FileName, "General", "Down"))
        .left = Val(GetVar(FileName, "General", "Left"))
        .Right = Val(GetVar(FileName, "General", "Right"))
        .BootMap = Val(GetVar(FileName, "General", "BootMap"))
        .BootX = Val(GetVar(FileName, "General", "BootX"))
        .BootY = Val(GetVar(FileName, "General", "BootY"))
        .MaxX = Val(GetVar(FileName, "General", "MaxX"))
        .MaxY = Val(GetVar(FileName, "General", "MaxY"))
        .BossNpc = Val(GetVar(FileName, "General", "BossNpc"))
        .Panorama = Val(GetVar(FileName, "General", "Panorama"))
        
        .Weather = Val(GetVar(FileName, "General", "Weather"))
        .WeatherIntensity = Val(GetVar(FileName, "General", "WeatherIntensity"))
        .Fog = Val(GetVar(FileName, "General", "Fog"))
        .FogSpeed = Val(GetVar(FileName, "General", "FogSpeed"))
        .FogOpacity = Val(GetVar(FileName, "General", "FogOpacity"))
        .Red = Val(GetVar(FileName, "General", "Red"))
        .Green = Val(GetVar(FileName, "General", "Green"))
        .Blue = Val(GetVar(FileName, "General", "Blue"))
        .Alpha = Val(GetVar(FileName, "General", "Alpha"))
        .Sun = Val(GetVar(FileName, "General", "Sun"))
        .DAYNIGHT = Val(GetVar(FileName, "General", "DayNight"))
        For I = 1 To MAX_MAP_NPCS
            .NPC(I) = Val(GetVar(FileName, "General", "Npc" & I))
        Next
    End With

    ' Events
    Map(MapNum).TileData.EventCount = Val(GetVar(FileName, "Events", "EventCount"))

    If Map(MapNum).TileData.EventCount > 0 Then
        ReDim Preserve Map(MapNum).TileData.Events(1 To Map(MapNum).TileData.EventCount)
        For I = 1 To Map(MapNum).TileData.EventCount
            With Map(MapNum).TileData.Events(I)
                .Name = GetVar(FileName, "Event" & I, "Name")
                .X = Val(GetVar(FileName, "Event" & I, "x"))
                .Y = Val(GetVar(FileName, "Event" & I, "y"))
                .PageCount = Val(GetVar(FileName, "Event" & I, "PageCount"))
            End With
            If Map(MapNum).TileData.Events(I).PageCount > 0 Then
                ReDim Preserve Map(MapNum).TileData.Events(I).EventPage(1 To Map(MapNum).TileData.Events(I).PageCount)
                For X = 1 To Map(MapNum).TileData.Events(I).PageCount
                    With Map(MapNum).TileData.Events(I).EventPage(X)
                        .chkPlayerVar = Val(GetVar(FileName, "Event" & I & "Page" & X, "chkPlayerVar"))
                        .chkSelfSwitch = Val(GetVar(FileName, "Event" & I & "Page" & X, "chkSelfSwitch"))
                        .chkHasItem = Val(GetVar(FileName, "Event" & I & "Page" & X, "chkHasItem"))
                        .PlayerVarNum = Val(GetVar(FileName, "Event" & I & "Page" & X, "PlayerVarNum"))
                        .SelfSwitchNum = Val(GetVar(FileName, "Event" & I & "Page" & X, "SelfSwitchNum"))
                        .HasItemNum = Val(GetVar(FileName, "Event" & I & "Page" & X, "HasItemNum"))
                        .PlayerVariable = Val(GetVar(FileName, "Event" & I & "Page" & X, "PlayerVariable"))
                        .GraphicType = Val(GetVar(FileName, "Event" & I & "Page" & X, "GraphicType"))
                        .Graphic = Val(GetVar(FileName, "Event" & I & "Page" & X, "Graphic"))
                        .GraphicX = Val(GetVar(FileName, "Event" & I & "Page" & X, "GraphicX"))
                        .GraphicY = Val(GetVar(FileName, "Event" & I & "Page" & X, "GraphicY"))
                        .MoveType = Val(GetVar(FileName, "Event" & I & "Page" & X, "MoveType"))
                        .MoveSpeed = Val(GetVar(FileName, "Event" & I & "Page" & X, "MoveSpeed"))
                        .MoveFreq = Val(GetVar(FileName, "Event" & I & "Page" & X, "MoveFreq"))
                        .WalkAnim = Val(GetVar(FileName, "Event" & I & "Page" & X, "WalkAnim"))
                        .StepAnim = Val(GetVar(FileName, "Event" & I & "Page" & X, "StepAnim"))
                        .DirFix = Val(GetVar(FileName, "Event" & I & "Page" & X, "DirFix"))
                        .WalkThrough = Val(GetVar(FileName, "Event" & I & "Page" & X, "WalkThrough"))
                        .Priority = Val(GetVar(FileName, "Event" & I & "Page" & X, "Priority"))
                        .Trigger = Val(GetVar(FileName, "Event" & I & "Page" & X, "Trigger"))
                        .CommandCount = Val(GetVar(FileName, "Event" & I & "Page" & X, "CommandCount"))
                    End With
                    If Map(MapNum).TileData.Events(I).EventPage(X).CommandCount > 0 Then
                        ReDim Preserve Map(MapNum).TileData.Events(I).EventPage(X).Commands(1 To Map(MapNum).TileData.Events(I).EventPage(X).CommandCount)
                        For Y = 1 To Map(MapNum).TileData.Events(I).EventPage(X).CommandCount
                            With Map(MapNum).TileData.Events(I).EventPage(X).Commands(Y)
                                .Type = Val(GetVar(FileName, "Event" & I & "Page" & X & "Command" & Y, "Type"))
                                .Text = GetVar(FileName, "Event" & I & "Page" & X & "Command" & Y, "Text")
                                .colour = Val(GetVar(FileName, "Event" & I & "Page" & X & "Command" & Y, "Colour"))
                                .Channel = Val(GetVar(FileName, "Event" & I & "Page" & X & "Command" & Y, "Channel"))
                                .TargetType = Val(GetVar(FileName, "Event" & I & "Page" & X & "Command" & Y, "TargetType"))
                                .target = Val(GetVar(FileName, "Event" & I & "Page" & X & "Command" & Y, "Target"))
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If

    ' dump tile data
    FileName = App.Path & "\data\maps\map" & MapNum & ".dat"
    F = FreeFile

    ' redim the map
    ReDim Map(MapNum).TileData.Tile(0 To Map(MapNum).MapData.MaxX, 0 To Map(MapNum).MapData.MaxY) As TileRec

    With Map(MapNum)
        Open FileName For Binary As #F
        For X = 0 To .MapData.MaxX
            For Y = 0 To .MapData.MaxY
                Get #F, , .TileData.Tile(X, Y).Type
                Get #F, , .TileData.Tile(X, Y).Data1
                Get #F, , .TileData.Tile(X, Y).Data2
                Get #F, , .TileData.Tile(X, Y).Data3
                Get #F, , .TileData.Tile(X, Y).Data4
                Get #F, , .TileData.Tile(X, Y).Data5
                Get #F, , .TileData.Tile(X, Y).Autotile
                Get #F, , .TileData.Tile(X, Y).DirBlock
                For I = 1 To MapLayer.Layer_Count - 1
                    Get #F, , .TileData.Tile(X, Y).Layer(I).Tileset
                    Get #F, , .TileData.Tile(X, Y).Layer(I).X
                    Get #F, , .TileData.Tile(X, Y).Layer(I).Y
                Next
            Next
        Next
        Close #F
    End With
End Sub

Sub CheckMaps()
    Dim I As Long

    For I = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & I & ".dat") Or Not FileExist("\Data\maps\map" & I & ".ini") Then
            Call SaveMap(I)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, Index)), LenB(MapItem(MapNum, Index)))
    MapItem(MapNum, Index).PlayerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).NPC(Index)), LenB(MapNpc(MapNum).NPC(Index)))
End Sub

Sub ClearMapNpcs()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, Y)
        Next
    Next

End Sub

Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).MapData.Name = vbNullString
    Map(MapNum).MapData.MaxX = MAX_MAPX
    Map(MapNum).MapData.MaxY = MAX_MAPY
    ReDim Map(MapNum).TileData.Tile(0 To Map(MapNum).MapData.MaxX, 0 To Map(MapNum).MapData.MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call ClearMap(I)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
    Case HP
        With Class(ClassNum)
            GetClassMaxVital = 100 + (.Stat(Endurance) * 5) + 2
        End With
    Case MP
        With Class(ClassNum)
            GetClassMaxVital = 30 + (.Stat(Intelligence) * 10) + 2
        End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Sub ClearParty(ByVal partynum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partynum)), LenB(Party(partynum)))
End Sub
