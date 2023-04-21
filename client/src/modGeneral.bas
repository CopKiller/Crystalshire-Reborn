Attribute VB_Name = "modGeneral"
Option Explicit
' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public Sub Main()
    Dim i As Long

    'verifica se o player ta rodando em modo de administrador
   ' If Not GetWindowsVersion = "Windows XP" Then
   '     If Not IsElevatedAccess And App.LogMode <> 0 Then
   '         Call MsgBox("Por Favor, Rode-o em modo Administrador.", vbOKOnly, GAME_NAME)
   '         End
   '     End If
   ' End If

    InitCRC32

    InitCryptographyKey
    
    'GenerateKeyAndIV
    
    ' This must be called before any Tick calls because it states what the values of Tick will be
    InitTime

    ' Check if the directory is there, if its not make it
    ChkDir App.path & "\data files\", "graphics"
    ChkDir App.path & "\data files\graphics\", "animations"
    ChkDir App.path & "\data files\graphics\", "characters"
    ChkDir App.path & "\data files\graphics\", "items"
    ChkDir App.path & "\data files\graphics\", "paperdolls"
    ChkDir App.path & "\data files\graphics\", "resources"
    ChkDir App.path & "\data files\graphics\", "spellicons"
    ChkDir App.path & "\data files\graphics\", "tilesets"
    ChkDir App.path & "\data files\graphics\", "faces"
    ChkDir App.path & "\data files\graphics\", "gui"
    ChkDir App.path & "\data files\", "logs"
    ChkDir App.path & "\data files\", "maps"
    ChkDir App.path & "\data files\", "music"
    ChkDir App.path & "\data files\", "sound"
    ChkDir App.path & "\data files\", "video"
    ChkDir App.path & "\data files\", "items"
    ChkDir App.path & "\data files\", "npcs"

    ' load options
    LoadOptions

    ' check the resolution
    CheckResolution

    ' load dx8
    If Options.Fullscreen Then
        frmMain.BorderStyle = 0
        frmMain.caption = frmMain.caption
    End If

    frmMain.Show
    InitDX8 frmMain.hWnd
    
    ' Do this only if we are running from a .exe. If run from IDE it messes up debugging
    If App.LogMode = 1 Then HookForMouseWheel frmMain.hWnd
    myHWnd = frmMain.hWnd

    If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents

    LoadTextures
    LoadFonts
    ' initialise the gui
    InitGUI
    ' Resize the GUI to screen size
    ResizeGUI
    ' initialise sound & music engines
    Init_Music
    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34)
    ' randomize rnd's seed
    Randomize
    
    Call SetCaption("Initializing TCP settings.")
    Call TcpInit(AUTH_SERVER_IP, AUTH_SERVER_PORT)
    Call InitMessages
    ' Reset values
    Ping = -1
    ' cache the buttons then reset & render them
    Call SetCaption("Caching map CRC32 checksums...")
    ' cache map crc32s
    For i = 1 To MAX_MAPS
        GetMapCRC32 i
    Next
    ' Load items
    Call SetCaption("Carregando Items...")
    Call LoadItems
    ' Load npcs
    Call SetCaption("Carregando Npcs...")
    Call LoadNpcs
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12    ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12    ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0    ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23    ' right
    DirArrowY(4) = 12
    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.Weapon
    ' set status
    SetCaption vbNullString
    ' show the main menu
    'frmMain.Show
    inMenu = True
    ' show login window
    HideWindows
    ShowWindow GetWindowIndex("winLogin")

    inSmallChat = True
    ' Set the loop going
    fadeAlpha = 255
    
    If Options.PlayIntro = 1 Then
        PlayIntro
    Else
        videoPlaying = False
        frmMain.picIntro.visible = False
        ' play the menu music
        If Len(Trim$(MenuMusic)) > 0 Then Play_Music Trim$(MenuMusic)
    End If
    
    MenuBG = Rand(1, Count_Panoramas)
    
    ' Colocar informações do tooltip no menu do jogador, sobre as teclas pré-configuradas no editor de controles de cada menu!
    RefreshMenu_Tooltip
    
    SetCaption vbNullString
    
    ' Set the active control
    If Not Len(Windows(GetWindowIndex("winLogin")).Controls(GetControlIndex("winLogin", "txtUser")).Text) > 0 Then
        SetActiveControl GetWindowIndex("winLogin"), GetControlIndex("winLogin", "txtUser")
    Else
        SetActiveControl GetWindowIndex("winLogin"), GetControlIndex("winLogin", "txtPass")
    End If
        
    MenuLoop
End Sub

Public Sub RefreshMenu_Tooltip()
Dim sString As String, tmpString() As String
' Colocar informações do tooltip no menu do jogador, sobre as teclas pré-configuradas no editor de controles de cada menu!
' Utiliza do split pra não acumular atalhos no nome do menu!
        With Windows(GetWindowIndex("winMenu"))
            sString = .Controls(GetControlIndex("winMenu", "btnInv")).tooltip
            tmpString = Split(sString, "(")
            .Controls(GetControlIndex("winMenu", "btnInv")).tooltip = tmpString(0) & "(" & KeycodeChar(Options.Bolsa) & ")"
            sString = .Controls(GetControlIndex("winMenu", "btnSkills")).tooltip
            tmpString = Split(sString, "(")
            .Controls(GetControlIndex("winMenu", "btnSkills")).tooltip = tmpString(0) & "(" & KeycodeChar(Options.Magias) & ")"
            sString = .Controls(GetControlIndex("winMenu", "btnChar")).tooltip
            tmpString = Split(sString, "(")
            .Controls(GetControlIndex("winMenu", "btnChar")).tooltip = tmpString(0) & "(" & KeycodeChar(Options.Personagem) & ")"
            sString = .Controls(GetControlIndex("winMenu", "btnGuild")).tooltip
            tmpString = Split(sString, "(")
            .Controls(GetControlIndex("winMenu", "btnGuild")).tooltip = tmpString(0) & "(" & KeycodeChar(Options.Guild) & ")"
            sString = .Controls(GetControlIndex("winMenu", "btnQuest")).tooltip
            tmpString = Split(sString, "(")
            .Controls(GetControlIndex("winMenu", "btnQuest")).tooltip = tmpString(0) & "(" & KeycodeChar(Options.Quests) & ")"
        End With
End Sub

Public Sub AddChar(Name As String, sex As Long, Class As Long, Sprite As Long)

    If ConnectToServer Then
        Call SetStatus("Sending character information.")
        Call SendAddChar(Name, sex, Class, Sprite)
        Exit Sub
    Else
        ShowWindow GetWindowIndex("winLogin")
        Dialogue "Connection Problem", "Cannot connect to game server.", "", TypeALERT
    End If

End Sub

Public Sub Login(Name As String, Password As String)

    TcpInit AUTH_SERVER_IP, AUTH_SERVER_PORT

    If ConnectToServer Then
        Call SetStatus("Sending login information.")
        Call SendAuthLogin(Name, Password)
        ' save details
        If Options.SaveUser Then Options.Username = Name Else Options.Username = vbNullString
        If Options.SavePass Then Options.Password = Password Else Options.Password = vbNullString
        SaveOptions
        
        ' Temporário pro auto reconnect!
        Options.TmpLogin = Name
        Options.TmpPassword = Password
        Exit Sub
    ElseIf Not isReconnect Then
        ShowWindow GetWindowIndex("winLogin")
        Dialogue "Connection Problem", "Cannot connect to login server.", "Please try again later.", TypeALERT
    End If

End Sub

Public Sub SendRegister(Name As String, Password As String, codigo As String, BirthDay As String)

    TcpInit AUTH_SERVER_IP, AUTH_SERVER_PORT

    If ConnectToServer Then
        Call SetStatus("Sending Register information.")
        Call SendNewAccount(Name, Password, codigo, BirthDay)
    Else
        ShowWindow GetWindowIndex("winregister")
        Dialogue "Connection Problem", "Cannot connect to login server.", "Please try again later.", TypeALERT
    End If

End Sub
Public Sub logoutGame()
    Dim i As Long
    InGame = False

    DestroyTCP

    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    If MyIndex > 0 Then
        Call ZeroMemory(ByVal VarPtr(Player(MyIndex)), LenB(Player(MyIndex)))
    End If

    ' destroy temp values
    DragInvSlotNum = 0
    LastItemDesc = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    Unload frmEditor_Conv
    Unload frmEditor_Serial
    Unload frmAdmin
    Unload frmEditor_Premium
    ' clear chat
    For i = 1 To ChatLines
        Chat(i).Text = vbNullString
    Next
    
    Call ClearAllGameData

    Call ResetReconnectVariables
    
    inMenu = True
    MenuLoop
End Sub

Sub ClearAllGameData()
    Call ClearQuests
    Call ClearSerials
    'Call ClearItems
    Call ClearSpells
    'Call ClearNpcs
    Call ClearResources
    Call ClearMap
    Call ClearShops
    Call ClearConvs
    Call ClearAnimations
End Sub

Sub GameInit()
    Dim musicFile As String
    ' hide gui
    InBank = False
    InTrade = False
    CloseShop
    ' get ping
    GetPing
    ' play music
    musicFile = Trim$(Map.MapData.Music)

    If Not musicFile = "None." Then
        Play_Music musicFile
    Else
        Stop_Music
    End If

    SetStatus vbNullString
End Sub

Public Sub DestroyGame()

    InGame = False
    StopIntro
    Call DestroyTCP
    ' destroy music & sound engines
    Destroy_Music
    Call UnloadAllForms
    
    Call UnHookMouseWheel
    'destroy objects in reverse order
    DestroyDX8
    
    Sleep 500    ' Give it some time ;)
    
    End
End Sub

Public Sub UnloadAllForms()
    Dim frm As Form

    For Each frm In VB.Forms
        Unload frm
    Next

End Sub

Public Sub SetStatus(ByVal caption As String)

' Não seta o status se tiver reconectando
If isReconnect Then Exit Sub

    HideWindows
    If Len(Trim$(caption)) > 0 Then
        ShowWindow GetWindowIndex("winLoading")
        Windows(GetWindowIndex("winLoading")).Controls(GetControlIndex("winLoading", "lblLoading")).Text = caption
    Else
        HideWindow GetWindowIndex("winLoading")
        Windows(GetWindowIndex("winLoading")).Controls(GetControlIndex("winLoading", "lblLoading")).Text = vbNullString
    End If
End Sub

Public Sub SetCaption(ByVal caption As String)

    If Len(Trim$(caption)) > 0 Then
        frmMain.caption = caption
        If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
    Else
        frmMain.caption = GAME_NAME
    End If
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)

    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If

    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean

    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
    Dim i As Long, tmpNum As Long
    ' Prevent high ascii chars
    tmpNum = Len(sInput)

    For i = 1 To tmpNum

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Dialogue "Illegal Characters", "This string contains illegal characters.", "", TypeALERT
            Exit Function
        End If

    Next

    isStringLegal = True
End Function

Public Sub PopulateLists()
    Dim strLoad As String, i As Long
    ' Cache music list
    strLoad = dir$(App.path & MUSIC_PATH & "*.*")
    i = 1

    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop

    ' Cache sound list
    strLoad = dir$(App.path & SOUND_PATH & "*.*")
    i = 1

    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop

End Sub
