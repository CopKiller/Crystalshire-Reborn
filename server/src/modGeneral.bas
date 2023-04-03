Attribute VB_Name = "modGeneral"
Option Explicit

Public Sub Main()
    Call InitServer
End Sub

Public Sub InitServer()
    Dim i As Long
    Dim F As Long
    Dim time1 As Currency
    Dim time2 As Currency

    ' log on by default
    ServerLog = True

    InitCRC32

    InitCryptographyKey

    ' This must be called before any Tick calls because it states what the values of Tick will be
    InitTime
    
    GameSecondsPerSecond = 30
    GameMinutesPerMinute = 1
    GameSeconds = Second(Now)
    GameMinutes = Minute(Now)
    GameHours = Hour(Now)
    
    ' Show Console Msgs
    frmServer.SSTab1.Tab = 0

    ' cache packet pointers
    Call InitMessages
    Call Auth_InitMessages
    Call Event_InitMessages

    ' time the load
    time1 = getTime
    frmServer.Show

    ' Initialize the random-number generator
    Randomize

    ' Check if the directory is there, if its not make it
    'ChkDir App.Path & "\Data\", "accounts"
    ChkDir App.Path & "\Data\", "animations"
    ChkDir App.Path & "\Data\", "items"
    ChkDir App.Path & "\Data\", "logs"
    ChkDir App.Path & "\Data\", "maps"
    ChkDir App.Path & "\Data\", "npcs"
    ChkDir App.Path & "\Data\", "resources"
    ChkDir App.Path & "\Data\", "shops"
    ChkDir App.Path & "\Data\", "spells"
    ChkDir App.Path & "\Data\", "convs"
    ChkDir App.Path & "\Data\", "serials"
    ChkDir App.Path & "\Data\", "quests"
    ChkDir App.Path & "\Data\", "guilds"
    ChkDir App.Path & "\Data\", "conjuntos"

    ' set quote character
    vbQuote = ChrW$(34)    ' "

    ' load options, set if they dont exist
    If Not FileExist(App.Path & "\data\options.ini", True) Then
        Options.MOTD = "Welcome to Crystalshire."
        Options.PartyBonus = 0
        Options.START_MAP = 1
        Options.START_X = 10
        Options.START_Y = 15
        SaveOptions
    Else
        LoadOptions
    End If

    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_SERVER_PORT

    ' Get the authentication socket going
    frmServer.AuthSocket.RemoteHost = AUTH_SERVER_IP
    frmServer.AuthSocket.LocalPort = SERVER_AUTH_PORT
    
    ' Get the event socket going
    frmServer.EventSocket.RemoteHost = EVENT_SERVER_IP
    frmServer.EventSocket.RemotePort = EVENT_SERVER_PORT

    ' Init all the player sockets
    Call SetStatus("Initializing player array...")

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    
    SetStatus "Caching map, items, npcs CRC32 checksums..."
    ' cache map crc32s
    For i = 1 To MAX_MAPS
        GetMapCRC32 i
    Next i
    ' cache item crc32s
    For i = 1 To MAX_ITEMS
        GetItemCRC32 i
    Next i
    ' cache npc crc32s
    For i = 1 To MAX_NPCS
        GetNpcCRC32 i
    Next i
    
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Creating All Cache Compress...")
    Call CreateFullCache
    Call SetStatus("Loading System Tray...")
    Call LoadSystemTray

    ' Start listening
    frmServer.Socket(0).Listen
    frmServer.AuthSocket.Listen
    Call UpdateCaption
    time2 = getTime
    
    Call SetStatus("Initialization complete. Server loaded in " & Int(time2 - time1) & "ms.")
    
    Call ConnectToEventServer

    ' reset shutdown value
    isShuttingDown = False
    
    ' Enable Controls
    frmServer.SSTab1.Enabled = True

    ' Starts the server loop
    ServerLoop
End Sub

Public Sub SendAllSaves()
    Dim sPath As String, sFile As String, sCount As Integer
    If IsConnectedAuthServer Then

        'Dim fso As Object


        ' ACCOUNT
        sPath = App.Path & "\data\accounts\"
        sFile = dir(sPath & "*.bin", vbDirectory)

        If dir(sPath, vbDirectory) <> vbNullString Then
            Do While sFile <> ""
                If InStr(sFile, ".bin") > 0 Then
                    Call LoadAccount_SendAuthServer(sPath & sFile)
                    Kill sPath & sFile
                    RmDir sPath
                    sCount = sCount + 1
                End If
                sFile = dir
            Loop

        End If

        If sCount > 0 Then
            Call TextLoginAdd("## " & sCount & " Dados de jogadores foram enviados com sucesso! ##")
        Else
            Call TextLoginAdd("## Nenhum dado de jogador perdido foi encontrado! ##")
        End If
    Else
        Call TextLoginAdd("## Dados dos jogadores não foram enviados, falha na comunicação com o servidor de autenticação! ##")
    End If
End Sub

Private Sub LoadAccount_SendAuthServer(ByVal FileName As String)
    Dim F As Long
    Dim Jogador As PlayerRec
    Dim Buffer As clsBuffer, DataSize As Long, tempData() As Byte

    If Trim$(FileName) = vbNullString Then Exit Sub
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , Jogador
    Close #F
    DataSize = LenB(Jogador)
    ReDim tempData(DataSize - 1)
    CopyMemory tempData(0), ByVal VarPtr(Jogador), DataSize
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong GSavePlayer
    Buffer.WriteString Trim$(Jogador.Login)
    Buffer.WriteBytes tempData

    Auth_SendDataTo Buffer.ToArray
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub DestroyServer()
    Dim i As Long
    ServerOnline = False
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")

    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next
    End
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTiles
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing Resources...")
    Call ClearResources
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing conversations...")
    Call ClearConvs
    Call SetStatus("Clearing guilds...")
    Call ClearGuilds
    Call SetStatus("Clearing seriais...")
    Call ClearSerials
    Call SetStatus("Clearing quests...")
    Call ClearQuests
    Call SetStatus("Clearing conjuntos...")
    Call ClearConjuntos
    Call SetStatus("Clearing Lottery and bets...")
    Call ClearLottery
    Call ClearBets
End Sub

Private Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading Resources...")
    Call LoadResources
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    Call SetStatus("Loading conversations...")
    Call LoadConvs
    Call SetStatus("Loading guilds...")
    Call LoadGuilds
    Call SetStatus("Loading serials...")
    Call LoadSerials
    Call SetStatus("Loading quests...")
    Call LoadQuests
    Call SetStatus("Loading Conjuntos...")
    Call LoadConjuntos
    Call SetStatus("Loading DayReward...")
    Call DayRewardInit
End Sub

Sub SetHighIndex()
    Dim i As Integer
    Dim x As Integer

    For i = 0 To MAX_PLAYERS
        x = MAX_PLAYERS - i

        If IsConnected(x) = True Then
            Player_HighIndex = x
            Exit Sub
        End If

    Next i

    Player_HighIndex = 0

End Sub

Public Sub TextAdd(Msg As String)
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
End Sub

Public Sub TextLoginAdd(Msg As String)
    LoginNumLines = LoginNumLines + 1

    If LoginNumLines >= MAX_LINES Then
        frmServer.txtLogin.Text = vbNullString
        LoginNumLines = 0
    End If

    frmServer.txtLogin.Text = frmServer.txtLogin.Text & vbNewLine & Msg
    frmServer.txtLogin.SelStart = Len(frmServer.txtLogin.Text)
End Sub

Public Sub TextEventAdd(Msg As String)
    EventNumLines = EventNumLines + 1

    If EventNumLines >= MAX_LINES Then
        frmServer.txtEvent.Text = vbNullString
        LoginNumLines = 0
    End If

    frmServer.txtEvent.Text = frmServer.txtEvent.Text & vbNewLine & Msg
    frmServer.txtEvent.SelStart = Len(frmServer.txtEvent.Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function

Public Function AryCount(ByRef Ary() As Byte) As Long
    On Error Resume Next

    AryCount = UBound(Ary) + 1
End Function
