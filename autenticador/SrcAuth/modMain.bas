Attribute VB_Name = "modMain"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type PlayerUDT
    Buffer As clsBuffer
    ' Network Data
    DataTimer As Currency
    DataBytes As Long
    DataPackets As Long
    PacketInIndex As Byte   ' Holds the index of what packetkey for incoming packets
    PacketOutIndex As Byte  ' Holds the index of what packetkey for outgoing packets
    
    ' Tempo que o usurio foi conectado.
    ConnectedTime As Currency
    TokenAccepted As Boolean
End Type

Public TempPlayer(1 To MAX_PLAYERS) As PlayerUDT
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public ServerOnline As Boolean
Public GameCPS As Long

Sub ClearPlayer(ByVal Index As Long)

    ZeroMemory ByVal VarPtr(Player(Index)), LenB(Player(Index))
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    Player(Index).Name = vbNullString
    Player(Index).StartPremium = vbNullString
    Player(Index).Class = 1
    
    ZeroMemory ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index))
    Set TempPlayer(Index).Buffer = New clsBuffer
End Sub

Sub Main()
Dim I As Long
Dim F As Long

    InitCryptographyKey
    
    ' This must be called before any Tick calls because it states what the values of Tick will be
    InitTime
    
    Randomize Timer                                                             ' Randomizes the system timer
    
    frmMain.Show

    frmMain.Socket(0).RemoteHost = frmMain.Socket(0).LocalIP                ' Sets up the server ip
    frmMain.Socket(0).LocalPort = AUTH_SERVER_PORT                           ' Sets up the default port
    frmMain.Socket(0).Listen                                                  ' Start listening
    
    ChkDir App.Path & "\", "accounts"
    
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("_charlist.txt") Then
        F = FreeFile
        Open App.Path & "\_charlist.txt" For Output As #F
        Close #F
    End If

    ' Setup our gameServerConnection
    ConnectToGameServer
        
    InitMessages                                                                ' Need to init messages for packets
    GS_InitMessages
    
    For I = 1 To MAX_PLAYERS
        ClearPlayer I
        Load frmMain.Socket(I)                                                ' load sockets
    Next
    
    LoadSystemTray
    
    CreateEmailObject
      
    SetStatus "Initialization complete. AuthServer loaded."
    
    Call ServerLoop
End Sub

Public Sub DestroyServer()
Dim I As Long

    On Error Resume Next
    
    ServerOnline = False
    
    For I = 1 To MAX_PLAYERS
        Set TempPlayer(I).Buffer = Nothing
        Unload frmMain.Socket(I)
    Next
    
    DestroySystemTray
    
    DestroyEmailObject
    
    Unload frmMain
    End
End Sub

Function RandomString(ByVal mask As String) As String
Dim I As Integer, acode As Integer, options As String, char As String
    
    ' initialize result with proper lenght
    RandomString = mask
    
    For I = 1 To Len(mask)
        ' get the character
        char = Mid$(mask, I, 1)
        Select Case char
            Case "?"
                char = Chr$(1 + Rnd * 127)
                options = ""
            Case "#"
                options = "0123456789"
            Case "A"
                options = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
            Case "N"
                options = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0" _
                    & "123456789"
            Case "H"
                options = "0123456789ABCDEF"
            Case Else
                ' don't modify the character
                options = ""
        End Select
    
        ' select a random char in the option string
        If Len(options) Then
            ' select a random char
            ' note that we add an extra char, in case RND returns 1
            char = Mid$(options & Right$(options, 1), 1 + Int(Rnd * Len(options)), 1)
        End If
        
        ' insert the character in result string
        Mid(RandomString, I, 1) = char
    Next
End Function

Public Sub AddText(ByVal rTxt As TextBox, ByVal Msg As String)
Dim s As String

    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmMain.txtLog.Text = vbNullString
        NumLines = 0
    End If
    s = Msg & vbCrLf
    rTxt.SelStart = Len(rTxt.Text)
    rTxt.SelText = s
    rTxt.SelStart = Len(rTxt.Text) - 1
    
    AddLog Msg
End Sub

Sub AddLog(ByVal Text As String)
Dim filename As String
Dim F As Long

    filename = App.Path & "/log.txt"

    If Not FileExist(filename) Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Append As #F
    Print #F, DateValue(Now) & " " & Time & ": " & Text
    Close #F

End Sub

Sub SetStatus(ByRef Status As String)
    AddText frmMain.txtLog, Format(Time, "HH:MM") & Space(1) & Status
End Sub

Public Function IsAlphaNumeric(s As String) As Boolean
    If Not s Like "*[!0-9A-Za-z]*" Then IsAlphaNumeric = True
End Function

Public Function IsAlpha(s As String) As Boolean
    If Not s Like "*[!A-Za-z]*" Then IsAlpha = True
End Function

Public Function FileExist(ByVal filename As String) As Boolean
    If dir$(filename) = vbNullString Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Public Sub ServerLoop()
    Dim tmr1000 As Currency
    Dim TickCPS As Currency
    Dim CPS As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = getTime
        
        If Tick > tmr1000 Then

            ' Fecha as conexoes em que o token ainda nao foi aceito.
            Call CheckConnectionTime

            ' reset timer
            tmr1000 = getTime + 1000
        End If

        'If Not CPSUnlock Then Sleep 1
        Sleep 1
        If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
        
        frmMain.lblCps.Caption = "CPS : " & GameCPS
        
    Loop

End Sub


' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function

Public Function isBanned_IP(ByVal IP As String) As Boolean
    Dim filename As String, fIP As String, F As Long

    filename = App.Path & "\banlist_ip.txt"

    ' Check if file exists
    If Not FileExist(filename) Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Input As #F

    Do While Not EOF(F)
        Input #F, fIP

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            isBanned_IP = True
            Close #F
            Exit Function
        End If
    Loop

    Close #F
End Function

Public Function isBanned_Account(ByVal Index As Long) As Boolean
    If Player(Index).isBanned = 1 Then
        isBanned_Account = True
    Else
        isBanned_Account = False
    End If
End Function

Function GetPlayerIP(ByVal Index As Long) As String

    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerIP = frmMain.Socket(Index).RemoteHostIP
End Function
