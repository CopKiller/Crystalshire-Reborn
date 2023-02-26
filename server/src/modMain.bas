Attribute VB_Name = "modMain"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type PlayerUDT
    Buffer As clsBuffer
    ' Network Data
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    PacketInIndex As Byte   ' Holds the index of what packetkey for incoming packets
    PacketOutIndex As Byte  ' Holds the index of what packetkey for outgoing packets
    
    ' Tempo que o usurio foi conectado.
    ConnectedTime As Long
    TokenAccepted As Boolean
End Type

Public Player(1 To MAX_PLAYERS) As PlayerUDT
Public ServerOnline As Boolean
Public GameCPS As Long

Sub ClearPlayer(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Player(Index)), LenB(Player(Index))
    Set Player(Index).Buffer = New clsBuffer
End Sub

Sub Main()
Dim i As Long

    Randomize Timer                                                             ' Randomizes the system timer
    
    frmMain.Show

    frmMain.Socket(0).RemoteHost = frmMain.Socket(0).LocalIP                ' Sets up the server ip
    frmMain.Socket(0).LocalPort = AUTH_SERVER_PORT                           ' Sets up the default port
    frmMain.Socket(0).Listen                                                  ' Start listening

    ' Setup our gameServerConnection
    ConnectToGameServer
        
    InitMessages                                                                ' Need to init messages for packets
    
    For i = 1 To MAX_PLAYERS
        ClearPlayer i
        Load frmMain.Socket(i)                                                ' load sockets
    Next

    Set classMD5 = New clsMD5
    Set cryptography = New clsCryptography
    
    LoadSystemTray
      
    SetStatus "Initialization complete. AuthServer loaded."
    
    Call ServerLoop
End Sub

Public Sub DestroyServer()
Dim i As Long

    On Error Resume Next
    
    ServerOnline = False

    'DisconnectFromSqlServer
    
    For i = 1 To MAX_PLAYERS
        Set Player(i).Buffer = Nothing
        Unload frmMain.Socket(i)
    Next
    
    DestroySystemTray
    
    Unload frmMain
    End
End Sub

Function RandomString(ByVal mask As String) As String
Dim i As Integer, acode As Integer, options As String, char As String
    
    ' initialize result with proper lenght
    RandomString = mask
    
    For i = 1 To Len(mask)
        ' get the character
        char = Mid$(mask, i, 1)
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
        Mid(RandomString, i, 1) = char
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
    AddText frmMain.txtLog, Status
End Sub

Public Function IsAlphaNumeric(s As String) As Boolean
    If Not s Like "*[!0-9A-Za-z]*" Then IsAlphaNumeric = True
End Function

Public Function IsAlpha(s As String) As Boolean
    If Not s Like "*[!A-Za-z]*" Then IsAlpha = True
End Function

Public Function FileExist(ByVal filename As String) As Boolean
    If Dir$(filename) = vbNullString Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Public Function SaltPassword(ByVal Password As String, ByVal salt As String) As String
    SaltPassword = MD5(MD5(salt) & Password)
End Function

Public Function MD5(ByVal theString As String) As String
    MD5 = LCase$(classMD5.DigestStrToHexStr(theString))
End Function

Public Sub ServerLoop()
    Dim tmr1000 As Long
    Dim tick As Long
    Dim TickCPS As Long
    Dim CPS As Long

    ServerOnline = True

    Do While ServerOnline
        tick = GetTickCount
        
        If tick > tmr1000 Then

            ' Fecha as conexoes em que o token ainda nao foi aceito.
            Call CheckConnectionTime

            ' reset timer
            tmr1000 = GetTickCount + 1000
        End If

        'If Not CPSUnlock Then Sleep 1
        Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < tick Then
            GameCPS = CPS
            TickCPS = tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
        
        frmMain.lblCps.Caption = "CPS : " & GameCPS
        
    Loop

End Sub
