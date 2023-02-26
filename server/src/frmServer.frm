VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock AuthSocket 
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblCPS"
      Tab(0).Control(1)=   "lblCpsLock"
      Tab(0).Control(2)=   "txtText"
      Tab(0).Control(3)=   "txtChat"
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(5)=   "chkMsgWindow"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwInfo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Control "
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraDatabase"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraServer"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "chkServerLog"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdShutDown"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CheckBox chkMsgWindow 
         Caption         =   "Msg Window"
         Height          =   375
         Left            =   -69600
         TabIndex        =   34
         Top             =   2880
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "Game Time"
         Height          =   615
         Left            =   -70440
         TabIndex        =   32
         Top             =   0
         Width           =   1815
         Begin VB.Label lblGameTime 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "xx:xx"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   0
            TabIndex        =   33
            Top             =   185
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdShutDown 
         Caption         =   "Desligar 30 Segs"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   0
         Width           =   1575
      End
      Begin VB.CheckBox chkServerLog 
         Caption         =   "Logs"
         Height          =   255
         Left            =   5760
         TabIndex        =   16
         Top             =   0
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.Frame fraServer 
         Caption         =   "Server"
         Height          =   2895
         Left            =   3120
         TabIndex        =   1
         Top             =   360
         Width           =   3135
         Begin VB.CommandButton Command3 
            Caption         =   "Open Lottery"
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   2160
            Width           =   2415
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Mais Configurações"
            Height          =   255
            Left            =   360
            TabIndex        =   36
            Top             =   2520
            Width           =   2415
         End
         Begin VB.TextBox txtGameSite 
            Height          =   285
            Left            =   1320
            TabIndex        =   30
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtGameName 
            Height          =   285
            Left            =   1320
            TabIndex        =   28
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2280
            TabIndex        =   23
            Text            =   "0"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Text            =   "0"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMap 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            TabIndex        =   21
            Text            =   "0"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMOTD 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Game Site:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Game Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label Label5 
            Caption         =   "Y"
            Height          =   255
            Left            =   2040
            TabIndex        =   26
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "X"
            Height          =   255
            Left            =   1200
            TabIndex        =   25
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Map"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "StartMap Location:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Boas Vindas:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraDatabase 
         Caption         =   "Reload"
         Height          =   2895
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2895
         Begin VB.CommandButton Command2 
            Caption         =   "Conjuntos"
            Height          =   255
            Left            =   1440
            TabIndex        =   39
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton cmdCheckIn 
            Caption         =   "CheckIn"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadQuests 
            Caption         =   "Quests"
            Height          =   255
            Left            =   1440
            TabIndex        =   35
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadSeriais 
            Caption         =   "Seriais"
            Height          =   255
            Left            =   1440
            TabIndex        =   31
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Height          =   255
            Left            =   1440
            TabIndex        =   13
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Height          =   255
            Left            =   1440
            TabIndex        =   12
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Classes"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   -74880
         TabIndex        =   3
         Top             =   2880
         Width           =   5175
      End
      Begin VB.TextBox txtText 
         Height          =   2175
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   600
         Width           =   6255
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4895
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label lblCpsLock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[Unlock]"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblCPS 
         Caption         =   "CPS: 0"
         Height          =   255
         Left            =   -74040
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
      Begin VB.Menu mnuMute 
         Caption         =   "Toggle Mute"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheckIn_Click()
    Call DayRewardInit
End Sub

Private Sub cmdReloadQuests_Click()
    Dim I As Long
    Call LoadQuests
    Call TextAdd("All Quests reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            SendQuests I
        End If
    Next
End Sub

Private Sub cmdReloadSeriais_Click()
    Dim I As Long
    Call LoadSerials
    Call TextAdd("All Serials reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            Call SendSerial(I)
        End If
    Next
End Sub

Private Sub Command1_Click()
    frmConfiguration.Show vbModeless, frmServer
End Sub

Private Sub Command2_Click()
    Dim I As Long
    Call LoadConjuntos
    Call TextAdd("All Conjuntos reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            SendConjuntos I
        End If
    Next
End Sub

Private Sub Command3_Click()
    Call StartLottery
End Sub

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
    Call AcceptConnection(Index, SocketId)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

' auth socket

Private Sub Auth_AcceptConnection(ByVal SocketId As Long)
    frmServer.AuthSocket.Close
    frmServer.AuthSocket.Accept SocketId

    Call TextAdd("Servidor de autenticação conectado...")
    
    ' Enviar dados de jogadores que foram salvos quando o autenticador estava desligado!
    Call SetStatus("## Verificando Dados dos jogadores pra enviar ao servidor de autenticação! ##")
    Call SendAllSaves
    Call SetStatus("## Enviando dados de classes... ##")
    Call Auth_ClassesData
End Sub

Private Sub AuthSocket_ConnectionRequest(ByVal requestID As Long)
    Call Auth_AcceptConnection(requestID)
End Sub

Private Sub AuthSocket_Accept(SocketId As Integer)
    Call Auth_AcceptConnection(SocketId)
End Sub

Private Sub AuthSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    frmServer.AuthSocket.Close
    frmServer.AuthSocket.Listen
    
    Call TextAdd("Authentication Server Disconnected")
End Sub

Private Sub AuthSocket_DataArrival(ByVal bytesTotal As Long)
    Auth_IncomingData bytesTotal
End Sub

Private Sub AuthSocket_Close()
    frmServer.AuthSocket.Close
    frmServer.AuthSocket.Listen
    
    Call TextAdd("Authentication Server Disconnected")
End Sub

' ********************
Private Sub chkServerLog_Click()

' if its not 0, then its true
    If Not chkServerLog.Value Then
        ServerLog = True
    End If

End Sub

Private Sub cmdReloadClasses_Click()
    Dim I As Long
    Call LoadClasses
    Call Auth_ClassesData
    Call TextAdd("All classes reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            SendClasses I
        End If
    Next
End Sub

Private Sub cmdReloadItems_Click()
    Dim I As Long
    Call LoadItems
    Call TextAdd("All items reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            SendItems I
        End If
    Next
End Sub

Private Sub cmdReloadMaps_Click()
    Dim I As Long
    Call LoadMaps
    Call TextAdd("All maps reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            PlayerWarp I, GetPlayerMap(I), GetPlayerX(I), GetPlayerY(I)
        End If
    Next
End Sub

Private Sub cmdReloadNPCs_Click()
    Dim I As Long
    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            SendNpcs I
        End If
    Next
End Sub

Private Sub cmdReloadShops_Click()
    Dim I As Long
    Call LoadShops
    Call TextAdd("All shops reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            SendShops I
        End If
    Next
End Sub

Private Sub cmdReloadSpells_Click()
    Dim I As Long
    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            SendSpells I
        End If
    Next
End Sub

Private Sub cmdReloadResources_Click()
    Dim I As Long
    Call LoadResources
    Call TextAdd("All Resources reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            SendResources I
        End If
    Next
End Sub

Private Sub cmdReloadAnimations_Click()
    Dim I As Long
    Call LoadAnimations
    Call TextAdd("All Animations reloaded.")
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            SendAnimations I
        End If
    Next
End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
    End If
    
    Call Auth_SendShutdown
End Sub

Private Sub Form_Load()
    Call UsersOnline_Start
End Sub

Private Sub Form_Resize()

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
'Set the SortKey to the Index of the ColumnHeader - 1
'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.Index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub txtMap_Change()

If Not IsNumeric(txtMap.Text) Then
    txtMap.Text = 1
End If

If txtMap.Text > MAX_MAPS Or txtMap.Text <= 0 Then
    txtMap.Text = 1
End If

Options.START_MAP = txtMap.Text
SaveOptions
End Sub

Private Sub txtMOTD_Change()
Options.MOTD = txtMOTD.Text
SaveOptions
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
        
            If chkMsgWindow.Value = NO Then
                Call GlobalMsg(txtChat.Text, BrightRed)
            Else
                Call SendMessageToAll("Mensagem do Servidor", txtChat.Text)
            End If
            
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Sub UsersOnline_Start()
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (I)

        If I < 10 Then
            frmServer.lvwInfo.ListItems(I).Text = "00" & I
        ElseIf I < 100 Then
            frmServer.lvwInfo.ListItems(I).Text = "0" & I
        Else
            frmServer.lvwInfo.ListItems(I).Text = I
        End If

        frmServer.lvwInfo.ListItems(I).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(I).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(I).SubItems(3) = vbNullString
    Next

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call AlertMsg(FindPlayer(Name), DIALOGUE_MSG_KICKED)
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        CloseSocket (FindPlayer(Name))
    End If
End Sub

Sub mnuMute_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call ToggleMute(FindPlayer(Name))
    End If
End Sub

Sub mnuBanPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call BanIndex(FindPlayer(Name))
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted administrator access.", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have had your administrator access revoked.", BrightRed)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lmsg As Long
    lmsg = X / Screen.TwipsPerPixelX

    Select Case lmsg
    Case WM_LBUTTONDBLCLK
        frmServer.WindowState = vbNormal
        frmServer.Show
        txtText.SelStart = Len(txtText.Text)
    End Select

End Sub

Private Sub txtX_Change()
If Not IsNumeric(txtX.Text) Then
    txtX.Text = 1
End If

If txtX.Text > MAX_MAPS Or txtX.Text <= 0 Then
    txtX.Text = 1
End If

Options.START_X = txtX.Text
SaveOptions
End Sub

Private Sub txtY_Change()
If Not IsNumeric(txtY.Text) Then
    txtY.Text = 1
End If

If txtY.Text > MAX_MAPS Or txtY.Text <= 0 Then
    txtY.Text = 1
End If

Options.START_Y = txtY.Text
SaveOptions
End Sub
