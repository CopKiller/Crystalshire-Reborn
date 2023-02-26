VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Authentication Server"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ServerSocket 
      Left            =   240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtLog 
      Height          =   3975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label lblGameState 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game Server State: Disconnected"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   2940
   End
   Begin VB.Label lblCps 
      BackStyle       =   0  'Transparent
      Caption         =   "CPS : 0"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Terminate()
    DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyServer
End Sub

Private Sub ServerSocket_Close()
    lblGameState.Caption = "Game Server State: Disconnected"
    lblGameState.ForeColor = &HFF&
    If ServerSocket.State <> sckConnected Then ConnectToGameServer
End Sub

Private Sub ServerSocket_Connect()
lblGameState.Caption = "Game Server State: Connected"
lblGameState.ForeColor = &HC000&
End Sub

Private Sub ServerSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Debug.Print Number, Description, Scode
    If ServerSocket.State <> sckConnected Then ConnectToGameServer
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    AcceptConnection Index, requestID
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    IncomingData Index, bytesTotal
End Sub

Private Sub Socket_Close(Index As Integer)
    CloseSocket Index
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long
    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmMain.WindowState = vbNormal
            frmMain.Show
    End Select
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState = vbMinimized Then
        frmMain.Hide
    End If
End Sub
