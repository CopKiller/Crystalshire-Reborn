VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crystalshire"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2670
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   178
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picIntro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ShowWindow GetWindowIndex("winClipboard"), True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

' Janela de troca de controles!
    If Windows(GetWindowIndex("winChangeControls")).Window.visible Then
        HandleKeyCodeControls KeyCode, Shift
        Exit Sub
    End If

    ' KeyDown
    HandleKeyDown KeyCode, Shift

End Sub

' Form
Private Sub Form_Unload(Cancel As Integer)
    DestroyGame
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeyPresses(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Not InGame Then Exit Sub

    If Windows(GetWindowIndex("winChangeControls")).Window.visible Then Exit Sub

    ' handles screenshot mode
    If KeyCode = vbKeyF11 Then
        If GetPlayerAccess(MyIndex) > 0 Then
            screenshotMode = Not screenshotMode
        End If
    End If

    ' handles form
    If KeyCode = vbKeyInsert Then

        If GetPlayerAccess(MyIndex) > ADMIN_MONITOR Then
            frmAdmin.Show
        Else
            If frmMain.BorderStyle = 0 Then
                frmMain.BorderStyle = 1
            Else
                frmMain.BorderStyle = 0
            End If
            frmMain.caption = frmMain.caption
        End If
    End If

    ' handles delete events
    If KeyCode = vbKeyDelete Then
        If InMapEditor Then DeleteEvent selTileX, selTileY
    End If

    ' handles copy + pasting events
    If KeyCode = vbKeyC Then
        If ControlDown Then
            If InMapEditor Then
                CopyEvent_Map selTileX, selTileY
            End If
        End If
    End If
    If KeyCode = vbKeyV Then
        If ControlDown Then
            If InMapEditor Then
                PasteEvent_Map selTileX, selTileY
            End If
        End If
    End If
End Sub

Private Sub Form_DblClick()
    HandleGuiMouse entStates.DblClick

    ' Handle events
    If currMouseX >= 0 And currMouseX <= frmMain.ScaleWidth Then
        If currMouseY >= 0 And currMouseY <= frmMain.ScaleHeight Then
            If InMapEditor Then
                If frmEditor_Map.optEvents.Value Then
                    AddEvent CurX, CurY
                End If
            End If
        End If
    End If
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

End Sub
