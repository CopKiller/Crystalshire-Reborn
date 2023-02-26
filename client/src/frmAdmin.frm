VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Painel Adminitrativo"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Editor"
      Height          =   5175
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command6 
         Caption         =   "Update Window"
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Show Window Test"
         Height          =   495
         Left            =   480
         TabIndex        =   35
         Top             =   4440
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Editar Conjuntos"
         Height          =   495
         Left            =   2040
         TabIndex        =   34
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmdAPremium 
         Caption         =   "Editar Premium"
         Height          =   495
         Left            =   2040
         TabIndex        =   32
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Editar Seriais"
         Height          =   495
         Left            =   2040
         TabIndex        =   31
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ScreenShotMap"
         Height          =   375
         Left            =   2760
         TabIndex        =   30
         Top             =   4800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CmdQuest 
         Caption         =   "Editar Quest's"
         Height          =   495
         Left            =   2040
         TabIndex        =   29
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton CmdResource 
         Caption         =   "Editar Recursos"
         Height          =   495
         Left            =   2040
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdAnimation 
         Caption         =   "Editar animações"
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton CmdShop 
         Caption         =   "Editar Lojas"
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton CmdNpc 
         Caption         =   "Editar Npc's"
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton CmdSpell 
         Caption         =   "Editar Magias"
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton CmdItem 
         Caption         =   "Editar Itens"
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton CmdMap 
         Caption         =   "Editar Mapa"
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdConv 
         Caption         =   "Editar Conversas"
         Height          =   495
         Left            =   2040
         TabIndex        =   21
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Administrator"
      Height          =   5175
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command3 
         Caption         =   "Mostrar Localização"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   4680
         Width           =   3495
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Teletransportar ao mapa"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Mudar personagem"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtASprite 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Trazer"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Expulsar"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Retirar"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "Ir Para"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdAtt 
         Caption         =   "Atualizar"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdLevel 
         Caption         =   "Evoluir de Lv."
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txtAmount 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "1"
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Dropar"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   3240
         Width           =   1335
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   1
         Top             =   3360
         Value           =   1
         Width           =   2055
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade:"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdABan_Click()
    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.Text)
End Sub

Private Sub cmdAKick_Click()
    If Len(Trim$(txtAName.Text)) < 1 Then Exit Sub

    SendKick Trim$(txtAName.Text)
End Sub

Private Sub CmdAnimation_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditAnimation
End Sub

Private Sub cmdAPremium_Click()
    ' Check Access
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    Call SendRequestEditPremium
End Sub

Private Sub cmdASpawn_Click()

    If Len(txtAmount.Text) = 0 Then Exit Sub
    If txtAmount.Text = 0 Then Exit Sub

    If scrlAItem.Value > 0 Then
        SendSpawnItem scrlAItem.Value, Trim$(txtAmount.Text)
    End If
End Sub

Private Sub cmdASprite_Click()

    If Len(Trim$(txtASprite.Text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtASprite.Text)) Then
        Exit Sub
    End If

    SendSetSprite CLng(Trim$(txtASprite.Text))

    Exit Sub
End Sub

Private Sub cmdAtt_Click()
    SendMapRespawn
End Sub

Private Sub cmdAWarp_Click()
    Dim n As Long

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then

        Exit Sub
    End If

    If Len(Trim$(txtAMap.Text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.Text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.Text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If

    ' Error handler
    Exit Sub
End Sub

Private Sub cmdAWarp2Me_Click()
    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.Text)
End Sub

Private Sub cmdAWarpMe2_Click()
    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.Text)
End Sub

Private Sub CmdConv_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditConv
End Sub

Private Sub CmdItem_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditItem
End Sub

Private Sub cmdLevel_Click()
    SendRequestLevelUp
End Sub

Private Sub CmdMap_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditMap
End Sub

Private Sub CmdNpc_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditNpc
End Sub

Private Sub CmdQuest_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
     SendRequestEditQuest
End Sub

Private Sub CmdResource_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditResource
End Sub

Private Sub cmdShop_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditShop
End Sub

Private Sub CmdSpell_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditSpell
End Sub

Private Sub Command1_Click()
    ScreenShotMap True, True, True
End Sub

Private Sub Command2_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditSerial
End Sub

Private Sub Command3_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    BLoc = Not BLoc
End Sub

Private Sub Command4_Click()
    ' Check Access
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    Call SendRequestEditConjunto
End Sub

Private Sub Command5_Click()
    ShowWindow GetWindowIndex("winCheckIn"), True
End Sub

Private Sub Command6_Click()
    Update_CheckInWindow
End Sub

Private Sub scrlAItem_Change()
    If scrlAItem.Value > 0 Then
        lblAItem.caption = " " & Trim$(Item(scrlAItem.Value).Name)
    Else
        lblAItem.caption = "Item: None"
    End If
End Sub

