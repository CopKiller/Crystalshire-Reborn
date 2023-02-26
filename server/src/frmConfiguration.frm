VERSION 5.00
Begin VB.Form frmConfiguration 
   BorderStyle     =   0  'None
   Caption         =   "Configurações Gerais do Servidor!"
   ClientHeight    =   3150
   ClientLeft      =   10875
   ClientTop       =   4830
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Bonus Gerais"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.HScrollBar scrlLottery 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   8
         Top             =   2640
         Width           =   2415
      End
      Begin VB.HScrollBar scrlPremiumDrop 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   5
         Top             =   1920
         Width           =   2415
      End
      Begin VB.HScrollBar scrlPremiumExp 
         Height          =   255
         Left            =   120
         Max             =   400
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.HScrollBar scrlPartyBonus 
         Height          =   255
         Left            =   120
         Max             =   400
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblLottery 
         AutoSize        =   -1  'True
         Caption         =   "Lottery Bonus: 0%"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   0
         Width           =   255
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   2640
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   2640
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblPremiumDrop 
         AutoSize        =   -1  'True
         Caption         =   "Premium Drop: 0%"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label lblPremiumExp 
         AutoSize        =   -1  'True
         Caption         =   "Premium Exp: 0%"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblPartyBonus 
         AutoSize        =   -1  'True
         Caption         =   "Party Members Bonus: 0%"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
    Me.Hide
End Sub

Private Sub scrlLottery_Change()
    lblLottery.Caption = "Lottery Bonus: " & scrlLottery.Value & "%"
    Options.LOTTERYBONUS = scrlLottery.Value
    SaveOptions
End Sub

Private Sub scrlPartyBonus_Change()
    lblPartyBonus.Caption = "Party Members Bonus: " & scrlPartyBonus.Value & "%"
    Options.PartyBonus = scrlPartyBonus.Value
    SaveOptions
End Sub

Private Sub scrlPremiumDrop_Change()
    lblPremiumDrop.Caption = "Premium Drop: " & scrlPremiumDrop.Value & "%"
    Options.PREMIUMDROP = scrlPremiumDrop.Value
    SaveOptions
End Sub

Private Sub scrlPremiumExp_Change()
    lblPremiumExp.Caption = "Premium Exp: " & scrlPremiumExp.Value & "%"
    Options.PREMIUMEXP = scrlPremiumExp.Value
    SaveOptions
End Sub
