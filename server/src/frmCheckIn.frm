VERSION 5.00
Begin VB.Form frmCheckIn 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   11070
   ClientTop       =   5640
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Editar Recompensas de Login"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtMonth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtQuant 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salvar"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   2760
         Width           =   1095
      End
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtReward 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mês:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   345
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Quant:"
         Height          =   195
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Recomp:"
         Height          =   195
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dia:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   150
      End
   End
End
Attribute VB_Name = "frmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    scrlItem.Max = MAX_ITEMS
End Sub

Private Sub Label1_Click()
    Me.Hide
End Sub


Private Sub scrlItem_Change()
    txtReward = scrlItem
    lblItemName = vbNullString

    If txtDay > 0 And txtDay <= 30 Then
        DayReward(txtDay).ItemNum = scrlItem
    End If

    If scrlItem > 0 Then
        lblItemName = Trim$(Item(scrlItem).Name)
    End If
End Sub

Private Sub txtDay_Change()
    If Not IsNumeric(txtDay) Then
        txtDay = 1
    End If
    
    If txtDay <= 0 Then
        txtDay = 1
    End If
    
    If txtMonth > 0 And txtMonth <= UBound(MonthReward) Then
        txtDay = 1
    End If
    
    If txtDay > 30 Then
        
    End If
    End If
End Sub

Private Sub txtQuant_Change()
    If Not IsNumeric(txtQuant) Then
        txtQuant = 0
        Exit Sub
    End If

    If txtQuant < 0 Then
        txtQuant = 0
        Exit Sub
    End If
    
    If txtDay > 0 And txtDay <= 30 Then
        DayReward(txtDay).ItemQuant = txtQuant
    End If
End Sub

Private Sub txtReward_Change()
    If Not IsNumeric(txtReward) Then
        scrlItem = 0
        txtQuant = 0
        Exit Sub
    End If
    
    If txtReward <= 0 Or txtReward > MAX_ITEMS Then
        scrlItem = 0
        txtQuant = 0
        Exit Sub
    End If
    
    scrlItem = txtReward
End Sub
