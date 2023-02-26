VERSION 5.00
Begin VB.Form frmCheckIn 
   Caption         =   "Editar Login Diario"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Paste Month"
      Height          =   255
      Left            =   5880
      TabIndex        =   55
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy Month"
      Height          =   255
      Left            =   5880
      TabIndex        =   54
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Paste Day"
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy Day"
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      Height          =   495
      Left            =   2760
      TabIndex        =   51
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Recompensa"
      Height          =   1095
      Left            =   120
      TabIndex        =   45
      Top             =   3000
      Width           =   7095
      Begin VB.TextBox txtQuant 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   49
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtItem 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   47
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "Item Name: None"
         Height          =   195
         Left            =   1800
         TabIndex        =   50
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Quant:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dia"
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   7095
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   30
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   29
         Left            =   6480
         TabIndex        =   43
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   28
         Left            =   5760
         TabIndex        =   42
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   27
         Left            =   5040
         TabIndex        =   41
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   26
         Left            =   4320
         TabIndex        =   40
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   25
         Left            =   3600
         TabIndex        =   39
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   24
         Left            =   2880
         TabIndex        =   38
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   23
         Left            =   2160
         TabIndex        =   37
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   22
         Left            =   1440
         TabIndex        =   36
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   21
         Left            =   720
         TabIndex        =   35
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   19
         Left            =   6480
         TabIndex        =   33
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   18
         Left            =   5760
         TabIndex        =   32
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   17
         Left            =   5040
         TabIndex        =   31
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   16
         Left            =   4320
         TabIndex        =   30
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   15
         Left            =   3600
         TabIndex        =   29
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   14
         Left            =   2880
         TabIndex        =   28
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   13
         Left            =   2160
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   12
         Left            =   1440
         TabIndex        =   26
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   11
         Left            =   720
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   9
         Left            =   6480
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   8
         Left            =   5760
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDay 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mês"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   11
         Left            =   5160
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   10
         Left            =   4080
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   9
         Left            =   3120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   8
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMonth 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Mes_Atual As Byte
Private Dia_Atual As Byte

Private Sub Command1_Click()
    SaveCheckIn
End Sub

Private Sub Command2_Click()
    CopyDayReward = MonthReward(Mes_Atual).DayReward(Dia_Atual)
End Sub

Private Sub Command3_Click()
    MonthReward(Mes_Atual).DayReward(Dia_Atual) = CopyDayReward
    optDay_Click Dia_Atual - 1
End Sub

Private Sub Command4_Click()
    CopyMonthReward = MonthReward(Mes_Atual)
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    MonthReward(Mes_Atual) = CopyMonthReward
End Sub

Private Sub optDay_Click(Index As Integer)
    Dim i As Byte

    txtItem.Enabled = True
    txtQuant.Enabled = True

    Dia_Atual = Index + 1

    txtItem = MonthReward(Mes_Atual).DayReward(Dia_Atual).ItemNum
    txtQuant = MonthReward(Mes_Atual).DayReward(Dia_Atual).ItemQuant

End Sub

Private Sub optMonth_Click(Index As Integer)
    Dim i As Byte, GetDaysInMonth As Byte
    
    Mes_Atual = Index + 1
    
    GetDaysInMonth = DaysInMonth(Year(Date), Mes_Atual)
    
    For i = 1 To 31
        If i <= GetDaysInMonth Then
            optDay(i - 1).Visible = True
            optDay(i - 1).value = False
        Else
            optDay(i - 1).Visible = False
            optDay(i - 1).value = False
        End If
    Next i
    
    ClearTexts
    
End Sub

Private Sub txtItem_Change()

    If Dia_Atual = 0 Then Exit Sub
    If Mes_Atual = 0 Then Exit Sub

    If Not IsNumeric(txtItem) Then
        txtItem = 0
    End If
    
    If txtItem = 0 Then
        lblItemName.Caption = "Item Name: None"
    End If

    If txtItem > MAX_ITEMS Then
        txtItem = MAX_ITEMS
    End If

    If txtItem < 0 Then
        txtItem = 0
    End If

    If txtItem > 0 Then
        lblItemName = "Item Name: " & Item(txtItem).Name
    End If
    
    MonthReward(Mes_Atual).DayReward(Dia_Atual).ItemNum = txtItem
End Sub

Private Sub ClearTexts()
    txtItem.Enabled = False
    txtQuant.Enabled = False
    txtItem = 0
    txtQuant = 0
End Sub

Private Sub txtQuant_Change()
    
    If Dia_Atual = 0 Then Exit Sub
    If Mes_Atual = 0 Then Exit Sub

    If Not IsNumeric(txtQuant) Then
        txtQuant = 0
    End If

    If txtQuant < 0 Then
        txtQuant = 0
    End If
    
    MonthReward(Mes_Atual).DayReward(Dia_Atual).ItemQuant = txtQuant
End Sub
