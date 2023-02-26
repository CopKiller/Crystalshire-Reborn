VERSION 5.00
Begin VB.Form frmEditor_Serial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serial Number Editor"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   355
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   586
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Mensagem"
      Height          =   1815
      Left            =   6840
      TabIndex        =   24
      Top             =   2880
      Width           =   1815
      Begin VB.TextBox txtMsg 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Adicionais"
      Height          =   2895
      Left            =   6840
      TabIndex        =   20
      Top             =   0
      Width           =   1815
      Begin VB.HScrollBar scrlGuildSlot 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   30
         Top             =   2400
         Width           =   1455
      End
      Begin VB.HScrollBar scrlTechnique 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   23
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtDias 
         Height          =   285
         Left            =   720
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblTecnica 
         BackStyle       =   0  'Transparent
         Caption         =   "-Nome Tecnica-"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblGuildSlot 
         Caption         =   "Aumentar Slots da Guild:"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Aprender Tecnica"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Tempo VIP"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Dias:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configurações"
      Height          =   1455
      Left            =   3240
      TabIndex        =   14
      Top             =   960
      Width           =   3495
      Begin VB.CheckBox chkBirthday 
         Caption         =   "Aniversario"
         Height          =   255
         Left            =   1920
         TabIndex        =   32
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkBlocked 
         Caption         =   "Blockeado?"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtPName 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkObtain 
         Caption         =   "Obter Uma Vez por Jogador"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Ligado ao Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraDrop 
      Caption         =   "Itens do Pacote"
      Height          =   2295
      Left            =   3240
      TabIndex        =   8
      Top             =   2400
      Width           =   3495
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   840
         TabIndex        =   12
         Text            =   "1"
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "ADD"
         Height          =   615
         Left            =   1920
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cmbItems 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Text            =   "No Itens"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ListBox lstItems 
         Height          =   1230
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Quant:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   4800
      Width           =   975
   End
   Begin VB.Frame fraName 
      Caption         =   "Init"
      Height          =   975
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   3495
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   18
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtSerial 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Nome :"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblSerial 
         Caption         =   "Serial :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Serial Code List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   4545
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Serial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkBirthday_Click()
    If EditorIndex = 0 Then Exit Sub
    
    Serial(EditorIndex).BirthDay = chkBirthday.Value
End Sub

Private Sub chkBlocked_Click()
    If EditorIndex = 0 Then Exit Sub
    
    Serial(EditorIndex).Blocked = chkBlocked.Value
End Sub

Private Sub chkObtain_Click()
    If EditorIndex = 0 Then Exit Sub
    
    Serial(EditorIndex).GiveOne = chkObtain.Value
End Sub

Private Sub cmdAddItem_Click()
    Dim tmpString() As String
    Dim X As Long, tmpIndex As Long

    ' exit out if needed
    If Not cmbItems.ListCount > 0 Then Exit Sub
    If Not lstItems.ListCount > 0 Then Exit Sub

    ' set the combo box properly
    tmpString = Split(cmbItems.list(cmbItems.ListIndex))
    ' make sure it's not a clear
    If Not cmbItems.list(cmbItems.ListIndex) = "No Items" Then
        Serial(EditorIndex).Item(lstItems.ListIndex + 1) = cmbItems.ListIndex
        Serial(EditorIndex).ItemValue(lstItems.ListIndex + 1) = txtAmount.Text
    Else
        Serial(EditorIndex).Item(lstItems.ListIndex + 1) = 0
        Serial(EditorIndex).ItemValue(lstItems.ListIndex + 1) = 0
    End If

    ' re-load the list
    tmpIndex = lstItems.ListIndex
    lstItems.Clear
    For X = 1 To MAX_SERIAL_ITEMS
        If Serial(EditorIndex).Item(X) > 0 Then
            lstItems.AddItem X & ": " & Serial(EditorIndex).ItemValue(X) & "x " & Trim$(Item(Serial(EditorIndex).Item(X)).Name)
        Else
            lstItems.AddItem X & ": No Items"
        End If
    Next X
    lstItems.ListIndex = tmpIndex
End Sub

Private Sub cmdCancel_Click()
    
    Call SerialEditorCancel

End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    If EditorIndex = 0 Or EditorIndex > MAX_SERIAL_NUMBER Then Exit Sub
    
    ClearSerial EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Trim$(Serial(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    SerialEditorInit
End Sub

Private Sub cmdSave_Click()
    
    Call SerialEditorOk

End Sub

Private Sub lstIndex_Click()
    
    SerialEditorInit

End Sub

Private Sub scrlGuildSlot_Change()
    If EditorIndex = 0 Then Exit Sub
    
    Serial(EditorIndex).GiveGuildSlot = scrlGuildSlot.Value
End Sub

Private Sub scrlTechnique_Change()
    If EditorIndex = 0 Then Exit Sub
    
    If scrlTechnique > 0 Then
        lblTecnica.caption = Trim$(Spell(scrlTechnique.Value).Name)
    Else
        lblTecnica.caption = "-Nome Tecnica-"
    End If
    
    Serial(EditorIndex).GiveSpell = scrlTechnique.Value
End Sub

Private Sub txtDias_Change()
    If EditorIndex = 0 Then Exit Sub
    
    If Not IsNumeric(txtDias.Text) Then
        txtDias.Text = vbNullString
        Exit Sub
    End If
    
    If Val(txtDias.Text) > MAX_INTEGER Or Val(txtDias.Text) < 0 Then
        txtDias.Text = vbNullString
        Exit Sub
    End If
    
    Serial(EditorIndex).VipDays = txtDias.Text
End Sub

Private Sub txtMsg_Change()
If EditorIndex = 0 Then Exit Sub

Serial(EditorIndex).Msg = Trim$(txtMsg.Text)
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Serial(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Serial(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex

End Sub

Private Sub txtPName_Change()
    If EditorIndex = 0 Then Exit Sub
    
    Serial(EditorIndex).NamePlayer = Trim$(txtPName.Text)
End Sub

Private Sub txtSerial_Change()
    If EditorIndex = 0 Then Exit Sub
    
    Serial(EditorIndex).Serial = Trim$(txtSerial.Text)
End Sub
