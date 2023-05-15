VERSION 5.00
Begin VB.Form frmEditor_Conjuntos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Editor"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Propriedades de Conjunto"
      Height          =   4695
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   7575
      Begin VB.Frame Frame4 
         Caption         =   "Actions On Equip Full Set"
         Height          =   1095
         Left            =   120
         TabIndex        =   34
         Top             =   3360
         Width           =   7335
         Begin VB.TextBox txtMsg 
            Height          =   615
            Left            =   5160
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   360
            Width           =   2055
         End
         Begin VB.HScrollBar scrlAnim 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   35
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Msg To Player:"
            Height          =   255
            Left            =   5640
            TabIndex        =   37
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lblAnim 
            AutoSize        =   -1  'True
            Caption         =   "Anim: None"
            Height          =   180
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bonus do Conjunto"
         Height          =   2655
         Left            =   3000
         TabIndex        =   12
         Top             =   600
         Width           =   4455
         Begin VB.TextBox txtDrop 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   2760
            TabIndex        =   42
            Text            =   "0"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtExp 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   2760
            TabIndex        =   39
            Text            =   "0"
            Top             =   1125
            Width           =   975
         End
         Begin VB.CheckBox chkPercentDefense 
            Caption         =   "%"
            Height          =   375
            Left            =   3840
            TabIndex        =   32
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtDefense 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   2760
            TabIndex        =   31
            Text            =   "0"
            Top             =   765
            Width           =   975
         End
         Begin VB.TextBox txtStatBonus 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   1
            Left            =   600
            TabIndex        =   24
            Text            =   "0"
            Top             =   405
            Width           =   975
         End
         Begin VB.TextBox txtStatBonus 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   2
            Left            =   600
            TabIndex        =   23
            Text            =   "0"
            Top             =   765
            Width           =   975
         End
         Begin VB.TextBox txtStatBonus 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   3
            Left            =   600
            TabIndex        =   22
            Text            =   "0"
            Top             =   1125
            Width           =   975
         End
         Begin VB.TextBox txtStatBonus 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   4
            Left            =   600
            TabIndex        =   21
            Text            =   "0"
            Top             =   1485
            Width           =   975
         End
         Begin VB.TextBox txtStatBonus 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   5
            Left            =   600
            TabIndex        =   20
            Text            =   "0"
            Top             =   1845
            Width           =   975
         End
         Begin VB.CheckBox chkPercentDamage 
            Caption         =   "%"
            Height          =   375
            Left            =   3840
            TabIndex        =   19
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtDamage 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   2760
            TabIndex        =   18
            Text            =   "0"
            Top             =   405
            Width           =   975
         End
         Begin VB.CheckBox chkPercentStats 
            Caption         =   "%"
            Height          =   375
            Index           =   1
            Left            =   1605
            TabIndex        =   17
            Top             =   345
            Width           =   495
         End
         Begin VB.CheckBox chkPercentStats 
            Caption         =   "%"
            Height          =   375
            Index           =   2
            Left            =   1605
            TabIndex        =   16
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox chkPercentStats 
            Caption         =   "%"
            Height          =   375
            Index           =   3
            Left            =   1605
            TabIndex        =   15
            Top             =   1065
            Width           =   495
         End
         Begin VB.CheckBox chkPercentStats 
            Caption         =   "%"
            Height          =   375
            Index           =   4
            Left            =   1605
            TabIndex        =   14
            Top             =   1440
            Width           =   495
         End
         Begin VB.CheckBox chkPercentStats 
            Caption         =   "%"
            Height          =   375
            Index           =   5
            Left            =   1605
            TabIndex        =   13
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Drop:"
            Height          =   180
            Left            =   2280
            TabIndex        =   44
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   420
         End
         Begin VB.Label Label6 
            Caption         =   "%"
            Height          =   255
            Left            =   3840
            TabIndex        =   43
            Top             =   1515
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "%"
            Height          =   255
            Left            =   3840
            TabIndex        =   41
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Exp:"
            Height          =   180
            Left            =   2400
            TabIndex        =   40
            Top             =   1125
            UseMnemonic     =   0   'False
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Def:"
            Height          =   180
            Left            =   2400
            TabIndex        =   33
            Top             =   765
            UseMnemonic     =   0   'False
            Width           =   315
         End
         Begin VB.Label lblStatBonus 
            AutoSize        =   -1  'True
            Caption         =   "+ Str:"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   405
            UseMnemonic     =   0   'False
            Width           =   435
         End
         Begin VB.Label lblDamage 
            AutoSize        =   -1  'True
            Caption         =   "Dano:"
            Height          =   180
            Left            =   2280
            TabIndex        =   29
            Top             =   405
            UseMnemonic     =   0   'False
            Width           =   450
         End
         Begin VB.Label lblStatBonus 
            AutoSize        =   -1  'True
            Caption         =   "+ End:"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   765
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label lblStatBonus 
            AutoSize        =   -1  'True
            Caption         =   "+ Int:"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   1125
            UseMnemonic     =   0   'False
            Width           =   435
         End
         Begin VB.Label lblStatBonus 
            AutoSize        =   -1  'True
            Caption         =   "+ Agi:"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   26
            Top             =   1485
            UseMnemonic     =   0   'False
            Width           =   465
         End
         Begin VB.Label lblStatBonus 
            AutoSize        =   -1  'True
            Caption         =   "+ Will:"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   25
            Top             =   1845
            UseMnemonic     =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame fraDrop 
         Caption         =   "Itens do Conjunto"
         Height          =   2655
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2775
         Begin VB.CommandButton cmdAddItem 
            Caption         =   "ADD"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   2160
            Width           =   2295
         End
         Begin VB.ComboBox cmbItems 
            Height          =   300
            Left            =   240
            TabIndex        =   10
            Text            =   "No Itens"
            Top             =   1800
            Width           =   2295
         End
         Begin VB.ListBox lstItems 
            Height          =   1500
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Set List"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         ItemData        =   "frmEditor_Conjuntos.frx":0000
         Left            =   120
         List            =   "frmEditor_Conjuntos.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Conjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPercentDamage_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub

    Conjunto(EditorIndex).Bonus.DanoPercent = chkPercentDamage.Value
End Sub

Private Sub chkPercentDefense_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub

    Conjunto(EditorIndex).Bonus.DefesaPercent = chkPercentDefense.Value
End Sub

Private Sub chkPercentStats_Click(Index As Integer)
    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub

    Conjunto(EditorIndex).Bonus.Add_Stat_Percent(Index) = CByte(chkPercentStats(Index).Value)
End Sub

Private Sub cmdAddItem_Click()
    Dim X As Long, tmpIndex As Long

    ' exit out if needed
    If Not cmbItems.ListCount > 0 Then Exit Sub
    If Not lstItems.ListCount > 0 Then Exit Sub

    ' make sure it's not a clear
    If Not cmbItems.list(cmbItems.ListIndex) = "No Items" Then
        Conjunto(EditorIndex).Item(lstItems.ListIndex + 1) = cmbItems.ListIndex
    Else
        Conjunto(EditorIndex).Item(lstItems.ListIndex + 1) = 0
    End If

    ' re-load the list
    tmpIndex = lstItems.ListIndex
    lstItems.Clear
    For X = 1 To Equipment.Equipment_Count - 1
        If Conjunto(EditorIndex).Item(X) > 0 Then
            lstItems.AddItem X & ": " & Trim$(Item(Conjunto(EditorIndex).Item(X)).Name)
        Else
            lstItems.AddItem X & ": No Items"
        End If
    Next


    If tmpIndex + 1 <> Equipment.Equipment_Count - 1 Then
        lstItems.ListIndex = tmpIndex + 1
    Else
        lstItems.ListIndex = tmpIndex
    End If

End Sub

Private Sub cmdCancel_Click()
    ConjuntoEditorCancel
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub
    ClearConjunto EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Conjunto(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    ConjuntoEditorInit
End Sub

Private Sub cmdSave_Click()
    ConjuntoEditorOk
End Sub

Private Sub Form_Load()
    txtMsg.MaxLength = ACTIONMSG_LENGTH
    scrlAnim.max = MAX_ANIMATIONS
End Sub

Private Sub lstIndex_Click()
    ConjuntoEditorInit
End Sub

Private Sub scrlAnim_Change()
    If scrlAnim = 0 Or scrlAnim > MAX_ANIMATIONS Then Exit Sub

    If Trim$(Animation(scrlAnim.Value).Name) <> vbNullString Then
        lblAnim.caption = "Anim: " & Trim$(Animation(scrlAnim.Value).Name)
    End If

    Conjunto(EditorIndex).Actions.Animation = scrlAnim.Value
End Sub

Private Sub txtDamage_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub

    If Not IsNumeric(txtDamage.text) Then
        txtDamage.text = 0
    End If

    Conjunto(EditorIndex).Bonus.Dano = CLng(txtDamage.text)
End Sub

Private Sub txtDefense_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub

    If Not IsNumeric(txtDefense.text) Then
        txtDefense.text = 0
    End If

    Conjunto(EditorIndex).Bonus.Defesa = CLng(txtDefense.text)
End Sub

Private Sub txtDrop_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub

    If Not IsNumeric(txtDrop) Then
        txtDrop = 0
    End If

    If txtDrop > 100 Then
        txtDrop = 100
    End If

    If txtDrop < 0 Then
        txtDrop = 0
    End If

    Conjunto(EditorIndex).Bonus.Drop = CByte(txtDrop)
End Sub

Private Sub txtEXP_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub

    If Not IsNumeric(txtExp.text) Then
        txtExp.text = 0
    End If

    Conjunto(EditorIndex).Bonus.EXP = CLng(txtExp.text)
End Sub

Private Sub txtMsg_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub

    Conjunto(EditorIndex).Actions.Msg = Trim$(txtMsg.text)
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Conjunto(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Conjunto(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtStatBonus_Change(Index As Integer)
    If EditorIndex = 0 Or EditorIndex > MAX_CONJUNTOS Then Exit Sub

    If Not IsNumeric(txtStatBonus(Index).text) Then
        txtStatBonus(Index).text = 0
    End If

    Conjunto(EditorIndex).Bonus.Add_Stat(Index) = CLng(txtStatBonus(Index).text)
End Sub
