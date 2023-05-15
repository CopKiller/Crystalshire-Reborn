VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
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
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Spawn"
      Height          =   735
      Left            =   3360
      TabIndex        =   50
      Top             =   5520
      Width           =   3015
      Begin VB.TextBox txtSpawnSecsMin 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   54
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   2160
         TabIndex        =   53
         Text            =   "0"
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox chkRndSpawn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Variavel"
         Height          =   180
         Left            =   120
         TabIndex        =   52
         ToolTipText     =   "Tiempo de Reaparicion (Respawn) al Azar."
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min:"
         Enabled         =   0   'False
         Height          =   180
         Left            =   1800
         TabIndex        =   56
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Secs:"
         Height          =   180
         Left            =   1200
         TabIndex        =   55
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn Rate:"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Experiência"
      Height          =   1935
      Left            =   6480
      TabIndex        =   40
      Top             =   4440
      Width           =   3015
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   48
         Text            =   "0"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtEXP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   46
         Text            =   "0"
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton opPercent_5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "5%"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton opPercent_10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "10%"
         Height          =   255
         Left            =   1080
         TabIndex        =   43
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton opPercent_20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "20%"
         Height          =   255
         Left            =   2160
         TabIndex        =   42
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkRndExp 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Experiencia Variavel"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "Experiencia a otorgar al Azar."
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   585
      End
      Begin VB.Label lblOutput 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Exp Varia: 0 - 0"
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.Frame fraDrop 
      Caption         =   "Loot"
      Height          =   2415
      Left            =   6480
      TabIndex        =   32
      Top             =   1920
      Width           =   3015
      Begin VB.ListBox lstItems 
         Height          =   1320
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cmbItems 
         Height          =   300
         Left            =   120
         TabIndex        =   36
         Text            =   "No Itens"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "ADD"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   2280
         TabIndex        =   34
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDrop 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   2280
         TabIndex        =   33
         Text            =   "0"
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Quant:"
         Height          =   255
         Left            =   1680
         TabIndex        =   39
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Chance:"
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   2040
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   6480
      TabIndex        =   31
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   375
      Left            =   7200
      TabIndex        =   30
      Top             =   6360
      Width           =   615
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell"
      Height          =   1455
      Left            =   3360
      TabIndex        =   25
      Top             =   4080
      Width           =   3015
      Begin VB.HScrollBar scrlSpellNum 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   28
         Top             =   1080
         Width           =   1695
      End
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   26
         Top             =   240
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label lblSpellNum 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblSpellName 
         Caption         =   "Spell: None"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   2775
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         X1              =   120
         X2              =   2880
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Info"
      Height          =   3975
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   3015
      Begin VB.HScrollBar scrlBalao 
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   3120
         Width           =   2775
      End
      Begin VB.CheckBox chkShadow 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sombra?"
         Height          =   180
         Left            =   1920
         TabIndex        =   67
         ToolTipText     =   "Tiempo de Reaparicion (Respawn) al Azar."
         Top             =   2400
         Width           =   975
      End
      Begin VB.HScrollBar scrlConv 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   2775
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2040
         Width           =   1695
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   2775
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2400
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   1200
         List            =   "frmEditor_NPC.frx":334E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1680
         Width           =   1695
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblBalao 
         Caption         =   "Balão:"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label lblConv 
         Caption         =   "Conv: None"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Atributos"
      Height          =   1815
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   3015
      Begin VB.HScrollBar scrlBlockChance 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         Max             =   100
         TabIndex        =   71
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtStat 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtStat 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   2
         Left            =   1080
         TabIndex        =   60
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtStat 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   3
         Left            =   2040
         TabIndex        =   59
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtStat 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   4
         Left            =   120
         TabIndex        =   58
         Text            =   "0"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtStat 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   5
         Left            =   1320
         TabIndex        =   57
         Text            =   "0"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblBlockChance 
         Caption         =   "Block Chance:"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "STR"
         Height          =   255
         Left            =   360
         TabIndex        =   66
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "RES"
         Height          =   255
         Left            =   1320
         TabIndex        =   65
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "INT"
         Height          =   255
         Left            =   2280
         TabIndex        =   64
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "AGI"
         Height          =   255
         Left            =   360
         TabIndex        =   63
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "FÉ"
         Height          =   255
         Left            =   1680
         TabIndex        =   62
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   5280
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5880
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DropIndex As Long
Private SpellIndex As Long

Private Sub chkRndExp_Click()
    Dim RangeLow As Long, RangeHigh As Long, ThisExp As Long
    NPC(EditorIndex).RandExp = chkRndExp.Value
    opPercent_5.visible = chkRndExp.Value
    opPercent_10.visible = chkRndExp.Value
    opPercent_20.visible = chkRndExp.Value
    lblOutput.visible = chkRndExp.Value

    'recheck varies text
    If Not chkRndExp.Value = vbChecked Then
        NPC(EditorIndex).Percent_5 = Abs(CInt(0))
        NPC(EditorIndex).Percent_10 = Abs(CInt(0))
        NPC(EditorIndex).Percent_20 = Abs(CInt(0))
    End If
End Sub

Private Sub chkRndSpawn_Click()
    txtSpawnSecsMin.enabled = chkRndSpawn.Value
    Label12.enabled = chkRndSpawn.Value
    NPC(EditorIndex).RndSpawn = chkRndSpawn.Value
    If Label12.enabled = True Then Label14.caption = "Max:"
    If Label12.enabled = False Then Label14.caption = "Spawn Secs:"
End Sub

Private Sub chkShadow_Click()
    NPC(EditorIndex).Shadow = chkShadow.Value
End Sub

Private Sub cmbBehaviour_Click()
    NPC(EditorIndex).Behaviour = cmbBehaviour.ListIndex
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
        NPC(EditorIndex).DropItem(lstItems.ListIndex + 1) = cmbItems.ListIndex
        NPC(EditorIndex).DropItemValue(lstItems.ListIndex + 1) = txtAmount.text
        NPC(EditorIndex).DropChance(lstItems.ListIndex + 1) = txtDrop.text
    Else
        NPC(EditorIndex).DropItem(lstItems.ListIndex + 1) = 0
        NPC(EditorIndex).DropItemValue(lstItems.ListIndex + 1) = 0
        NPC(EditorIndex).DropChance(lstItems.ListIndex + 1) = 0
    End If

    ' re-load the list
    tmpIndex = lstItems.ListIndex
    lstItems.Clear
    For X = 1 To MAX_NPC_DROPS
        If NPC(EditorIndex).DropItem(X) > 0 Then
            lstItems.AddItem X & ": " & NPC(EditorIndex).DropItemValue(X) & "x " & Trim$(Item(NPC(EditorIndex).DropItem(X)).Name) & " : 1 em " & NPC(EditorIndex).DropChance(X)
        Else
            lstItems.AddItem X & ": No Items"
        End If
    Next
    lstItems.ListIndex = tmpIndex

End Sub

Private Sub cmdCopy_Click()
    NpcEditorCopy
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    ClearNPC EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    NpcEditorInit
End Sub

Private Sub cmdPaste_Click()
    NpcEditorPaste
End Sub

Private Sub Form_Load()
    scrlSprite.max = Count_Char
    scrlAnimation.max = MAX_ANIMATIONS
    scrlConv.max = MAX_CONVS
End Sub

Private Sub opPercent_10_Click()
    Dim RangeLow As Long, RangeHigh As Long, ThisExp As Long

    NPC(EditorIndex).Percent_5 = Abs(CInt(opPercent_5.Value))
    NPC(EditorIndex).Percent_10 = Abs(CInt(opPercent_10.Value))
    NPC(EditorIndex).Percent_20 = Abs(CInt(opPercent_20.Value))

    If Not IsNumeric(txtExp.text) Then Exit Sub
    If lblOutput.visible Then
        ThisExp = CLng(txtExp.text)
        RangeLow = ThisExp - (ThisExp * 0.1)
        RangeHigh = ThisExp + (ThisExp * 0.1)
        lblOutput.caption = "Variação Exp: " & RangeLow & " - " & RangeHigh
    End If
End Sub

Private Sub opPercent_20_Click()
    Dim RangeLow As Long, RangeHigh As Long, ThisExp As Long

    NPC(EditorIndex).Percent_5 = Abs(CInt(opPercent_5.Value))
    NPC(EditorIndex).Percent_10 = Abs(CInt(opPercent_10.Value))
    NPC(EditorIndex).Percent_20 = Abs(CInt(opPercent_20.Value))

    If Not IsNumeric(txtExp.text) Then Exit Sub
    If lblOutput.visible Then
        ThisExp = CLng(txtExp.text)
        RangeLow = ThisExp - (ThisExp * 0.2)
        RangeHigh = ThisExp + (ThisExp * 0.2)
        lblOutput.caption = "Variação Exp: " & RangeLow & " - " & RangeHigh
    End If
End Sub

Private Sub opPercent_5_Click()
    Dim RangeLow As Long, RangeHigh As Long, ThisExp As Long

    NPC(EditorIndex).Percent_5 = Abs(CInt(opPercent_5.Value))
    NPC(EditorIndex).Percent_10 = Abs(CInt(opPercent_10.Value))
    NPC(EditorIndex).Percent_20 = Abs(CInt(opPercent_20.Value))

    If Not IsNumeric(txtExp.text) Then Exit Sub
    If lblOutput.visible Then
        ThisExp = CLng(txtExp.text)
        RangeLow = ThisExp - (ThisExp * 0.05)
        RangeHigh = ThisExp + (ThisExp * 0.05)
        lblOutput.caption = "Variacion Exp: " & RangeLow & " - " & RangeHigh
    End If
End Sub

Private Sub scrlBalao_Change()

    Select Case scrlBalao.Value
    Case 0
        lblBalao.caption = "Balão: Nenhum"
    Case 1
        lblBalao.caption = "Balão: !"
    Case 2
        lblBalao.caption = "Balão: ?"
    Case 3
        lblBalao.caption = "Balão: Music"
    Case 4
        lblBalao.caption = "Balão: Bravo"
    Case 5
        lblBalao.caption = "Balão: Exausto"
    Case 6
        lblBalao.caption = "Balão: Confuso"
    Case 7
        lblBalao.caption = "Balão: Digitando"
    Case 8
        lblBalao.caption = "Balão: Solução"
    Case 9
        lblBalao.caption = "Balão: Inativo"
    Case 10
        lblBalao.caption = "Balão: Cegado"
    End Select

    NPC(EditorIndex).Balao = scrlBalao.Value

    'Important = 1    ' !
    'Question      ' ?
    'Music         ' (6
    'Love          ' <3
    'Angry         ' Bravo
    'Exhausted     ' Exausto
    'Confused      ' Confuso
    'typing        ' Digitando
    'Idea          ' Solução
    'Afk           ' Inativo
    'Flashed       ' Cegado
End Sub

Private Sub scrlBlockChance_Change()
    NPC(EditorIndex).BlockChance = scrlBlockChance
    lblBlockChance = "Block Chance: " & scrlBlockChance & "%"
End Sub

Private Sub scrlConv_Change()

    If scrlConv.Value > 0 Then
        lblConv.caption = "Conv: " & Trim$(Conv(scrlConv.Value).Name)
    Else
        lblConv.caption = "Conv: None"
    End If

    NPC(EditorIndex).Conv = scrlConv.Value
End Sub

Private Sub cmdSave_Click()
    Call NpcEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub lstIndex_Click()
    NpcEditorInit
End Sub

Private Sub scrlAnimation_Change()
    Dim SString As String

    If scrlAnimation.Value = 0 Then SString = "None" Else SString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnimation.caption = "Anim: " & SString
    NPC(EditorIndex).Animation = scrlAnimation.Value
End Sub

Private Sub scrlSpell_Change()
    SpellIndex = scrlSpell.Value
    fraSpell.caption = "Spell - " & SpellIndex
    scrlSpellNum.Value = NPC(EditorIndex).Spell(SpellIndex)
End Sub

Private Sub scrlSpellNum_Change()
    lblSpellNum.caption = "Num: " & scrlSpellNum.Value

    If scrlSpellNum.Value > 0 Then
        lblSpellName.caption = "Spell: " & Trim$(Spell(scrlSpellNum.Value).Name)
    Else
        lblSpellName.caption = "Spell: None"
    End If

    NPC(EditorIndex).Spell(SpellIndex) = scrlSpellNum.Value
End Sub

Private Sub scrlSprite_Change()
    lblSprite.caption = "Sprite: " & scrlSprite.Value
    NPC(EditorIndex).Sprite = scrlSprite.Value
End Sub

Private Sub scrlRange_Change()
    lblRange.caption = "Range: " & scrlRange.Value
    NPC(EditorIndex).Range = scrlRange.Value
End Sub

Private Sub txtAttackSay_Change()
    NPC(EditorIndex).AttackSay = txtAttackSay.text
End Sub

Private Sub txtEXP_Change()

    If Not Len(txtExp.text) > 0 Then Exit Sub
    If IsNumeric(txtExp.text) Then NPC(EditorIndex).EXP = Val(txtExp.text)
End Sub

Private Sub txtLevel_Change()

    If Not Len(txtLevel.text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.text) Then NPC(EditorIndex).Level = Val(txtLevel.text)
End Sub

Public Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    NPC(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtSpawnSecs_Change()

    If Not Len(txtSpawnSecs.text) > 0 Then Exit Sub
    NPC(EditorIndex).SpawnSecs = Val(txtSpawnSecs.text)
End Sub

Private Sub cmbSound_Click()

    If cmbSound.ListIndex >= 0 Then
        NPC(EditorIndex).sound = cmbSound.list(cmbSound.ListIndex)
    Else
        NPC(EditorIndex).sound = "None."
    End If

End Sub

Private Sub txtSpawnSecsMin_Change()
    If Not Len(txtSpawnSecsMin.text) > 0 Then
        txtSpawnSecsMin.text = 0
        Exit Sub
    End If

    If Not IsNumeric(txtSpawnSecsMin.text) Then
        txtSpawnSecsMin.text = 0
        Exit Sub
    End If

    If Val(txtSpawnSecsMin.text) > MAX_LONG Or Val(txtSpawnSecsMin.text) < 0 Then
        txtSpawnSecsMin.text = 0
        Exit Sub
    End If

    NPC(EditorIndex).SpawnSecsMin = Val(txtSpawnSecsMin.text)
End Sub

Private Sub txtStat_Change(Index As Integer)
    If Not Len(txtStat(Index).text) > 0 Then
        txtStat(Index).text = 0
        Exit Sub
    End If

    If Not IsNumeric(txtStat(Index).text) Then
        txtStat(Index).text = 0
        Exit Sub
    End If

    If Val(txtStat(Index).text) > MAX_LONG Or Val(txtStat(Index).text) < 0 Then
        txtStat(Index).text = 0
        Exit Sub
    End If

    NPC(EditorIndex).Stat(Index) = txtStat(Index).text
End Sub
