VERSION 5.00
Begin VB.Form frmEditor_Resource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
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
   Icon            =   "frmEditor_Resource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   627
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   28
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resource Properties"
      Height          =   8535
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.CheckBox chkShadow 
         Caption         =   "Sombra?"
         Height          =   255
         Left            =   3480
         TabIndex        =   33
         Top             =   4800
         Width           =   1095
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   7920
         Width           =   3975
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   29
         Top             =   7560
         Width           =   4815
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   1920
         Width           =   2295
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   2400
         Left            =   2640
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   23
         Top             =   2280
         Width           =   2280
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":3332
         Left            =   960
         List            =   "frmEditor_Resource.frx":3342
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   3975
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   5160
         Width           =   4815
      End
      Begin VB.HScrollBar scrlTool 
         Height          =   255
         Left            =   120
         Max             =   3
         TabIndex        =   9
         Top             =   5760
         Width           =   4815
      End
      Begin VB.HScrollBar scrlHealth 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   8
         Top             =   6360
         Width           =   4815
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   2400
         Left            =   120
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   7
         Top             =   2280
         Width           =   2280
      End
      Begin VB.HScrollBar scrlRespawn 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   6
         Top             =   6960
         Width           =   4815
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtMessage2 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   7920
         Width           =   1455
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   7320
         Width           =   1260
      End
      Begin VB.Label lblExhaustedPic 
         AutoSize        =   -1  'True
         Caption         =   "Exhausted Image: 0"
         Height          =   180
         Left            =   2640
         TabIndex        =   25
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lblNormalPic 
         AutoSize        =   -1  'True
         Caption         =   "Normal Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1470
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         Caption         =   "Item Reward: None"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   4920
         Width           =   1440
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   5520
         Width           =   1530
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         Caption         =   "Health: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   6120
         Width           =   705
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time (Seconds): 0"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   6720
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   7800
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resource List"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkShadow_Click()
    Resource(EditorIndex).Shadow = chkShadow.Value
End Sub

Private Sub cmbType_Click()
    Resource(EditorIndex).ResourceType = cmbType.ListIndex
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    ClearResource EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    ResourceEditorInit
End Sub

Private Sub cmdSave_Click()
    Call ResourceEditorOk
End Sub

Private Sub Form_Load()
    scrlReward.max = MAX_ITEMS
End Sub

Private Sub cmdCancel_Click()
    Call ResourceEditorCancel
End Sub

Private Sub lstIndex_Click()
    ResourceEditorInit
End Sub

Private Sub scrlAnimation_Change()
    Dim SString As String

    If scrlAnimation.Value = 0 Then SString = "None" Else SString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnim.caption = "Animation: " & SString
    Resource(EditorIndex).Animation = scrlAnimation.Value
End Sub

Private Sub scrlExhaustedPic_Change()
    lblExhaustedPic.caption = "Exhausted Image: " & scrlExhaustedPic.Value
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.Value
End Sub

Private Sub scrlHealth_Change()
    lblHealth.caption = "Health: " & scrlHealth.Value
    Resource(EditorIndex).health = scrlHealth.Value
End Sub

Private Sub scrlNormalPic_Change()
    lblNormalPic.caption = "Normal Image: " & scrlNormalPic.Value
    Resource(EditorIndex).ResourceImage = scrlNormalPic.Value
End Sub

Private Sub scrlRespawn_Change()
    lblRespawn.caption = "Respawn Time (Seconds): " & scrlRespawn.Value
    Resource(EditorIndex).RespawnTime = scrlRespawn.Value
End Sub

Private Sub scrlReward_Change()

    If scrlReward.Value > 0 Then
        lblReward.caption = "Item Reward: " & Trim$(Item(scrlReward.Value).Name)
    Else
        lblReward.caption = "Item Reward: None"
    End If

    Resource(EditorIndex).ItemReward = scrlReward.Value
End Sub

Private Sub scrlTool_Change()
    Dim Name As String

    Select Case scrlTool.Value

    Case 0
        Name = "None"

    Case 1
        Name = "Hatchet"

    Case 2
        Name = "Rod"

    Case 3
        Name = "Pickaxe"
    End Select

    lblTool.caption = "Tool Required: " & Name
    Resource(EditorIndex).ToolRequired = scrlTool.Value
End Sub

Private Sub txtMessage_Change()
    Resource(EditorIndex).SuccessMessage = Trim$(txtMessage.text)
End Sub

Private Sub txtMessage2_Change()
    Resource(EditorIndex).EmptyMessage = Trim$(txtMessage2.text)
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub cmbSound_Click()

    If cmbSound.ListIndex >= 0 Then
        Resource(EditorIndex).sound = cmbSound.list(cmbSound.ListIndex)
    Else
        Resource(EditorIndex).sound = "None."
    End If

End Sub
