VERSION 5.00
Begin VB.Form frmEditor_Conv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversation Editor"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7710
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
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   22
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Frame fraConv 
      Caption         =   "Conversation - 1"
      Height          =   6495
      Left            =   3360
      TabIndex        =   6
      Top             =   1200
      Width           =   4215
      Begin VB.HScrollBar scrlData3 
         Height          =   255
         Left            =   1680
         Max             =   1000
         TabIndex        =   30
         Top             =   6120
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar scrlData2 
         Height          =   255
         Left            =   1680
         Max             =   1000
         TabIndex        =   28
         Top             =   5760
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar scrlData1 
         Height          =   255
         Left            =   1680
         Max             =   1000
         TabIndex        =   26
         Top             =   5400
         Value           =   1
         Width           =   2415
      End
      Begin VB.ComboBox cmbEvent 
         Height          =   315
         ItemData        =   "frmEditor_Conv.frx":0000
         Left            =   120
         List            =   "frmEditor_Conv.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   5040
         Width           =   3975
      End
      Begin VB.HScrollBar scrlConv 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   20
         Top             =   240
         Value           =   1
         Width           =   3975
      End
      Begin VB.ComboBox cmbReply 
         Height          =   315
         Index           =   4
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   4335
         Width           =   1095
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   4440
         Width           =   2775
      End
      Begin VB.ComboBox cmbReply 
         Height          =   315
         Index           =   3
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3975
         Width           =   1095
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   3960
         Width           =   2775
      End
      Begin VB.ComboBox cmbReply 
         Height          =   315
         Index           =   2
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3615
         Width           =   1095
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   2775
      End
      Begin VB.ComboBox cmbReply 
         Height          =   315
         Index           =   1
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3225
         Width           =   1095
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtConv 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label lblData3 
         AutoSize        =   -1  'True
         Caption         =   "Data3: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   6120
         UseMnemonic     =   0   'False
         Width           =   750
      End
      Begin VB.Label lblData2 
         AutoSize        =   -1  'True
         Caption         =   "Data2: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   5760
         UseMnemonic     =   0   'False
         Width           =   750
      End
      Begin VB.Label lblData1 
         AutoSize        =   -1  'True
         Caption         =   "Data1: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   5400
         UseMnemonic     =   0   'False
         Width           =   750
      End
      Begin VB.Label Label4 
         Caption         =   "Event:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Replies:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   975
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.HScrollBar scrlChatCount 
         Height          =   255
         Left            =   1680
         Max             =   100
         Min             =   1
         TabIndex        =   19
         Top             =   600
         Value           =   1
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblChatCount 
         AutoSize        =   -1  'True
         Caption         =   "Chat Count: 1"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Conversation List"
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7080
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   7800
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_Conv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curConv As Long
Public IsIniting As Boolean

Private Sub cmbEvent_Click()
    If IsIniting Then Exit Sub

    Call UpdateText

    Conv(EditorIndex).Conv(curConv).Event = cmbEvent.ListIndex
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    ClearConv EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Conv(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    ConvEditorInit
End Sub

Private Sub cmdSave_Click()
    Call ConvEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ConvEditorCancel
End Sub

Public Sub UpdateText()
    Select Case cmbEvent.ListIndex

    Case 0, 2    ' None, Bank
        ' set max values
        scrlData1.max = 1
        scrlData2.max = 1
        scrlData3.max = 1
        ' hide / unhide
        scrlData1.visible = False
        scrlData2.visible = False
        scrlData3.visible = False
        lblData1.visible = False
        lblData2.visible = False
        lblData3.visible = False

    Case 1    ' Shop
        ' set max values
        scrlData1.max = MAX_SHOPS
        scrlData2.max = 1
        scrlData3.max = 1
        ' hide / unhide
        scrlData1.visible = True
        scrlData2.visible = False
        scrlData3.visible = False
        lblData1.visible = True
        lblData2.visible = False
        lblData3.visible = False
        ' set strings
        lblData1.caption = "Shop: None"

    Case 3    ' Give Item
        ' set max values
        scrlData1.max = MAX_ITEMS
        scrlData2.max = 32000
        scrlData3.max = 1
        ' hide / unhide
        scrlData1.visible = True
        scrlData2.visible = True
        scrlData3.visible = False
        lblData1.visible = True
        lblData2.visible = True
        lblData3.visible = False
        ' set strings
        lblData1.caption = "Item: None"
        lblData2.caption = "Amount: " & scrlData2.Value

    Case 4    ' Unique
        scrlData1.max = 32000
        scrlData2.max = 32000
        scrlData3.max = 32000
        ' hide
        scrlData1.visible = True
        scrlData2.visible = True
        scrlData3.visible = True
        lblData1.visible = True
        lblData2.visible = True
        lblData3.visible = True
        ' set the strings
        lblData1.caption = "Data1: 0"
        lblData2.caption = "Data2: 0"
        lblData3.caption = "Data3: 0"
    Case 5    ' Quest
        scrlData1.max = MAX_QUESTS
        scrlData2.max = 1
        scrlData3.max = 1
        ' hide
        scrlData1.visible = True
        scrlData2.visible = False
        scrlData3.visible = False
        lblData1.visible = True
        lblData2.visible = False
        lblData3.visible = False
        ' set the strings
        lblData1.caption = "Quest: 0"

    Case 5    ' ProtectDrop
        scrlData1.max = 1
        scrlData2.max = 1
        scrlData3.max = 1
        ' hide
        scrlData1.visible = False
        scrlData2.visible = False
        scrlData3.visible = False
        lblData1.visible = False
        lblData2.visible = False
        lblData3.visible = False
    End Select

End Sub

Private Sub lstIndex_Click()
    Call ConvEditorInit
End Sub

Private Sub scrlChatCount_Change()
    Dim n As Long, i As Long

    lblChatCount.caption = "Chat Count: " & scrlChatCount.Value
    Conv(EditorIndex).chatCount = scrlChatCount.Value
    scrlConv.max = scrlChatCount.Value

    ReDim Preserve Conv(EditorIndex).Conv(1 To scrlChatCount.Value) As ConvRec

    If scrlConv.Value > scrlConv.max Then
        scrlConv.Value = scrlConv.max
    End If

    For n = 1 To 4
        cmbReply(n).Clear
        cmbReply(n).AddItem "None"

        For i = 1 To Conv(EditorIndex).chatCount
            cmbReply(n).AddItem i
        Next

    Next

    If scrlConv.Value > 0 Then scrlConv_Change
End Sub

Private Sub scrlConv_Change()
    Dim X As Long
    curConv = scrlConv.Value
    fraConv.caption = "Conversation - " & curConv

    With Conv(EditorIndex).Conv(curConv)
        txtConv.text = .Conv

        For X = 1 To 4
            If .rTarget(X) > Conv(EditorIndex).chatCount Then .rTarget(X) = 0
            txtReply(X).text = .rText(X)
            cmbReply(X).ListIndex = .rTarget(X)
        Next

        cmbEvent.ListIndex = .Event
        scrlData1.Value = .Data1
        scrlData2.Value = .Data2
        scrlData3.Value = .Data3
    End With

End Sub

Private Sub scrlData1_Change()

    Select Case cmbEvent.ListIndex

    Case 1    ' shop

        If scrlData1.Value > 0 Then
            lblData1.caption = "Shop: " & Trim$(Shop(scrlData1.Value).Name)
        Else
            lblData1.caption = "Shop: None"
        End If

    Case 3    ' Give item

        If scrlData1.Value > 0 Then
            lblData1.caption = "Item: " & Trim$(Shop(scrlData1.Value).Name)
        Else
            lblData1.caption = "Item: None"
        End If

    Case 4    ' Unique
        lblData1.caption = "Data1: " & scrlData1.Value
    Case 5    ' Quest
        If scrlData1.Value > 0 Then
            lblData1.caption = "Quest: " & scrlData1.Value & "-" & Trim$(Quest(scrlData1.Value).Name)
        Else
            lblData1.caption = "Quest: 0"
        End If
    End Select

    Conv(EditorIndex).Conv(curConv).Data1 = scrlData1.Value
End Sub

Private Sub scrlData2_Change()

    Select Case cmbEvent.ListIndex

    Case 3    ' Give item
        lblData2.caption = "Amount: " & scrlData2.Value

    Case 4    ' Unique
        lblData1.caption = "Data2: " & scrlData2.Value
    End Select

    Conv(EditorIndex).Conv(curConv).Data2 = scrlData2.Value
End Sub

Private Sub scrlData3_Change()

    Select Case cmbEvent.ListIndex

    Case 4    ' Unique
        lblData1.caption = "Data3: " & scrlData3.Value
    End Select

    Conv(EditorIndex).Conv(curConv).Data3 = scrlData3.Value
End Sub

Private Sub txtConv_Change()
    Conv(EditorIndex).Conv(curConv).Conv = txtConv.text
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Conv(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Conv(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtReply_Change(Index As Integer)
    Conv(EditorIndex).Conv(curConv).rText(Index) = txtReply(Index).text
End Sub

Private Sub cmbReply_Click(Index As Integer)
    Conv(EditorIndex).Conv(curConv).rTarget(Index) = cmbReply(Index).ListIndex
End Sub
