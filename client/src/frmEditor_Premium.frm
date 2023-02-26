VERSION 5.00
Begin VB.Form frmEditor_Premium 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor Premium"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Sair"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdRPremium 
      Caption         =   "Retirar Premium"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdPremium 
      Caption         =   "Dar Premium"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtDPremium 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtSPremium 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtPlayer 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Nome / StartData / Dias"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmEditor_Premium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

Me.visible = False

End Sub

Private Sub cmdPremium_Click()
    'Check Access
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    
    'Check for blanks fields
    If txtPlayer.Text = vbNullString Or txtSPremium.Text = vbNullString Or txtDPremium.Text = vbNullString Then
        MsgBox ("There are blank fields, please fill out.")
        Exit Sub
    End If
    
    If Not IsNumeric(txtDPremium.Text) Then
        MsgBox "Dias inválidos, digite um número!"
        Exit Sub
    End If
    
    'If all right, go for the Premium
    If Val(txtDPremium.Text) > 999 Then
        MsgBox "Excedeu limite de dias."
        Exit Sub
    End If
    
    Call SendChangePremium(txtPlayer.Text, txtSPremium.Text, txtDPremium.Text)
End Sub

Private Sub cmdRPremium_Click()

    'Check Access
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    
    'Check for blanks fields
    If txtPlayer.Text = vbNullString Then
        MsgBox ("The name of the player is required for this operation.")
        Exit Sub
    End If
    
    'If all is right, remove the Premium
    Call SendRemovePremium(txtPlayer.Text)
End Sub

