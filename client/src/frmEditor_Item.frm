VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   375
      Left            =   7320
      TabIndex        =   80
      Top             =   7920
      Width           =   735
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   6480
      TabIndex        =   79
      Top             =   7920
      Width           =   735
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   3375
      Left            =   3360
      TabIndex        =   26
      Top             =   4560
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "?"
         Height          =   255
         Left            =   3600
         TabIndex        =   110
         Top             =   1320
         Width           =   255
      End
      Begin VB.Frame Frame4 
         Caption         =   "% Base Damage/Defense"
         Height          =   855
         Left            =   2160
         TabIndex        =   104
         Top             =   480
         Width           =   2175
         Begin VB.OptionButton optBase 
            Caption         =   "Null"
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   111
            Top             =   480
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optBase 
            Caption         =   "Will"
            Height          =   180
            Index           =   5
            Left            =   720
            TabIndex        =   109
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton optBase 
            Caption         =   "Agi"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   108
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton optBase 
            Caption         =   "Int"
            Height          =   180
            Index           =   3
            Left            =   1440
            TabIndex        =   107
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optBase 
            Caption         =   "End"
            Height          =   180
            Index           =   2
            Left            =   720
            TabIndex        =   106
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optBase 
            Caption         =   "Str"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.HScrollBar scrlGiveSpell 
         Height          =   255
         LargeChange     =   10
         Left            =   2160
         Max             =   1000
         TabIndex        =   101
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chkPercentStats 
         Caption         =   "%"
         Height          =   375
         Index           =   5
         Left            =   1600
         TabIndex        =   100
         Top             =   2475
         Width           =   495
      End
      Begin VB.CheckBox chkPercentStats 
         Caption         =   "%"
         Height          =   375
         Index           =   4
         Left            =   1600
         TabIndex        =   99
         Top             =   2115
         Width           =   495
      End
      Begin VB.CheckBox chkPercentStats 
         Caption         =   "%"
         Height          =   375
         Index           =   3
         Left            =   1600
         TabIndex        =   98
         Top             =   1740
         Width           =   495
      End
      Begin VB.CheckBox chkPercentStats 
         Caption         =   "%"
         Height          =   375
         Index           =   2
         Left            =   1600
         TabIndex        =   97
         Top             =   1395
         Width           =   495
      End
      Begin VB.CheckBox chkPercentStats 
         Caption         =   "%"
         Height          =   375
         Index           =   1
         Left            =   1600
         TabIndex        =   96
         Top             =   1020
         Width           =   495
      End
      Begin VB.TextBox txtDamage 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   600
         TabIndex        =   95
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkPercentDamage 
         Caption         =   "%"
         Height          =   375
         Left            =   1600
         TabIndex        =   94
         Top             =   675
         Width           =   495
      End
      Begin VB.TextBox txtStatBonus 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   5
         Left            =   600
         TabIndex        =   93
         Text            =   "0"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtStatBonus 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   4
         Left            =   600
         TabIndex        =   92
         Text            =   "0"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtStatBonus 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   3
         Left            =   600
         TabIndex        =   91
         Text            =   "0"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtStatBonus 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   2
         Left            =   600
         TabIndex        =   90
         Text            =   "0"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtStatBonus 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   1
         Left            =   600
         TabIndex        =   89
         Text            =   "0"
         Top             =   1080
         Width           =   975
      End
      Begin VB.HScrollBar scrlBlockChance 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   82
         Top             =   3015
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.HScrollBar scrlProf 
         Height          =   255
         Left            =   2160
         Max             =   2
         TabIndex        =   78
         Top             =   2040
         Width           =   1695
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   3960
         ScaleHeight     =   72
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   144
         TabIndex        =   46
         Top             =   2160
         Width           =   2160
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   2160
         TabIndex        =   45
         Top             =   2520
         Width           =   1695
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   2280
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   28
         Top             =   3000
         Value           =   100
         Width           =   1575
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3332
         Left            =   1200
         List            =   "frmEditor_Item.frx":3342
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   200
         Width           =   4935
      End
      Begin VB.Label lblGiveSpell 
         AutoSize        =   -1  'True
         Caption         =   "Give Spell: 0"
         Height          =   180
         Left            =   2160
         TabIndex        =   102
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblBlockChance 
         AutoSize        =   -1  'True
         Caption         =   "Block Chance(Shield): 0 %"
         Height          =   180
         Left            =   120
         TabIndex        =   83
         Top             =   2820
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label lblProf 
         Caption         =   "Proficiency: None"
         Height          =   255
         Left            =   2160
         TabIndex        =   77
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   2160
         TabIndex        =   44
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   2400
         TabIndex        =   36
         Top             =   2760
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will:"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   35
         Top             =   2520
         UseMnemonic     =   0   'False
         Width           =   480
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi:"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int:"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End:"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Dano:"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str:"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   3495
      Left            =   3360
      TabIndex        =   12
      Top             =   120
      Width           =   6255
      Begin VB.HScrollBar scrlChance 
         Height          =   255
         Left            =   4320
         Max             =   100
         TabIndex        =   113
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CheckBox chkDropDead 
         Caption         =   "Drop on Dead"
         Height          =   255
         Left            =   2880
         TabIndex        =   112
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   4200
         TabIndex        =   103
         Text            =   "0"
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkStackable 
         Caption         =   "Stackable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   960
         Width           =   1215
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   62
         Top             =   2760
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   60
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   19
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3363
         Left            =   4200
         List            =   "frmEditor_Item.frx":3370
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3399
         Left            =   120
         List            =   "frmEditor_Item.frx":33D0
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblChance 
         Caption         =   "0 %"
         Height          =   255
         Left            =   5520
         TabIndex        =   114
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   63
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   61
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   2880
         TabIndex        =   59
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   56
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   25
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   24
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   2880
         TabIndex        =   23
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   22
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   975
      Left            =   3360
      TabIndex        =   6
      Top             =   3600
      Width           =   6255
      Begin VB.TextBox txtStatReq 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   5
         Left            =   2640
         TabIndex        =   88
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtStatReq 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   4
         Left            =   480
         TabIndex        =   87
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtStatReq 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   3
         Left            =   4920
         TabIndex        =   86
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtStatReq 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   2
         Left            =   2640
         TabIndex        =   85
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtStatReq 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   1
         Left            =   480
         TabIndex        =   84
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str:"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   285
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End:"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int:"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   9
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   285
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi:"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   315
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will:"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   7
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   7695
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
   Begin VB.Frame fraUnique 
      Caption         =   "Unique"
      Height          =   615
      Left            =   3360
      TabIndex        =   64
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlUnique 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   65
         Top             =   240
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblUnique 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   41
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Frame fraFood 
      Caption         =   "Food"
      Height          =   3135
      Left            =   3360
      TabIndex        =   67
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlFoodInterval 
         Height          =   255
         LargeChange     =   100
         Left            =   120
         Max             =   30000
         TabIndex        =   76
         Top             =   2280
         Width           =   3375
      End
      Begin VB.HScrollBar scrlFoodTick 
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   1680
         Width           =   3375
      End
      Begin VB.HScrollBar scrlFoodHeal 
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1080
         Width           =   3375
      End
      Begin VB.OptionButton optSP 
         Caption         =   "SP"
         Height          =   255
         Left            =   840
         TabIndex        =   70
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optHP 
         Caption         =   "HP"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblFoodInterval 
         Caption         =   "Interval: 0(ms)"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label lblFoodTick 
         Caption         =   "Tick Count: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblFoodHeal 
         Caption         =   "Heal per Tick: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Heals HP or SP"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Consume Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   37
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CheckBox chkInstant 
         Caption         =   "Instant Cast?"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   53
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   51
         Top             =   1800
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   49
         Top             =   1200
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   38
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         Caption         =   "Cast Spell: None"
         Height          =   180
         Left            =   120
         TabIndex        =   54
         Top             =   2160
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   52
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   50
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDropDead_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).DropDead = chkDropDead.Value
End Sub

Private Sub chkPercentDamage_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data2_Percent = chkPercentDamage.Value
End Sub

Private Sub chkPercentStats_Click(Index As Integer)
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Stat_Percent(Index) = chkPercentStats(Index).Value
End Sub

Private Sub chkStackable_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Stackable = chkStackable.Value
End Sub

Private Sub cmbBind_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
End Sub

Private Sub cmbClassReq_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
End Sub

Private Sub cmbSound_Click()

    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).sound = cmbSound.list(cmbSound.ListIndex)
    Else
        Item(EditorIndex).sound = "None."
    End If

End Sub

Private Sub cmbTool_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
End Sub

Private Sub cmdCopy_Click()
    ItemEditorCopy
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    ClearItem EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    ItemEditorInit
End Sub

Private Sub cmdPaste_Click()
    ItemEditorPaste
End Sub

Private Sub Command1_Click()
    If scrlGiveSpell.Value > 0 And scrlGiveSpell.Value <= MAX_SPELLS Then
        MsgBox "Spell Name: " & Trim$(Spell(scrlGiveSpell.Value).Name)
    End If
End Sub

Private Sub Form_Load()
    scrlPic.max = Count_Item
    scrlAnim.max = MAX_ANIMATIONS
    scrlLevelReq.max = MAX_LEVELS
    scrlPaperdoll.max = Count_Paperdoll
End Sub

Private Sub cmdSave_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_RINGRIGHT) Then
        fraEquipment.visible = True

        'Att a label
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            lblDamage.caption = "Dano:"
        Else
            lblDamage.caption = "Def:"
        End If
    Else
        fraEquipment.visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SHIELD) Then
        scrlBlockChance.visible = True
        lblBlockChance.visible = True
    Else
        scrlBlockChance.visible = False
        lblBlockChance.visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.visible = True
    Else
        fraSpell.visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_UNIQUE Then
        fraUnique.visible = True
    Else
        fraUnique.visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_FOOD Then
        fraFood.visible = True
    Else
        fraFood.visible = False
    End If

    Item(EditorIndex).Type = cmbType.ListIndex
End Sub

Private Sub lstIndex_Click()
    ItemEditorInit
End Sub

Private Sub optBase_Click(Index As Integer)
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).AtributeBase = Index
End Sub

Private Sub optHP_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).HPorSP = 1    ' hp
End Sub

Private Sub optSP_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).HPorSP = 2    ' sp
End Sub

Private Sub scrlAccessReq_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.caption = "Access Req: " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
End Sub

Private Sub scrlAddHp_Change()
    lblAddHP.caption = "Add HP: " & scrlAddHp.Value
    Item(EditorIndex).AddHP = scrlAddHp.Value
End Sub

Private Sub scrlAddMp_Change()
    lblAddMP.caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
End Sub

Private Sub scrlAddExp_Change()
    lblAddExp.caption = "Add Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddEXP = scrlAddExp.Value
End Sub

Private Sub scrlAnim_Change()
    Dim SString As String

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.Value = 0 Then
        SString = "None"
    Else
        SString = Trim$(Animation(scrlAnim.Value).Name)
    End If

    lblAnim.caption = "Anim: " & SString
    Item(EditorIndex).Animation = scrlAnim.Value
End Sub

Private Sub scrlBlockChance_Change()
    lblBlockChance.caption = "Block Chance(Shield): " & scrlBlockChance.Value & " %"
    Item(EditorIndex).BlockChance = scrlBlockChance.Value
End Sub

Private Sub scrlChance_Change()
lblChance.caption = scrlChance.Value & " %"
Item(EditorIndex).DropDeadChance = scrlChance.Value
End Sub

Private Sub scrlFoodHeal_Change()
    lblFoodHeal.caption = "Heal Per Tick: " & scrlFoodHeal.Value
    Item(EditorIndex).FoodPerTick = scrlFoodHeal.Value
End Sub

Private Sub scrlFoodInterval_Change()
    lblFoodInterval.caption = "Interval: " & scrlFoodInterval.Value & "(ms)"
    Item(EditorIndex).FoodInterval = scrlFoodInterval.Value
End Sub

Private Sub scrlFoodTick_Change()
    lblFoodTick.caption = "Tick Count: " & scrlFoodTick.Value
    Item(EditorIndex).FoodTickCount = scrlFoodTick.Value
End Sub

Private Sub scrlGiveSpell_Change()
    lblGiveSpell.caption = "Give Spell: " & scrlGiveSpell.Value
    Item(EditorIndex).GiveSpellNum = scrlGiveSpell.Value
End Sub

Private Sub scrlLevelReq_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
End Sub

Private Sub scrlPaperdoll_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.Value
End Sub

Private Sub scrlPic_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.caption = "Pic: " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value
End Sub

Private Sub scrlProf_Change()
    Dim theProf As String

    Select Case scrlProf.Value

    Case 0    ' None
        theProf = "None"

    Case 1    ' Sword/Armour
        theProf = "Sword/Armour"

    Case 2    ' Staff/Cloth
        theProf = "Staff/Cloth"
    End Select

    lblProf.caption = "Proficiency: " & theProf
    Item(EditorIndex).proficiency = scrlProf.Value
End Sub

Private Sub scrlRarity_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.caption = "Rarity: " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
End Sub

Private Sub scrlSpeed_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.caption = "Speed: " & scrlSpeed.Value / 1000 & " sec"
    Item(EditorIndex).speed = scrlSpeed.Value
End Sub

Private Sub txtDamage_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If Not Len(txtDamage.Text) > 0 Then
        txtDamage.Text = 0
        Exit Sub
    End If

    If Not IsNumeric(txtDamage.Text) Then
        txtDamage.Text = 0
        Exit Sub
    End If

    If Val(txtDamage.Text) > MAX_LONG Or Val(txtDamage.Text) < 0 Then
        txtDamage.Text = 0
        Exit Sub
    End If

    Item(EditorIndex).Data2 = txtDamage.Text

End Sub

Private Sub txtPrice_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If Not Len(txtPrice.Text) > 0 Then
        txtPrice.Text = 0
        Exit Sub
    End If

    If Not IsNumeric(txtPrice.Text) Then
        txtPrice.Text = 0
        Exit Sub
    End If

    If Val(txtPrice.Text) > MAX_LONG Or Val(txtPrice.Text) < 0 Then
        txtPrice.Text = 0
        Exit Sub
    End If

    Item(EditorIndex).Price = txtPrice.Text
End Sub

Private Sub txtStatBonus_Change(Index As Integer)

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If Not Len(txtStatBonus(Index).Text) > 0 Then
        txtStatBonus(Index).Text = 0
        Exit Sub
    End If

    If Not IsNumeric(txtStatBonus(Index).Text) Then
        txtStatBonus(Index).Text = 0
        Exit Sub
    End If

    If Val(txtStatBonus(Index).Text) > MAX_LONG Or Val(txtStatBonus(Index).Text) < 0 Then
        txtStatBonus(Index).Text = 0
        Exit Sub
    End If

    Item(EditorIndex).Add_Stat(Index) = txtStatBonus(Index).Text
End Sub

Private Sub scrlSpell_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If Len(Trim$(Spell(scrlSpell.Value).Name)) > 0 Then
        lblSpellName.caption = "Name: " & Trim$(Spell(scrlSpell.Value).Name)
    Else
        lblSpellName.caption = "Name: None"
    End If

    lblSpell.caption = "Spell: " & scrlSpell.Value
    Item(EditorIndex).Data1 = scrlSpell.Value
End Sub

Private Sub scrlUnique_Change()
    lblUnique.caption = "Num: " & scrlUnique.Value
    Item(EditorIndex).Data1 = scrlUnique.Value
End Sub

Private Sub txtDesc_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Desc = txtDesc.Text
End Sub

Public Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtStatReq_Change(Index As Integer)

    If Not Len(txtStatReq(Index).Text) > 0 Then
        txtStatReq(Index).Text = 0
        Exit Sub
    End If

    If Not IsNumeric(txtStatReq(Index).Text) Then
        txtStatReq(Index).Text = 0
        Exit Sub
    End If

    If Val(txtStatReq(Index).Text) > MAX_LONG Or Val(txtStatReq(Index).Text) < 0 Then
        txtStatReq(Index).Text = 0
        Exit Sub
    End If

    Item(EditorIndex).Stat_Req(Index) = txtStatReq(Index).Text
End Sub
