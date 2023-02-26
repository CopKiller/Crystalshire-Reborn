VERSION 5.00
Begin VB.Form frmEditor_MapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
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
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "Weather"
      Height          =   1575
      Left            =   120
      TabIndex        =   50
      Top             =   5880
      Width           =   2055
      Begin VB.ComboBox CmbWeather 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0000
         Left            =   120
         List            =   "frmMapProperties.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   240
         Width           =   1815
      End
      Begin VB.HScrollBar scrlWeatherIntensity 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   52
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   1920
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblWeatherIntensity 
         Caption         =   "Power: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Boss"
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   4800
      Width           =   2055
      Begin VB.HScrollBar scrlBoss 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblBoss 
         Caption         =   "Boss: None"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Effects"
      Height          =   2895
      Left            =   2280
      TabIndex        =   35
      Top             =   5040
      Width           =   4215
      Begin VB.HScrollBar scrlSun 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   58
         Top             =   2400
         Width           =   1695
      End
      Begin VB.HScrollBar scrlPanorama 
         Height          =   255
         Left            =   2040
         Max             =   20
         TabIndex        =   53
         Top             =   1920
         Width           =   2055
      End
      Begin VB.HScrollBar scrlA 
         Height          =   255
         Left            =   3120
         Max             =   255
         TabIndex        =   49
         Top             =   1320
         Width           =   975
      End
      Begin VB.HScrollBar ScrlB 
         Height          =   255
         Left            =   3120
         Max             =   255
         TabIndex        =   47
         Top             =   960
         Width           =   975
      End
      Begin VB.HScrollBar ScrlG 
         Height          =   255
         Left            =   3120
         Max             =   255
         TabIndex        =   45
         Top             =   600
         Width           =   975
      End
      Begin VB.HScrollBar ScrlR 
         Height          =   255
         Left            =   3120
         Max             =   255
         TabIndex        =   43
         Top             =   240
         Width           =   975
      End
      Begin VB.HScrollBar ScrlFogSpeed 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   41
         Top             =   1080
         Width           =   1695
      End
      Begin VB.HScrollBar scrlFogOpacity 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   39
         Top             =   1680
         Width           =   1695
      End
      Begin VB.HScrollBar ScrlFog 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   37
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblSun 
         AutoSize        =   -1  'True
         Caption         =   "Sun: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   2160
         Width           =   570
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   1920
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000000&
         X1              =   1920
         X2              =   1920
         Y1              =   120
         Y2              =   2400
      End
      Begin VB.Label lblPanorama 
         Caption         =   "Panorama:"
         Height          =   255
         Left            =   2040
         TabIndex        =   54
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Opacity: 0"
         Height          =   195
         Left            =   2040
         TabIndex        =   48
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label lblB 
         AutoSize        =   -1  'True
         Caption         =   "Blue: 0"
         Height          =   195
         Left            =   2040
         TabIndex        =   46
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblG 
         AutoSize        =   -1  'True
         Caption         =   "Green: 0"
         Height          =   195
         Left            =   2040
         TabIndex        =   44
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblR 
         AutoSize        =   -1  'True
         Caption         =   "Red: 0"
         Height          =   195
         Left            =   2040
         TabIndex        =   42
         Top             =   240
         Width           =   570
      End
      Begin VB.Label lblFogSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Fog Speed:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblFogOpacity 
         AutoSize        =   -1  'True
         Caption         =   "Fog Opacity: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label lblFog 
         Caption         =   "Fog: None"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Music"
      Height          =   3255
      Left            =   4440
      TabIndex        =   27
      Top             =   1800
      Width           =   2055
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ListBox lstMusic 
         Height          =   2205
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmMaxSizes 
      Caption         =   "Max Sizes"
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   2055
      Begin VB.TextBox txtMaxX 
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Max X:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Max Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   2055
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   20
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   19
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   18
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Current map: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      Height          =   1215
      Left            =   2280
      TabIndex        =   13
      Top             =   480
      Width           =   4215
      Begin VB.ComboBox cmbDayNight 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0045
         Left            =   1080
         List            =   "frmMapProperties.frx":0052
         TabIndex        =   56
         Text            =   "cmbDayNight"
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0078
         Left            =   1080
         List            =   "frmMapProperties.frx":0085
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Day/Night:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moral:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Boot Map:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Boot X:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Boot Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   630
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      Height          =   3255
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
      Begin VB.ListBox lstNpcs 
         Height          =   2400
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEditor_MapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPlay_Click()
    Stop_Music
    Play_Music lstMusic.list(lstMusic.ListIndex)
End Sub

Private Sub cmdStop_Click()
    Stop_Music
End Sub

Private Sub cmdOk_Click()
    Dim X As Long, X2 As Long
    Dim Y As Long, Y2 As Long
    Dim tempArr() As TileRec

    If Not IsNumeric(txtMaxX.Text) Then txtMaxX.Text = Map.MapData.MaxX
    If Val(txtMaxX.Text) < 1 Then txtMaxX.Text = 1
    If Val(txtMaxX.Text) > MAX_BYTE Then txtMaxX.Text = MAX_BYTE
    If Not IsNumeric(txtMaxY.Text) Then txtMaxY.Text = Map.MapData.MaxY
    If Val(txtMaxY.Text) < 1 Then txtMaxY.Text = 1
    If Val(txtMaxY.Text) > MAX_BYTE Then txtMaxY.Text = MAX_BYTE

    With Map.MapData
        .Name = Trim$(txtName.Text)

        If lstMusic.ListIndex >= 0 Then
            .Music = lstMusic.list(lstMusic.ListIndex)
        Else
            .Music = vbNullString
        End If

        .Up = Val(txtUp.Text)
        .Down = Val(txtDown.Text)
        .Left = Val(txtLeft.Text)
        .Right = Val(txtRight.Text)
        .Moral = cmbMoral.ListIndex
        .BootMap = Val(txtBootMap.Text)
        .BootX = Val(txtBootX.Text)
        .BootY = Val(txtBootY.Text)
        .BossNpc = scrlBoss.Value
        .Panorama = scrlPanorama.Value
        
        .Weather = CmbWeather.ListIndex
        .WeatherIntensity = scrlWeatherIntensity.Value
        
        .Fog = ScrlFog.Value
        .FogSpeed = ScrlFogSpeed.Value
        .FogOpacity = scrlFogOpacity.Value
        
        .Red = scrlR.Value
        .Green = scrlG.Value
        .Blue = scrlB.Value
        .Alpha = scrlA.Value
        
        .Sun = scrlSun.Value
        
        .DayNight = cmbDayNight.ListIndex
        
        ' set the data before changing it
        tempArr = Map.TileData.Tile
        X2 = Map.MapData.MaxX
        Y2 = Map.MapData.MaxY
        ' change the data
        .MaxX = Val(txtMaxX.Text)
        .MaxY = Val(txtMaxY.Text)

        If X2 > .MaxX Then X2 = .MaxX
        If Y2 > .MaxY Then Y2 = .MaxY
        ' redim the map size
        ReDim Map.TileData.Tile(0 To .MaxX, 0 To .MaxY)

        For X = 0 To X2
            For Y = 0 To Y2
                Map.TileData.Tile(X, Y) = tempArr(X, Y)
            Next
        Next

    End With

    ' cache the shit
    initAutotiles
    Unload frmEditor_MapProperties
    ClearTempTile
End Sub

Private Sub cmdCancel_Click()
    Unload frmEditor_MapProperties
End Sub

Private Sub lstNpcs_Click()
    Dim tmpString() As String
    Dim NpcNum As Long

    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    ' set the combo box properly
    tmpString = Split(lstNpcs.list(lstNpcs.ListIndex))
    NpcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
    cmbNpc.ListIndex = Map.MapData.NPC(NpcNum)
End Sub

Private Sub cmbNpc_Click()
    Dim tmpString() As String
    Dim NpcNum As Long
    Dim X As Long, tmpIndex As Long

    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    ' set the combo box properly
    tmpString = Split(cmbNpc.list(cmbNpc.ListIndex))

    ' make sure it's not a clear
    If Not cmbNpc.list(cmbNpc.ListIndex) = "No NPC" Then
        NpcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        Map.MapData.NPC(lstNpcs.ListIndex + 1) = NpcNum
    Else
        Map.MapData.NPC(lstNpcs.ListIndex + 1) = 0
    End If

    ' re-load the list
    tmpIndex = lstNpcs.ListIndex
    lstNpcs.Clear

    For X = 1 To MAX_MAP_NPCS

        If Map.MapData.NPC(X) > 0 Then
            lstNpcs.AddItem X & ": " & Trim$(NPC(Map.MapData.NPC(X)).Name)
        Else
            lstNpcs.AddItem X & ": No NPC"
        End If

    Next

    lstNpcs.ListIndex = tmpIndex
End Sub

Private Sub scrlA_Change()
    lblA.caption = "Opacity: " & scrlA.Value
End Sub

Private Sub scrlB_Change()
    lblB.caption = "Blue: " & scrlB.Value
End Sub

Private Sub scrlBoss_Change()

    If scrlBoss.Value > 0 Then
        If Map.MapData.NPC(scrlBoss.Value) > 0 Then
            lblBoss.caption = "Boss Npc: " & Trim$(NPC(Map.MapData.NPC(scrlBoss.Value)).Name)
        Else
            lblBoss.caption = "Boss Npc: None"
        End If
    Else
        lblBoss.caption = "Boss Npc: None"
    End If

End Sub

Private Sub ScrlFog_Change()
    If ScrlFog.Value = 0 Then
        lblFog.caption = "None."
    Else
        lblFog.caption = "Fog: " & ScrlFog.Value
    End If
End Sub

Private Sub scrlFogOpacity_Change()
    lblFogOpacity.caption = "Fog Opacity: " & scrlFogOpacity.Value
End Sub

Private Sub ScrlFogSpeed_Change()
    lblFogSpeed.caption = "Fog Speed: " & ScrlFogSpeed.Value
End Sub

Private Sub scrlG_Change()
    lblG.caption = "Green: " & scrlG.Value
End Sub

Private Sub scrlPanorama_Change()
    lblPanorama.caption = "Panorama: " & scrlPanorama.Value
End Sub

Private Sub scrlR_Change()
    lblR.caption = "Red: " & scrlR.Value
End Sub

Private Sub scrlSun_Change()
    If scrlSun > 0 Then
        lblSun = "Sun: " & scrlSun
    Else
        lblSun = "Sub: None"
    End If
End Sub

Private Sub scrlWeatherIntensity_Change()
    lblWeatherIntensity.caption = "Intensity: " & scrlWeatherIntensity.Value
End Sub
