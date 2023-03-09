Attribute VB_Name = "modWindow_Character"
Option Explicit

Public Sub CreateWindow_Character()
Dim i As Byte, X As Long, Y As Long
Dim Coluns_MaxLines As Byte
Dim Coluns_MaxBonus As Byte

' Create window
    CreateWindow "winCharacter", "Character Status", zOrder_Win, 0, 0, 170, 382, Tex_Item(62), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' RENDER IN ALL TABS COUNT 1 TO 5
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Char)
    ' Parchment and wood background
    CreatePictureBox WindowCount, "picWood", 0, 21, 170, 360, , , , , , , , DesignTypes.desWood, DesignTypes.desWood, DesignTypes.desWood
    CreatePictureBox WindowCount, "picParchment", 6, 46, 162, 327, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Window Tab CheckBoxes Atributes/Equips
    CreateCheckbox WindowCount, "chkAtributos", 7, 26, 79, 20, 1, "Atributos", rockwellDec_10, , , , , DesignTypes.desChkCharacter, , , GetAddress(AddressOf chkCharacters_Atributos)
    CreateCheckbox WindowCount, "chkEquipamentos", 84, 26, 79, 20, 0, "Armaduras", rockwellDec_10, , , , , DesignTypes.desChkCharacter, , , GetAddress(AddressOf chkCharacters_Equipamentos)

    ' RENDER EQUIPMENT WINDOW COUNT 6 TO 27
    ' Equipamentos
    CreatePictureBox WindowCount, "picEquips", 18, 59, 138, 9, False, , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblEquips", 18, 56, 138, , "Equipamentos", rockwellDec_15, , Alignment.alignCentre, False
    CreatePictureBox WindowCount, "picBoxEquip1", 69, 79, 34, 34, False, , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_DblClick)
    CreatePictureBox WindowCount, "picBoxEquip2", 69, 131, 34, 34, False, , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_DblClick)
    CreatePictureBox WindowCount, "picBoxEquip3", 69, 183, 34, 34, False, , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_DblClick)
    CreatePictureBox WindowCount, "picBoxEquip4", 69, 235, 34, 34, False, , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_DblClick)
    CreatePictureBox WindowCount, "picBoxEquip5", 17, 131, 34, 34, False, , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_DblClick)
    CreatePictureBox WindowCount, "picBoxEquip6", 121, 131, 34, 34, False, , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_DblClick)
    CreatePictureBox WindowCount, "picBoxEquip7", 17, 79, 34, 34, False, , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_DblClick)
    CreatePictureBox WindowCount, "picBoxEquip8", 17, 183, 34, 34, False, , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_DblClick)
    CreatePictureBox WindowCount, "picBoxEquip9", 121, 183, 34, 34, False, , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_DblClick), GetAddress(AddressOf DrawCharacter)
    ' Bonus Set
    CreatePictureBox WindowCount, "picTextBackground", 18, 273, 138, 9, False, , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblSet", 18, 270, 138, , "Bonus de Conjunto", rockwellDec_15, , Alignment.alignCentre, False
    
    ' INIT Bonus Window
    X = 18
    Y = 270
    Coluns_MaxBonus = 9
    Coluns_MaxLines = 6
    ' LOGIC Bonus Window
    For i = 1 To Coluns_MaxBonus
        Y = Y + 14
        If i = Coluns_MaxLines Then Y = 284: X = 90
        CreateLabel WindowCount, "lblBonus" & i, X, Y, 138, , "Bonus" & i, rockwellDec_15, , Alignment.alignLeft, False
    Next i
    
    ' RENDER ATRIBUTES WINDOW 28 TO MAX .CONTROLSCOUNT
    ' White boxes
    CreatePictureBox WindowCount, "picWhiteBox", 13, 54, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 74, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 94, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 114, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 134, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 154, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 174, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 194, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 214, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    ' Labels
    CreateLabel WindowCount, "lblName", 18, 56, 147, 16, "Name", rockwellDec_10
    CreateLabel WindowCount, "lblClass", 18, 76, 147, 16, "Class", rockwellDec_10
    CreateLabel WindowCount, "lblLevel", 18, 96, 147, 16, "Level", rockwellDec_10
    CreateLabel WindowCount, "lblGuild", 18, 116, 147, 16, "Guild", rockwellDec_10
    CreateLabel WindowCount, "lblHealth", 18, 136, 147, 16, "Health", rockwellDec_10
    CreateLabel WindowCount, "lblSpirit", 18, 156, 147, 16, "Spirit", rockwellDec_10
    CreateLabel WindowCount, "lblExperience", 18, 176, 147, 16, "Exp", rockwellDec_10
    CreateLabel WindowCount, "lblVip", 18, 196, 147, 16, "Vip", rockwellDec_10
    CreateLabel WindowCount, "lblVipD", 18, 216, 147, 16, "Days", rockwellDec_10
    ' Attributes
    CreatePictureBox WindowCount, "picShadow", 18, 236, 138, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblLabel", 18, 233, 138, , "Character Attributes", rockwellDec_15, , Alignment.alignCentre
    ' Black boxes
    CreatePictureBox WindowCount, "picBlackBox", 13, 246, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    CreatePictureBox WindowCount, "picBlackBox", 13, 266, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    CreatePictureBox WindowCount, "picBlackBox", 13, 286, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    CreatePictureBox WindowCount, "picBlackBox", 13, 306, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    CreatePictureBox WindowCount, "picBlackBox", 13, 326, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    CreatePictureBox WindowCount, "picBlackBox", 13, 346, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    ' Labels
    CreateLabel WindowCount, "lblLabel", 18, 248, 138, , "Strength", rockwellDec_10, Gold, Alignment.alignRight
    CreateLabel WindowCount, "lblLabel", 18, 268, 138, , "Endurance", rockwellDec_10, Gold, Alignment.alignRight
    CreateLabel WindowCount, "lblLabel", 18, 288, 138, , "Intelligence", rockwellDec_10, Gold, Alignment.alignRight
    CreateLabel WindowCount, "lblLabel", 18, 308, 138, , "Agility", rockwellDec_10, Gold, Alignment.alignRight
    CreateLabel WindowCount, "lblLabel", 18, 328, 138, , "Willpower", rockwellDec_10, Gold, Alignment.alignRight
    CreateLabel WindowCount, "lblLabel", 18, 348, 138, , "Unused Stat Points", rockwellDec_10, LightGreen, Alignment.alignRight
    ' Buttons
    CreateButton WindowCount, "btnStat_1", 15, 248, 15, 15, , , , , , , Tex_GUI(48), Tex_GUI(49), Tex_GUI(50), , , , , , GetAddress(AddressOf Character_SpendPoint1)
    CreateButton WindowCount, "btnStat_2", 15, 268, 15, 15, , , , , , , Tex_GUI(48), Tex_GUI(49), Tex_GUI(50), , , , , , GetAddress(AddressOf Character_SpendPoint2)
    CreateButton WindowCount, "btnStat_3", 15, 288, 15, 15, , , , , , , Tex_GUI(48), Tex_GUI(49), Tex_GUI(50), , , , , , GetAddress(AddressOf Character_SpendPoint3)
    CreateButton WindowCount, "btnStat_4", 15, 308, 15, 15, , , , , , , Tex_GUI(48), Tex_GUI(49), Tex_GUI(50), , , , , , GetAddress(AddressOf Character_SpendPoint4)
    CreateButton WindowCount, "btnStat_5", 15, 328, 15, 15, , , , , , , Tex_GUI(48), Tex_GUI(49), Tex_GUI(50), , , , , , GetAddress(AddressOf Character_SpendPoint5)
    ' fake buttons
    CreatePictureBox WindowCount, "btnGreyStat_1", 15, 248, 15, 15, , , , , Tex_GUI(47), Tex_GUI(47), Tex_GUI(47)
    CreatePictureBox WindowCount, "btnGreyStat_2", 15, 268, 15, 15, , , , , Tex_GUI(47), Tex_GUI(47), Tex_GUI(47)
    CreatePictureBox WindowCount, "btnGreyStat_3", 15, 288, 15, 15, , , , , Tex_GUI(47), Tex_GUI(47), Tex_GUI(47)
    CreatePictureBox WindowCount, "btnGreyStat_4", 15, 308, 15, 15, , , , , Tex_GUI(47), Tex_GUI(47), Tex_GUI(47)
    CreatePictureBox WindowCount, "btnGreyStat_5", 15, 328, 15, 15, , , , , Tex_GUI(47), Tex_GUI(47), Tex_GUI(47)
    ' Labels
    CreateLabel WindowCount, "lblStat_1", 32, 248, 100, , "255", rockwellDec_10
    CreateLabel WindowCount, "lblStat_2", 32, 268, 100, , "255", rockwellDec_10
    CreateLabel WindowCount, "lblStat_3", 32, 288, 100, , "255", rockwellDec_10
    CreateLabel WindowCount, "lblStat_4", 32, 308, 100, , "255", rockwellDec_10
    CreateLabel WindowCount, "lblStat_5", 32, 328, 100, , "255", rockwellDec_10
    CreateLabel WindowCount, "lblPoints", 18, 348, 100, , "255", rockwellDec_10
End Sub

Public Sub DrawCharacter()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, i As Long, Sprite As Long, itemNum As Long, ItemPic As Long, Value As Byte
    Dim Colour As Long, rec As RECT

    'Chk Equipamentos non checked? exit
    If Windows(GetWindowIndex("winCharacter")).Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = 0 Then Exit Sub

    ' loop through equipment
    For i = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(MyIndex, i)

        ' Calc the position of gui itembox null
        If i > Equipment.Shield Then
            Value = i + 26
        Else
            Value = i
        End If

        ' get the item sprite
        If itemNum > 0 Then
            ItemPic = Tex_Item(Item(itemNum).Pic)
        Else
            ' no item equiped - use blank image
            ItemPic = Tex_GUI(37 + Value)
        End If

        yO = Windows(GetWindowIndex("winCharacter")).Window.top
        xO = Windows(GetWindowIndex("winCharacter")).Window.Left

        Select Case i
        Case Equipment.Helmet
            xO = xO + 70
            yO = yO + 80
        Case Equipment.Armor
            xO = xO + 70
            yO = yO + 132
        Case Equipment.Legs
            xO = xO + 70
            yO = yO + 184
        Case Equipment.Boots
            xO = xO + 70
            yO = yO + 236
        Case Equipment.Weapon
            xO = xO + 18
            yO = yO + 132
        Case Equipment.Shield
            xO = xO + 122
            yO = yO + 132
        Case Equipment.Amulet
            xO = xO + 18
            yO = yO + 80
        Case Equipment.RingLeft
            xO = xO + 18
            yO = yO + 184
        Case Equipment.RingRight
            xO = xO + 122
            yO = yO + 184
        End Select

        ' Default color
        Colour = -1
        If DragBox.Origin = origin_Inventory Then
            If DragBox.Type = Part_Item Then
                If GetPlayerInvItemNum(MyIndex, DragBox.Slot) > 0 Then
                    If Item(GetPlayerInvItemNum(MyIndex, DragBox.Slot)).Type = i Then
                        If IsPlayerItemRequerimentsOK(MyIndex, GetPlayerInvItemNum(MyIndex, DragBox.Slot)) Then
                            Colour = DX8Colour(Green, 255)
                        Else
                            Colour = DX8Colour(Red, 255)
                        End If
                    End If
                End If
            End If
        End If

        If Options.ItemAnimation = YES Then
            rec.top = 0
            rec.Left = EquipmentFrames(i) * PIC_X
        End If
        RenderTexture ItemPic, xO, yO, rec.Left, rec.top, 32, 32, 32, 32, Colour
    Next
End Sub

Public Function IsPlayerItemRequerimentsOK(ByVal Index As Long, ByVal itemNum As Long) As Boolean
    IsPlayerItemRequerimentsOK = True
    Dim i As Byte

    ' stat requirements
    For i = 1 To Stats.Stat_Count - 1
        If GetPlayerRawStat(Index, i) < Item(itemNum).Stat_Req(i) Then
            IsPlayerItemRequerimentsOK = False
        End If
    Next

    ' level requirement
    If GetPlayerLevel(Index) < Item(itemNum).LevelReq Then
        IsPlayerItemRequerimentsOK = False
    End If

    ' class requirement
    If Item(itemNum).ClassReq > 0 Then
        If Not GetPlayerClass(Index) = Item(itemNum).ClassReq Then
            IsPlayerItemRequerimentsOK = False
        End If
    End If

    ' access requirement
    If Not GetPlayerAccess(Index) >= Item(itemNum).AccessReq Then
        IsPlayerItemRequerimentsOK = False
    End If

    ' prociency requirement
    If Not hasProficiency(Index, Item(itemNum).proficiency) Then
        IsPlayerItemRequerimentsOK = False
    End If

End Function

Public Function GetPlayerRawStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    If Index > Player_HighIndex Then Exit Function

    GetPlayerRawStat = Player(Index).Stat(Stat)
End Function

Public Sub chkCharacters_Atributos()
    Dim i As Integer
    With Windows(GetWindowIndex("winCharacter"))
        ' Hide Controls = 6 to 16 (Equipments)
        For i = 6 To 27
            .Controls(i).visible = False
        Next i
        
        ' Show Controls > Count 17 (Atributos)
        For i = 28 To .ControlCount
            .Controls(i).visible = True
        Next i
        ' grey out buttons
        If GetPlayerPOINTS(MyIndex) = 0 Then
            For i = 1 To Stats.Stat_Count - 1
                .Controls(GetControlIndex("winCharacter", "btnGreyStat_" & i)).visible = True
            Next
        Else
            For i = 1 To Stats.Stat_Count - 1
                .Controls(GetControlIndex("winCharacter", "btnGreyStat_" & i)).visible = False
            Next
        End If
        .Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = 0
        .Controls(GetControlIndex("winCharacter", "chkAtributos")).Value = 1
    End With
End Sub

Public Sub chkCharacters_Equipamentos()
    Dim i As Integer
    With Windows(GetWindowIndex("winCharacter"))
        ' Hide Controls > Count 18 (Atributos)
        For i = 28 To .ControlCount
            .Controls(i).visible = False
        Next i
        ' Show Controls = 6 to 17 (Equipments)
        For i = 6 To 19
            .Controls(i).visible = True
        Next i
        
        ' Atualizar a janela do conjunto
        UpdateConjuntoWindow UsingSet
        
        .Controls(GetControlIndex("winCharacter", "chkAtributos")).Value = 0
        .Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = 1
    End With
End Sub

' ###############
' ## Character ##
' ###############

Public Sub Character_MouseDown()
    Dim EqItem As Long, winIndex As Long, i As Long

    ' is there an item?
    EqItem = IsEqItem(Windows(GetWindowIndex("winCharacter")).Window.Left, Windows(GetWindowIndex("winCharacter")).Window.top)

    If EqItem > 0 Then
        ' exit out if we're offering that item

        ' drag it
        With DragBox
            .Type = Part_Item
            .Value = GetPlayerEquipment(MyIndex, EqItem)
            .Origin = origin_Equip
            .Slot = EqItem
        End With

        winIndex = GetWindowIndex("winDragBox")
        With Windows(winIndex).Window
            .state = MouseDown
            .Left = lastMouseX - 16
            .top = lastMouseY - 16
            .movedX = clickedX - .Left
            .movedY = clickedY - .top
        End With

        ShowWindow winIndex, , False
        ' stop dragging inventory
        Windows(GetWindowIndex("winCharacter")).Window.state = Normal
    End If

    ' show desc. if needed
    Character_MouseMove
End Sub

Public Sub Character_DblClick()
    Dim itemNum As Long

    'Chk Equipamentos non checked? exit
    If Windows(GetWindowIndex("winCharacter")).Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = 0 Then Exit Sub

    itemNum = IsEqItem(Windows(GetWindowIndex("winCharacter")).Window.Left, Windows(GetWindowIndex("winCharacter")).Window.top)

    If itemNum Then
        SendUnequip itemNum
    End If

    ' show desc. if needed
    Character_MouseMove
End Sub

Public Sub Character_MouseMove()
    Dim itemNum As Long, X As Long, Y As Long

    'Chk Equipamentos non checked? exit
    If Windows(GetWindowIndex("winCharacter")).Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = 0 Then Exit Sub

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

    itemNum = IsEqItem(Windows(GetWindowIndex("winCharacter")).Window.Left, Windows(GetWindowIndex("winCharacter")).Window.top)

    If itemNum Then
        ' calc position
        X = Windows(GetWindowIndex("winCharacter")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        Y = Windows(GetWindowIndex("winCharacter")).Window.top - 4
        ' offscreen?
        If X < 0 Then
            ' switch to right
            X = Windows(GetWindowIndex("winCharacter")).Window.Left + Windows(GetWindowIndex("winCharacter")).Window.Width
        End If
        ' go go go
        ShowEqDesc X, Y, itemNum
    End If
End Sub

Public Sub Character_SpendPoint1()
    SendTrainStat 1
End Sub

Public Sub Character_SpendPoint2()
    SendTrainStat 2
End Sub

Public Sub Character_SpendPoint3()
    SendTrainStat 3
End Sub

Public Sub Character_SpendPoint4()
    SendTrainStat 4
End Sub

Public Sub Character_SpendPoint5()
    SendTrainStat 5
End Sub
