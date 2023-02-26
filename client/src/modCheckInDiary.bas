Attribute VB_Name = "modCheckIn"
Option Explicit

Public DayReward() As DayRewardRec
Public DayCheckIn As Byte

Private Type DayRewardRec
    itemNum As Integer
    ItemQuant As Long
End Type

Public Sub CreateWindow_CheckIn()
    Dim i As Byte, X As Long, Y As Long, sString As String
    ' Create window
    CreateWindow "winCheckIn", "Realizar Login Diario", zOrder_Win, 0, 0, 489, 245, Tex_Item(105), , , , , , DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , , , , , False, False
    ' Centralise it
    CentraliseWindow WindowCount
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Close Button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf Close_CheckInWindow)

    ' Parchment and Wood Background
    CreatePictureBox WindowCount, "picWood", 0, 20, 489, 230, , , , , , , , DesignTypes.desWood, DesignTypes.desWood, DesignTypes.desWood
    'CreatePictureBox WindowCount, "picParchment", 6, 25, 788, 424, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment, , , GetAddress(AddressOf ChangeControls_Unselect)
    
    ' Pictures

    X = 17
    Y = 37
    For i = 1 To 31
        CreatePictureBox WindowCount, "picDay" & i, X, Y, 32, 32, False, , , , , , , DesignTypes.desDescPic, DesignTypes.desDescPic, DesignTypes.desDescPic, , GetAddress(AddressOf CheckIn_MouseMove), , GetAddress(AddressOf CheckIn_MouseMove)
        X = X + 47
        
        sString = "Dia " & i
        CreateLabel WindowCount, "lblDay" & i, X - 45, Y - 11, 64, , sString, rockwellDec_15, White, Alignment.alignLeft, False
        CreateLabel WindowCount, "lblQuant" & i, X - 62, Y + 24, 64, , , rockwellDec_15, White, Alignment.alignCentre, False
        
        If i = 10 Or i = 20 Or i = 30 Then Y = Y + 47: X = 17
    Next i
    
    CreatePictureBox WindowCount, "picSelect", 0, 0, 16, 16, False, , , , Tex_GUI(5), Tex_GUI(5), Tex_GUI(5)
    
    CreateButton WindowCount, "btnReivindicar", 183, 217, 122, 22, "Check-In", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf Reivindicar_CheckInWindow)
End Sub

Public Sub Update_CheckInWindow()
    Dim i As Byte, z As Byte ', TextureBlack As TextureStruct, Tex_Texture As Long

    With Windows(GetWindowIndex("winCheckIn"))
        For i = 1 To UBound(DayReward)
            .Controls(GetControlIndex("winCheckIn", "picDay" & i)).visible = True
            .Controls(GetControlIndex("winCheckIn", "lblDay" & i)).visible = True

            If DayReward(i).itemNum > 0 Then
                .Controls(GetControlIndex("winCheckIn", "lblQuant" & i)).Text = DayReward(i).ItemQuant
                .Controls(GetControlIndex("winCheckIn", "lblQuant" & i)).visible = True
                
                For z = 0 To 2
                    'TextureBlack = mTexture(Tex_Item(Item(DayReward(i).ItemNum).Pic))
                    'Set TextureBlack.Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, mTexture(Tex_Item(Item(DayReward(i).ItemNum).Pic)).data(0), AryCount(mTexture(Tex_Item(Item(DayReward(i).ItemNum).Pic)).data), mTexture(Tex_Item(Item(DayReward(i).ItemNum).Pic)).w, mTexture(Tex_Item(Item(DayReward(i).ItemNum).Pic)).h, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, D3DColorARGB(255, 0, 0, 0), ByVal 0, ByVal 0)
                    'Tex_Texture = LoadTextureFile(App.path & Path_Item & Item(DayReward(i).ItemNum).Pic & ".png")
                    'Set mTexture(Tex_Texture).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, mTexture(Tex_Texture).data(0), AryCount(mTexture(Tex_Texture).data), mTexture(Tex_Texture).w, mTexture(Tex_Texture).h, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 5, ByVal 0, ByVal 0)
                    
                    .Controls(GetControlIndex("winCheckIn", "picDay" & i)).image(z) = Tex_Item(Item(DayReward(i).itemNum).Pic)
                Next z
            End If
        Next i
    End With
End Sub

Private Sub Graphic_ArrowInDay()

    If DayCheckIn = 0 Then
        Windows(GetWindowIndex("winCheckIn")).Controls(GetControlIndex("winCheckIn", "picSelect")).visible = False
        Exit Sub
    End If

    With Windows(GetWindowIndex("winCheckIn"))
        .Controls(GetControlIndex("winCheckIn", "picSelect")).visible = True
        .Controls(GetControlIndex("winCheckIn", "picSelect")).top = .Controls(GetControlIndex("winCheckIn", "picDay" & DayCheckIn)).top - 30
        .Controls(GetControlIndex("winCheckIn", "picSelect")).Left = .Controls(GetControlIndex("winCheckIn", "picDay" & DayCheckIn)).Left + 10
        
        .Controls(GetControlIndex("winCheckIn", "lblDay" & DayCheckIn)).textColour = Yellow
        
        ' Fiz pra usar na animação
        .Controls(GetControlIndex("winCheckIn", "picSelect")).Min = .Controls(GetControlIndex("winCheckIn", "picSelect")).top
        
    End With
End Sub

Public Sub Graphic_ArrowInDay_Animated()

    If DayCheckIn = 0 Then Exit Sub
    
    With Windows(GetWindowIndex("winCheckIn"))
        If .Window.visible = False Then Exit Sub

        If (.Controls(GetControlIndex("winCheckIn", "picSelect")).top) >= .Controls(GetControlIndex("winCheckIn", "picSelect")).Min Then
            If (.Controls(GetControlIndex("winCheckIn", "picSelect")).top - .Controls(GetControlIndex("winCheckIn", "picSelect")).Min) > 10 Then
                .Controls(GetControlIndex("winCheckIn", "picSelect")).top = .Controls(GetControlIndex("winCheckIn", "picSelect")).Min
            Else
                .Controls(GetControlIndex("winCheckIn", "picSelect")).top = .Controls(GetControlIndex("winCheckIn", "picSelect")).top + 1
            End If
        End If
    End With
End Sub

Private Sub CheckIn_MouseMove()
    Dim X As Integer, Y As Integer
    Dim Width As Integer, Height As Integer
    Dim i As Byte

    With Windows(GetWindowIndex("winCheckIn"))
        If .Window.visible = False Then Exit Sub

        For i = 1 To 31
            If .Controls(GetControlIndex("winCheckIn", "picDay" & i)).visible Then
                If .Controls(GetControlIndex("winCheckIn", "picDay" & i)).image(0) <> 0 Then
                    X = .Window.Left + .Controls(GetControlIndex("winCheckIn", "picDay" & i)).Left
                    Y = .Window.top + .Controls(GetControlIndex("winCheckIn", "picDay" & i)).top
                    Width = .Controls(GetControlIndex("winCheckIn", "picDay" & i)).Width
                    Height = .Controls(GetControlIndex("winCheckIn", "picDay" & i)).Height
                    If GlobalX >= X And GlobalX <= X + Width Then
                        If GlobalY >= Y And GlobalY <= Y + Height Then
                            ShowItemDesc GlobalX, GlobalY, CLng(DayReward(i).itemNum), False
                        End If
                    End If
                End If
            End If
        Next i
    End With

End Sub

Private Sub Close_CheckInWindow()
    HideWindow GetWindowIndex("winCheckIn")
End Sub

Private Sub Reivindicar_CheckInWindow()
    HideWindow GetWindowIndex("winCheckIn")
    
    Call SendCheckIn
End Sub

Private Sub SendCheckIn()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCheckIn
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub HandleUpdateDayReward(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Days As Byte, i As Byte, ActualDay As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    DayCheckIn = Buffer.ReadByte
    
    Days = Buffer.ReadByte
    
    ReDim DayReward(1 To Days)
    
    For i = 1 To Days
        DayReward(i).itemNum = Buffer.ReadInteger
        DayReward(i).ItemQuant = Buffer.ReadLong
    Next i
    
    Set Buffer = Nothing
    
    Call Update_CheckInWindow
    ShowWindow GetWindowIndex("winCheckIn"), True
    Call Graphic_ArrowInDay
End Sub
