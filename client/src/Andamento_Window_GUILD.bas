Attribute VB_Name = "modGuild_Window"
Option Explicit

Public Const MAX_GUILD_CAPACITY As Byte = 10

Public Sub CreateWindow_Guild()
' Create window
    CreateWindow "winGuild", "Guild", zOrder_Win, 0, 0, 175, 340, Tex_Item(1), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Guild)

    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 162, 306, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment

    ' Attributes
    CreatePictureBox WindowCount, "picShadow", 18, 38, 138, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblGuild", 18, 35, 138, , "Nenhuma Guild", rockwellDec_15, , Alignment.alignCentre

    ' White boxes
    CreatePictureBox WindowCount, "picWhiteBox", 13, 51, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 71, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 91, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 111, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 131, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 151, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    ' Labels
    CreateLabel WindowCount, "lblRank", 18, 54, 147, 16, "Lider:", rockwellDec_10, , Alignment.alignLeft
    CreateLabel WindowCount, "lblKills", 18, 74, 147, 16, "Kills:", rockwellDec_10
    CreateLabel WindowCount, "lblVictory", 18, 94, 147, 16, "Vitorias:", rockwellDec_10
    CreateLabel WindowCount, "lblLose", 18, 114, 124, 16, "Derrotas:", rockwellDec_10
    CreateLabel WindowCount, "lblHonra", 18, 134, 144, 16, "Honra:", rockwellDec_10
    CreateLabel WindowCount, "lblMembers", 18, 154, 147, 16, "Membros:", rockwellDec_10

    'Label Shadow
    CreatePictureBox WindowCount, "picShadow", 18, 178, 138, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblGuild", 18, 175, 138, , "Anuncio", rockwellDec_15, , Alignment.alignCentre

    'Anuncio
    CreatePictureBox WindowCount, "picWhiteBox", 13, 190, 148, 98, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreateLabel WindowCount, "lblAnnouncement", 20, 198, 138, , "Você não tem guild!", rockwell_15, , Alignment.alignLeft

    'Btn Menu
    CreateButton WindowCount, "btnReturn", 15, 293, 55, 23, "Menu", rockwellDec_15, , , , , , , , DesignTypes.desBlue, DesignTypes.desBlue_Hover, DesignTypes.desBlue_Click, , , GetAddress(AddressOf btnGuildMenu_Save)

    'Btn Leave
    CreateButton WindowCount, "btnLeave", 105, 293, 55, 23, "Leave", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf SendLeaveGuild)

End Sub

Public Sub CreateWindow_GuildMenu()
    Dim i As Byte
    Dim X As Long, Y As Long
    ' Create window
    CreateWindow "winGuildMenu", "Guild Menu", zOrder_Win, 0, 0, 375, 450, Tex_Item(1), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, , , , , , , , , , False
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnGuildMenu_Close)

    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 362, 416, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment, , , GetAddress(AddressOf GuildTextboxes_Unselect)

    ' Anuncio
    CreatePictureBox WindowCount, "picShadow1", 18, 38, 170, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblGuild1", 35, 35, 138, , "Anuncio", rockwellDec_15, , Alignment.alignCentre
    CreateTextbox WindowCount, "txtAnnouncement", 18, 55, 170, 18, , rockwell_15, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , , 1, , , 122

    ' Cor da Guild
    CreatePictureBox WindowCount, "picShadow2", 18, 148, 170, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblGuild2", 18, 145, 170, , "Cor da Guild", rockwellDec_15, , Alignment.alignCentre
    CreateComboBox WindowCount, "cmbColor", 18, 165, 170, 18, DesignTypes.desComboNorm, verdana_12

    ' Membros da Guild
    CreatePictureBox WindowCount, "picShadow4", 190, 38, 170, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblGuild4", 207, 35, 138, , "Membros da Guild", rockwellDec_15, , Alignment.alignCentre

    ' Render Guild Members Box
    X = 200
    Y = 55
    For i = 1 To MAX_GUILD_CAPACITY
        CreateCheckbox WindowCount, "chkPlayer" & i, X, Y, 80, , , "Vazio", rockwell_15, , , , , DesignTypes.desChkNorm, , , GetAddress(AddressOf chkPlayer_Click)
        Y = Y + 18
    Next i
    CreateButton WindowCount, "btnPromover", X, Y, 80, 18, "Promover", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf BtnGuildPromote)
    CreateButton WindowCount, "btnRebaixar", X + 85, Y, 80, 18, "Rebaixar", rockwellDec_15, , , , , , , , DesignTypes.desOrange, DesignTypes.desOrange_Hover, DesignTypes.desOrange_Click, , , GetAddress(AddressOf BtnGuildRebaixar)
    CreateButton WindowCount, "btnKick", X + 40, Y + 25, 80, 18, "Kick", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf BtnGuildKick)

    ' Convidar pelo nome
    CreatePictureBox WindowCount, "picShadow5", 18, 208, 170, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblGuild5", 18, 205, 170, , "Convidar Jogador", rockwellDec_15, , Alignment.alignCentre
    CreateTextbox WindowCount, "txtInvite", 18, 225, 170, 18, , rockwell_15, , Alignment.alignCentre, , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , , , , , , , 1, , , NAME_LENGTH
    CreateButton WindowCount, "btnInvite", 18, 255, 170, 18, "Enviar Convite", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf BtnGuildInvite)

    ' Ícone da Guild
    CreatePictureBox WindowCount, "picShadow3", 18, 293, 340, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblGuild3", 100, 290, 170, , "Icone da Guild", rockwellDec_15, , Alignment.alignCentre
    Y = 310
    X = 18
    ' render grid - row
    For i = 1 To Count_Flags

        If i = Count_Flags Then
            CreatePictureBox WindowCount, "picFlags" & i, X, Y, 16, 12, , , , , Tex_Flags(i), Tex_Flags(i), Tex_Flags(i), , , , , , GetAddress(AddressOf btnSelect_Flag), , , GetAddress(AddressOf DrawFlagBox)
        Else
            CreatePictureBox WindowCount, "picFlags" & i, X, Y, 16, 12, , , , , Tex_Flags(i), Tex_Flags(i), Tex_Flags(i), , , , , , GetAddress(AddressOf btnSelect_Flag)
        End If
        X = X + 18

        If i = 19 Then Y = Y + 24: X = 18
        If i = 38 Then Y = Y + 24: X = 18
        If i = 57 Then Y = Y + 24: X = 18
    Next

    CreatePictureBox WindowCount, "picSelect", 0, 0, 16, 16, , , , , Tex_GUI(5), Tex_GUI(5), Tex_GUI(5)

    'Btn Menu
    CreateButton WindowCount, "btnGuildMenuOk", 40, (Windows(WindowCount).Window.Height - 40), 100, 23, "Salvar", rockwellDec_15, , , , , , , , DesignTypes.desBlue, DesignTypes.desBlue_Hover, DesignTypes.desBlue_Click, , , GetAddress(AddressOf btnGuildMenu_Save)
    'Btn Disband Guild
    CreateButton WindowCount, "btnGuildDisband", 235, (Windows(WindowCount).Window.Height - 40), 100, 23, "Destruir Guild", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf BtnGuildDestroy)
End Sub

Private Sub chkPlayer_Click()
    Dim i As Byte
    Dim Selected As Byte

    If Player(MyIndex).Guild_ID = 0 Then Exit Sub
    With Windows(GetWindowIndex("winGuildMenu"))
        For i = 1 To MAX_GUILD_CAPACITY

            ' Verifica o controle que vai selecionar
            If GlobalX >= .Window.left + .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).left And GlobalX <= .Window.left + .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).left + .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Width Then
                If GlobalY >= .Window.top + .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).top And GlobalY <= .Window.top + .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).top + .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Height Then
                    Selected = i
                End If
            End If
            If i <> Selected Then
                .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Value = NO
            End If
        Next i
    End With

End Sub

Private Sub BtnGuildKick()
    Dim i As Byte
    Dim Selected As Byte

    If Player(MyIndex).Guild_ID = 0 Then Exit Sub
    
    If GuildMembers(Player(MyIndex).Guild_ID).Membro(Player(MyIndex).Guild_MembroID).Admin = False Then
        AddText "Apenas Líderes da guild podem fazer isso!", BrightRed
        Exit Sub
    End If

    With Windows(GetWindowIndex("winGuildMenu"))
        For i = 1 To MAX_GUILD_CAPACITY
           If i <= Guild(Player(MyIndex).Guild_ID).Capacidade Then
            If GuildMembers(Player(MyIndex).Guild_ID).Membro(i).MembroID <> 0 Then
                If .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Value Then
                    Selected = i: Exit For
                End If
            End If
           End If
        Next i

        If Selected > 0 Then
            If MsgBox("Tem certeza que quer expulsar " & GuildMembers(Player(MyIndex).Guild_ID).Membro(Selected).Name & " ?", vbYesNo) = vbYes Then
                SendGuildKick Selected
            End If
        End If

    End With

End Sub

Private Sub BtnGuildRebaixar()
    Dim i As Byte
    Dim Selected As Byte

    If Player(MyIndex).Guild_ID = 0 Then Exit Sub
    
    If GuildMembers(Player(MyIndex).Guild_ID).Membro(Player(MyIndex).Guild_MembroID).Admin = False Then
        AddText "Apenas donos da guild podem fazer isso!", BrightRed
        Exit Sub
    End If

    With Windows(GetWindowIndex("winGuildMenu"))
        For i = 1 To MAX_GUILD_CAPACITY
            If i <= Guild(Player(MyIndex).Guild_ID).Capacidade Then
                If GuildMembers(Player(MyIndex).Guild_ID).Membro(i).MembroDisponivel = False Then
                    If .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Value Then
                        Selected = i: Exit For
                    End If
                End If
            End If
        Next i

        If Selected > 0 Then
            If MsgBox("Tem certeza que quer rebaixar " & GuildMembers(Player(MyIndex).Guild_ID).Membro(Selected).Name & " à membro da Guild?", vbYesNo) = vbYes Then
                Call SendGuildPromote(NO, Selected)
            End If
        End If

    End With
End Sub

Private Sub BtnGuildPromote()
    Dim i As Byte
    Dim Selected As Byte

    If Player(MyIndex).Guild_ID = 0 Then Exit Sub
    If GuildMembers(Player(MyIndex).Guild_ID).Membro(Player(MyIndex).Guild_MembroID).Admin = False Then Exit Sub

    With Windows(GetWindowIndex("winGuildMenu"))
        For i = 1 To MAX_GUILD_CAPACITY
            If i <= Guild(Player(MyIndex).Guild_ID).Capacidade Then
                If GuildMembers(Player(MyIndex).Guild_ID).Membro(i).MembroDisponivel = False Then
                    If .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Value Then
                        Selected = i: Exit For
                    End If
                End If
            End If
        Next i

        If Selected > 0 Then
            If MsgBox("Tem certeza que quer promover " & GuildMembers(Player(MyIndex).Guild_ID).Membro(Selected).Name & " à Admin da Guild?", vbYesNo) = vbYes Then
                Call SendGuildPromote(YES, Selected)
            End If
        End If

    End With

End Sub

Private Sub BtnGuildInvite()
    SendGuildInvite Windows(GetWindowIndex("winGuildMenu")).Controls(GetControlIndex("winGuildMenu", "txtInvite")).Text

    Windows(GetWindowIndex("winGuildMenu")).Controls(GetControlIndex("winGuildMenu", "txtInvite")).Text = vbNullString
End Sub

Private Sub BtnGuildDestroy()
    If MsgBox("Tem certeza que quer destruir sua guild?", vbYesNo) = vbYes Then
        SendGuildDestroy
        HideWindow GetWindowIndex("winGuildMenu")
    End If
End Sub

Private Sub DrawFlagBox()

    If Player(MyIndex).Guild_Icon = 0 Or Player(MyIndex).Guild_Icon > Count_Flags Then Exit Sub

    With Windows(GetWindowIndex("winGuildMenu"))

        'RenderEntity_Square Tex_Design(17), .Window.Left + .Controls(GetControlIndex("winGuildMenu", "picFlags" & Player(MyIndex).Guild_Icon)).Left, .Window.top + .Controls(GetControlIndex("winGuildMenu", "picFlags" & Player(MyIndex).Guild_Icon)).top, 17, 13, 1, 100
        RenderDesign DesignTypes.desTileBox, .Window.left + .Controls(GetControlIndex("winGuildMenu", "picFlags" & Player(MyIndex).Guild_Icon)).left, .Window.top + .Controls(GetControlIndex("winGuildMenu", "picFlags" & Player(MyIndex).Guild_Icon)).top, 17, 13
    End With
End Sub

Private Sub SetPicSelect()

    If Player(MyIndex).Guild_Icon = 0 Then
        Windows(GetWindowIndex("winGuildMenu")).Controls(GetControlIndex("winGuildMenu", "picSelect")).visible = False
        Exit Sub
    End If

    With Windows(GetWindowIndex("winGuildMenu"))
        .Controls(GetControlIndex("winGuildMenu", "picSelect")).visible = True
        .Controls(GetControlIndex("winGuildMenu", "picSelect")).top = .Controls(GetControlIndex("winGuildMenu", "picFlags" & Player(MyIndex).Guild_Icon)).top - 12
        .Controls(GetControlIndex("winGuildMenu", "picSelect")).left = .Controls(GetControlIndex("winGuildMenu", "picFlags" & Player(MyIndex).Guild_Icon)).left
    End With
End Sub

Private Sub btnSelect_Flag()
    Dim i As Byte
    If Player(MyIndex).Guild_ID = 0 Then Exit Sub
    With Windows(GetWindowIndex("winGuildMenu"))
        For i = 1 To Count_Flags
            If GlobalX >= .Window.left + .Controls(GetControlIndex("winGuildMenu", "picFlags" & i)).left And GlobalX <= .Window.left + .Controls(GetControlIndex("winGuildMenu", "picFlags" & i)).left + .Controls(GetControlIndex("winGuildMenu", "picFlags" & i)).Width Then
                If GlobalY >= .Window.top + .Controls(GetControlIndex("winGuildMenu", "picFlags" & i)).top And GlobalY <= .Window.top + .Controls(GetControlIndex("winGuildMenu", "picFlags" & i)).top + .Controls(GetControlIndex("winGuildMenu", "picFlags" & i)).Height Then
                    Player(MyIndex).Guild_Icon = i
                    SetPicSelect
                End If
            End If
        Next i
    End With
End Sub

' Ao Clicar no parchment do background, tira a seleção de controles!
Public Sub GuildTextboxes_Unselect()
    Dim i As Byte
    With Windows(GetWindowIndex("winGuildMenu"))
        For i = 1 To .ControlCount
            If .activeControl = i Then
                .activeControl = 0
            End If
        Next i
    End With
End Sub

Private Sub AddComboColors()

' Proteção
    If Player(MyIndex).Guild_ID = 0 Then Exit Sub

    ' clear the combolists
    Erase Windows(GetWindowIndex("winGuildMenu")).Controls(GetControlIndex("winGuildMenu", "cmbColor")).list
    ReDim Windows(GetWindowIndex("winGuildMenu")).Controls(GetControlIndex("winGuildMenu", "cmbColor")).list(0)

    ' add combobox options
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Preto"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Azul"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Verde"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Ciano"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Vermelho"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Magenta"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Marrom"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Cinza"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Cinza Escuro"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Azul Claro"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Verde Claro"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Ciano Claro"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Vermelho Claro"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Rosa"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Amarelo"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Branco"
    Combobox_AddItem GetWindowIndex("winGuildMenu"), GetControlIndex("winGuildMenu", "cmbColor"), "Marrom Escuro"

    Windows(GetWindowIndex("winGuildMenu")).Controls(GetControlIndex("winGuildMenu", "cmbColor")).Value = 1 + Guild(Player(MyIndex).Guild_ID).Color
End Sub

Public Sub btnGuildMenu_Save()

    If Player(MyIndex).Guild_ID = 0 Then Exit Sub

    With Windows(GetWindowIndex("winGuildMenu"))
        If .Window.visible Then
            Call SendSaveGuild
            HideWindow GetWindowIndex("winGuildMenu")
            ShowWindow GetWindowIndex("winGuild"), , False
        Else
            'If GuildMembers(Player(MyIndex).Guild_ID).Membro(Player(MyIndex).Guild_MembroID).Admin = True Then
            Player(MyIndex).Guild_Icon = Guild(Player(MyIndex).Guild_ID).Icon
            SetPicSelect
            HideWindow GetWindowIndex("winGuild")
            ShowWindow GetWindowIndex("winGuildMenu"), , False
            ' End If
        End If
    End With
End Sub

Public Sub btnGuildMenu_Close()
    HideWindow GetWindowIndex("winGuildMenu")
    ShowWindow GetWindowIndex("winGuild"), , False
End Sub

Public Sub SendSaveGuild()
    Dim Anuncio As String, Color As Byte, IconChanged As Byte
    Dim Buffer As clsBuffer

    If GuildMembers(Player(MyIndex).Guild_ID).Membro(Player(MyIndex).Guild_MembroID).Admin = False Then
        AddText "Apenas donos da guild podem fazer isso!", BrightRed
        UpdateWindowGuild
        HideWindow GetWindowIndex("winGuildMenu")
        Exit Sub
    End If

    Anuncio = Windows(GetWindowIndex("winGuildMenu")).Controls(GetControlIndex("winGuildMenu", "txtAnnouncement")).Text
    Color = (Windows(GetWindowIndex("winGuildMenu")).Controls(GetControlIndex("winGuildMenu", "cmbColor")).Value - 1)
    IconChanged = Player(MyIndex).Guild_Icon

    ' Pra não alterar caso seja o mesmo que já está na guild
    If Anuncio = Trim$(Guild(Player(MyIndex).Guild_ID).Motd) And Color = Guild(Player(MyIndex).Guild_ID).Color And IconChanged = Guild(Player(MyIndex).Guild_ID).Icon Then
        Exit Sub
    End If

    If IconChanged = 0 Then IconChanged = Guild(Player(MyIndex).Guild_ID).Icon

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSaveGuild

    Buffer.WriteString Anuncio
    Buffer.WriteByte Color
    Buffer.WriteByte IconChanged

    SendData Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub UpdateWindowGuild()
    Dim GuildID As Byte
    Dim i As Byte
    Dim GuildLeader As String
    Dim GuildMembros As Byte

    ' Tem Guild? Se não tiver faz a limpeza da janela!
    If Player(MyIndex).Guild_ID = 0 Then
        ClearWindowGuild
        Exit Sub
    End If

    ' Pega o ID da Guild
    GuildID = Player(MyIndex).Guild_ID

    With Windows(GetWindowIndex("winGuild"))

        ' Partiu atualizar os textos
        ' Procura o lider da guild
        For i = 1 To Guild(GuildID).Capacidade
            If GuildMembers(GuildID).Membro(i).Dono Then
                GuildLeader = GuildMembers(GuildID).Membro(i).Name
            End If

            ' Conta quantos membros tem na guild
            If Not GuildMembers(GuildID).Membro(i).MembroDisponivel Then
                GuildMembros = GuildMembros + 1
            End If
        Next i

        'Painel Geral
        .Controls(GetControlIndex("winGuild", "lblGuild")).Text = Guild(GuildID).Name
        .Controls(GetControlIndex("winGuild", "lblRank")).Text = "Lider: " & GuildLeader
        .Controls(GetControlIndex("winGuild", "lblKills")).Text = "Kills: " & Guild(GuildID).Kills
        .Controls(GetControlIndex("winGuild", "lblVictory")).Text = "Vitorias: " & Guild(GuildID).Victory
        .Controls(GetControlIndex("winGuild", "lblLose")).Text = "Derrotas: " & Guild(GuildID).Lose
        .Controls(GetControlIndex("winGuild", "lblHonra")).Text = "Honra: " & Guild(GuildID).Honra
        .Controls(GetControlIndex("winGuild", "lblMembers")).Text = "Membros: " & GuildMembros
        .Controls(GetControlIndex("winGuild", "lblAnnouncement")).Text = Guild(GuildID).Motd
    End With

    ' Painel Guild Menu
    With Windows(GetWindowIndex("winGuildMenu"))
        .Controls(GetControlIndex("winGuildMenu", "txtAnnouncement")).Height = (TextHeight(font(verdanaBold_12)) * 5)
        .Controls(GetControlIndex("winGuildMenu", "txtAnnouncement")).Text = Guild(GuildID).Motd
        AddComboColors
        ' Coloca o nome do jogador na lista de players da guild no guild menu
        For i = 1 To Guild(GuildID).Capacidade
            If GuildMembers(GuildID).Membro(i).Name <> vbNullString Then
                If GuildMembers(GuildID).Membro(i).Admin = True Then
                    If GuildMembers(GuildID).Membro(i).Dono = True Then
                        If GuildMembers(GuildID).Membro(i).Online = True Then
                            .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Text = GuildMembers(GuildID).Membro(i).Name & "[Leader] (On)"
                        Else
                            .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Text = GuildMembers(GuildID).Membro(i).Name & "[Leader] (Off)"
                        End If
                    Else
                        If GuildMembers(GuildID).Membro(i).Online = True Then
                            .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Text = GuildMembers(GuildID).Membro(i).Name & "[Admin] (On)"
                        Else
                            .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Text = GuildMembers(GuildID).Membro(i).Name & "[Admin] (Off)"
                        End If
                    End If
                Else
                    If GuildMembers(GuildID).Membro(i).Online = True Then
                        .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Text = GuildMembers(GuildID).Membro(i).Name & " (On)"
                    Else
                        .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Text = GuildMembers(GuildID).Membro(i).Name & " (Off)"
                    End If
                End If
                ' Enable Checkboxes
                .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).enabled = True
            Else
                ' Disable Checkboxes caso nao exista o jogador
                .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).enabled = False
                .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Text = "Vazio"
                .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Value = 0
                .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).textColour = White
            End If
        Next i

        ' Adiciona o nome ''Bloqueado'' em slots que a guild ainda não possui adquirido
        If Guild(GuildID).Capacidade < MAX_GUILD_CAPACITY Then
            For i = (Guild(GuildID).Capacidade + 1) To MAX_GUILD_CAPACITY
                .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Text = "Locked (Unlock Donate)"
                .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).textColour = BrightRed
                .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).enabled = False
            Next i
        End If
    End With

End Sub

Public Sub ClearWindowGuild()
    Dim i As Integer

    ' Painel da Guild
    With Windows(GetWindowIndex("winGuild"))
        .Controls(GetControlIndex("winGuild", "lblGuild")).Text = "Nenhuma Guild"
        .Controls(GetControlIndex("winGuild", "lblRank")).Text = "Lider: "
        .Controls(GetControlIndex("winGuild", "lblKills")).Text = "Kills: "
        .Controls(GetControlIndex("winGuild", "lblVictory")).Text = "Vitorias: "
        .Controls(GetControlIndex("winGuild", "lblLose")).Text = "Derrotas: "
        .Controls(GetControlIndex("winGuild", "lblHonra")).Text = "Honra: "
        .Controls(GetControlIndex("winGuild", "lblMembers")).Text = "Membros: "
        .Controls(GetControlIndex("winGuild", "lblAnnouncement")).Text = vbNullString
        .Controls(GetControlIndex("winGuildMenu", "txtAnnouncement")).Text = vbNullString
    End With

    ' Painel Guild Menu
    With Windows(GetWindowIndex("winGuildMenu"))
        .Controls(GetControlIndex("winGuildMenu", "txtAnnouncement")).Height = (TextHeight(font(verdanaBold_12)) * 5)
        .Controls(GetControlIndex("winGuildMenu", "txtAnnouncement")).Text = vbNullString

        ' Adiciona o nome ''Bloqueado'' em slots que a guild ainda não possui adquirido
        For i = 1 To MAX_GUILD_CAPACITY
            .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Text = "Vazio"
            .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).textColour = White
            .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).enabled = False
            .Controls(GetControlIndex("winGuildMenu", "chkPlayer" & i)).Value = 0
        Next i

        Player(MyIndex).Guild_Icon = 0
    End With
End Sub

Public Sub CreateWindow_GuildMaker()
' Create window
    CreateWindow "winGuildMaker", "Guild Maker", zOrder_Win, 0, 0, 200, 170, Tex_Item(1), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnCloseCreateGuild)

    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 187, 136, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment

    CreateLabel WindowCount, "lblGuild", 18, 45, 138, , "Guild Nome:", rockwellDec_15, , Alignment.alignLeft

    CreateTextbox WindowCount, "txtGuildName", 18, 65, 170, 18, , Fonts.rockwellDec_15, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , , 1, , GetAddress(AddressOf btnCreateGuild)
    'Btn Create
    CreateButton WindowCount, "btnReturn", 37, 120, 100, 23, "Create", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnCreateGuild)
End Sub

Public Sub btnCreateGuild()
    With Windows(GetWindowIndex("winGuildMaker"))
        ' Window visible?
        If .Window.visible = False Then Exit Sub
        ' Have Guild Name?
        If .Controls(GetControlIndex("winGuildMaker", "txtGuildName")).Text = vbNullString Then Exit Sub
        ' Ok, Send It.
        Call SendCriarGuild(.Controls(GetControlIndex("winGuildMaker", "txtGuildName")).Text)
        ' Hide the window.
        HideWindow GetWindowIndex("winGuildMaker")
    End With
End Sub

Public Sub btnCloseCreateGuild()
    Windows(GetWindowIndex("winGuildMaker")).Controls(GetControlIndex("winGuildMaker", "txtGuildName")).Text = vbNullString
    HideWindow GetWindowIndex("winGuildMaker")
End Sub

Public Sub DrawGuild(ByVal Index As Long)
    Dim textX As Long, textY As Long, Text As String, textSize As Long, Colour As Long, Icon As Byte

    If Player(Index).Guild_ID = 0 Then Exit Sub

    Icon = Guild(Player(Index).Guild_ID).Icon
    Text = Trim$(Guild(Player(Index).Guild_ID).Name)
    textSize = TextWidth(font(Fonts.rockwell_15), Text)
    ' get the colour
    Colour = Guild(Player(Index).Guild_ID).Color

    textX = Player(Index).X * PIC_X + Player(Index).xOffset + (PIC_X \ 2) - (textSize \ 2)
    textY = Player(Index).Y * PIC_Y + Player(Index).yOffset - 32

    If GetPlayerSprite(Index) >= 1 And GetPlayerSprite(Index) <= Count_Char Then
        textY = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - (mTexture(Tex_Char(GetPlayerSprite(Index))).h / 4)
    End If

    Call RenderText(font(Fonts.rockwell_15), Text, ConvertMapX(textX), ConvertMapY(textY), Colour)

    If Icon > 0 Then
        textX = textX - 18
        textY = textY + 2
        RenderTexture Tex_Flags(Icon), ConvertMapX(textX), ConvertMapY(textY), 0, 0, 16, 12, 16, 12
    End If
End Sub
