Attribute VB_Name = "modWindow_ChangeControls"
Option Explicit

Public Sub HandleKeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Byte

    ' exit out if chatting
    If Not inSmallChat Then Exit Sub
    ' exit out if dialog
    If diaIndex > 0 Then Exit Sub
    ' exit out if talking
    If Windows(GetWindowIndex("winChat")).Window.visible Then Exit Sub
    ' exit out if creating guild
    If Windows(GetWindowIndex("winGuildMaker")).Window.visible Then Exit Sub
    ' exit out if validade serial number
    If Windows(GetWindowIndex("winSerial")).Window.visible Then Exit Sub
    ' exit out if guild menu
    If Windows(GetWindowIndex("winGuildMenu")).Window.visible Then Exit Sub
    ' exit with changing controls
    If Windows(GetWindowIndex("winChangeControls")).Window.visible Then Exit Sub

    If InGame Then
        Select Case KeyCode
        Case Options.Options
            ' hide options screen
            HideWindow GetWindowIndex("winOptions")
            CloseComboMenu
            ' hide/show chat window
            If Windows(GetWindowIndex("winChat")).Window.visible Then
                Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).Text = vbNullString
                HideChat
                inSmallChat = True
                Exit Sub
            End If

            If Windows(GetWindowIndex("winEscMenu")).Window.visible Then
                ' hide it
                HideWindow GetWindowIndex("winBlank")
                HideWindow GetWindowIndex("winEscMenu")
            Else
                ' show them
                ShowWindow GetWindowIndex("winBlank"), True
                ShowWindow GetWindowIndex("winEscMenu"), True
            End If
            ' exit out early
            Exit Sub
        Case Options.Bolsa
            ' hide/show inventory
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Inv
        Case Options.Personagem
            ' hide/show inventory
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Char
        Case Options.Magias
            ' hide/show skills
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Skills
        Case Options.Guild
            ' hide/show guild
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Guild
        Case Options.Quests
            ' hide/show quest
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Quest
        End Select

        ' handles hotbar
        For i = 1 To 10
            If KeyCode = Options.Hotbar(i) Then
                SendHotbarUse i
            End If
        Next
    End If
    ' check if we're skipping video
    If KeyCode = Options.Options Then
        ' hide options screen
        HideWindow GetWindowIndex("winOptions")
        CloseComboMenu
        ' handle the video
        If videoPlaying Then
            videoPlaying = False
            fadeAlpha = 0
            frmMain.picIntro.visible = False
            StopIntro
            Exit Sub
        End If
        If Windows(GetWindowIndex("winEscMenu")).Window.visible Then
            ' hide it
            HideWindow GetWindowIndex("winBlank")
            HideWindow GetWindowIndex("winEscMenu")
            Exit Sub
        Else
            ' show them
            ShowWindow GetWindowIndex("winBlank"), True
            ShowWindow GetWindowIndex("winEscMenu"), True
            Exit Sub
        End If
    End If
End Sub

' Botões que tratam visibilidade do formulário!
' Change Controls
Public Sub btnChangeControls_Open()
    HideWindow GetWindowIndex("winOptions")
    ShowWindow GetWindowIndex("winChangeControls"), True, True
    ChangeControls_Unselect
End Sub

' Salvando os dados!
Public Sub btnSalvar_Return()

    ' Colocar informações do tooltip no menu do jogador, sobre as teclas pré-configuradas no editor de controles de cada menu!
    RefreshMenu_Tooltip
    
    SaveOptions
    HideWindow GetWindowIndex("winChangeControls")
    ShowWindow GetWindowIndex("winOptions"), True, True
End Sub

' Ao Clicar no parchment do background, tira a seleção de controles!
Public Sub ChangeControls_Unselect()
    Dim i As Byte
    With Windows(GetWindowIndex("winChangeControls"))
        For i = 1 To .ControlCount
            If .activeControl = i Then
                .activeControl = 0
                .Controls(i).textColour = White
            End If
        Next i
    End With
End Sub

' Adicionar cor amarela quando o controle for selecionado!
Public Sub ChangeControls_Select()
    Dim i As Byte
    With Windows(GetWindowIndex("winChangeControls"))
        For i = 1 To .ControlCount
            If .activeControl = i Then
                .Controls(i).textColour = Yellow
            Else
                .Controls(i).textColour = White
            End If
        Next i
    End With
End Sub

Public Sub CreateWindow_ChangeControls()
    Dim i As Byte
    ' Create window
    CreateWindow "winChangeControls", "Alterar Teclas de Controles", zOrder_Win, 0, 0, 800, 450, Tex_Item(38), , , , , , DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , , , , , False, False
    ' Centralise it
    CentraliseWindow WindowCount
    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Parchment and Wood Background
    CreatePictureBox WindowCount, "picWood", 0, 20, 800, 435, , , , , , , , DesignTypes.desWood, DesignTypes.desWood, DesignTypes.desWood
    CreatePictureBox WindowCount, "picParchment", 6, 25, 788, 424, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment, , , GetAddress(AddressOf ChangeControls_Unselect)
    ' Buttons
    CreateButton WindowCount, "btnSalvar", 186, 406, 178, 22, "Salvar", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnSalvar_Return)
    CreateButton WindowCount, "btnRestore", 426, 406, 178, 22, "Restaurar Padrao", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf ChangeControls_Restore)

    ' Actions Backgrounds
    CreatePictureBox WindowCount, "picBlackBox1", 50, 50, 340, 260, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , , GetAddress(AddressOf ChangeControls_Unselect)
    CreatePictureBox WindowCount, "picWhiteBox1", 55, 55, 330, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreateLabel WindowCount, "lblActions1", 195, 58, 173, , "ACTIONS", Fonts.georgiaDec_16, White
    ' Movimentation Backgrounds
    CreatePictureBox WindowCount, "picBlackBox2", 400, 50, 340, 260, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , , GetAddress(AddressOf ChangeControls_Unselect)
    CreatePictureBox WindowCount, "picWhiteBox2", 405, 55, 330, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreateLabel WindowCount, "lblActions2", 555, 58, 173, , "MOVES", Fonts.georgiaDec_16, White
    ' Name Hotbar background
    CreatePictureBox WindowCount, "picWhiteBox2", 79, 322, 640, 7, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack

    ' Directions Moves Image
    CreatePictureBox WindowCount, "picUp", 555, 85, 16, 16, , , , , Tex_GUI(4), Tex_GUI(4), Tex_GUI(4)
    CreatePictureBox WindowCount, "picDown", 557, 280, 16, 16, , , , , Tex_GUI(5), Tex_GUI(5), Tex_GUI(5)
    CreatePictureBox WindowCount, "picLeft", 410, 180, 16, 16, , , , , Tex_GUI(12), Tex_GUI(12), Tex_GUI(12)
    CreatePictureBox WindowCount, "picRight", 720, 180, 16, 16, , , , , Tex_GUI(13), Tex_GUI(13), Tex_GUI(13)

    'labels idents controls
    CreateLabel WindowCount, "lblRun", 88, 80, 173, , "CORRER", Fonts.georgiaDec_16, White
    CreateLabel WindowCount, "lblFight", 188, 80, 173, , "ATACAR", Fonts.georgiaDec_16, White
    CreateLabel WindowCount, "lblTargetEnemy", 280, 80, 173, , "TARGET ENEMY", Fonts.georgiaDec_16, White
    
    CreateLabel WindowCount, "lblCharacter", 98, 140, 173, , "CHAR", Fonts.georgiaDec_16, White
    CreateLabel WindowCount, "lblBag", 198, 140, 173, , "BOLSA", Fonts.georgiaDec_16, White
    CreateLabel WindowCount, "lblSpells", 288, 140, 173, , "TECNICAS", Fonts.georgiaDec_16, White
    
    CreateLabel WindowCount, "lblGuild", 98, 200, 173, , "GUILD", Fonts.georgiaDec_16, White
    CreateLabel WindowCount, "lblQuest", 198, 200, 173, , "QUEST", Fonts.georgiaDec_16, White
    CreateLabel WindowCount, "lblOptions", 288, 200, 173, , "OPTIONS", Fonts.georgiaDec_16, White
    
    CreateLabel WindowCount, "lblGetItem", 88, 260, 173, , "GET ITEM", Fonts.georgiaDec_16, White
    CreateLabel WindowCount, "lblChat", 175, 260, 173, , "CHAT OPEN/SEND", Fonts.georgiaDec_16, White
    
    CreateLabel WindowCount, "lblHotbar", 388, 310, 173, , "HOTBARS", Fonts.georgiaDec_16, White

    ' Texts Boxs Actions / Key Controls - 28 To .ControlCount
    CreateTextbox WindowCount, "txtRun", 70, 95, 100, 32, KeycodeChar(Options.Correr), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtFight", 170, 95, 100, 32, KeycodeChar(Options.Atacar), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtTargetEnemy", 270, 95, 100, 32, KeycodeChar(Options.Target), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    
    CreateTextbox WindowCount, "txtCharacter", 70, 155, 100, 32, KeycodeChar(Options.Personagem), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtBag", 170, 155, 100, 32, KeycodeChar(Options.Bolsa), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtSpells", 270, 155, 100, 32, KeycodeChar(Options.Magias), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    
    CreateTextbox WindowCount, "txtGuild", 70, 215, 100, 32, KeycodeChar(Options.Guild), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtQuest", 170, 215, 100, 32, KeycodeChar(Options.Quests), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtOptions", 270, 215, 100, 32, KeycodeChar(Options.Options), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    
    CreateTextbox WindowCount, "txtGetItem", 70, 275, 100, 32, KeycodeChar(Options.PegarItem), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtChat", 170, 275, 100, 32, KeycodeChar(Options.Chat), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    
    'CreateTextbox WindowCount, "txtTarget", 270, 270, 100, 32, KeycodeChar(Options.Target), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtUp", 520, 115, 80, 32, KeycodeChar(Options.Up), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtDown", 520, 230, 80, 32, KeycodeChar(Options.Down), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtLeft", 430, 170, 80, 32, KeycodeChar(Options.Left), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    CreateTextbox WindowCount, "txtRight", 620, 170, 80, 32, KeycodeChar(Options.Right), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    ' Hotbars
    For i = 1 To MAX_HOTBAR
        CreateTextbox WindowCount, "txtHotbar" & i, 16 + (64 * i), 330, 64, 64, KeycodeChar(Options.Hotbar(i)), Fonts.georgiaDec_16, , Alignment.alignCentre, , , , , , DesignTypes.desTextBlack_Sq, DesignTypes.desTextBlack_Sq, DesignTypes.desTextBlack_Sq, , , GetAddress(AddressOf ChangeControls_Select), , , , , 12
    Next i
    
    ' Fora da contagem
    CreateCheckbox WindowCount, "chkSetas", 415, 280, 142, , Val(Options.UsarSetas), "Usar Setas.", Fonts.georgiaDec_16, , , , , DesignTypes.desChkNorm, , , GetAddress(AddressOf chkSetas_MouseDown)
    
End Sub

Public Sub ChangeControls_Restore()
    Dim i As Byte

    ' Reset Change Controls to Default
    Options.Correr = 16     'Shift
    Options.Atacar = 17     'Ctrl
    Options.Target = 9      'Tab
    
    Options.Personagem = 80 'P
    Options.Bolsa = 66      'B
    Options.Magias = 84     'T
    
    Options.Guild = 71      'G
    Options.Quests = 77     'M
    Options.Options = 27    'Esc
    
    Options.PegarItem = 32  'Space
    Options.Chat = 13       'Enter
    
    Options.Up = 87         'Arrow Up
    Options.Down = 83       'Arrow Down
    Options.Left = 65       'Arrow Left
    Options.Right = 68      'Arrow Right
    Options.UsarSetas = 1
    For i = 1 To MAX_HOTBAR
        If i < MAX_HOTBAR Then
            Options.Hotbar(i) = 48 + i
        Else
            Options.Hotbar(i) = 48
        End If
    Next i

    ' Change in form
    With Windows(GetWindowIndex("winChangeControls"))
        If .Window.visible Then
            .Controls(GetControlIndex("winChangeControls", "txtRun")).Text = KeycodeChar(Options.Correr)
            .Controls(GetControlIndex("winChangeControls", "txtFight")).Text = KeycodeChar(Options.Atacar)
            .Controls(GetControlIndex("winChangeControls", "txtTargetEnemy")).Text = KeycodeChar(Options.Target)
            
            .Controls(GetControlIndex("winChangeControls", "txtCharacter")).Text = KeycodeChar(Options.Personagem)
            .Controls(GetControlIndex("winChangeControls", "txtBag")).Text = KeycodeChar(Options.Bolsa)
            .Controls(GetControlIndex("winChangeControls", "txtSpells")).Text = KeycodeChar(Options.Magias)
            
            .Controls(GetControlIndex("winChangeControls", "txtGuild")).Text = KeycodeChar(Options.Guild)
            .Controls(GetControlIndex("winChangeControls", "txtQuest")).Text = KeycodeChar(Options.Quests)
            .Controls(GetControlIndex("winChangeControls", "txtOptions")).Text = KeycodeChar(Options.Options)
            
            .Controls(GetControlIndex("winChangeControls", "txtGetItem")).Text = KeycodeChar(Options.PegarItem)
            .Controls(GetControlIndex("winChangeControls", "txtChat")).Text = KeycodeChar(Options.Chat)
            For i = 1 To MAX_HOTBAR
                .Controls(GetControlIndex("winChangeControls", "txtHotbar" & i)).Text = KeycodeChar(Options.Hotbar(i))
            Next i
            .Controls(GetControlIndex("winChangeControls", "txtUp")).Text = KeycodeChar(Options.Up)
            .Controls(GetControlIndex("winChangeControls", "txtDown")).Text = KeycodeChar(Options.Down)
            .Controls(GetControlIndex("winChangeControls", "txtLeft")).Text = KeycodeChar(Options.Left)
            .Controls(GetControlIndex("winChangeControls", "txtRight")).Text = KeycodeChar(Options.Right)
            .Controls(GetControlIndex("winChangeControls", "chkSetas")).Value = Options.UsarSetas
        End If
    End With
End Sub

Public Sub chkSetas_MouseDown()
    Options.UsarSetas = Windows(GetWindowIndex("winChangeControls")).Controls(GetControlIndex("winChangeControls", "chkSetas")).Value
End Sub

' Conversão dos dados em variáveis e tratamento do texto a mostrar!
Public Sub HandleKeyCodeControls(KeyCode As Integer, Shift As Integer)

    With Windows(GetWindowIndex("winChangeControls"))
        If .activeControl = 0 Then Exit Sub

        .Controls(.activeControl).Text = KeycodeChar(KeyCode)

        Select Case .activeControl
        Case 28
            Options.Correr = KeyCode
        Case 29
            Options.Atacar = KeyCode
        Case 30
            Options.Target = KeyCode
        Case 31
            Options.Personagem = KeyCode
        Case 32
            Options.Bolsa = KeyCode
        Case 33
            Options.Magias = KeyCode
        Case 34
            Options.Guild = KeyCode
        Case 35
            Options.Quests = KeyCode
        Case 36
            Options.Options = KeyCode
        Case 37
            Options.PegarItem = KeyCode
        Case 38
            Options.Chat = KeyCode
        Case 39
            Options.Up = KeyCode
        Case 40
            Options.Down = KeyCode
        Case 41
            Options.Left = KeyCode
        Case 42
            Options.Right = KeyCode
        Case 43
            Options.Hotbar(1) = KeyCode
        Case 44
            Options.Hotbar(2) = KeyCode
        Case 45
            Options.Hotbar(3) = KeyCode
        Case 46
            Options.Hotbar(4) = KeyCode
        Case 47
            Options.Hotbar(5) = KeyCode
        Case 48
            Options.Hotbar(6) = KeyCode
        Case 49
            Options.Hotbar(7) = KeyCode
        Case 50
            Options.Hotbar(8) = KeyCode
        Case 51
            Options.Hotbar(9) = KeyCode
        Case 52
            Options.Hotbar(10) = KeyCode
        End Select
    End With
End Sub

' Função que obtém strings mais agradáveis p/ identação!
Public Function KeycodeChar(ByVal NumKeyCode As Integer) As String
    Select Case NumKeyCode
    Case 8
        KeycodeChar = "BACKSPACE"
        Exit Function
    Case 9
        KeycodeChar = "TAB"
        Exit Function
    Case 13
        KeycodeChar = "ENTER"
        Exit Function
    Case 16
        KeycodeChar = "SHIFT"
        Exit Function
    Case 17
        KeycodeChar = "CTRL"
        Exit Function
    Case 18
        KeycodeChar = "ALT"
        Exit Function
    Case 20
        KeycodeChar = "CAPS"
        Exit Function
    Case 27
        KeycodeChar = "ESC"
        Exit Function
    Case 32
        KeycodeChar = "SPACE"
        Exit Function
    Case 33
        KeycodeChar = "PG UP"
        Exit Function
    Case 34
        KeycodeChar = "PG DOWN"
        Exit Function
    Case 35
        KeycodeChar = "END"
        Exit Function
    Case 36
        KeycodeChar = "HOME"
        Exit Function
    Case 37
        KeycodeChar = "LEFT"
        Exit Function
    Case 38
        KeycodeChar = "UP"
        Exit Function
    Case 39
        KeycodeChar = "RIGHT"
        Exit Function
    Case 40
        KeycodeChar = "DOWN"
        Exit Function
    Case 45
        KeycodeChar = "INSERT"
        Exit Function
    Case 46
        KeycodeChar = "DELETE"
        Exit Function
    Case 91
        KeycodeChar = "WINDOWS"
        Exit Function
    Case 92
        KeycodeChar = "WINDOWS"
        Exit Function
    Case 93
        KeycodeChar = "LIST"
        Exit Function
    Case 96
        KeycodeChar = "NumPad 0"
        Exit Function
    Case 97
        KeycodeChar = "NumPad 1"
        Exit Function
    Case 98
        KeycodeChar = "NumPad 2"
        Exit Function
    Case 99
        KeycodeChar = "NumPad 3"
        Exit Function
    Case 100
        KeycodeChar = "NumPad 4"
        Exit Function
    Case 101
        KeycodeChar = "NumPad 5"
        Exit Function
    Case 102
        KeycodeChar = "NumPad 6"
        Exit Function
    Case 103
        KeycodeChar = "NumPad 7"
        Exit Function
    Case 104
        KeycodeChar = "NumPad 8"
        Exit Function
    Case 105
        KeycodeChar = "NumPad 9"
        Exit Function
    Case 106
        KeycodeChar = "*"
        Exit Function
    Case 107
        KeycodeChar = "+"
        Exit Function
    Case 109
        KeycodeChar = "-"
        Exit Function
    Case 110
        KeycodeChar = ","
        Exit Function
    Case 111
        KeycodeChar = "/"
        Exit Function
    Case 112
        KeycodeChar = "F1"
        Exit Function
    Case 113
        KeycodeChar = "F2"
        Exit Function
    Case 114
        KeycodeChar = "F3"
        Exit Function
    Case 115
        KeycodeChar = "F4"
        Exit Function
    Case 116
        KeycodeChar = "F5"
        Exit Function
    Case 117
        KeycodeChar = "F6"
        Exit Function
    Case 118
        KeycodeChar = "F7"
        Exit Function
    Case 119
        KeycodeChar = "F8"
        Exit Function
    Case 120
        KeycodeChar = "F9"
        Exit Function
    Case 121
        KeycodeChar = "F10"
        Exit Function
    Case 122
        KeycodeChar = "F11"
        Exit Function
    Case 123
        KeycodeChar = "F12"
        Exit Function
    Case 187
        KeycodeChar = "="
        Exit Function
    Case 188
        KeycodeChar = ","
        Exit Function
    Case 189
        KeycodeChar = "-"
        Exit Function
    Case 190
        KeycodeChar = "."
        Exit Function
    Case 191
        KeycodeChar = ";"
        Exit Function
    Case 192
        KeycodeChar = "'"
        Exit Function
    Case 193
        KeycodeChar = "/"
        Exit Function
    Case 194
        KeycodeChar = "."
        Exit Function
    Case 219
        KeycodeChar = "´"
        Exit Function
    Case 220
        KeycodeChar = "]"
        Exit Function
    Case 221
        KeycodeChar = "["
        Exit Function
    Case 222
        KeycodeChar = "~"
        Exit Function
    End Select

    KeycodeChar = ChrW$(NumKeyCode)
End Function

'vbKeyLButton    Left Mouse Button
'vbKeyRButton    Right Mouse Button
'vnKeyCancel     Cancel Key
'vbKeyMButton    Middle Mouse button
'vbKeyBack       Back Space Key
'vbKeyTab        Tab Key
'vbKeyClear      Clear Key
'vbKeyReturn     Enter Key
'vbKeyShift      Shift Key
'vbKeyControl    Ctrl Key
'vbKeyMenu       Menu Key
'vbKeyPause      Pause Key
'vbKeyCapital    Caps Lock Key
'vbKeyEscape     Escape Key
'vbKeySpace      Spacebar Key
'vbKeyPageUp     Page Up Key
'vbKeyPageDown   Page Down Key
'vbKeyEnd        End Key
'vbKeyHome       Home Key
'vbKeyLeft       Left Arrow Key
'vbKeyUp         Up Arrow Key
'vbKeyRight      Right Arrow Key
'vbKeyDown       Down Arrow Key
'vbKeySelect     Select Key
'vbKeyPrint      Print Screen Key
'vbKeyExecute    Execute Key
'vbKeySnapshot   Snapshot Key
'vbKeyInsert     Insert Key
'vbKeyDelete     Delete Key
'vbKeyHelp       Help Key
'vbKeyNumlock    Delete Key

'vbKeyA through vbKeyZ are the key code constants for the alphabet
'vbKey0 through vbKey9 are the key code constants for numbers
'vbKeyF1 through vbKeyF16 are the key code constants for the function keys
'vbKeyNumpad0 through vbKeyNumpad9 are the key code constants for the numeric key pad

'Math signs are:
'vbKeyMultiply      -  Multiplication Sign (*)
'vbKeyAdd             - Addition Sign (+)
'vbKeySubtract     - Minus Sign (-)
'vbKeyDecimal    - Decimal Point (.)
'vbKeyDivide        - Division sign (/)
'vbKeySeparator  - Enter (keypad) sign
