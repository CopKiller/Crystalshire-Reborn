Attribute VB_Name = "modAutoLogin"
Option Explicit

Public isReconnect As Boolean
Private tmr13000 As Long
Private tmr1000 As Long
Private Reconnects As Long

Public Sub CreateWindow_Reconnect()
' Cria a janela de reconexão
    CreateWindow "winReconnect", "Problemas na conexão...", zOrder_Win, 0, 0, 278, 130, Tex_Item(104), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centraliza a janela
    CentraliseWindow WindowCount

    ' Define o índice para criar controles
    zOrder_Con = 1

    ' Pergaminho
    CreatePictureBox WindowCount, "picParchment", 6, 26, 266, 98, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Fundo do texto
    CreatePictureBox WindowCount, "picReconnect", 26, 39, 226, 52, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    ' Rótulo
    CreateLabel WindowCount, "lblReconnect", 6, 43, 266, , "Por favor aguarde, reconectando!", rockwell_15, , Alignment.alignCentre
    CreateLabel WindowCount, "lblTentativas", 6, 58, 266, , "Tentativas realizadas: 0", rockwell_15, , Alignment.alignCentre
    CreateLabel WindowCount, "lblTentativas1", 6, 70, 266, , "Tentando novamente em: 10 Segs", rockwell_15, , Alignment.alignCentre
    ' Botão
    CreateButton WindowCount, "btnCancel", ((Windows(WindowCount).Window.Width / 2) - (68 / 2)), (Windows(WindowCount).Window.Height - 40), 68, 24, "Cancelar", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf CancelReconnect)
End Sub

Public Sub InitReconnect()
' Reconectar caso tenha ativado o reconectar automático
    If Options.Reconnect = YES Then
        HideWindows
        ShowWindow GetWindowIndex("winReconnect"), True
        isReconnect = True
        logoutGame
    Else
        ' Servidor inativo, não reconectar. Fazer logout do jogo
        HideWindows
        ShowWindow GetWindowIndex("winLogin")
        logoutGame
    End If
End Sub

Public Sub ProccessReconnect()
    Dim Tick As Currency

    Tick = getTime

    ' Está tentando reconectar?
    With Windows(GetWindowIndex("winReconnect"))
        If isReconnect Then
            If tmr13000 < Tick Then
                If Reconnects > 0 Then
                    Call Login(Options.TmpLogin, Options.TmpPassword)
                End If
                .Controls(GetControlIndex("winReconnect", "lblTentativas")).text = "Tentativas realizadas: " & Reconnects
                .Controls(GetControlIndex("winReconnect", "lblTentativas1")).text = "Tentando novamente em: " & ((tmr13000 - Tick) \ 1000) & " Segs"
                Reconnects = Reconnects + 1
                tmr13000 = Tick + 13000    '13 segundos uma nova tentativa de login
            ElseIf tmr1000 < Tick Then
                .Controls(GetControlIndex("winReconnect", "lblTentativas1")).text = "Tentando novamente em: " & ((tmr13000 - Tick) \ 1000) & " Segs"
                tmr1000 = Tick + 1000      '1 segundo atualiza a contagem na label
            End If
        End If
    End With
End Sub

Private Sub CancelReconnect()
    isReconnect = False
    ResetReconnectVariables
    HideWindow GetWindowIndex("winReconnect")
    ShowWindow GetWindowIndex("winLogin")
End Sub

Public Sub ResetReconnectVariables()
    tmr13000 = 0
    tmr1000 = 0
    Reconnects = 0
End Sub
