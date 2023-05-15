Attribute VB_Name = "modClipBoard"
Option Explicit

Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyW Lib "kernel32.dll" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long

Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const CF_UNICODETEXT As Long = 13

'KeyState
Public Const VK_CONTROL As Long = &H11
Public Const VK_V As Long = &H56
Public Ctrl_V As Boolean

Private Type EntityClipboard
    winIndex As Long
    controlIndex As Long
End Type

Private currentControl As EntityClipboard

Public Function GetClipboardText() As String
    Dim hwnd As Long
    Dim hData As Long
    Dim lpData As Long

    ' Abre a área de transferência
    OpenClipboard 0&

    ' Verifica se o formato CF_UNICODETEXT está disponível
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        ' Obtém o handle do dado da área de transferência
        hData = GetClipboardData(CF_UNICODETEXT)

        ' Verifica se obteve o handle do dado com sucesso
        If hData <> 0 Then
            ' Trava o handle para obter acesso à memória
            lpData = GlobalLock(hData)

            ' Verifica se obteve acesso à memória com sucesso
            If lpData <> 0 Then
                ' Obtém o comprimento do texto
                Dim length As Long
                length = lstrlenW(lpData)

                ' Cria uma variável para armazenar o texto
                Dim buffer As String
                buffer = String$(length, vbNullChar)

                ' Copia o texto para a variável
                lstrcpyW StrPtr(buffer), lpData

                ' Libera o acesso à memória
                GlobalUnlock hData

                ' Define o resultado como o conteúdo do buffer
                GetClipboardText = buffer
            End If
        End If
    End If

    ' Fecha a área de transferência
    CloseClipboard
End Function

Public Sub CopyToClipboard(ByVal text As String)
' Abre a área de transferência
    OpenClipboard 0&
    ' Limpa o conteúdo atual da área de transferência
    EmptyClipboard

    ' Calcula o comprimento necessário para o texto em bytes
    Dim textLength As Long
    textLength = (Len(text) + 1) * 2    ' Multiplica por 2 para acomodar caracteres Unicode

    ' Aloca memória para o texto na área de transferência
    Dim hMem As Long
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, textLength)

    If hMem <> 0 Then
        ' Bloqueia a memória para obter acesso
        Dim lpMem As Long
        lpMem = GlobalLock(hMem)

        If lpMem <> 0 Then
            ' Copia o texto para a memória
            lstrcpyW lpMem, StrPtr(text)
            ' Desbloqueia a memória
            GlobalUnlock hMem

            ' Define o texto como o conteúdo da área de transferência
            SetClipboardData CF_UNICODETEXT, hMem
        End If
    End If

    ' Fecha a área de transferência
    CloseClipboard
End Sub

Public Sub CreateWindow_Clipboard()
    CreateWindow "winClipboard", "CopyPaste", zOrder_Win, 0, 0, 80, 74, Tex_Item(1), True, Fonts.rockwellDec_15, , 2, 5, , , , , , , , , , , , False, False
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 17, 4, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnClip_Close)

    CreateButton WindowCount, "btnCopy", 2, 17, 76, 15, "Copy", Fonts.georgiaBold_16, , , , , , , , DesignTypes.desBlue, DesignTypes.desBlue_Hover, DesignTypes.desBlue_Click, , , GetAddress(AddressOf btnCopy_Click)
    CreateButton WindowCount, "btnPaste", 2, 33, 76, 15, "Paste", Fonts.georgiaBold_16, , , , , , , , DesignTypes.desBlue, DesignTypes.desBlue_Hover, DesignTypes.desBlue_Click, , , GetAddress(AddressOf btnPaste_Click)
    CreateButton WindowCount, "btnClear", 2, 49, 76, 15, "Clear", Fonts.georgiaBold_16, , , , , , , , DesignTypes.desBlue, DesignTypes.desBlue_Hover, DesignTypes.desBlue_Click, , , GetAddress(AddressOf btnClear_Click)

End Sub

Private Sub btnClip_Close()
    HideWindow GetWindowIndex("winClipboard")
End Sub

Private Sub btnCopy_Click()
    CopyToClipboard Windows(currentControl.winIndex).Controls(currentControl.controlIndex).text
End Sub

Private Sub btnClear_Click()
    Windows(currentControl.winIndex).Controls(currentControl.controlIndex).text = vbNullString
End Sub

Private Sub btnPaste_Click()
    ' Obter o texto atual do controle
    Dim SString As String
    SString = Windows(currentControl.winIndex).Controls(currentControl.controlIndex).text

    ' Obter o texto da área de transferência
    Dim ClipboardText As String
    ClipboardText = GetClipboardText

    ' Obter o comprimento do texto atual do controle
    Dim lText As Integer
    lText = Len(SString)

    ' Obter o comprimento do texto na área de transferência
    Dim lClip As Integer
    lClip = Len(ClipboardText)

    ' Calcular o espaço disponível para colar o texto
    Dim availableSpace As Integer
    availableSpace = Windows(currentControl.winIndex).Controls(currentControl.controlIndex).max - lText

    ' Verificar se há espaço suficiente para colar o texto
    If availableSpace > 0 Then
        ' Atualizar o espaço disponível para o comprimento do texto na área de transferência
        If availableSpace > lClip Then
            availableSpace = lClip
        End If

        ' Obter uma parte do texto da área de transferência com base no espaço disponível
        Dim pasteText As String
        pasteText = Left$(ClipboardText, availableSpace)

        ' Colar o texto no controle
        Windows(currentControl.winIndex).Controls(currentControl.controlIndex).text = SString & pasteText
    End If
End Sub

Public Sub Clipboard(ByVal winIndex As Long, ByVal controlIndex As Long)
    
    ShowWindow GetWindowIndex("winClipboard"), True
    
    Windows(GetWindowIndex("winClipboard")).Window.top = GlobalY
    Windows(GetWindowIndex("winClipboard")).Window.Left = GlobalX
    
    currentControl.winIndex = winIndex
    currentControl.controlIndex = controlIndex
End Sub
