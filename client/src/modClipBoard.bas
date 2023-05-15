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

    ' Abre a �rea de transfer�ncia
    OpenClipboard 0&

    ' Verifica se o formato CF_UNICODETEXT est� dispon�vel
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        ' Obt�m o handle do dado da �rea de transfer�ncia
        hData = GetClipboardData(CF_UNICODETEXT)

        ' Verifica se obteve o handle do dado com sucesso
        If hData <> 0 Then
            ' Trava o handle para obter acesso � mem�ria
            lpData = GlobalLock(hData)

            ' Verifica se obteve acesso � mem�ria com sucesso
            If lpData <> 0 Then
                ' Obt�m o comprimento do texto
                Dim length As Long
                length = lstrlenW(lpData)

                ' Cria uma vari�vel para armazenar o texto
                Dim buffer As String
                buffer = String$(length, vbNullChar)

                ' Copia o texto para a vari�vel
                lstrcpyW StrPtr(buffer), lpData

                ' Libera o acesso � mem�ria
                GlobalUnlock hData

                ' Define o resultado como o conte�do do buffer
                GetClipboardText = buffer
            End If
        End If
    End If

    ' Fecha a �rea de transfer�ncia
    CloseClipboard
End Function

Public Sub CopyToClipboard(ByVal text As String)
' Abre a �rea de transfer�ncia
    OpenClipboard 0&
    ' Limpa o conte�do atual da �rea de transfer�ncia
    EmptyClipboard

    ' Calcula o comprimento necess�rio para o texto em bytes
    Dim textLength As Long
    textLength = (Len(text) + 1) * 2    ' Multiplica por 2 para acomodar caracteres Unicode

    ' Aloca mem�ria para o texto na �rea de transfer�ncia
    Dim hMem As Long
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, textLength)

    If hMem <> 0 Then
        ' Bloqueia a mem�ria para obter acesso
        Dim lpMem As Long
        lpMem = GlobalLock(hMem)

        If lpMem <> 0 Then
            ' Copia o texto para a mem�ria
            lstrcpyW lpMem, StrPtr(text)
            ' Desbloqueia a mem�ria
            GlobalUnlock hMem

            ' Define o texto como o conte�do da �rea de transfer�ncia
            SetClipboardData CF_UNICODETEXT, hMem
        End If
    End If

    ' Fecha a �rea de transfer�ncia
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

    ' Obter o texto da �rea de transfer�ncia
    Dim ClipboardText As String
    ClipboardText = GetClipboardText

    ' Obter o comprimento do texto atual do controle
    Dim lText As Integer
    lText = Len(SString)

    ' Obter o comprimento do texto na �rea de transfer�ncia
    Dim lClip As Integer
    lClip = Len(ClipboardText)

    ' Calcular o espa�o dispon�vel para colar o texto
    Dim availableSpace As Integer
    availableSpace = Windows(currentControl.winIndex).Controls(currentControl.controlIndex).max - lText

    ' Verificar se h� espa�o suficiente para colar o texto
    If availableSpace > 0 Then
        ' Atualizar o espa�o dispon�vel para o comprimento do texto na �rea de transfer�ncia
        If availableSpace > lClip Then
            availableSpace = lClip
        End If

        ' Obter uma parte do texto da �rea de transfer�ncia com base no espa�o dispon�vel
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
