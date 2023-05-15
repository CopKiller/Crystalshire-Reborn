Attribute VB_Name = "modMouseWheel"
Option Explicit

'***********************************************************************************
'*****THIS SOURCE HAS BEEN OBTAINED FROM http://www.andreavb.com/tip060008.html*****
'***********************************************************************************

'************************************************************
'API
'************************************************************
Private Declare Function CallWindowProc _
                          Lib "user32.dll" _
                              Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                       ByVal hwnd As Long, _
                                                       ByVal Msg As Long, _
                                                       ByVal wParam As Long, _
                                                       ByVal lParam As Long) As Long

Private Declare Function SetWindowLong _
                          Lib "user32.dll" _
                              Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

'************************************************************
'Constants
'************************************************************

Public Const MK_CONTROL = &H8

Public Const MK_LBUTTON = &H1

Public Const MK_RBUTTON = &H2

Public Const MK_MBUTTON = &H10

Public Const MK_SHIFT = &H4

Private Const GWL_WNDPROC = -4

Private Const WM_MOUSEWHEEL = &H20A

'************************************************************
'Variables
'************************************************************

Private hControl As Long

Private lPrevWndProc As Long

' Chat options

Public ChatMouseScroll As Boolean

Public IsLeftMouseButtonDown As Boolean
Public Const ChatScroll_MinY As Byte = 20
Public Const ChatScroll_MaxY As Byte = 60

'*************************************************************
'WindowProc
'*************************************************************

Private Function WindowProc(ByVal lWnd As Long, _
                            ByVal lMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

'Test if the message is WM_MOUSEWHEEL

    If lMsg = WM_MOUSEWHEEL And frmMain.hwnd = myHWnd Then

        'Add event handling code here
        'this will be universal to all forms that are 'hooked' to this code
        If wParam > 0 Then    ' Moved up
            HandleMouseWheelMove 0
        Else    ' Moved down
            HandleMouseWheelMove 1
        End If
    End If

    'Sends message to previous procedure if not MOUSEWHEEL
    'This is VERY IMPORTANT!!!
    If lMsg <> WM_MOUSEWHEEL Then
        WindowProc = CallWindowProc(lPrevWndProc, lWnd, lMsg, wParam, lParam)
    End If

End Function

'*************************************************************
'Hook
'All forms that call this procedure must implement this procedure in their module:
'   Public Sub MouseWheelRolled()
'        <your code>
'   End Sub
'*************************************************************
Public Sub HookForMouseWheel(ByVal hControl_ As Long)

    hControl = hControl_
    lPrevWndProc = SetWindowLong(hControl, GWL_WNDPROC, AddressOf WindowProc)

End Sub

'*************************************************************
'UnHook
'*************************************************************
Public Sub UnHookMouseWheel()

    If hControl = 0 Then Exit Sub    ' Don't do anything if we haven't hooked yet
    Call SetWindowLong(hControl, GWL_WNDPROC, lPrevWndProc)

End Sub

Public Sub HandleMouseWheelMove(ByVal UpOrDown As Byte)

    If InGame = True Then
        Chat_Scroll UpOrDown
    End If

End Sub

Public Sub Chat_Scroll(ByVal UpOrDown As Byte)

    With Windows(GetWindowIndex("winChat"))

        If (GlobalX >= .Window.Left And GlobalX <= .Window.Left + .Window.Width) And (GlobalY >= .Window.top And GlobalY <= .Window.top + .Window.Height) Then
            If UpOrDown = 0 Then  ' Down.
                ChatButtonUp = True    ' Don't know why but this boolean is inverted. It works though
                ChatMouseScroll = True
            Else
                ChatButtonDown = True    ' Don't know why but this boolean is inverted. It works though
                ChatMouseScroll = True
            End If
        End If

    End With

End Sub

'GetAddress(AddressOf ChatScroll_MouseDown), GetAddress(AddressOf ChatScroll_MouseMove)
Public Sub ChatScroll_MouseDown()
'    Dim winIndex As Long, xO As Long, yO As Long

'    winIndex = GetWindowIndex("winChat")
'    If Not Windows(winIndex).Window.visible Then Exit Sub

    IsLeftMouseButtonDown = True

    'Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkEvent")).Value
End Sub

Public Sub ChatScroll_MouseMove()
    Dim winIndex As Long, controlIndex As Long, xO As Long, yO As Long

    winIndex = GetWindowIndex("winChat")
    controlIndex = GetControlIndex("winChat", "btnScroll")

    ' A janela de chat não está aberta.
    If Not Windows(winIndex).Window.visible Then Exit Sub
    ' O controle de scroll não está visível.
    If Not Windows(winIndex).Controls(controlIndex).visible Then Exit Sub
    ' A quantidade de texto no chat não necessita do scroll.
    If Chat_HighIndex <= ChatHeight_Lines Then Exit Sub

    ' O Botão Left do mouse está pressionado?
    If IsLeftMouseButtonDown Then
        ' Down
        If GlobalY > (Windows(winIndex).Window.top + Windows(winIndex).Controls(controlIndex).top + (Windows(winIndex).Controls(controlIndex).Height / 2)) Then
            If Windows(winIndex).Controls(controlIndex).top <= ChatScroll_MaxY Then
                Windows(winIndex).Controls(controlIndex).top = Windows(winIndex).Controls(controlIndex).top + 2
                ScrollChatBox 1
            End If
            ' Up
        Else
            If Windows(winIndex).Controls(controlIndex).top >= ChatScroll_MinY Then
                Windows(winIndex).Controls(controlIndex).top = Windows(winIndex).Controls(controlIndex).top - 2
                ScrollChatBox 0
            End If
        End If
    End If
End Sub

'Public Sub ScrollChatBox(ByVal direction As Byte)
'    If direction = 0 Then    ' up
'        If ChatScroll < Chat_HighIndex - ChatHeight_Lines Then
'            ChatScroll = ChatScroll + 1
'        End If
'    Else
'        If ChatScroll > 0 Then
'            ChatScroll = ChatScroll - 1
'        End If
'    End If
'End Sub

'Public Function IsLeftMouseButtonDown() As Boolean
' Verifica se o botão esquerdo do mouse está pressionado
'    If GetAsyncKeyState(VK_LBUTTON) And &H8000 Then
'        IsLeftMouseButtonDown = True
'    Else
'        IsLeftMouseButtonDown = False
'    End If
'End Function

'Public Const VK_LBUTTON As Long = &H1   ' Botão esquerdo do mouse
'Public Const VK_RBUTTON As Long = &H2   ' Botão direito do mouse
'Public Const VK_MBUTTON As Long = &H4   ' Botão do meio do mouse
