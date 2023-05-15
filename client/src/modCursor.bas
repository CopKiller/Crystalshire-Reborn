Attribute VB_Name = "modCursor"
Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As CURSORINFO) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type CURSORINFO
    cbSize As Long
    flags As Long
    hCursor As Long
    ptScreenPos As POINTAPI
End Type

Public Enum CustomCursores
    Default = 1
    Button
    text

    Cursores_Count
End Enum

Private Const CURSOR_SHOWING As Long = &H1

Public MouseIcon As Byte

Public Sub CursorVisibility(ByVal isVisible As Boolean)
    If App.LogMode <> 1 Then Exit Sub
    Call ShowCursor(isVisible)
End Sub

Public Sub SetCursor(ByVal Icon As CustomCursores)
    MouseIcon = Icon
End Sub

Public Sub ResetMouseIcon()
    Dim i As Byte

    For i = 2 To Cursores_Count - 1
        If MouseIcon = i Then
            SetCursor Default
        End If
    Next i
End Sub

Public Sub RenderCursor()
    Dim XDiff As Integer

    'If Not IsCursorVisible Then
    '    Call CursorVisibility(True)
    'End If
    If App.LogMode <> 1 Then Exit Sub

    ' Check if the MouseIcon is set to 0 (indicating cursor outside screen)
    If MouseIcon = 0 Then
        ' Check if the mouse is not outside the screen
        If Not IsMouseOutsideScreen Then
            ' Set the MouseIcon to the default cursor
            MouseIcon = CustomCursores.Default
            ' Hide the cursor
            CursorVisibility False
        End If
    Else    ' MouseIcon <> 0 (indicating cursor inside screen)
        ' Check if the mouse is outside the screen
        If IsMouseOutsideScreen Then
            ' Set the MouseIcon to 0 (no cursor)
            MouseIcon = 0
            ' Show the cursor
            CursorVisibility True
        End If

        ' Set Cursor Fix Position
        If MouseIcon = CustomCursores.Button Then XDiff = -5 Else XDiff = 0
        ' Render the cursor texture
        RenderTexture Tex_Cursor(MouseIcon), GlobalX + XDiff, GlobalY, 0, 0, 32, 32, 32, 32
    End If
End Sub

Private Function IsMouseOutsideScreen() As Boolean
' Check if the GlobalX is outside the screen width boundaries
    If GlobalX < 0 Or GlobalX > screenWidth Then
        ' If so, return True indicating mouse is outside the screen
        IsMouseOutsideScreen = True
        Exit Function    ' Exit the function early
    End If

    ' Check if the GlobalY is outside the screen height boundaries
    If GlobalY < 0 Or GlobalY > screenHeight Then
        ' If so, return True indicating mouse is outside the screen
        IsMouseOutsideScreen = True
        Exit Function    ' Exit the function early
    End If

    ' If the mouse is within the screen boundaries, return False
    IsMouseOutsideScreen = False
End Function

Public Function IsCursorVisible() As Boolean
    Dim ci As CURSORINFO

    ci.cbSize = Len(ci)

    If GetCursorInfo(ci) Then
        If (ci.flags And CURSOR_SHOWING) <> 0 Then
            IsCursorVisible = True
        End If
    End If
End Function
