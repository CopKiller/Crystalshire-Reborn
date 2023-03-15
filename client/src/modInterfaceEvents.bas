Attribute VB_Name = "modInterfaceEvents"
Option Explicit
Public Declare Sub GetCursorPos Lib "user32" (lpPoint As POINTAPI)
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function EntCallBack Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Window As Long, ByRef Control As Long, ByVal forced As Long, ByVal lParam As Long) As Long
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public lastMouseX As Long, lastMouseY As Long
Attribute lastMouseY.VB_VarUserMemId = 1073741824
Public currMouseX As Long, currMouseY As Long
Attribute currMouseX.VB_VarUserMemId = 1073741826
Attribute currMouseY.VB_VarUserMemId = 1073741826
Public clickedX As Long, clickedY As Long
Attribute clickedX.VB_VarUserMemId = 1073741828
Attribute clickedY.VB_VarUserMemId = 1073741828
Public mouseClick(1 To 2) As Long
Attribute mouseClick.VB_VarUserMemId = 1073741830
Public lastMouseClick(1 To 2) As Long
Attribute lastMouseClick.VB_VarUserMemId = 1073741831

Public GlobalCaptcha As Long
Attribute GlobalCaptcha.VB_VarUserMemId = 1073741832


Public Function MouseX(Optional ByVal hWnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint

    If hWnd Then ScreenToClient hWnd, lpPoint
    MouseX = lpPoint.X
End Function

Public Function MouseY(Optional ByVal hWnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint

    If hWnd Then ScreenToClient hWnd, lpPoint
    MouseY = lpPoint.Y
End Function

Public Sub HandleMouseInput()
    Dim entState As entStates, i As Long, X As Long

    ' exit out if we're playing video
    If videoPlaying Then Exit Sub

    ' set values
    lastMouseX = currMouseX
    lastMouseY = currMouseY
    currMouseX = MouseX(frmMain.hWnd)
    currMouseY = MouseY(frmMain.hWnd)
    GlobalX = currMouseX
    GlobalY = currMouseY
    lastMouseClick(VK_LBUTTON) = mouseClick(VK_LBUTTON)
    lastMouseClick(VK_RBUTTON) = mouseClick(VK_RBUTTON)
    mouseClick(VK_LBUTTON) = GetAsyncKeyState(VK_LBUTTON)
    mouseClick(VK_RBUTTON) = GetAsyncKeyState(VK_RBUTTON)

    ' Hover
    entState = entStates.Hover

    ' MouseDown
    If (mouseClick(VK_LBUTTON) And lastMouseClick(VK_LBUTTON) = 0) Or (mouseClick(VK_RBUTTON) And lastMouseClick(VK_RBUTTON) = 0) Then
        clickedX = currMouseX
        clickedY = currMouseY
        entState = entStates.MouseDown
        ' MouseUp
    ElseIf (mouseClick(VK_LBUTTON) = 0 And lastMouseClick(VK_LBUTTON)) Or (mouseClick(VK_RBUTTON) = 0 And lastMouseClick(VK_RBUTTON)) Then
        entState = entStates.MouseUp
        ' MouseMove
    ElseIf (currMouseX <> lastMouseX) Or (currMouseY <> lastMouseY) Then
        entState = entStates.MouseMove
    End If

    ' Handle everything else
    If Not HandleGuiMouse(entState) Then
        ' reset /all/ control mouse events
        For i = 1 To WindowCount
            For X = 1 To Windows(i).ControlCount
                Windows(i).Controls(X).state = Normal
            Next
        Next
        If InGame Then
            If entState = entStates.MouseDown Then
                ' Handle events
                If currMouseX >= 0 And currMouseX <= frmMain.ScaleWidth Then
                    If currMouseY >= 0 And currMouseY <= frmMain.ScaleHeight Then
                        If InMapEditor Then
                            If (mouseClick(VK_LBUTTON) And lastMouseClick(VK_LBUTTON) = 0) Then
                                If frmEditor_Map.optEvents.Value Then
                                    selTileX = CurX
                                    selTileY = CurY
                                Else
                                    Call MapEditorMouseDown(vbLeftButton, GlobalX, GlobalY, False)
                                End If
                            ElseIf (mouseClick(VK_RBUTTON) And lastMouseClick(VK_RBUTTON) = 0) Then
                                If Not frmEditor_Map.optEvents.Value Then Call MapEditorMouseDown(vbRightButton, GlobalX, GlobalY, False)
                            End If
                        Else
                            ' left click
                            If (mouseClick(VK_LBUTTON) And lastMouseClick(VK_LBUTTON) = 0) Then
                                ' targetting
                                FindTarget
                                ' right click
                            ElseIf (mouseClick(VK_RBUTTON) And lastMouseClick(VK_RBUTTON) = 0) Then
                                If ShiftDown Then
                                    ' admin warp if we're pressing shift and right clicking
                                    If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
                                    Exit Sub
                                End If
                                ' right-click menu
                                For i = 1 To Player_HighIndex
                                    If IsPlaying(i) Then
                                        If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                                            If GetPlayerX(i) = CurX And GetPlayerY(i) = CurY Then
                                                ShowPlayerMenu i, currMouseX, currMouseY
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            ElseIf entState = entStates.MouseMove Then
                GlobalX_Map = GlobalX + (TileView.Left * PIC_X) + Camera.Left
                GlobalY_Map = GlobalY + (TileView.top * PIC_Y) + Camera.top
                ' Handle the events
                CurX = TileView.Left + ((currMouseX + Camera.Left) \ PIC_X)
                CurY = TileView.top + ((currMouseY + Camera.top) \ PIC_Y)

                If InMapEditor Then
                    If (mouseClick(VK_LBUTTON)) Then
                        If Not frmEditor_Map.optEvents.Value Then Call MapEditorMouseDown(vbLeftButton, CurX, CurY, False)
                    ElseIf (mouseClick(VK_RBUTTON)) Then
                        If Not frmEditor_Map.optEvents.Value Then Call MapEditorMouseDown(vbRightButton, CurX, CurY, False)
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Function HandleGuiMouse(entState As entStates) As Boolean
    Dim i As Long, curWindow As Long, curControl As Long, Callback As Long, X As Long

    ' if hiding gui
    If hideGUI = True Or InMapEditor Then Exit Function

    ' Find the container
    For i = 1 To WindowCount
        With Windows(i).Window
            If .enabled And .visible Then
                If .state <> entStates.MouseDown Then .state = entStates.Normal
                If currMouseX >= .Left And currMouseX <= .Width + .Left Then
                    If currMouseY >= .top And currMouseY <= .Height + .top Then
                        ' set the combomenu
                        If .design(0) = DesignTypes.desComboMenuNorm Then
                            ' set the hover menu
                            If entState = MouseMove Or entState = Hover Then
                                ComboMenu_MouseMove i
                            ElseIf entState = MouseDown Then
                                ComboMenu_MouseDown i
                            End If
                        End If
                        ' everything else
                        If curWindow = 0 Then curWindow = i
                        If .zOrder > Windows(curWindow).Window.zOrder Then curWindow = i
                    End If
                End If
                If entState = entStates.MouseMove Then
                    If .canDrag Then
                        If .state = entStates.MouseDown Then
                            .Left = Clamp(.Left + ((currMouseX - .Left) - .movedX), 0, ScreenWidth - .Width)
                            .top = Clamp(.top + ((currMouseY - .top) - .movedY), 0, ScreenHeight - .Height)
                        End If
                    End If
                End If
            End If
        End With
    Next

    ' Handle any controls first
    If curWindow Then
        ' reset /all other/ control mouse events
        For i = 1 To WindowCount
            If i <> curWindow Then
                For X = 1 To Windows(i).ControlCount
                    Windows(i).Controls(X).state = Normal
                Next
            End If
        Next
        For i = 1 To Windows(curWindow).ControlCount
            With Windows(curWindow).Controls(i)
                If Not .clickThrough Then    ' Skip if it's clickthrough
                    If .enabled And .visible Then
                        If .state <> entStates.MouseDown Then .state = entStates.Normal
                        If currMouseX >= .Left + Windows(curWindow).Window.Left And currMouseX <= .Left + .Width + Windows(curWindow).Window.Left Then
                            If currMouseY >= .top + Windows(curWindow).Window.top And currMouseY <= .top + .Height + Windows(curWindow).Window.top Then
                                If curControl = 0 Then curControl = i
                                If .zOrder > Windows(curWindow).Controls(curControl).zOrder Then curControl = i
                            End If
                        End If
                        If entState = entStates.MouseMove Then
                            If .canDrag Then
                                If .state = entStates.MouseDown Then
                                    .Left = Clamp(.Left + ((currMouseX - .Left) - .movedX), 0, Windows(curWindow).Window.Width - .Width)
                                    .top = Clamp(.top + ((currMouseY - .top) - .movedY), 0, Windows(curWindow).Window.Height - .Height)
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next
        ' Handle control
        If curControl Then
            HandleGuiMouse = True
            With Windows(curWindow).Controls(curControl)
                If .state <> entStates.MouseDown Then
                    If entState <> entStates.MouseMove Then
                        .state = entState
                    Else
                        .state = entStates.Hover
                    End If
                End If
                If entState = entStates.MouseDown Then
                    If .canDrag Then
                        .movedX = clickedX - .Left
                        .movedY = clickedY - .top
                    End If
                    ' toggle boxes
                    Select Case .Type
                    Case EntityTypes.entCheckbox
                        ' grouped boxes
                        If .group > 0 Then
                            If .Value = 0 Then
                                For i = 1 To Windows(curWindow).ControlCount
                                    If Windows(curWindow).Controls(i).Type = EntityTypes.entCheckbox Then
                                        If Windows(curWindow).Controls(i).group = .group Then
                                            Windows(curWindow).Controls(i).Value = 0
                                        End If
                                    End If
                                Next
                                .Value = 1
                            End If
                        Else
                            If .Value = 0 Then
                                .Value = 1
                            Else
                                .Value = 0
                            End If
                        End If
                    Case EntityTypes.entCombobox
                        ShowComboMenu curWindow, curControl
                    End Select
                    ' set active input
                    SetActiveControl curWindow, curControl
                End If
                Callback = .EntCallBack(entState)
            End With
        Else
            ' Handle container
            With Windows(curWindow).Window
                HandleGuiMouse = True
                If .state <> entStates.MouseDown Then
                    If entState <> entStates.MouseMove Then
                        .state = entState
                    Else
                        .state = entStates.Hover
                    End If
                End If
                If entState = entStates.MouseDown Then
                    If .canDrag Then
                        .movedX = clickedX - .Left
                        .movedY = clickedY - .top
                    End If
                End If
                Callback = .EntCallBack(entState)
            End With
        End If
        ' bring to front
        If entState = entStates.MouseDown Then
            UpdateZOrder curWindow
            activeWindow = curWindow
        End If
        ' call back
        If Callback <> 0 Then EntCallBack Callback, curWindow, curControl, 0, 0
    End If

    ' Reset
    If entState = entStates.MouseUp Then ResetMouseDown
End Function

Public Sub ResetGUI()
    Dim i As Long, X As Long

    For i = 1 To WindowCount

        If Windows(i).Window.state <> MouseDown Then Windows(i).Window.state = Normal

        For X = 1 To Windows(i).ControlCount

            If Windows(i).Controls(X).state <> MouseDown Then Windows(i).Controls(X).state = Normal
        Next
    Next

End Sub

Public Sub ResetMouseDown()
    Dim Callback As Long
    Dim i As Long, X As Long

    For i = 1 To WindowCount

        With Windows(i)
            .Window.state = entStates.Normal
            Callback = .Window.EntCallBack(entStates.Normal)

            If Callback <> 0 Then EntCallBack Callback, i, 0, 0, 0

            For X = 1 To .ControlCount
                .Controls(X).state = entStates.Normal
                Callback = .Controls(X).EntCallBack(entStates.Normal)

                If Callback <> 0 Then EntCallBack Callback, i, X, 0, 0
            Next

        End With

    Next

End Sub
' ################## ##
' ## REGISTER WINDOW ##
' #####################
Public Sub btnRegister_Click()
    HideWindows
    RenCaptcha
    ClearRegisterTexts
    ShowWindow GetWindowIndex("winRegister")
End Sub
Sub ClearRegisterTexts()
    Dim i As Long
    With Windows(GetWindowIndex("winRegister"))
        .Controls(GetControlIndex("winRegister", "txtAccount")).Text = vbNullString
        .Controls(GetControlIndex("winRegister", "txtPass")).Text = vbNullString
        .Controls(GetControlIndex("winRegister", "txtPass2")).Text = vbNullString
        .Controls(GetControlIndex("winRegister", "txtCode")).Text = vbNullString
        .Controls(GetControlIndex("winRegister", "txtCaptcha")).Text = vbNullString
        For i = 0 To 6
            .Controls(GetControlIndex("winRegister", "picCaptcha")).image(i) = Tex_Captcha(GlobalCaptcha)
        Next
    End With
End Sub
Sub RenCaptcha()
    Dim n As Long
    n = Int(Rnd * (Count_Captcha - 1)) + 1
    GlobalCaptcha = n
End Sub
Public Sub btnSendRegister_Click()
    Dim user As String, Pass As String, pass2 As String, Code As String, BirthDay As String, Captcha As String

    With Windows(GetWindowIndex("winRegister"))
        user = .Controls(GetControlIndex("winRegister", "txtAccount")).Text
        Pass = .Controls(GetControlIndex("winRegister", "txtPass")).Text
        pass2 = .Controls(GetControlIndex("winRegister", "txtPass2")).Text
        Code = .Controls(GetControlIndex("winRegister", "txtCode")).Text
        BirthDay = .Controls(GetControlIndex("winRegister", "txtBirthDay")).Text
        Captcha = .Controls(GetControlIndex("winRegister", "txtCaptcha")).Text
    End With

    If Trim$(Pass) <> Trim$(pass2) Then
        DialogueAlert DialogueMsg.MsgPASSCONFIRM
        Exit Sub
    End If
    If user = vbNullString Then
        DialogueAlert DialogueMsg.MsgUSERNULL
        Exit Sub
    End If
    If Pass = vbNullString Or pass2 = vbNullString Then
        DialogueAlert DialogueMsg.MsgPASSNULL
        Exit Sub
    End If
    If Len(Trim$(user)) < 3 Or Len(Trim$(user)) > ACCOUNT_LENGTH Then
        DialogueAlert DialogueMsg.MsgUSERLENGTH
        Exit Sub
    End If
    If Len(Trim$(Pass)) < 3 Or Len(Trim$(Pass)) > NAME_LENGTH Then
        DialogueAlert DialogueMsg.MsgWRONGPASS
        Exit Sub
    End If
    If InStr(1, Code, "@") = 0 Or Len(Trim$(Code)) < 4 Or Len(Trim$(Code)) > EMAIL_LENGTH Then
        DialogueAlert DialogueMsg.MsgEMAILINVALID
        Exit Sub
    End If
    
    If ConvertBirthDayToLng(BirthDay) = NO Then
        DialogueAlert DialogueMsg.MsgINVALIDBIRTHDAY
        Exit Sub
    End If

    If Trim$(Captcha) <> Trim$(GetCaptcha) Then
        RenCaptcha
        ClearRegisterTexts
        DialogueAlert DialogueMsg.MsgCAPTCHAINCORRECT
        Exit Sub
    End If

    SendRegister user, Pass, Code, BirthDay
End Sub

' Convert date string to long value
Private Function ConvertBirthDayToLng(ByVal BirthDay As String) As Long
    Dim tmpString() As String, tmpChar As String
    Dim i As Integer, CountCharacters As Integer

    ' Return 0 if have a problem in format
    ConvertBirthDayToLng = 0

    '00/00/0000 <- Length = 10, Verify if is a valid Date Format!   '
    If Len(BirthDay) < 10 Or Len(BirthDay) > 10 Then Exit Function  '

    ' Verify if have 8 numerics characters! \
    For i = 1 To Len(BirthDay)                 '\
        If IsNumeric(Mid$(BirthDay, i, 1)) Then    '\
            CountCharacters = CountCharacters + 1       '\
        End If                                              '\
    Next i                                                      '\
    If CountCharacters < 8 Or CountCharacters > 8 Then Exit Function    '\

    ' Verify if have 3 "/" or "\" or "-"
    If InStr(1, BirthDay, "/") = 3 Then
        tmpChar = "/"
    ElseIf InStr(1, BirthDay, "\") = 3 Then
        tmpChar = "\"
    ElseIf InStr(1, BirthDay, "-") = 3 Then
        tmpChar = "-"
    Else: Exit Function
    End If
    
    ' All OK, Go Split the string, to convert to long!
    tmpString = Split(BirthDay, tmpChar)

    tmpChar = vbNullString
    For i = 0 To UBound(tmpString)
        tmpChar = tmpChar + tmpString(i)
    Next i

    ConvertBirthDayToLng = tmpChar

End Function

Public Sub btnReturnMain_Click()
    HideWindows
    ShowWindow GetWindowIndex("winLogin")
End Sub
Function GetCaptcha() As String

    Select Case GlobalCaptcha
    Case 1
        GetCaptcha = "EVqu"
    Case 2
        GetCaptcha = "8NmV"
    Case 3
        GetCaptcha = "1Swi"
    Case 4
        GetCaptcha = "Vynk"
    Case 5
        GetCaptcha = "isKD"
    Case 6
        GetCaptcha = "eYX2"
    End Select

End Function

' ##################
' ## Login Window ##
' ##################

Public Sub btnLogin_Click()
    Dim user As String, Pass As String

    With Windows(GetWindowIndex("winLogin"))
        user = .Controls(GetControlIndex("winLogin", "txtUser")).Text
        Pass = .Controls(GetControlIndex("winLogin", "txtPass")).Text
    End With

    If user = vbNullString Then
        DialogueAlert DialogueMsg.MsgUSERNULL
        ClearUserAndPass
        Exit Sub
    End If
    If Pass = vbNullString Then
        DialogueAlert DialogueMsg.MsgPASSNULL
        ClearUserAndPass
        Exit Sub
    End If
    If Len(Trim$(user)) < 3 Or Len(Trim$(user)) > ACCOUNT_LENGTH Then
        DialogueAlert DialogueMsg.MsgUSERLENGTH
        ClearUserAndPass
        Exit Sub
    End If
    If Len(Trim$(Pass)) < 3 Or Len(Trim$(Pass)) > NAME_LENGTH Then
        DialogueAlert DialogueMsg.MsgPASSLENGTH
        ClearUserAndPass
        Exit Sub
    End If

    Login user, Pass
End Sub

Public Sub ClearUserAndPass()
    With Windows(GetWindowIndex("winLogin"))
        .Controls(GetControlIndex("winLogin", "txtUser")).Text = vbNullString
        .Controls(GetControlIndex("winLogin", "txtPass")).Text = vbNullString
    End With
End Sub

Public Sub chkSaveUser_Click()

    With Windows(GetWindowIndex("winLogin")).Controls(GetControlIndex("winLogin", "chkSaveUser"))
        If .Value = 0 Then    ' set as false
            Options.SaveUser = 0
            Options.Username = vbNullString
            SaveOptions
        Else
            Options.SaveUser = 1
            SaveOptions
        End If
    End With
End Sub

Public Sub chkSavePass_Click()

    With Windows(GetWindowIndex("winLogin")).Controls(GetControlIndex("winLogin", "chkSavePass"))
        If .Value = 0 Then    ' set as false
            Options.SavePass = 0
            Options.Password = vbNullString
            SaveOptions
        Else
            Options.SavePass = 1
            SaveOptions
        End If
    End With
End Sub

' #####################
' ## Dialogue Window ##
' #####################

Public Sub btnDialogue_Close()
    If diaStyle = StyleOKAY Then
        dialogueHandler 1
    ElseIf diaStyle = StyleYesNo Then
        dialogueHandler 3
    Else
        dialogueHandler 0
    End If
End Sub

Public Sub Dialogue_Okay()
    dialogueHandler 1
End Sub

Public Sub Dialogue_Yes()
    dialogueHandler 2
End Sub

Public Sub Dialogue_No()
    dialogueHandler 3
End Sub

' ####################
' ## Classes Window ##
' ####################

Public Sub Classes_DrawFace()
    Dim imageFace As Long, xO As Long, yO As Long

    xO = Windows(GetWindowIndex("winClasses")).Window.Left
    yO = Windows(GetWindowIndex("winClasses")).Window.top

    Max_Classes = 3

    If newCharClass = 0 Then newCharClass = 1

    Select Case newCharClass
    Case 1    ' Warrior
        imageFace = Tex_GUI(18)
    Case 2    ' Wizard
        imageFace = Tex_GUI(19)
    Case 3    ' Whisperer
        imageFace = Tex_GUI(20)
    End Select

    ' render face
    RenderTexture imageFace, xO + 14, yO - 41, 0, 0, 256, 256, 256, 256
End Sub

Public Sub Classes_DrawText()
    Dim image As Long, Text As String, xO As Long, yO As Long, textArray() As String, i As Long, Count As Long, Y As Long, X As Long

    xO = Windows(GetWindowIndex("winClasses")).Window.Left
    yO = Windows(GetWindowIndex("winClasses")).Window.top

    Select Case newCharClass
    Case 1    ' Warrior
        Text = "The way of a warrior has never been an easy one. Skilled use of a sword is not something learnt overnight. Being able to take a decent amount of hits is important for these characters and as such they weigh a lot of importance on endurance and strength."
    Case 2    ' Wizard
        Text = "Wizards are often mistrusted characters who have mastered the practise of using their own spirit to create elemental entities. Generally seen as playful and almost childish because of the huge amounts of pleasure they take from setting things on fire."
    Case 3    ' Whisperer
        Text = "The art of healing is one which comes with tremendous amounts of pressure and guilt. Constantly being put under high-pressure situations where their abilities could mean the difference between life and death leads many Whisperers to insanity."
    End Select

    ' wrap text
    WordWrap_Array Text, 200, textArray()
    ' render text
    Count = UBound(textArray)
    Y = yO + 60
    For i = 1 To Count
        X = xO + 132 + (200 \ 2) - (TextWidth(font(Fonts.rockwell_15), textArray(i)) \ 2)
        RenderText font(Fonts.rockwell_15), textArray(i), X, Y, White
        Y = Y + 14
    Next
End Sub

Public Sub btnClasses_Left()
    Dim Text As String
    newCharClass = newCharClass - 1
    If newCharClass <= 0 Then
        newCharClass = Max_Classes
    End If
    Windows(GetWindowIndex("winClasses")).Controls(GetControlIndex("winClasses", "lblClassName")).Text = Trim$(Class(newCharClass).Name)
End Sub

Public Sub btnClasses_Right()
    Dim Text As String
    newCharClass = newCharClass + 1
    If newCharClass > Max_Classes Then
        newCharClass = 1
    End If
    Windows(GetWindowIndex("winClasses")).Controls(GetControlIndex("winClasses", "lblClassName")).Text = Trim$(Class(newCharClass).Name)
End Sub

Public Sub btnClasses_Accept()
    HideWindow GetWindowIndex("winClasses")
    ShowWindow GetWindowIndex("winNewChar")
End Sub

Public Sub btnClasses_Close()
    HideWindows
    ShowWindow GetWindowIndex("winNewChar")
End Sub

' ###################
' ## New Character ##
' ###################

Public Sub NewChar_OnDraw()
    Dim imageFace As Long, imageChar As Long, xO As Long, yO As Long

    xO = Windows(GetWindowIndex("winNewChar")).Window.Left
    yO = Windows(GetWindowIndex("winNewChar")).Window.top

    If newCharGender = SEX_MALE Then
        imageFace = Tex_Face(Class(newCharClass).MaleSprite(newCharSprite))
        imageChar = Tex_Char(Class(newCharClass).MaleSprite(newCharSprite))
    Else
        imageFace = Tex_Face(Class(newCharClass).FemaleSprite(newCharSprite))
        imageChar = Tex_Char(Class(newCharClass).FemaleSprite(newCharSprite))
    End If

    ' render face
    RenderTexture imageFace, xO + 166, yO + 56, 0, 0, 94, 94, 94, 94
    ' render char
    RenderTexture imageChar, xO + 166, yO + 116, 32, 0, 32, 32, 32, 32
End Sub

Public Sub btnNewChar_Left()
    Dim spriteCount As Long

    If newCharGender = SEX_MALE Then
        spriteCount = UBound(Class(newCharClass).MaleSprite)
    Else
        spriteCount = UBound(Class(newCharClass).FemaleSprite)
    End If

    If newCharSprite <= 0 Then
        newCharSprite = spriteCount
    Else
        newCharSprite = newCharSprite - 1
    End If
End Sub

Public Sub btnNewChar_Right()
    Dim spriteCount As Long

    If newCharGender = SEX_MALE Then
        spriteCount = UBound(Class(newCharClass).MaleSprite)
    Else
        spriteCount = UBound(Class(newCharClass).FemaleSprite)
    End If

    If newCharSprite >= spriteCount Then
        newCharSprite = 0
    Else
        newCharSprite = newCharSprite + 1
    End If
End Sub

Public Sub chkNewChar_Male()
    newCharSprite = 1
    newCharGender = SEX_MALE
End Sub

Public Sub chkNewChar_Female()
    newCharSprite = 1
    newCharGender = SEX_FEMALE
End Sub

Public Sub btnNewChar_Cancel()
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "txtName")).Text = vbNullString
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "chkMale")).Value = 1
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "chkFemale")).Value = 0
    newCharSprite = 1
    newCharGender = SEX_MALE
    HideWindows
    ShowWindow GetWindowIndex("winClasses")
End Sub

Public Sub btnNewChar_Accept()
    Dim Name As String, i As Long, n As Long

    Name = Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "txtName")).Text

    ' Prevent hacking
    If Len(Trim$(Name)) < 3 Or Len(Trim$(Name)) > NAME_LENGTH Then
        DialogueAlert DialogueMsg.MsgNAMELENGTH
        Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Name)
        n = AscW(Mid$(Name, i, 1))
        If Not isNameLegal(n) Then
            DialogueAlert DialogueMsg.MsgNAMEILLEGAL
            Exit Sub
        End If
    Next

    ' Prevent hacking
    If (newCharGender < SEX_MALE) Or (newCharGender > SEX_FEMALE) Then
        DialogueAlert DialogueMsg.MsgCONNECTION
        Exit Sub
    End If

    ' Prevent hacking
    If newCharClass < 1 Or newCharClass > Max_Classes Then
        Exit Sub
    End If

    HideWindows
    AddChar Name, newCharGender, newCharClass, newCharSprite
End Sub

' ##############
' ## Esc Menu ##
' ##############

Public Sub btnEscMenu_Return()
    HideWindow GetWindowIndex("winBlank")
    HideWindow GetWindowIndex("winEscMenu")
End Sub

Public Sub btnEscMenu_Options()
    HideWindow GetWindowIndex("winEscMenu")
    ShowWindow GetWindowIndex("winOptions"), True, True
End Sub

Public Sub btnEscMenu_MainMenu()
    HideWindows
    ShowWindow GetWindowIndex("winLogin")
    Stop_Music
    ' play the menu music
    If Len(Trim$(MenuMusic)) > 0 Then Play_Music Trim$(MenuMusic)
    logoutGame
End Sub

Public Sub btnEscMenu_Exit()
    HideWindow GetWindowIndex("winBlank")
    HideWindow GetWindowIndex("winEscMenu")
    DestroyGame
End Sub

' ##########
' ## Bars ##
' ##########

Public Sub Bars_OnDraw()
    Dim xO As Long, yO As Long, Width As Long

    xO = Windows(GetWindowIndex("winBars")).Window.Left
    yO = Windows(GetWindowIndex("winBars")).Window.top

    ' Bars
    RenderTexture Tex_GUI(27), xO + 15, yO + 15, 0, 0, BarWidth_GuiHP, 13, BarWidth_GuiHP, 13
    RenderTexture Tex_GUI(28), xO + 15, yO + 32, 0, 0, BarWidth_GuiSP, 13, BarWidth_GuiSP, 13
    RenderTexture Tex_GUI(29), xO + 15, yO + 49, 0, 0, BarWidth_GuiEXP, 13, BarWidth_GuiEXP, 13
End Sub

' ##########
' ## Menu ##
' ##########

Public Sub btnMenu_Char()
    Dim curWindow As Long
    curWindow = GetWindowIndex("winCharacter")
    If Windows(curWindow).Window.visible Then
        HideWindow curWindow
    Else
        ShowWindow curWindow, , False
    End If
End Sub

Public Sub btnMenu_Inv()
    Dim curWindow As Long
    curWindow = GetWindowIndex("winInventory")
    If Windows(curWindow).Window.visible Then
        HideWindow curWindow
    Else
        ShowWindow curWindow, , False
    End If
End Sub

Public Sub btnMenu_Skills()
    Dim curWindow As Long
    curWindow = GetWindowIndex("winSkills")
    If Windows(curWindow).Window.visible Then
        HideWindow curWindow
    Else
        ShowWindow curWindow, , False
    End If
End Sub

Public Sub btnMenu_Map()
'Windows(GetWindowIndex("winCharacter")).Window.visible = Not Windows(GetWindowIndex("winCharacter")).Window.visible
End Sub

Public Sub btnMenu_Guild()
    Dim curWindow As Long
    curWindow = GetWindowIndex("winGuild")

    If Windows(curWindow).Window.visible Then
        HideWindow curWindow
    Else
        ShowWindow curWindow, , False
    End If
End Sub

Public Sub btnMenu_Quest()
    Dim curWindow As Long
    curWindow = GetWindowIndex("winQuest")

    If Windows(curWindow).Window.visible Then
        HideWindow curWindow
    Else
        ShowWindow curWindow, , False
    End If
End Sub

' ###############
' ## Inventory ##
' ###############

Public Sub Inventory_MouseDown()
    Dim invNum As Long, winIndex As Long, i As Long

    ' is there an item?
    invNum = IsItem(Windows(GetWindowIndex("winInventory")).Window.Left, Windows(GetWindowIndex("winInventory")).Window.top)

    If invNum Then
        ' exit out if we're offering that item
        If InTrade > 0 Then
            For i = 1 To MAX_INV
                If TradeYourOffer(i).num = invNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Stackable > 0 Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            ' currency handler
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                Dialogue "Select Amount", "Please choose how many to offer", "", TypeTRADEAMOUNT, StyleINPUT, invNum
                Exit Sub
            End If
            ' trade the normal item
            Call TradeItem(invNum, 0)
            Exit Sub
        End If

        ' drag it
        With DragBox
            .Type = Part_Item
            .Value = GetPlayerInvItemNum(MyIndex, invNum)
            .Origin = origin_Inventory
            .Slot = invNum
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
        Windows(GetWindowIndex("winInventory")).Window.state = Normal
    End If

    ' show desc. if needed
    Inventory_MouseMove
End Sub

Public Sub Inventory_DblClick()
    Dim itemNum As Long, i As Long

    If InTrade > 0 Then Exit Sub

    itemNum = IsItem(Windows(GetWindowIndex("winInventory")).Window.Left, Windows(GetWindowIndex("winInventory")).Window.top)

    If itemNum Then
        SendUseItem itemNum
    End If

    ' show desc. if needed
    Inventory_MouseMove
End Sub

Public Sub Inventory_MouseMove()
    Dim itemNum As Long, X As Long, Y As Long, i As Long

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

    itemNum = IsItem(Windows(GetWindowIndex("winInventory")).Window.Left, Windows(GetWindowIndex("winInventory")).Window.top)

    If itemNum Then
        ' exit out if we're offering that item
        If InTrade > 0 Then
            For i = 1 To MAX_INV
                If TradeYourOffer(i).num = itemNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Stackable > 0 Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
        End If
        ' make sure we're not dragging the item
        If DragBox.Type = Part_Item And DragBox.Value = itemNum Then Exit Sub
        ' calc position
        X = Windows(GetWindowIndex("winInventory")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        Y = Windows(GetWindowIndex("winInventory")).Window.top - 4
        ' offscreen?
        If X < 0 Then
            ' switch to right
            X = Windows(GetWindowIndex("winInventory")).Window.Left + Windows(GetWindowIndex("winInventory")).Window.Width
        End If
        ' go go go
        ShowInvDesc X, Y, itemNum
    End If
End Sub

' #################
' ## Description ##
' #################

Public Sub Description_OnDraw()
    Dim xO As Long, yO As Long, texNum As Long, Y As Long, i As Long, Count As Long
    ' dim rec As RECT

    ' exit out if we don't have a num
    If descItem = 0 Or descType = 0 Then Exit Sub

    xO = Windows(GetWindowIndex("winDescription")).Window.Left
    yO = Windows(GetWindowIndex("winDescription")).Window.top

    Select Case descType
    Case 1    ' Inventory Item
        texNum = Tex_Item(Item(descItem).Pic)
        'rec.top = 0
        'rec.Left = mTexture(Tex_Item(Item(descItem).Pic)).LeftFrames * PIC_X
    Case 2    ' Spell Icon
        texNum = Tex_Spellicon(Spell(descItem).Icon)
        ' render bar
        With Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "picBar"))
            If .visible Then RenderTexture Tex_GUI(45), xO + .Left, yO + .top, 0, 12, .Value, 12, .Value, 12
        End With
    End Select

    ' render sprite
    RenderTexture texNum, xO + 20, yO + 34, 0, 0, 64, 64, 32, 32

    ' render text array
    Y = 18
    Count = UBound(descText)
    For i = 1 To Count
        RenderText font(Fonts.verdana_12), descText(i).Text, xO + 141 - (TextWidth(font(Fonts.verdana_12), descText(i).Text) \ 2), yO + Y, descText(i).Colour
        Y = Y + 12
    Next

    ' close
    HideWindow GetWindowIndex("winDescription")
End Sub

' ##############
' ## Drag Box ##
' ##############

Public Sub DragBox_OnDraw()
    Dim xO As Long, yO As Long, texNum As Long, winIndex As Long

    winIndex = GetWindowIndex("winDragBox")
    xO = Windows(winIndex).Window.Left
    yO = Windows(winIndex).Window.top

    ' get texture num
    With DragBox
        Select Case .Type
        Case Part_Item
            If .Value Then
                texNum = Tex_Item(Item(.Value).Pic)
            End If
        Case Part_spell
            If .Value Then
                texNum = Tex_Spellicon(Spell(.Value).Icon)
            End If
        End Select
    End With

    ' draw texture
    RenderTexture texNum, xO, yO, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DragBox_Check()
    Dim winIndex As Long, i As Long, curWindow As Long, curControl As Long, tmpRec As RECT
    Dim xO As Integer, yO As Integer

    winIndex = GetWindowIndex("winDragBox")

    ' can't drag nuthin'
    If DragBox.Type = part_None Then Exit Sub

    ' check for other windows
    For i = 1 To WindowCount
        With Windows(i).Window
            If .visible Then
                ' can't drag to self
                If .Name <> "winDragBox" Then
                    If currMouseX >= .Left And currMouseX <= .Left + .Width Then
                        If currMouseY >= .top And currMouseY <= .top + .Height Then
                            If curWindow = 0 Then curWindow = i
                            If .zOrder > Windows(curWindow).Window.zOrder Then curWindow = i
                        End If
                    End If
                End If
            End If
        End With
    Next

    ' we have a window - check if we can drop
    If curWindow Then
        Select Case Windows(curWindow).Window.Name

        Case "winCharacter"
            If DragBox.Origin = origin_Inventory Then
                If DragBox.Type = Part_Item Then
                    If Windows(GetWindowIndex("winCharacter")).Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = 1 Then
                            If IsEqSlot(Windows(GetWindowIndex("winCharacter")).Window.Left, Windows(GetWindowIndex("winCharacter")).Window.top, DragBox.Slot) Then
                                SendUseItem DragBox.Slot
                            End If
                    End If
                End If
            End If

        Case "winBank"
            If DragBox.Origin = origin_Bank Then
                ' it's from the inventory!
                If DragBox.Type = Part_Item Then
                    ' find the slot to switch with
                    For i = 1 To MAX_BANK
                        With tmpRec
                            .top = Windows(curWindow).Window.top + BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                            .bottom = .top + 32
                            .Left = Windows(curWindow).Window.Left + BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                            .Right = .Left + 32
                        End With

                        If currMouseX >= tmpRec.Left And currMouseX <= tmpRec.Right Then
                            If currMouseY >= tmpRec.top And currMouseY <= tmpRec.bottom Then
                                ' switch the slots
                                If DragBox.Slot <> i Then ChangeBankSlots DragBox.Slot, i
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If

            If DragBox.Origin = origin_Inventory Then
                If DragBox.Type = Part_Item Then

                    If Item(GetPlayerInvItemNum(MyIndex, DragBox.Slot)).Stackable = 0 Then
                        DepositItem DragBox.Slot, 1
                    Else
                        Dialogue "Depositar Item", "Insira a quantidade para depósito.", "", TypeDEPOSITITEM, StyleINPUT, DragBox.Slot
                    End If

                End If
            End If

        Case "winInventory"
            If DragBox.Origin = origin_Inventory Then
                ' it's from the inventory!
                If DragBox.Type = Part_Item Then
                    ' find the slot to switch with
                    For i = 1 To MAX_INV
                        With tmpRec
                            .top = Windows(curWindow).Window.top + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                            .bottom = .top + 32
                            .Left = Windows(curWindow).Window.Left + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                            .Right = .Left + 32
                        End With

                        If currMouseX >= tmpRec.Left And currMouseX <= tmpRec.Right Then
                            If currMouseY >= tmpRec.top And currMouseY <= tmpRec.bottom Then
                                ' switch the slots
                                If DragBox.Slot <> i Then SendChangeInvSlots DragBox.Slot, i
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If

            If DragBox.Origin = origin_Bank Then
                If DragBox.Type = Part_Item Then

                    If Item(Bank.Item(DragBox.Slot).num).Stackable = 0 Then
                        WithdrawItem DragBox.Slot, 1
                    Else
                        Dialogue "Retirar Item", "Insira a quantidade que deseja retirar", "", TypeWITHDRAWITEM, StyleINPUT, DragBox.Slot
                    End If
                End If
            End If

            If DragBox.Origin = origin_Equip Then
                If DragBox.Type = Part_Item Then
                    If Windows(GetWindowIndex("winCharacter")).Controls(GetControlIndex("winCharacter", "chkEquipamentos")).Value = 1 Then
                        SendUnequip DragBox.Slot
                    End If
                End If
            End If
        Case "winSkills"
            If DragBox.Origin = origin_Spells Then
                ' it's from the spells!
                If DragBox.Type = Part_spell Then
                    ' find the slot to switch with
                    For i = 1 To MAX_PLAYER_SPELLS
                        With tmpRec
                            .top = Windows(curWindow).Window.top + SkillTop + ((SkillOffsetY + 32) * ((i - 1) \ SkillColumns))
                            .bottom = .top + 32
                            .Left = Windows(curWindow).Window.Left + SkillLeft + ((SkillOffsetX + 32) * (((i - 1) Mod SkillColumns)))
                            .Right = .Left + 32
                        End With

                        If currMouseX >= tmpRec.Left And currMouseX <= tmpRec.Right Then
                            If currMouseY >= tmpRec.top And currMouseY <= tmpRec.bottom Then
                                ' switch the slots
                                If DragBox.Slot <> i Then SendChangeSpellSlots DragBox.Slot, i
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
        Case "winHotbar"
            If DragBox.Origin <> origin_None Then
                If DragBox.Type <> part_None Then
                    ' find the slot
                    For i = 1 To MAX_HOTBAR
                        With tmpRec
                            .top = Windows(curWindow).Window.top + HotbarTop
                            .bottom = .top + 32
                            .Left = Windows(curWindow).Window.Left + HotbarLeft + ((i - 1) * HotbarOffsetX)
                            .Right = .Left + 32
                        End With

                        If currMouseX >= tmpRec.Left And currMouseX <= tmpRec.Right Then
                            If currMouseY >= tmpRec.top And currMouseY <= tmpRec.bottom Then
                                ' set the hotbar slot
                                If DragBox.Origin <> origin_Hotbar Then
                                    If DragBox.Type = Part_Item Then
                                        SendHotbarChange 1, DragBox.Slot, i
                                    ElseIf DragBox.Type = Part_spell Then
                                        SendHotbarChange 2, DragBox.Slot, i
                                    End If
                                Else
                                    ' SWITCH the hotbar slots
                                    If DragBox.Slot <> i Then SwitchHotbar DragBox.Slot, i
                                End If
                                ' exit early
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
        End Select
    Else
        ' no windows found - dropping on bare map
        Select Case DragBox.Origin
        Case PartTypeOrigins.origin_Inventory
            If GetPlayerInvItemNum(MyIndex, DragBox.Slot) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, DragBox.Slot)).Stackable = 0 Then
                SendDropItem DragBox.Slot, 1
            Else
                Dialogue "Derrubar Item", "Insira a quantidade que deseja derrubar", "", TypeDROPITEM, StyleINPUT, DragBox.Slot
            End If
            End If
        Case PartTypeOrigins.origin_Spells
            ' dialogue
        Case PartTypeOrigins.origin_Hotbar
            SendHotbarChange 0, 0, DragBox.Slot
        End Select
    End If

    ' close window
    HideWindow winIndex
    With DragBox
        .Type = part_None
        .Slot = 0
        .Origin = origin_None
        .Value = 0
    End With
End Sub

' ############
' ## Skills ##
' ############

Public Sub Skills_MouseDown()
    Dim slotNum As Long, winIndex As Long

    ' is there an item?
    slotNum = IsSkill(Windows(GetWindowIndex("winSkills")).Window.Left, Windows(GetWindowIndex("winSkills")).Window.top)

    If slotNum Then
        With DragBox
            .Type = Part_spell
            .Value = PlayerSpells(slotNum).Spell
            .Origin = origin_Spells
            .Slot = slotNum
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
        Windows(GetWindowIndex("winSkills")).Window.state = Normal
    End If

    ' show desc. if needed
    Skills_MouseMove
End Sub

Public Sub Skills_DblClick()
    Dim slotNum As Long

    slotNum = IsSkill(Windows(GetWindowIndex("winSkills")).Window.Left, Windows(GetWindowIndex("winSkills")).Window.top)

    If slotNum Then
        CastSpell slotNum
    End If

    ' show desc. if needed
    Skills_MouseMove
End Sub

Public Sub Skills_MouseMove()
    Dim slotNum As Long, X As Long, Y As Long

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

    slotNum = IsSkill(Windows(GetWindowIndex("winSkills")).Window.Left, Windows(GetWindowIndex("winSkills")).Window.top)

    If slotNum Then
        ' make sure we're not dragging the item
        If DragBox.Type = Part_Item And DragBox.Value = slotNum Then Exit Sub
        ' calc position
        X = Windows(GetWindowIndex("winSkills")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        Y = Windows(GetWindowIndex("winSkills")).Window.top - 4
        ' offscreen?
        If X < 0 Then
            ' switch to right
            X = Windows(GetWindowIndex("winSkills")).Window.Left + Windows(GetWindowIndex("winSkills")).Window.Width
        End If
        ' go go go
        ShowPlayerSpellDesc X, Y, slotNum
    End If
End Sub

' ############
' ## Hotbar ##
' ############

Public Sub Hotbar_MouseDown()
    Dim slotNum As Long, winIndex As Long

    ' is there an item?
    slotNum = IsHotbar(Windows(GetWindowIndex("winHotbar")).Window.Left, Windows(GetWindowIndex("winHotbar")).Window.top)

    If slotNum Then
        With DragBox
            If Hotbar(slotNum).sType = 1 Then    ' inventory
                .Type = Part_Item
            ElseIf Hotbar(slotNum).sType = 2 Then    ' spell
                .Type = Part_spell
            End If
            .Value = Hotbar(slotNum).Slot
            .Origin = origin_Hotbar
            .Slot = slotNum
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
        Windows(GetWindowIndex("winHotbar")).Window.state = Normal
    End If

    ' show desc. if needed
    Hotbar_MouseMove
End Sub

Public Sub Hotbar_DblClick()
    Dim slotNum As Long

    slotNum = IsHotbar(Windows(GetWindowIndex("winHotbar")).Window.Left, Windows(GetWindowIndex("winHotbar")).Window.top)

    If slotNum Then
        SendHotbarUse slotNum
    End If

    ' show desc. if needed
    Hotbar_MouseMove
End Sub

Public Sub Hotbar_MouseMove()
    Dim slotNum As Long, X As Long, Y As Long

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

    slotNum = IsHotbar(Windows(GetWindowIndex("winHotbar")).Window.Left, Windows(GetWindowIndex("winHotbar")).Window.top)

    If slotNum Then
        ' make sure we're not dragging the item
        If DragBox.Origin = origin_Hotbar And DragBox.Slot = slotNum Then Exit Sub
        ' calc position
        X = Windows(GetWindowIndex("winHotbar")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        Y = Windows(GetWindowIndex("winHotbar")).Window.top - 4
        ' offscreen?
        If X < 0 Then
            ' switch to right
            X = Windows(GetWindowIndex("winHotbar")).Window.Left + Windows(GetWindowIndex("winHotbar")).Window.Width
        End If
        ' go go go
        Select Case Hotbar(slotNum).sType
        Case 1    ' inventory
            ShowItemDesc X, Y, Hotbar(slotNum).Slot, False
        Case 2    ' spells
            ShowSpellDesc X, Y, Hotbar(slotNum).Slot, 0
        End Select
    End If
End Sub

' Chat
Public Sub btnSay_Click()
    HandleKeyPresses vbKeyReturn
End Sub

Public Sub OnDraw_Chat()
    Dim winIndex As Long, xO As Long, yO As Long

    winIndex = GetWindowIndex("winChat")
    xO = Windows(winIndex).Window.Left
    yO = Windows(winIndex).Window.top + 16

    ' draw the box
    RenderDesign DesignTypes.desWin_Desc, xO, yO, 352, 152
    ' draw the input box
    RenderTexture Tex_GUI(46), xO + 7, yO + 123, 0, 0, 171, 22, 171, 22
    RenderTexture Tex_GUI(46), xO + 174, yO + 123, 0, 22, 171, 22, 171, 22
    ' call the chat render
    RenderChat
End Sub

Public Sub OnDraw_ChatSmall()
    Dim winIndex As Long, xO As Long, yO As Long

    winIndex = GetWindowIndex("winChatSmall")

    If actChatWidth < 160 Then actChatWidth = 160
    If actChatHeight < 10 Then actChatHeight = 10

    xO = Windows(winIndex).Window.Left + 10
    yO = ScreenHeight - 16 - actChatHeight - 8

    ' draw the background
    RenderDesign DesignTypes.desWin_Shadow, xO, yO, actChatWidth, actChatHeight
    ' call the chat render
    RenderChat
End Sub

Public Sub chkChat_Event()
    Options.channelState(ChatChannel.chEvent) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkEvent")).Value
    UpdateChat
End Sub

Public Sub chkChat_Game()
    Options.channelState(ChatChannel.chGame) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkGame")).Value
    UpdateChat
End Sub

Public Sub chkChat_Map()
    Options.channelState(ChatChannel.chMap) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkMap")).Value
    UpdateChat
End Sub

Public Sub chkChat_Global()
    Options.channelState(ChatChannel.chGlobal) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkGlobal")).Value
    UpdateChat
End Sub

Public Sub chkChat_Party()
    Options.channelState(ChatChannel.chParty) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkParty")).Value
    UpdateChat
End Sub

Public Sub chkChat_Guild()
    Options.channelState(ChatChannel.chGuild) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkGuild")).Value
    UpdateChat
End Sub

Public Sub chkChat_Private()
    Options.channelState(ChatChannel.chPrivate) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkPrivate")).Value
    UpdateChat
End Sub

Public Sub chkChat_Quest()
    Options.channelState(ChatChannel.chQuest) = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "chkQuest")).Value
    UpdateChat
End Sub

Public Sub btnChat_Up()
    ChatButtonUp = True
End Sub

Public Sub btnChat_Down()
    ChatButtonDown = True
End Sub

Public Sub btnChat_Up_MouseUp()
    ChatButtonUp = False
End Sub

Public Sub btnChat_Down_MouseUp()
    ChatButtonDown = False
End Sub

' Options
Public Sub btnOptions_Close()
    HideWindow GetWindowIndex("winOptions")
    ShowWindow GetWindowIndex("winEscMenu")
End Sub

Sub btnOptions_Confirm()
    Dim i As Long, Value As Long, Width As Long, Height As Long, message As Boolean, musicFile As String

    ' music
    Value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkMusic")).Value
    If Options.Music <> Value Then
        Options.Music = Value
        ' let them know
        If Value = 0 Then
            AddText "Music turned off.", BrightGreen
            Stop_Music
        Else
            AddText "Music tured on.", BrightGreen
            ' play music
            If InGame Then musicFile = Trim$(Map.MapData.Music) Else musicFile = Trim$(MenuMusic)
            If Not musicFile = "None." Then
                Play_Music musicFile
            Else
                Stop_Music
            End If
        End If
    End If

    ' sound
    Value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkSound")).Value
    If Options.sound <> Value Then
        Options.sound = Value
        ' let them know
        If Value = 0 Then
            AddText "Sound turned off.", BrightGreen
        Else
            AddText "Sound tured on.", BrightGreen
        End If
    End If

    ' autotiles
    Value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkAutotiles")).Value
    If Value = 1 Then Value = 0 Else Value = 1
    If Options.NoAuto <> Value Then
        Options.NoAuto = Value
        ' let them know
        If Value = 0 Then
            If InGame Then
                AddText "Autotiles turned on.", BrightGreen
                initAutotiles
            End If
        Else
            If InGame Then
                AddText "Autotiles turned off.", BrightGreen
                initAutotiles
            End If
        End If
    End If

    ' fullscreen
    Value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkFullscreen")).Value
    If Options.Fullscreen <> Value Then
        Options.Fullscreen = Value
        message = True
    End If

    ' resolution
    With Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRes"))
        If .Value > 0 And .Value <= RES_COUNT Then
            If Options.Resolution <> .Value Then
                Options.Resolution = .Value
                If Not isFullscreen Then
                    SetResolution
                Else
                    message = True
                End If
            End If
        End If
    End With

    ' render
    With Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRender"))
        If .Value > 0 And .Value <= 3 Then
            If Options.Render <> .Value - 1 Then
                Options.Render = .Value - 1
                message = True
            End If
        End If
    End With
    
    ' reconnect
    Value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkReconnect")).Value
    If Options.Reconnect <> Value Then
        Options.Reconnect = Value
    End If
    
    ' item name
    Value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkItemName")).Value
    If Options.ItemName <> Value Then
        Options.ItemName = Value
    End If
    
    ' item animation
    Value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkItemAnimation")).Value
    If Options.ItemAnimation <> Value Then
        Options.ItemAnimation = Value
    End If
    
    ' Fps & Ping
    Value = Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "chkFPSConection")).Value
    If Options.FPSConection <> Value Then
        Options.FPSConection = Value
    End If

    ' save options
    SaveOptions
    ' let them know
    If InGame Then
        If message Then AddText "Some changes will take effect next time you load the game.", BrightGreen
    End If
    ' close
    btnOptions_Close
End Sub

' Npc Chat
Public Sub btnMessage_Close()
    HideWindow GetWindowIndex("winMessage")
End Sub

' Npc Chat
Public Sub btnNpcChat_Close()
    HideWindow GetWindowIndex("winNpcChat")
End Sub

Public Sub btnOpt1()
    SendChatOption 1
End Sub
Public Sub btnOpt2()
    SendChatOption 2
End Sub
Public Sub btnOpt3()
    SendChatOption 3
End Sub
Public Sub btnOpt4()
    SendChatOption 4
End Sub

' Shop
Public Sub btnShop_Close()
    CloseShop
End Sub

Public Sub chkShopBuying()
    With Windows(GetWindowIndex("winShop"))
        If .Controls(GetControlIndex("winShop", "chkBuying")).Value = 1 Then
            .Controls(GetControlIndex("winShop", "chkSelling")).Value = 0
        Else
            .Controls(GetControlIndex("winShop", "chkSelling")).Value = 0
            .Controls(GetControlIndex("winShop", "chkBuying")).Value = 1
            Exit Sub
        End If
    End With
    ' show buy button, hide sell
    With Windows(GetWindowIndex("winShop"))
        .Controls(GetControlIndex("winShop", "btnSell")).visible = False
        .Controls(GetControlIndex("winShop", "btnBuy")).visible = True
    End With
    ' update the shop
    shopIsSelling = False
    shopSelectedSlot = 1
    UpdateShop
End Sub

Public Sub chkShopSelling()
    With Windows(GetWindowIndex("winShop"))
        If .Controls(GetControlIndex("winShop", "chkSelling")).Value = 1 Then
            .Controls(GetControlIndex("winShop", "chkBuying")).Value = 0
        Else
            .Controls(GetControlIndex("winShop", "chkBuying")).Value = 0
            .Controls(GetControlIndex("winShop", "chkSelling")).Value = 1
            Exit Sub
        End If
    End With
    ' show sell button, hide buy
    With Windows(GetWindowIndex("winShop"))
        .Controls(GetControlIndex("winShop", "btnBuy")).visible = False
        .Controls(GetControlIndex("winShop", "btnSell")).visible = True
    End With
    ' update the shop
    shopIsSelling = True
    shopSelectedSlot = 1
    UpdateShop
End Sub

Public Sub btnShopBuy()
    BuyItem shopSelectedSlot
End Sub

Public Sub btnShopSell()
    SellItem shopSelectedSlot
End Sub

Public Sub Shop_MouseDown()
    Dim shopNum As Long

    ' is there an item?
    shopNum = IsShopSlot(Windows(GetWindowIndex("winShop")).Window.Left, Windows(GetWindowIndex("winShop")).Window.top)

    If shopNum Then
        ' set the active slot
        shopSelectedSlot = shopNum
        UpdateShop
    End If

    Shop_MouseMove
End Sub

Public Sub Shop_MouseMove()
    Dim shopSlot As Long, itemNum As Long, X As Long, Y As Long

    If InShop = 0 Then Exit Sub

    shopSlot = IsShopSlot(Windows(GetWindowIndex("winShop")).Window.Left, Windows(GetWindowIndex("winShop")).Window.top)

    If shopSlot Then
        ' calc position
        X = Windows(GetWindowIndex("winShop")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        Y = Windows(GetWindowIndex("winShop")).Window.top - 4
        ' offscreen?
        If X < 0 Then
            ' switch to right
            X = Windows(GetWindowIndex("winShop")).Window.Left + Windows(GetWindowIndex("winShop")).Window.Width
        End If
        ' selling/buying
        If Not shopIsSelling Then
            ' get the itemnum
            itemNum = Shop(InShop).TradeItem(shopSlot).Item
            If itemNum = 0 Then Exit Sub
            ShowShopDesc X, Y, itemNum
        Else
            ' get the itemnum
            itemNum = GetPlayerInvItemNum(MyIndex, shopSlot)
            If itemNum = 0 Then Exit Sub
            ShowShopDesc X, Y, itemNum
        End If
    End If
End Sub

' Right Click Menu
Sub RightClick_Close()
' close all menus
    HideWindow GetWindowIndex("winRightClickBG")
    HideWindow GetWindowIndex("winPlayerMenu")
End Sub

' Player Menu
Sub PlayerMenu_Party()
    RightClick_Close
    If PlayerMenuIndex = 0 Then Exit Sub
    SendPartyRequest PlayerMenuIndex
End Sub

Sub PlayerMenu_Trade()
    RightClick_Close
    If PlayerMenuIndex = 0 Then Exit Sub
    SendTradeRequest PlayerMenuIndex
End Sub

Sub PlayerMenu_Guild()
    RightClick_Close
    If PlayerMenuIndex = 0 Then Exit Sub
    SendGuildInvite Player(PlayerMenuIndex).Name
End Sub

Sub PlayerMenu_PM()
    RightClick_Close
    If PlayerMenuIndex = 0 Then Exit Sub
    AddText "System not yet in place.", BrightRed
End Sub

' Invitations
Sub btnInvite_Party()
    Dim top As Integer

    top = ScreenHeight - 80

    HideWindow GetWindowIndex("winInvite_Party")

    ' First
    If Windows(GetWindowIndex("winInvite_Trade")).Window.visible Then
        Windows(GetWindowIndex("winInvite_Trade")).Window.top = top
        If Windows(GetWindowIndex("winInvite_Guild")).Window.visible Then
            top = top - 37
            Windows(GetWindowIndex("winInvite_Guild")).Window.top = top
        End If
    ElseIf Windows(GetWindowIndex("winInvite_Guild")).Window.visible Then
        Windows(GetWindowIndex("winInvite_Guild")).Window.top = top
        If Windows(GetWindowIndex("winInvite_Trade")).Window.visible Then
            top = top - 37
            Windows(GetWindowIndex("winInvite_Trade")).Window.top = top
        End If
    End If

        Dialogue "Party Invitation", diaDataString & " has invited you to a party.", "Would you like to join?", TypePARTY, StyleYesNo
    End Sub

Sub btnInvite_Trade()
Dim top As Integer
    HideWindow GetWindowIndex("winInvite_Trade")
    
    top = ScreenHeight - 80
    
    ' First
    If Windows(GetWindowIndex("winInvite_party")).Window.visible Then
        Windows(GetWindowIndex("winInvite_party")).Window.top = top
        If Windows(GetWindowIndex("winInvite_Guild")).Window.visible Then
            top = top - 37
            Windows(GetWindowIndex("winInvite_Guild")).Window.top = top
        End If
    ElseIf Windows(GetWindowIndex("winInvite_Guild")).Window.visible Then
        Windows(GetWindowIndex("winInvite_Guild")).Window.top = top
        If Windows(GetWindowIndex("winInvite_party")).Window.visible Then
            top = top - 37
            Windows(GetWindowIndex("winInvite_party")).Window.top = top
        End If
    End If
    
    Dialogue "Trade Invitation", diaDataString & " has invited you to trade.", "Would you like to accept?", TypeTRADE, StyleYesNo
End Sub

Sub btnInvite_Guild()
Dim top As Integer
    HideWindow GetWindowIndex("winInvite_Guild")
    
    top = ScreenHeight - 80
    
    ' First
    If Windows(GetWindowIndex("winInvite_party")).Window.visible Then
        Windows(GetWindowIndex("winInvite_party")).Window.top = top
        If Windows(GetWindowIndex("winInvite_Trade")).Window.visible Then
            top = top - 37
            Windows(GetWindowIndex("winInvite_Trade")).Window.top = top
        End If
    ElseIf Windows(GetWindowIndex("winInvite_Trade")).Window.visible Then
        Windows(GetWindowIndex("winInvite_Trade")).Window.top = top
        If Windows(GetWindowIndex("winInvite_party")).Window.visible Then
            top = top - 37
            Windows(GetWindowIndex("winInvite_party")).Window.top = top
        End If
    End If

    Dialogue "Guild Invitation", diaDataString & " has invited you to Guild.", "Would you like to accept?", TypeGUILD, StyleYesNo
End Sub


' Trade
Sub btnTrade_Close()
    HideWindow GetWindowIndex("winTrade")
    DeclineTrade
End Sub

Sub btnTrade_Accept()
    AcceptTrade
End Sub

Sub TradeMouseDown_Your()
    Dim xO As Long, yO As Long, itemNum As Long
    xO = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Left
    yO = Windows(GetWindowIndex("winTrade")).Window.top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).top
    itemNum = IsTrade(xO, yO)

    ' make sure it exists
    If itemNum > 0 Then
        If TradeYourOffer(itemNum).num = 0 Then Exit Sub
        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(itemNum).num) = 0 Then Exit Sub

        ' unoffer the item
        UntradeItem itemNum
    End If
End Sub

Sub TradeMouseMove_Your()
    Dim xO As Long, yO As Long, itemNum As Long, X As Long, Y As Long
    xO = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Left
    yO = Windows(GetWindowIndex("winTrade")).Window.top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).top
    itemNum = IsTrade(xO, yO)

    ' make sure it exists
    If itemNum > 0 Then
        If TradeYourOffer(itemNum).num = 0 Then Exit Sub
        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(itemNum).num) = 0 Then Exit Sub

        ' calc position
        X = Windows(GetWindowIndex("winTrade")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        Y = Windows(GetWindowIndex("winTrade")).Window.top - 4
        ' offscreen?
        If X < 0 Then
            ' switch to right
            X = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Window.Width
        End If
        ' go go go
        ShowItemDesc X, Y, GetPlayerInvItemNum(MyIndex, TradeYourOffer(itemNum).num), False
    End If
End Sub

Sub TradeMouseMove_Their()
    Dim xO As Long, yO As Long, itemNum As Long, X As Long, Y As Long
    xO = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).Left
    yO = Windows(GetWindowIndex("winTrade")).Window.top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).top
    itemNum = IsTrade(xO, yO)

    ' make sure it exists
    If itemNum > 0 Then
        If TradeTheirOffer(itemNum).num = 0 Then Exit Sub

        ' calc position
        X = Windows(GetWindowIndex("winTrade")).Window.Left - Windows(GetWindowIndex("winDescription")).Window.Width
        Y = Windows(GetWindowIndex("winTrade")).Window.top - 4
        ' offscreen?
        If X < 0 Then
            ' switch to right
            X = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Window.Width
        End If
        ' go go go
        ShowItemDesc X, Y, TradeTheirOffer(itemNum).num, False
    End If
End Sub

' combobox
Sub CloseComboMenu()
    HideWindow GetWindowIndex("winComboMenuBG")
    HideWindow GetWindowIndex("winComboMenu")
End Sub

Public Sub SendTradeGold()
    With Windows(GetWindowIndex("winInventory"))
        If InTrade > 0 Then
            'If GetPlayerGold(MyIndex) > 0 Then
                Dialogue "Select Amount", "Please choose how many to offer", "", TypeTRADEGOLD, StyleINPUT
                Exit Sub
            'End If
        End If
    End With
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function
