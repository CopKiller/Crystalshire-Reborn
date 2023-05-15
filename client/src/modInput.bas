Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

' Actual input
Public Sub CheckKeys()

' exit out if dialogue
    If diaIndex > 0 Then Exit Sub
    If GetAsyncKeyState(Options.Up) >= 0 Then upDown = False
    If GetAsyncKeyState(Options.Down) >= 0 Then downDown = False
    If GetAsyncKeyState(Options.Left) >= 0 Then leftDown = False
    If GetAsyncKeyState(Options.Right) >= 0 Then rightDown = False
    If GetAsyncKeyState(Options.Atacar) >= 0 Then ControlDown = False
    If GetAsyncKeyState(Options.Correr) >= 0 Then ShiftDown = False
    If GetAsyncKeyState(Options.Target) >= 0 Then TargetDown = False

    If GetAsyncKeyState(VK_UP) >= 0 Then SetaUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then SetaDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then SetaLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then SetaRight = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then SetaRight = False

    If GetAsyncKeyState(VK_LBUTTON) >= 0 Then IsLeftMouseButtonDown = False

    'If GetAsyncKeyState(VK_CONTROL) >= 0 And GetAsyncKeyState(VK_V) >= 0 Then
    '    Ctrl_V = False
    'End If
End Sub

Public Sub CheckInputKeys()

' exit out if dialogue
    If diaIndex > 0 Then Exit Sub

    ' exit out if talking
    If Windows(GetWindowIndex("winChat")).Window.visible Then Exit Sub

    ' exit out if creating guild
    If Windows(GetWindowIndex("winGuildMaker")).Window.visible Then Exit Sub

    ' exit out if validade serial number
    If Windows(GetWindowIndex("winSerial")).Window.visible Then Exit Sub

    ' exit with changing controls
    If Windows(GetWindowIndex("winChangeControls")).Window.visible Then Exit Sub

    ' exit if
    If Windows(GetWindowIndex("winGuildMenu")).Window.visible Then Exit Sub

    ' continue
    If GetKeyState(Options.Correr) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(Options.Atacar) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If

    If GetKeyState(Options.Target) < 0 Then
        TargetDown = True
    Else
        TargetDown = False
    End If

    If GetKeyState(Options.PegarItem) < 0 Then
        CheckMapGetItem
    End If
    ' move up
    If GetKeyState(Options.Up) < 0 Then
        upDown = True
        downDown = False
        leftDown = False
        rightDown = False
        Exit Sub
    Else
        upDown = False
    End If
    'Move Right
    If GetKeyState(Options.Right) < 0 Then
        upDown = False
        downDown = False
        leftDown = False
        rightDown = True
        Exit Sub
    Else
        rightDown = False
    End If
    'Move down
    If GetKeyState(Options.Down) < 0 Then
        upDown = False
        downDown = True
        leftDown = False
        rightDown = False
        Exit Sub
    Else
        downDown = False
    End If
    'Move left
    If GetKeyState(Options.Left) < 0 Then
        upDown = False
        downDown = False
        leftDown = True
        rightDown = False
        Exit Sub
    Else
        leftDown = False
    End If

    If Options.UsarSetas > 0 Then
        ' move up
        If GetKeyState(VK_UP) < 0 Then
            SetaUp = True
            SetaDown = False
            SetaLeft = False
            SetaRight = False
            Exit Sub
        Else
            SetaUp = False
        End If
        'Move Right
        If GetKeyState(VK_RIGHT) < 0 Then
            SetaUp = False
            SetaDown = False
            SetaLeft = False
            SetaRight = True
            Exit Sub
        Else
            SetaRight = False
        End If
        'Move down
        If GetKeyState(VK_DOWN) < 0 Then
            SetaUp = False
            SetaDown = True
            SetaLeft = False
            SetaRight = False
            Exit Sub
        Else
            SetaDown = False
        End If
        'Move left
        If GetKeyState(VK_LEFT) < 0 Then
            SetaUp = False
            SetaDown = False
            SetaLeft = True
            SetaRight = False
            Exit Sub
        Else
            leftDown = False
        End If
    End If

End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
    Dim chatText As String, Name As String, i As Long, n As Long, Command() As String, buffer As clsBuffer, tmpNum As Long, tmpText As String

    ' Exit if
    If Windows(GetWindowIndex("winChangeControls")).Window.visible Then Exit Sub

    If InGame Then
        chatText = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text
    End If

    If KeyAscii = vbKeyEscape Then Exit Sub

    ' Do we have an active window
    If activeWindow > 0 Then
        ' make sure it's visible
        If Windows(activeWindow).Window.visible Then
            ' Do we have an active control
            If Windows(activeWindow).activeControl > 0 Then
                ' Do our thing
                With Windows(activeWindow).Controls(Windows(activeWindow).activeControl)
                    ' Handle input
                    Select Case KeyAscii
                    Case vbKeyBack
                        If LenB(.text) > 0 Then
                            .text = Left$(.text, Len(.text) - 1)
                        End If
                    Case vbKeyReturn
                        ' override for function callbacks
                        If .EntCallBack(entStates.Enter) > 0 Then
                            EntCallBack .EntCallBack(entStates.Enter), activeWindow, Windows(activeWindow).activeControl, 0, 0
                            Exit Sub
                        Else
                            n = 0
                            For i = Windows(activeWindow).ControlCount To 1 Step -1
                                If i > Windows(activeWindow).activeControl Then
                                    If SetActiveControl(activeWindow, i) Then n = i
                                End If
                            Next
                            If n = 0 Then
                                For i = Windows(activeWindow).ControlCount To 1 Step -1
                                    SetActiveControl activeWindow, i
                                Next
                            End If
                        End If
                    Case vbKeyTab
                        n = 0
                        For i = 1 To Windows(activeWindow).ControlCount
                            If i > Windows(activeWindow).activeControl Then
                                If SetActiveControl(activeWindow, i) Then n = i: Exit Sub
                            End If
                        Next
                        If n = 0 Then
                            For i = Windows(activeWindow).ControlCount To 1 Step -1
                                SetActiveControl activeWindow, i
                            Next
                        End If
                    Case Else
                        'chatLineIndex
                        If Ctrl_V Then
                            Ctrl_V = False
                            tmpText = GetClipboardText
                            If LenB(Trim$(tmpText)) > 0 Then
                                If (Len(.text) + Len(tmpText)) < .max Then
                                    .text = .text & GetClipboardText
                                Else
                                    AddText "String length exceeded limit", BrightRed, , ChatChannel.chGame
                                End If
                            End If
                            ' Respeita o maximo de letras do controle, caso nao tenha limita em 255 characteres
                        ElseIf Len(.text) < .max Then
                            .text = .text & ChrW$(KeyAscii)
                        End If
                    End Select
                    ' exit out early - if not chatting
                    If Windows(activeWindow).Window.Name <> "winChat" Then Exit Sub
                End With
            End If
        End If
    End If

    ' exit out early if we're not ingame
    If Not InGame Then Exit Sub

    ' Handle when the player presses the return key
    If KeyAscii = Options.Chat Then
        If Windows(GetWindowIndex("winChatSmall")).Window.visible Then
            ShowChat
            inSmallChat = False

            SendStatusDigitando YES
            Exit Sub
        Else
            SendStatusDigitando NO
        End If

        ' Broadcast message
        If Left$(chatText, 1) = "'" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call BroadcastMsg(chatText)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Guild message
        If Left$(chatText, 1) = "=" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call GuildMsg(chatText)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Party message
        If Left$(chatText, 1) = "+" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call PartyMsg(chatText)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Emote message
        If Left$(chatText, 1) = "-" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call EmoteMsg(chatText)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Player message
        If Left$(chatText, 1) = "!" Then
            Exit Sub
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
            Name = vbNullString
            ' Get the desired player from the user text
            tmpNum = Len(chatText)

            For i = 1 To tmpNum

                If Mid$(chatText, i, 1) <> Space$(1) Then
                    Name = Name & Mid$(chatText, i, 1)
                Else
                    Exit For
                End If

            Next

            chatText = Mid$(chatText, i, Len(chatText) - 1)

            ' Make sure they are actually sending something
            If Len(chatText) - i > 0 Then
                chatText = Mid$(chatText, i + 1, Len(chatText) - i)
                ' Send the message to the player
                Call PlayerMsg(chatText, Name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        If Left$(chatText, 1) = "/" Then
            Command = Split(chatText, Space$(1))

            Select Case Command(0)

            Case "/help"
                Call AddText("Social Commands:", HelpColor)
                Call AddText("'msghere = Global Message", HelpColor)
                Call AddText("-msghere = Emote Message", HelpColor)
                Call AddText("!namehere msghere = Player Message", HelpColor)
                Call AddText("Available Commands: /who, /fps, /fpslock, /gui, /maps", HelpColor)

            Case "/maps"
                ClearMapCache

            Case "/gui"
                hideGUI = Not hideGUI

            Case "/info"

                ' Checks to make sure we have more than one string in the array
                If UBound(Command) < 1 Then
                    AddText "Usage: /info (name)", AlertColor
                    GoTo continue
                End If

                If IsNumeric(Command(1)) Then
                    AddText "Usage: /info (name)", AlertColor
                    GoTo continue
                End If

                Set buffer = New clsBuffer
                buffer.WriteLong CPlayerInfoRequest
                buffer.WriteString Command(1)
                SendData buffer.ToArray()
                Set buffer = Nothing

                ' Whos Online
            Case "/who"
                SendWhosOnline

                ' toggle fps lock
            Case "/fpslock"
                FPS_Lock = Not FPS_Lock

                ' Request stats
            Case "/stats"
                Set buffer = New clsBuffer
                buffer.WriteLong CGetStats
                SendData buffer.ToArray()
                Set buffer = Nothing

                ' // Monitor Admin Commands //
                ' Kicking a player
            Case "/kick"

                If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo continue
                If UBound(Command) < 1 Then
                    AddText "Usage: /kick (name)", AlertColor
                    GoTo continue
                End If

                If IsNumeric(Command(1)) Then
                    AddText "Usage: /kick (name)", AlertColor
                    GoTo continue
                End If

                SendKick Command(1)

                ' // Mapper Admin Commands //
                ' Map Editor
            Case "/editmap"

                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                SendRequestEditMap

                ' Warping to a player
            Case "/warpmeto"

                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                If UBound(Command) < 1 Then
                    AddText "Usage: /warpmeto (name)", AlertColor
                    GoTo continue
                End If

                If IsNumeric(Command(1)) Then
                    AddText "Usage: /warpmeto (name)", AlertColor
                    GoTo continue
                End If

                GettingMap = True
                WarpMeTo Command(1)

                ' Warping a player to you
            Case "/warptome"

                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                If UBound(Command) < 1 Then
                    AddText "Usage: /warptome (name)", AlertColor
                    GoTo continue
                End If

                If IsNumeric(Command(1)) Then
                    AddText "Usage: /warptome (name)", AlertColor
                    GoTo continue
                End If

                WarpToMe Command(1)

                ' Warping to a map
            Case "/warpto"

                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                If UBound(Command) < 1 Then
                    AddText "Usage: /warpto (map #)", AlertColor
                    GoTo continue
                End If

                If Not IsNumeric(Command(1)) Then
                    AddText "Usage: /warpto (map #)", AlertColor
                    GoTo continue
                End If

                n = CLng(Command(1))

                ' Check to make sure its a valid map #
                If n > 0 And n <= MAX_MAPS Then
                    GettingMap = True
                    Call WarpTo(n)
                Else
                    Call AddText("Invalid map number.", Red)
                End If

                ' Setting sprite
            Case "/setsprite"

                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                If UBound(Command) < 1 Then
                    AddText "Usage: /setsprite (sprite #)", AlertColor
                    GoTo continue
                End If

                If Not IsNumeric(Command(1)) Then
                    AddText "Usage: /setsprite (sprite #)", AlertColor
                    GoTo continue
                End If

                SendSetSprite CLng(Command(1))

                ' Map report
            Case "/mapreport"

                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                SendMapReport

                ' Respawn request
            Case "/respawn"

                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                SendMapRespawn

                ' MOTD change
            Case "/motd"

                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                If UBound(Command) < 1 Then
                    AddText "Usage: /motd (new motd)", AlertColor
                    GoTo continue
                End If

                SendMOTDChange Right$(chatText, Len(chatText) - 5)
                ' Banning a player
            Case "/ban"

                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                If UBound(Command) < 1 Then
                    AddText "Usage: /ban (name)", AlertColor
                    GoTo continue
                End If

                SendBan Command(1)

                ' // Developer Admin Commands //
                ' Editing item request
            Case "/edititem"

                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                SendRequestEditItem

                ' editing conv request
            Case "/editconv"

                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                SendRequestEditConv

                ' Editing animation request
            Case "/editanimation"

                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                SendRequestEditAnimation

                ' Editing npc request
            Case "/editnpc"

                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                SendRequestEditNpc

            Case "/editresource"

                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                SendRequestEditResource

                ' Editing shop request
            Case "/editshop"

                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                SendRequestEditShop

                ' Editing spell request
            Case "/editspell"

                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                SendRequestEditSpell

            Case "/editquest"
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                SendRequestEditQuest

            Case "/editserial"
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
                SendRequestEditSerial

            Case "/editpremium"
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
                SendRequestEditPremium

                ' // Creator Admin Commands //
                ' Giving another player access
            Case "/setaccess"

                If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue
                If UBound(Command) < 2 Then
                    AddText "Usage: /setaccess (name) (access)", AlertColor
                    GoTo continue
                End If

                If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                    AddText "Usage: /setaccess (name) (access)", AlertColor
                    GoTo continue
                End If

                SendSetAccess Command(1), CLng(Command(2))

                ' Packet debug mode
            Case "/debug"

                If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue
                DEBUG_MODE = (Not DEBUG_MODE)

            Case Else
                AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
continue:
            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Say message
        If Len(chatText) > 0 Then
            Call SayMsg(chatText)
        End If

        Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString

        ' hide/show chat window
        If Windows(GetWindowIndex("winChat")).Window.visible Then HideChat
        Exit Sub
    End If

    ' hide/show chat window
    If Windows(GetWindowIndex("winChatSmall")).Window.visible Then
        Exit Sub
    End If
End Sub
