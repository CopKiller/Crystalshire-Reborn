Attribute VB_Name = "modInterface"
Option Explicit

' Entity Types
Public Enum EntityTypes
    entLabel = 1
    entWindow
    entButton
    entTextBox
    entScrollbar
    entPictureBox
    entCheckbox
    entCombobox
    entCombomenu
End Enum

' Design Types
Public Enum DesignTypes
    ' Boxes
    desWood = 1
    desWood_Small
    desWood_Empty
    desGreen
    desGreen_Hover
    desGreen_Click
    desRed
    desRed_Hover
    desRed_Click
    desBlue
    desBlue_Hover
    desBlue_Click
    desOrange
    desOrange_Hover
    desOrange_Click
    desGrey
    desDescPic
    ' Windows
    desWin_Black
    desWin_Norm
    desWin_NoBar
    desWin_Empty
    desWin_Desc
    desWin_Shadow
    desWin_Party
    ' Pictures
    desParchment
    desBlackOval
    ' Textboxes
    desTextBlack
    desTextWhite
    desTextBlack_Sq
    ' Checkboxes
    desChkNorm
    desChkChat
    desChkCustom_Buying
    desChkCustom_Selling
    desChkCharacter
    ' Right-click Menu
    desMenuHeader
    desMenuOption
    ' Comboboxes
    desComboNorm
    desComboMenuNorm
    ' tile Selection
    desTileBox
End Enum

' Button States
Public Enum entStates
    Normal = 0
    hover
    MouseDown
    MouseMove
    MouseUp
    DblClick
    Enter
    ' Count
    state_Count
End Enum

' Alignment
Public Enum Alignment
    alignLeft = 0
    alignRight
    alignCentre
End Enum

' Part Types
Public Enum PartType
    part_None = 0
    Part_Item
    Part_spell
End Enum

' Origins
Public Enum PartTypeOrigins
    origin_None = 0
    origin_Inventory
    origin_Hotbar
    origin_Spells
    origin_Bank
    origin_Equip
End Enum

' Entity UDT
Public Type EntityRec
    ' constants
    Name As String
    ' values
Type As Byte
    top As Long
    Left As Long
    Width As Long
    Height As Long
    enabled As Boolean
    visible As Boolean
    canDrag As Boolean
    max As Long
    Min As Long
    Value As Long
    text As String
    image(0 To entStates.state_Count - 1) As Long
    design(0 To entStates.state_Count - 1) As Long
    EntCallBack(0 To entStates.state_Count - 1) As Long
    Alpha As Long
    clickThrough As Boolean
    xOffset As Long
    yOffset As Long
    align As Byte
    font As Long
    textColour As Long
    textColour_Hover As Long
    textColour_Click As Long
    zChange As Byte
    onDraw As Long
    origLeft As Long
    origTop As Long
    tooltip As String
    group As Long
    list() As String
    activated As Boolean
    linkedToWin As Long
    linkedToCon As Long
    ' window
    Icon As Long
    ' textbox
    isCensor As Boolean
    ' temp
    state As entStates
    movedX As Long
    movedY As Long
    zOrder As Long
End Type

' For small parts
Public Type EntityPartRec
Type As PartType
    Origin As PartTypeOrigins
    Value As Long
    Slot As Long
End Type

' Window UDT
Public Type WindowRec
    Window As EntityRec
    Controls() As EntityRec
    ControlCount As Long
    activeControl As Long
End Type

' actual GUI
Public Windows() As WindowRec
Public WindowCount As Long
Public activeWindow As Long
' GUI parts
Public DragBox As EntityPartRec
' Used for automatically the zOrder
Public zOrder_Win As Long
Public zOrder_Con As Long

Public Sub CreateEntity(winNum As Long, zOrder As Long, Name As String, tType As EntityTypes, ByRef design() As Long, ByRef image() As Long, ByRef EntCallBack() As Long, _
                        Optional Left As Long, Optional top As Long, Optional Width As Long, Optional Height As Long, Optional visible As Boolean = True, Optional canDrag As Boolean, Optional max As Long, _
                        Optional Min As Long, Optional Value As Long, Optional text As String, Optional align As Byte, Optional font As Long = Fonts.georgia_16, Optional textColour As Long = White, _
                        Optional Alpha As Long = 255, Optional clickThrough As Boolean, Optional xOffset As Long, Optional yOffset As Long, Optional zChange As Byte, Optional ByVal Icon As Long, _
                        Optional ByVal onDraw As Long, Optional isActive As Boolean, Optional isCensor As Boolean, Optional textColour_Hover As Long, Optional textColour_Click As Long, _
                        Optional tooltip As String, Optional group As Long)
    Dim i As Long

    ' check if it's a legal number
    If winNum <= 0 Or winNum > WindowCount Then
        Exit Sub
    End If

    ' re-dim the control array
    With Windows(winNum)
        .ControlCount = .ControlCount + 1
        ReDim Preserve .Controls(1 To .ControlCount) As EntityRec
    End With

    ' Set the new control values
    With Windows(winNum).Controls(Windows(winNum).ControlCount)
        .Name = Name
        .Type = tType

        ' loop through states
        For i = 0 To entStates.state_Count - 1
            .design(i) = design(i)
            .image(i) = image(i)
            .EntCallBack(i) = EntCallBack(i)
        Next

        .Left = Left
        .top = top
        .origLeft = Left
        .origTop = top
        .Width = Width
        .Height = Height
        .visible = visible
        .canDrag = canDrag
        .max = max
        .Min = Min
        .Value = Value
        .text = text
        .align = align
        .font = font
        .textColour = textColour
        .textColour_Hover = textColour_Hover
        .textColour_Click = textColour_Click
        .Alpha = Alpha
        .clickThrough = clickThrough
        .xOffset = xOffset
        .yOffset = yOffset
        .zChange = zChange
        .zOrder = zOrder
        .enabled = True
        .Icon = Icon
        .onDraw = onDraw
        .isCensor = isCensor
        .tooltip = tooltip
        .group = group
        ReDim .list(0 To 0) As String
    End With

    ' set the active control
    If isActive Then Windows(winNum).activeControl = Windows(winNum).ControlCount

    ' set the zOrder
    zOrder_Con = zOrder_Con + 1
End Sub

Public Sub UpdateZOrder(winNum As Long, Optional forced As Boolean = False)
    Dim i As Long
    Dim oldZOrder As Long

    With Windows(winNum).Window

        If Not forced Then If .zChange = 0 Then Exit Sub
        If .zOrder = WindowCount Then Exit Sub
        oldZOrder = .zOrder

        For i = 1 To WindowCount

            If Windows(i).Window.zOrder > oldZOrder Then
                Windows(i).Window.zOrder = Windows(i).Window.zOrder - 1
            End If

        Next

        .zOrder = WindowCount
    End With

End Sub

Public Sub SortWindows()
    Dim tempWindow As WindowRec
    Dim i As Long, X As Long
    X = 1

    While X <> 0
        X = 0

        For i = 1 To WindowCount - 1

            If Windows(i).Window.zOrder > Windows(i + 1).Window.zOrder Then
                tempWindow = Windows(i)
                Windows(i) = Windows(i + 1)
                Windows(i + 1) = tempWindow
                X = 1
            End If

        Next

    Wend

End Sub

Public Sub RenderEntities()
    Dim i As Long, X As Long, curZOrder As Long

    ' don't render anything if we don't have any containers
    If WindowCount = 0 Then Exit Sub
    ' reset zOrder
    curZOrder = 1

    ' loop through windows
    Do While curZOrder <= WindowCount
        For i = 1 To WindowCount
            If curZOrder = Windows(i).Window.zOrder Then
                ' increment
                curZOrder = curZOrder + 1
                ' make sure it's visible
                If Windows(i).Window.visible Then
                    ' render container
                    RenderWindow i
                    ' render controls
                    For X = 1 To Windows(i).ControlCount
                        If Windows(i).Controls(X).visible Then RenderEntity i, X
                    Next
                End If
            End If
        Next
    Loop
End Sub

Public Sub RenderEntity(winNum As Long, entNum As Long)
    Dim xO As Integer, yO As Integer, hor_centre As Long, ver_centre As Long, Height As Long, Width As Long, Left As Long, texNum As Long, xOffset As Long
    Dim Callback As Long, taddText As String, Colour As Long, textArray() As String, Count As Long, yOffset As Long, i As Long, Y As Long, X As Long
    Dim TextOrigin As String, top As Integer

    ' check if the window exists
    If winNum <= 0 Or winNum > WindowCount Then
        Exit Sub
    End If

    ' check if the entity exists
    If entNum <= 0 Or entNum > Windows(winNum).ControlCount Then
        Exit Sub
    End If

    ' check the container's position
    xO = Windows(winNum).Window.Left
    yO = Windows(winNum).Window.top

    With Windows(winNum).Controls(entNum)

        ' find the control type
        Select Case .Type
            ' picture box
        Case EntityTypes.entPictureBox
            ' render specific designs
            If .design(.state) > 0 Then RenderDesign .design(.state), .Left + xO, .top + yO, .Width, .Height, .Alpha
            ' render image
            If .image(.state) > 0 Then RenderTexture .image(.state), .Left + xO, .top + yO, 0, 0, .Width, .Height, .Width, .Height, DX8Colour(White, .Alpha)
        Case EntityTypes.entTextBox
            ' render specific designs
            If .design(.state) > 0 Then RenderDesign .design(.state), .Left + xO, .top + yO, .Width, .Height, .Alpha
            ' render image
            If .image(.state) > 0 Then RenderTexture .image(.state), .Left + xO, .top + yO, 0, 0, .Width, .Height, .Width, .Height, DX8Colour(White, .Alpha)
            ' render text
            If activeWindow = winNum And Windows(winNum).activeControl = entNum Then taddText = chatShowLine
            ' if it's censored then render censored
            If Not .isCensor Then
                TextOrigin = .text
            Else
                TextOrigin = CensorWord(.text)
            End If
            Select Case .align
            Case Alignment.alignLeft
                ' check if need to word wrap
                If TextWidth(font(.font), TextOrigin) > .Width Then
                    ' wrap text
                    WordWrap_Array TextOrigin, .Width, textArray()
                    ' render text
                    Count = UBound(textArray)
                    For i = 1 To Count
                        If i = UBound(textArray) Then
                            RenderText font(.font), textArray(i) & taddText, .Left + xO + .xOffset, .top + yO + yOffset + .yOffset, .textColour
                        Else
                            RenderText font(.font), textArray(i), .Left + xO + .xOffset, .top + yO + yOffset + .yOffset, .textColour
                        End If

                        yOffset = yOffset + 14
                    Next
                Else
                    ' just one line
                    RenderText font(.font), TextOrigin & taddText, .Left + xO + .xOffset, .top + yO + .yOffset, .textColour
                End If
            Case Alignment.alignRight
                ' check if need to word wrap
                If TextWidth(font(.font), TextOrigin) > .Width Then
                    ' wrap text
                    WordWrap_Array TextOrigin, .Width, textArray()
                    ' render text
                    Count = UBound(textArray)
                    For i = 1 To Count
                        Left = .Left + .Width - TextWidth(font(.font), textArray(i))
                        If i = UBound(textArray) Then
                            RenderText font(.font), textArray(i) & taddText, Left + xO, .top + yO + yOffset + .yOffset, .textColour, .Alpha
                        Else
                            RenderText font(.font), textArray(i), Left + xO, .top + yO + yOffset + .yOffset, .textColour, .Alpha
                        End If
                        yOffset = yOffset + 14
                    Next
                Else
                    ' just one line
                    Left = .Left + .Width - TextWidth(font(.font), TextOrigin)
                    RenderText font(.font), TextOrigin & taddText, Left + xO + .xOffset, .top + yO + .yOffset, .textColour
                End If
            Case Alignment.alignCentre
                ' check if need to word wrap
                If TextWidth(font(.font), TextOrigin) > .Width Then
                    ' wrap text
                    WordWrap_Array TextOrigin, .Width, textArray()
                    ' render text
                    Count = UBound(textArray)
                    For i = 1 To Count
                        Left = .Left + (.Width \ 2) - (TextWidth(font(.font), textArray(i)) \ 2)
                        If i = UBound(textArray) Then
                            RenderText font(.font), textArray(i) & taddText, Left + xO + .xOffset, .top + yO + yOffset + .yOffset, .textColour, .Alpha
                        Else
                            RenderText font(.font), textArray(i), Left + xO + .xOffset, .top + yO + yOffset + .yOffset, .textColour, .Alpha
                        End If
                        yOffset = yOffset + 14
                    Next
                Else
                    ' just one line
                    Left = .Left + (.Width \ 2) - (TextWidth(font(.font), TextOrigin) \ 2)
                    RenderText font(.font), TextOrigin & taddText, Left + xO + .xOffset, .top + yO + .yOffset, .textColour
                End If
            End Select

            ' buttons
        Case EntityTypes.entButton
            ' render specific designs

            If .design(.state) > 0 Then
                If .design(.state) > 0 Then
                    RenderDesign .design(.state), .Left + xO, .top + yO, .Width, .Height
                End If
            End If
            ' render image
            If .image(.state) > 0 Then
                If .image(.state) > 0 Then
                    RenderTexture .image(.state), .Left + xO, .top + yO, 0, 0, .Width, .Height, .Width, .Height
                End If
            End If
            ' render icon
            If .Icon > 0 Then
                Width = mTexture(.Icon).w
                Height = mTexture(.Icon).h
                RenderTexture .Icon, .Left + xO + .xOffset, .top + yO + .yOffset, 0, 0, Width, Height, Width, Height
            End If
            ' for changing the text space
            xOffset = Width
            ' calculate the vertical centre
            Height = TextHeight(font(Fonts.georgiaDec_16))
            If Height > .Height Then
                ver_centre = .top + yO
            Else
                ver_centre = .top + yO + ((.Height - Height) \ 2) + 1
            End If
            ' calculate the horizontal centre
            Width = TextWidth(font(.font), .text)
            If Width > .Width Then
                hor_centre = .Left + xO + xOffset
            Else
                hor_centre = .Left + xO + xOffset + ((.Width - Width - xOffset) \ 2)
            End If
            ' get the colour
            If .state = hover Then
                Colour = .textColour_Hover
                If .tooltip <> vbNullString Then
                    Call RenderEntity_Square(Tex_Design(6), (hor_centre - (TextWidth(font(.font), .tooltip) / 2)) - 22, (ver_centre - PIC_Y) - 2, TextWidth(font(.font), .tooltip) + 10, 20, 5, 200)
                    RenderText font(.font), .tooltip, (hor_centre - (TextWidth(font(.font), .tooltip) / 2)) - 16, (ver_centre - PIC_Y), .textColour
                End If
            ElseIf .state = MouseDown Then
                Colour = .textColour_Click
            Else
                Colour = .textColour
            End If

            RenderText font(.font), .text, hor_centre, ver_centre, Colour

            ' labels
        Case EntityTypes.entLabel
            If Len(.text) > 0 Then

                If .state = hover Then
                    Colour = .textColour_Hover
                ElseIf .state = MouseDown Then
                    Colour = .textColour_Click
                Else
                    Colour = .textColour
                End If

                Select Case .align
                Case Alignment.alignLeft
                    ' check if need to word wrap
                    If TextWidth(font(.font), .text) > .Width Then
                        ' wrap text
                        WordWrap_Array .text, .Width, textArray()
                        ' render text
                        Count = UBound(textArray)
                        For i = 1 To Count
                            RenderText font(.font), textArray(i), .Left + xO, .top + yO + yOffset, Colour, .Alpha
                            yOffset = yOffset + 14
                        Next
                    Else
                        ' just one line
                        RenderText font(.font), .text, .Left + xO, .top + yO, Colour, .Alpha
                    End If
                Case Alignment.alignRight
                    ' check if need to word wrap
                    If TextWidth(font(.font), .text) > .Width Then
                        ' wrap text
                        WordWrap_Array .text, .Width, textArray()
                        ' render text
                        Count = UBound(textArray)
                        For i = 1 To Count
                            Left = .Left + .Width - TextWidth(font(.font), textArray(i))
                            RenderText font(.font), textArray(i), Left + xO, .top + yO + yOffset, Colour, .Alpha
                            yOffset = yOffset + 14
                        Next
                    Else
                        ' just one line
                        Left = .Left + .Width - TextWidth(font(.font), .text)
                        RenderText font(.font), .text, Left + xO, .top + yO, Colour, .Alpha
                    End If
                Case Alignment.alignCentre
                    ' check if need to word wrap
                    If TextWidth(font(.font), .text) > .Width Then
                        ' wrap text
                        WordWrap_Array .text, .Width, textArray()
                        ' render text
                        Count = UBound(textArray)
                        For i = 1 To Count
                            Left = .Left + (.Width \ 2) - (TextWidth(font(.font), textArray(i)) \ 2)
                            RenderText font(.font), textArray(i), Left + xO, .top + yO + yOffset, Colour, .Alpha
                            yOffset = yOffset + 14
                        Next
                    Else
                        ' just one line
                        Left = .Left + (.Width \ 2) - (TextWidth(font(.font), .text) \ 2)
                        RenderText font(.font), .text, Left + xO, .top + yO, Colour, .Alpha
                    End If
                End Select
            End If

            ' checkboxes
        Case EntityTypes.entCheckbox

            Select Case .design(0)
            Case DesignTypes.desChkNorm
                ' empty?
                If .Value = 0 Then texNum = Tex_GUI(2) Else texNum = Tex_GUI(3)
                ' render box
                RenderTexture texNum, .Left + xO, .top + yO, 0, 0, 14, 14, 14, 14
                ' find text position
                Select Case .align
                Case Alignment.alignLeft
                    Left = .Left + 18 + xO
                Case Alignment.alignRight
                    Left = .Left + 18 + (.Width - 18) - TextWidth(font(.font), .text) + xO
                Case Alignment.alignCentre
                    Left = .Left + 18 + ((.Width - 18) / 2) - (TextWidth(font(.font), .text) / 2) + xO
                End Select
                ' render text
                RenderText font(.font), .text, Left, .top + yO, .textColour, .Alpha
            Case DesignTypes.desChkChat
                If .Value = 0 Then .Alpha = 150 Else .Alpha = 255
                ' render box
                RenderTexture Tex_GUI(51), .Left + xO, .top + yO, 0, 0, 49, 23, 49, 23, DX8Colour(White, .Alpha)
                ' render text
                Left = .Left + (49 / 2) - (TextWidth(font(.font), .text) / 2) + xO
                ' render text
                RenderText font(.font), .text, Left, .top + yO + 4, .textColour, .Alpha
            Case DesignTypes.desChkCustom_Buying
                If .Value = 0 Then texNum = Tex_GUI(58) Else texNum = Tex_GUI(56)
                RenderTexture texNum, .Left + xO, .top + yO, 0, 0, 49, 20, 49, 20
            Case DesignTypes.desChkCustom_Selling
                If .Value = 0 Then texNum = Tex_GUI(59) Else texNum = Tex_GUI(57)
                RenderTexture texNum, .Left + xO, .top + yO, 0, 0, 49, 20, 49, 20
            Case DesignTypes.desChkCharacter
                If .Value = 0 Then .Alpha = 150 Else .Alpha = 255
                ' render box
                RenderTexture Tex_GUI(51), .Left + xO, .top + yO, 0, 0, .Width, .Height, 49, 23, DX8Colour(White, .Alpha)
                ' render text
                Left = .Left + (.Width / 2) - (TextWidth(font(.font), .text) / 2) + xO
                ' render text
                RenderText font(.font), .text, Left, .top + yO + 4, .textColour, .Alpha
            End Select

            ' comboboxes
        Case EntityTypes.entCombobox
            Select Case .design(0)
            Case DesignTypes.desComboNorm
                ' draw the background
                RenderDesign DesignTypes.desTextBlack, .Left + xO, .top + yO, .Width, .Height
                ' render the text
                If .Value > 0 Then
                    If .Value <= UBound(.list) Then
                        RenderText font(.font), .list(.Value), .Left + xO + 5, .top + yO + 3, White
                    End If
                End If
                ' draw the little arow
                RenderTexture Tex_GUI(66), .Left + xO + .Width - 11, .top + yO + 7, 0, 0, 5, 4, 5, 4
            End Select
        End Select

        ' callback draw
        Callback = .onDraw

        If Callback <> 0 Then EntCallBack Callback, winNum, entNum, 0, 0
    End With

End Sub

Public Sub RenderWindow(winNum As Long)
    Dim Width As Long, Height As Long, Callback As Long, X As Long, Y As Long, i As Long, Left As Long

    ' check if the window exists
    If winNum <= 0 Or winNum > WindowCount Then
        Exit Sub
    End If

    With Windows(winNum).Window

        Select Case .design(0)
        Case DesignTypes.desComboMenuNorm
            RenderTexture Tex_Blank, .Left, .top, 0, 0, .Width, .Height, 1, 1, DX8Colour(Black, 157)
            ' text
            If UBound(.list) > 0 Then
                Y = .top + 2
                X = .Left
                For i = 1 To UBound(.list)
                    ' render select
                    If i = .Value Or i = .group Then RenderTexture Tex_Blank, X, Y - 1, 0, 0, .Width, 15, 1, 1, DX8Colour(Black, 255)
                    ' render text
                    Left = X + (.Width \ 2) - (TextWidth(font(.font), .list(i)) \ 2)
                    If i = .Value Or i = .group Then
                        RenderText font(.font), .list(i), Left, Y, Yellow
                    Else
                        RenderText font(.font), .list(i), Left, Y, White
                    End If
                    Y = Y + 16
                Next
            End If
            Exit Sub
        End Select

        Select Case .design(.state)

        Case DesignTypes.desWin_Black
            RenderTexture Tex_Fader, .Left, .top, 0, 0, .Width, .Height, 32, 32, DX8Colour(Black, 190)

        Case DesignTypes.desWin_Norm
            ' render window
            RenderDesign DesignTypes.desWood, .Left, .top, .Width, .Height
            RenderDesign DesignTypes.desGreen, .Left + 2, .top + 2, .Width - 4, 21
            ' render the icon
            Width = mTexture(.Icon).w
            Height = mTexture(.Icon).h
            RenderTexture .Icon, .Left + .xOffset, .top - (Width - 18) + .yOffset, 0, 0, Width, Height, Width, Height
            ' render the caption
            RenderText font(.font), Trim$(.text), .Left + Height + 2, .top + 5, .textColour

        Case DesignTypes.desWin_NoBar
            ' render window
            RenderDesign DesignTypes.desWood, .Left, .top, .Width, .Height

        Case DesignTypes.desWin_Empty
            ' render window
            RenderDesign DesignTypes.desWood_Empty, .Left, .top, .Width, .Height
            RenderDesign DesignTypes.desGreen, .Left + 2, .top + 2, .Width - 4, 21
            ' render the icon
            Width = mTexture(.Icon).w
            Height = mTexture(.Icon).h
            RenderTexture .Icon, .Left + .xOffset, .top - (Width - 18) + .yOffset, 0, 0, Width, Height, Width, Height
            ' render the caption
            RenderText font(.font), Trim$(.text), .Left + Height + 2, .top + 5, .textColour

        Case DesignTypes.desWin_Desc
            ' render window
            RenderDesign DesignTypes.desWin_Desc, .Left, .top, .Width, .Height

        Case desWin_Shadow
            ' render window
            RenderDesign DesignTypes.desWin_Shadow, .Left, .top, .Width, .Height

        Case desWin_Party
            ' render window
            RenderDesign DesignTypes.desWin_Party, .Left, .top, .Width, .Height
        End Select

        ' OnDraw call back
        Callback = .onDraw

        If Callback <> 0 Then EntCallBack Callback, winNum, 0, 0, 0
    End With

End Sub

Public Sub RenderDesign(design As Long, Left As Long, top As Long, Width As Long, Height As Long, Optional Alpha As Long = 255)
    Dim bs As Long, Colour As Long
    ' change colour for alpha
    Colour = DX8Colour(White, Alpha)

    Select Case design

    Case DesignTypes.desMenuHeader
        ' render the header
        RenderTexture Tex_Blank, Left, top, 0, 0, Width, Height, 32, 32, D3DColorARGB(200, 47, 77, 29)

    Case DesignTypes.desMenuOption
        ' render the option
        RenderTexture Tex_Blank, Left, top, 0, 0, Width, Height, 32, 32, D3DColorARGB(200, 98, 98, 98)

    Case DesignTypes.desWood
        bs = 4
        ' render the wood box
        RenderEntity_Square Tex_Design(1), Left, top, Width, Height, bs, Alpha
        ' render wood texture
        RenderTexture Tex_GUI(1), Left + bs, top + bs, 100, 100, Width - (bs * 2), Height - (bs * 2), Width - (bs * 2), Height - (bs * 2), Colour

    Case DesignTypes.desWood_Small
        bs = 2
        ' render the wood box
        RenderEntity_Square Tex_Design(8), Left, top, Width, Height, bs, Alpha
        ' render wood texture
        RenderTexture Tex_GUI(1), Left + bs, top + bs, 100, 100, Width - (bs * 2), Height - (bs * 2), Width - (bs * 2), Height - (bs * 2), Colour

    Case DesignTypes.desWood_Empty
        bs = 4
        ' render the wood box
        RenderEntity_Square Tex_Design(9), Left, top, Width, Height, bs, Alpha

    Case DesignTypes.desGreen
        bs = 2
        ' render the green box
        RenderEntity_Square Tex_Design(2), Left, top, Width, Height, bs, Alpha
        ' render green gradient overlay
        RenderTexture Tex_Gradient(1), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desGreen_Hover
        bs = 2
        ' render the green box
        RenderEntity_Square Tex_Design(2), Left, top, Width, Height, bs, Alpha
        ' render green gradient overlay
        RenderTexture Tex_Gradient(2), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desGreen_Click
        bs = 2
        ' render the green box
        RenderEntity_Square Tex_Design(2), Left, top, Width, Height, bs, Alpha
        ' render green gradient overlay
        RenderTexture Tex_Gradient(3), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desRed
        bs = 2
        ' render the red box
        RenderEntity_Square Tex_Design(3), Left, top, Width, Height, bs, Alpha
        ' render red gradient overlay
        RenderTexture Tex_Gradient(4), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desRed_Hover
        bs = 2
        ' render the red box
        RenderEntity_Square Tex_Design(3), Left, top, Width, Height, bs, Alpha
        ' render red gradient overlay
        RenderTexture Tex_Gradient(5), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desRed_Click
        bs = 2
        ' render the red box
        RenderEntity_Square Tex_Design(3), Left, top, Width, Height, bs, Alpha
        ' render red gradient overlay
        RenderTexture Tex_Gradient(6), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desBlue
        bs = 2
        ' render the Blue box
        RenderEntity_Square Tex_Design(14), Left, top, Width, Height, bs, Alpha
        ' render Blue gradient overlay
        RenderTexture Tex_Gradient(8), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desBlue_Hover
        bs = 2
        ' render the Blue box
        RenderEntity_Square Tex_Design(14), Left, top, Width, Height, bs, Alpha
        ' render Blue gradient overlay
        RenderTexture Tex_Gradient(9), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desBlue_Click
        bs = 2
        ' render the Blue box
        RenderEntity_Square Tex_Design(14), Left, top, Width, Height, bs, Alpha
        ' render Blue gradient overlay
        RenderTexture Tex_Gradient(10), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desOrange
        bs = 2
        ' render the Orange box
        RenderEntity_Square Tex_Design(15), Left, top, Width, Height, bs, Alpha
        ' render Orange gradient overlay
        RenderTexture Tex_Gradient(11), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desOrange_Hover
        bs = 2
        ' render the Orange box
        RenderEntity_Square Tex_Design(15), Left, top, Width, Height, bs, Alpha
        ' render Orange gradient overlay
        RenderTexture Tex_Gradient(12), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desOrange_Click
        bs = 2
        ' render the Orange box
        RenderEntity_Square Tex_Design(15), Left, top, Width, Height, bs, Alpha
        ' render Orange gradient overlay
        RenderTexture Tex_Gradient(13), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desGrey
        bs = 2
        ' render the Orange box
        RenderEntity_Square Tex_Design(17), Left, top, Width, Height, bs, Alpha
        ' render Orange gradient overlay
        RenderTexture Tex_Gradient(14), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desParchment
        bs = 20
        ' render the parchment box
        RenderEntity_Square Tex_Design(4), Left, top, Width, Height, bs, Alpha

    Case DesignTypes.desBlackOval
        bs = 4
        ' render the black oval
        RenderEntity_Square Tex_Design(5), Left, top, Width, Height, bs, Alpha

    Case DesignTypes.desTextBlack
        bs = 5
        ' render the black oval
        RenderEntity_Square Tex_Design(6), Left, top, Width, Height, bs, Alpha

    Case DesignTypes.desTextWhite
        bs = 5
        ' render the black oval
        RenderEntity_Square Tex_Design(7), Left, top, Width, Height, bs, Alpha

    Case DesignTypes.desTextBlack_Sq
        bs = 4
        ' render the black oval
        RenderEntity_Square Tex_Design(10), Left, top, Width, Height, bs, Alpha

    Case DesignTypes.desWin_Desc
        bs = 8
        ' render black square
        RenderEntity_Square Tex_Design(11), Left, top, Width, Height, bs, Alpha

    Case DesignTypes.desDescPic
        bs = 3
        ' render the green box
        RenderEntity_Square Tex_Design(12), Left, top, Width, Height, bs, Alpha
        ' render green gradient overlay
        RenderTexture Tex_Gradient(7), Left + bs, top + bs, 0, 0, Width - (bs * 2), Height - (bs * 2), 128, 128, Colour

    Case DesignTypes.desWin_Shadow
        bs = 35
        ' render the green box
        RenderEntity_Square Tex_Design(13), Left - bs, top - bs, Width + (bs * 2), Height + (bs * 2), bs, Alpha

    Case DesignTypes.desWin_Party
        bs = 12
        ' render black square
        RenderEntity_Square Tex_Design(16), Left, top, Width, Height, bs, Alpha

    Case DesignTypes.desTileBox
        bs = 2
        ' render box
        RenderEntity_Square Tex_Design(18), Left, top, Width, Height, bs, Alpha
    End Select

End Sub

Public Sub RenderEntity_Square(texNum As Long, X As Long, Y As Long, Width As Long, Height As Long, borderSize As Long, Optional Alpha As Long = 255)
    Dim bs As Long, Colour As Long
    ' change colour for alpha
    Colour = DX8Colour(White, Alpha)
    ' Set the border size
    bs = borderSize
    ' Draw centre
    RenderTexture texNum, X + bs, Y + bs, bs + 1, bs + 1, Width - (bs * 2), Height - (bs * 2), 1, 1, Colour
    ' Draw top side
    RenderTexture texNum, X + bs, Y, bs, 0, Width - (bs * 2), bs, 1, bs, Colour
    ' Draw left side
    RenderTexture texNum, X, Y + bs, 0, bs, bs, Height - (bs * 2), bs, 1, Colour
    ' Draw right side
    RenderTexture texNum, X + Width - bs, Y + bs, bs + 3, bs, bs, Height - (bs * 2), bs, 1, Colour
    ' Draw bottom side
    RenderTexture texNum, X + bs, Y + Height - bs, bs, bs + 3, Width - (bs * 2), bs, 1, bs, Colour
    ' Draw top left corner
    RenderTexture texNum, X, Y, 0, 0, bs, bs, bs, bs, Colour
    ' Draw top right corner
    RenderTexture texNum, X + Width - bs, Y, bs + 3, 0, bs, bs, bs, bs, Colour
    ' Draw bottom left corner
    RenderTexture texNum, X, Y + Height - bs, 0, bs + 3, bs, bs, bs, bs, Colour
    ' Draw bottom right corner
    RenderTexture texNum, X + Width - bs, Y + Height - bs, bs + 3, bs + 3, bs, bs, bs, bs, Colour
End Sub

Sub Combobox_AddItem(winIndex As Long, controlIndex As Long, text As String)
    Dim Count As Long
    Count = UBound(Windows(winIndex).Controls(controlIndex).list)
    ReDim Preserve Windows(winIndex).Controls(controlIndex).list(0 To Count + 1)
    Windows(winIndex).Controls(controlIndex).list(Count + 1) = text
End Sub

Public Sub CreateWindow(Name As String, caption As String, zOrder As Long, Left As Long, top As Long, Width As Long, Height As Long, Icon As Long, _
                        Optional visible As Boolean = True, Optional font As Long = Fonts.georgia_16, Optional textColour As Long = White, Optional xOffset As Long, _
                        Optional yOffset As Long, Optional design_norm As Long, Optional design_hover As Long, Optional design_mousedown As Long, Optional image_norm As Long, _
                        Optional image_hover As Long, Optional image_mousedown As Long, Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, _
                        Optional entCallBack_mousemove As Long, Optional entCallBack_dblclick As Long, Optional canDrag As Boolean = True, Optional zChange As Byte = True, Optional ByVal onDraw As Long, _
                        Optional isActive As Boolean, Optional clickThrough As Boolean)

    Dim i As Long
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim EntCallBack(0 To entStates.state_Count - 1) As Long

    ' fill temp arrays
    design(entStates.Normal) = design_norm
    design(entStates.hover) = design_hover
    design(entStates.MouseDown) = design_mousedown
    design(entStates.DblClick) = design_norm
    design(entStates.MouseUp) = design_norm
    image(entStates.Normal) = image_norm
    image(entStates.hover) = image_hover
    image(entStates.MouseDown) = image_mousedown
    image(entStates.DblClick) = image_norm
    image(entStates.MouseUp) = image_norm
    EntCallBack(entStates.Normal) = entCallBack_norm
    EntCallBack(entStates.hover) = entCallBack_hover
    EntCallBack(entStates.MouseDown) = entCallBack_mousedown
    EntCallBack(entStates.MouseMove) = entCallBack_mousemove
    EntCallBack(entStates.DblClick) = entCallBack_dblclick
    ' redim the windows
    WindowCount = WindowCount + 1
    ReDim Preserve Windows(1 To WindowCount) As WindowRec

    ' set the properties
    With Windows(WindowCount).Window
        .Name = Name
        .Type = EntityTypes.entWindow

        ' loop through states
        For i = 0 To entStates.state_Count - 1
            .design(i) = design(i)
            .image(i) = image(i)
            .EntCallBack(i) = EntCallBack(i)
        Next

        .Left = Left
        .top = top
        .origLeft = Left
        .origTop = top
        .Width = Width
        .Height = Height
        .visible = visible
        .canDrag = canDrag
        .text = caption
        .font = font
        .textColour = textColour
        .xOffset = xOffset
        .yOffset = yOffset
        .Icon = Icon
        .enabled = True
        .zChange = zChange
        .zOrder = zOrder
        .onDraw = onDraw
        .clickThrough = clickThrough
        ' set active
        If .visible Then activeWindow = WindowCount
    End With

    ' set the zOrder
    zOrder_Win = zOrder_Win + 1
End Sub

Public Sub CreateTextbox(winNum As Long, Name As String, Left As Long, top As Long, Width As Long, Height As Long, Optional text As String, Optional font As Long = Fonts.georgia_16, _
                         Optional textColour As Long = White, Optional align As Byte = Alignment.alignLeft, Optional visible As Boolean = True, Optional Alpha As Long = 255, Optional image_norm As Long, _
                         Optional image_hover As Long, Optional image_mousedown As Long, Optional design_norm As Long, Optional design_hover As Long, Optional design_mousedown As Long, _
                         Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, Optional entCallBack_dblclick As Long, _
                         Optional isActive As Boolean, Optional xOffset As Long, Optional yOffset As Long, Optional isCensor As Boolean, Optional entCallBack_enter As Long, Optional MaxLenght As Long = 255)
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim EntCallBack(0 To entStates.state_Count - 1) As Long
    ' fill temp arrays
    design(entStates.Normal) = design_norm
    design(entStates.hover) = design_hover
    design(entStates.MouseDown) = design_mousedown
    image(entStates.Normal) = image_norm
    image(entStates.hover) = image_hover
    image(entStates.MouseDown) = image_mousedown
    EntCallBack(entStates.Normal) = entCallBack_norm
    EntCallBack(entStates.hover) = entCallBack_hover
    EntCallBack(entStates.MouseDown) = entCallBack_mousedown
    EntCallBack(entStates.MouseMove) = entCallBack_mousemove
    EntCallBack(entStates.DblClick) = entCallBack_dblclick
    EntCallBack(entStates.Enter) = entCallBack_enter
    ' create the textbox
    CreateEntity winNum, zOrder_Con, Name, entTextBox, design(), image(), EntCallBack(), Left, top, Width, Height, visible, , MaxLenght, , , text, align, font, textColour, Alpha, , xOffset, yOffset, , , , isActive, isCensor
End Sub

Public Sub CreatePictureBox(winNum As Long, Name As String, Left As Long, top As Long, Width As Long, Height As Long, Optional visible As Boolean = True, Optional canDrag As Boolean, _
                            Optional Alpha As Long = 255, Optional clickThrough As Boolean, Optional image_norm As Long, Optional image_hover As Long, Optional image_mousedown As Long, Optional design_norm As Long, _
                            Optional design_hover As Long, Optional design_mousedown As Long, Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, _
                            Optional entCallBack_mousemove As Long, Optional entCallBack_dblclick As Long, Optional onDraw As Long)
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim EntCallBack(0 To entStates.state_Count - 1) As Long
    ' fill temp arrays
    design(entStates.Normal) = design_norm
    design(entStates.hover) = design_hover
    design(entStates.MouseDown) = design_mousedown
    image(entStates.Normal) = image_norm
    image(entStates.hover) = image_hover
    image(entStates.MouseDown) = image_mousedown
    EntCallBack(entStates.Normal) = entCallBack_norm
    EntCallBack(entStates.hover) = entCallBack_hover
    EntCallBack(entStates.MouseDown) = entCallBack_mousedown
    EntCallBack(entStates.MouseMove) = entCallBack_mousemove
    EntCallBack(entStates.DblClick) = entCallBack_dblclick
    ' create the box
    CreateEntity winNum, zOrder_Con, Name, entPictureBox, design(), image(), EntCallBack(), Left, top, Width, Height, visible, canDrag, , , , , , , , Alpha, clickThrough, , , , , onDraw
End Sub

Public Sub CreateButton(winNum As Long, Name As String, Left As Long, top As Long, Width As Long, Height As Long, Optional text As String, Optional font As Fonts = Fonts.georgia_16, _
                        Optional textColour As Long = White, Optional Icon As Long, Optional visible As Boolean = True, Optional Alpha As Long = 255, Optional image_norm As Long, Optional image_hover As Long, _
                        Optional image_mousedown As Long, Optional design_norm As Long, Optional design_hover As Long, Optional design_mousedown As Long, Optional entCallBack_norm As Long, _
                        Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, Optional entCallBack_dblclick As Long, Optional xOffset As Long, _
                        Optional yOffset As Long, Optional textColour_Hover As Long = -1, Optional textColour_Click As Long = -1, Optional tooltip As String)
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim EntCallBack(0 To entStates.state_Count - 1) As Long
    ' default the colours
    If textColour_Hover = -1 Then textColour_Hover = textColour
    If textColour_Click = -1 Then textColour_Click = textColour
    ' fill temp arrays
    design(entStates.Normal) = design_norm
    design(entStates.hover) = design_hover
    design(entStates.MouseDown) = design_mousedown
    image(entStates.Normal) = image_norm
    image(entStates.hover) = image_hover
    image(entStates.MouseDown) = image_mousedown
    EntCallBack(entStates.Normal) = entCallBack_norm
    EntCallBack(entStates.hover) = entCallBack_hover
    EntCallBack(entStates.MouseDown) = entCallBack_mousedown
    EntCallBack(entStates.MouseMove) = entCallBack_mousemove
    EntCallBack(entStates.DblClick) = entCallBack_dblclick
    ' create the box
    CreateEntity winNum, zOrder_Con, Name, entButton, design(), image(), EntCallBack(), Left, top, Width, Height, visible, , , , , text, , font, textColour, Alpha, , xOffset, yOffset, , Icon, , , , textColour_Hover, textColour_Click, tooltip
End Sub

Public Sub CreateLabel(winNum As Long, Name As String, Left As Long, top As Long, Width As Long, Optional Height As Long, Optional text As String, Optional font As Fonts = Fonts.georgia_16, _
                       Optional textColour As Long = White, Optional align As Byte = Alignment.alignLeft, Optional visible As Boolean = True, Optional Alpha As Long = 255, Optional clickThrough As Boolean, _
                       Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, Optional entCallBack_dblclick As Long, _
                       Optional textColour_Hover As Long, Optional textColour_Click As Long)
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim EntCallBack(0 To entStates.state_Count - 1) As Long
    ' fill temp arrays
    EntCallBack(entStates.Normal) = entCallBack_norm
    EntCallBack(entStates.hover) = entCallBack_hover
    EntCallBack(entStates.MouseDown) = entCallBack_mousedown
    EntCallBack(entStates.MouseMove) = entCallBack_mousemove
    EntCallBack(entStates.DblClick) = entCallBack_dblclick
    ' create the box
    CreateEntity winNum, zOrder_Con, Name, entLabel, design(), image(), EntCallBack(), Left, top, Width, Height, visible, , , , , text, align, font, textColour, Alpha, clickThrough, , , , , , , , textColour_Hover, textColour_Click
End Sub

Public Sub CreateCheckbox(winNum As Long, Name As String, Left As Long, top As Long, Width As Long, Optional Height As Long = 15, Optional Value As Long, Optional text As String, _
                          Optional font As Fonts = Fonts.georgia_16, Optional textColour As Long = White, Optional align As Byte = Alignment.alignLeft, Optional visible As Boolean = True, Optional Alpha As Long = 255, _
                          Optional theDesign As Long, Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, _
                          Optional entCallBack_dblclick As Long, Optional group As Long)
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim EntCallBack(0 To entStates.state_Count - 1) As Long
    ' fill temp arrays
    EntCallBack(entStates.Normal) = entCallBack_norm
    EntCallBack(entStates.hover) = entCallBack_hover
    EntCallBack(entStates.MouseDown) = entCallBack_mousedown
    EntCallBack(entStates.MouseMove) = entCallBack_mousemove
    EntCallBack(entStates.DblClick) = entCallBack_dblclick
    ' fill temp array
    design(0) = theDesign
    ' create the box
    CreateEntity winNum, zOrder_Con, Name, entCheckbox, design(), image(), EntCallBack(), Left, top, Width, Height, visible, , , , Value, text, align, font, textColour, Alpha, , , , , , , , , , , , group
End Sub

Public Sub CreateComboBox(winNum As Long, Name As String, Left As Long, top As Long, Width As Long, Height As Long, design As Long, Optional font As Fonts = Fonts.georgia_16)
    Dim theDesign(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim EntCallBack(0 To entStates.state_Count - 1) As Long
    theDesign(0) = design
    ' create the box
    CreateEntity winNum, zOrder_Con, Name, entCombobox, theDesign(), image(), EntCallBack(), Left, top, Width, Height, , , , , , , , font
End Sub

Public Function GetWindowIndex(winName As String) As Long
    Dim i As Long

    For i = 1 To WindowCount

        If LCase$(Windows(i).Window.Name) = LCase$(winName) Then
            GetWindowIndex = i
            Exit Function
        End If

    Next

    GetWindowIndex = 0
End Function

Public Function GetControlIndex(winName As String, controlName As String) As Long
    Dim i As Long, winIndex As Long

    winIndex = GetWindowIndex(winName)

    If Not winIndex > 0 Or Not winIndex <= WindowCount Then Exit Function

    For i = 1 To Windows(winIndex).ControlCount

        If LCase$(Windows(winIndex).Controls(i).Name) = LCase$(controlName) Then
            GetControlIndex = i
            Exit Function
        End If

    Next

    GetControlIndex = 0
End Function

Public Function SetActiveControl(curWindow As Long, curControl As Long) As Boolean
' make sure it's something which CAN be active
    Select Case Windows(curWindow).Controls(curControl).Type
    Case EntityTypes.entTextBox
        Windows(curWindow).activeControl = curControl
        SetActiveControl = True
    End Select
End Function

Public Sub CentraliseWindow(curWindow As Long)
    With Windows(curWindow).Window
        .Left = (screenWidth / 2) - (.Width / 2)
        .top = (screenHeight / 2) - (.Height / 2)
        .origLeft = .Left
        .origTop = .top
    End With
End Sub

Public Sub HideWindows()
    Dim i As Long
    For i = 1 To WindowCount
        HideWindow i
    Next
End Sub

Public Sub ShowWindow(curWindow As Long, Optional forced As Boolean, Optional resetPosition As Boolean = True)
    Windows(curWindow).Window.visible = True

    If forced Then
        UpdateZOrder curWindow, forced
        activeWindow = curWindow
    ElseIf Windows(curWindow).Window.zChange Then
        UpdateZOrder curWindow
        activeWindow = curWindow
    End If
    If resetPosition Then
        With Windows(curWindow).Window
            .Left = .origLeft
            .top = .origTop
        End With
    End If
End Sub

Public Sub HideWindow(curWindow As Long)
    Dim i As Long
    Windows(curWindow).Window.visible = False

    ' find next window to set as active
    For i = WindowCount To 1 Step -1
        If Windows(i).Window.visible And Windows(i).Window.zChange Then
            'UpdateZOrder i
            activeWindow = i
            Exit Sub
        End If
    Next
End Sub

Public Sub CreateWindow_Login()
' Create the window
    CreateWindow "winLogin", "Login", zOrder_Win, 0, 0, 276, 212, Tex_Item(45), , Fonts.rockwellDec_15, , 3, 5, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf DestroyGame)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 264, 180, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Shadows
    CreatePictureBox WindowCount, "picShadow_1", 67, 43, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreatePictureBox WindowCount, "picShadow_2", 67, 79, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    ' Buttons
    CreateButton WindowCount, "btnAccept", 68, 134, 67, 22, "Accept", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnLogin_Click)
    CreateButton WindowCount, "btnExit", 142, 134, 67, 22, "Exit", rockwellDec_15, White, , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf DestroyGame)
    ' Labels
    CreateLabel WindowCount, "lblUsername", 66, 39, 142, , "Username", rockwellDec_15, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblPassword", 66, 75, 142, , "Password", rockwellDec_15, White, Alignment.alignCentre
    ' Textboxes
    CreateTextbox WindowCount, "txtUser", 67, 55, 142, 19, Options.Username, Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3, , , ACCOUNT_LENGTH
    CreateTextbox WindowCount, "txtPass", 67, 91, 142, 19, Options.Password, Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3, True, GetAddress(AddressOf btnLogin_Click), NAME_LENGTH
    ' Checkbox
    CreateCheckbox WindowCount, "chkSaveUser", 67, 114, 142, , Options.SaveUser, "Save User?", rockwell_15, , , , , DesignTypes.desChkNorm, , , GetAddress(AddressOf chkSaveUser_Click)
    CreateCheckbox WindowCount, "chkSavePass", 150, 114, 142, , Options.SavePass, "Save Pass?", rockwell_15, , , , , DesignTypes.desChkNorm, , , GetAddress(AddressOf chkSavePass_Click)

    ' Register Button
    CreateButton WindowCount, "btnRegister", 12, Windows(WindowCount).Window.Height - 35, 252, 22, "Create Account", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnRegister_Click)
    
    ' Recovery Account
    CreateLabel WindowCount, "lblRecovery", 0, 28, 142, 16, "Recovery Account", rockwellDec_15, White, Alignment.alignCentre, , , , , , GetAddress(AddressOf lblRecovery_Click), , , Yellow, Black
End Sub

Private Sub lblRecovery_Click()
    Dim Email As String
    ' Inputbox provisorio, pois causa travamento do loop.
    Email = InputBox("Digite seu e-mail", "Esqueci minha senha")
    
    TcpInit AUTH_SERVER_IP, AUTH_SERVER_PORT

    If ConnectToServer Then
        Call SetStatus("Sending email informations.")
        Call SendAccountRecovery(Email)
    Else
        ShowWindow GetWindowIndex("winregister")
        Dialogue "Connection Problem", "Cannot connect to login server.", "Please try again later.", TypeALERT
    End If
End Sub

'Email = InputBox("Digite seu e-mail", "Esqueci minha senha")

Public Sub CreateWindow_Register()

' Create the window
    CreateWindow "winRegister", "Register", zOrder_Win, 0, 0, 276, 347, Tex_Item(45), , Fonts.rockwellDec_15, , 3, 5, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, , , , , , , , , , , GetAddress(AddressOf CheckBirthDayFormat)
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnReturnMain_Click)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 264, 315, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment

    ' Shadows
    CreatePictureBox WindowCount, "picShadow_1", 67, 43, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreatePictureBox WindowCount, "picShadow_2", 67, 79, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreatePictureBox WindowCount, "picShadow_3", 67, 115, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreatePictureBox WindowCount, "picShadow_4", 67, 151, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreatePictureBox WindowCount, "picShadow_5", 67, 187, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreatePictureBox WindowCount, "picShadow_6", 67, 232, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval

    ' Buttons
    CreateButton WindowCount, "btnAccept", 68, 307, 67, 22, "Create", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnSendRegister_Click)
    CreateButton WindowCount, "btnExit", 142, 307, 67, 22, "Back", rockwellDec_15, White, , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnReturnMain_Click)

    ' Labels
    CreateLabel WindowCount, "lblUsername", 66, 39, 142, , "Username", rockwellDec_15, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblPassword", 66, 75, 142, , "Password", rockwellDec_15, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblPassword2", 66, 111, 142, , "Retype Password", rockwellDec_15, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblCode", 66, 147, 142, , "Email", rockwellDec_15, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblBirthday", 66, 183, 142, , "Data de Nascimento", rockwellDec_15, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblCaptcha", 66, 228, 142, , "Captcha", rockwellDec_15, White, Alignment.alignCentre

    ' Textboxes
    CreateTextbox WindowCount, "txtAccount", 67, 55, 142, 19, vbNullString, Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , True, 5, 3, False, GetAddress(AddressOf btnSendRegister_Click), ACCOUNT_LENGTH
    CreateTextbox WindowCount, "txtPass", 67, 91, 142, 19, vbNullString, Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3, True, GetAddress(AddressOf btnSendRegister_Click), NAME_LENGTH
    CreateTextbox WindowCount, "txtPass2", 67, 127, 142, 19, vbNullString, Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3, True, GetAddress(AddressOf btnSendRegister_Click), NAME_LENGTH
    CreateTextbox WindowCount, "txtCode", 67, 163, 142, 19, vbNullString, Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3, False, GetAddress(AddressOf btnSendRegister_Click), EMAIL_LENGTH
    CreateTextbox WindowCount, "txtBirthDay", 67, 199, 142, 19, vbNullString, Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3, False, GetAddress(AddressOf btnSendRegister_Click), BIRTHDAY_LENGTH
    CreateTextbox WindowCount, "txtCaptcha", 67, 280, 142, 19, vbNullString, Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3, False, GetAddress(AddressOf btnSendRegister_Click), CAPTCHA_LENGTH

    CreatePictureBox WindowCount, "picCaptcha", 67, 244, 156, 30, , , , , Tex_Captcha(GlobalCaptcha), Tex_Captcha(GlobalCaptcha), Tex_Captcha(GlobalCaptcha), DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
End Sub

Private Sub CheckBirthDayFormat()
    Dim SString As String, StrArray() As String

    If GetKeyState(vbKeyBack) < 0 Then Exit Sub

    With Windows(GetWindowIndex("winRegister"))

        SString = Trim$(.Controls(GetControlIndex("winRegister", "txtBirthDay")).text)

        If SString <> vbNullString Then
            StrArray = Split(SString, "/")
            If Len(SString) = 2 Then
                If UBound(StrArray) = 0 Then
                    SString = SString & "/"
                End If
            ElseIf Len(SString) = 5 Then
                If UBound(StrArray) = 1 Then
                    SString = SString & "/"
                End If
            End If

            If SString <> Trim$(.Controls(GetControlIndex("winRegister", "txtBirthDay")).text) Then
                .Controls(GetControlIndex("winRegister", "txtBirthDay")).text = SString
            End If
        End If
    End With

End Sub

Public Sub CreateWindow_Loading()
' Create the window
    CreateWindow "winLoading", "Loading", zOrder_Win, 0, 0, 278, 79, Tex_Item(104), True, Fonts.rockwellDec_15, , 2, 7, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 266, 47, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Text background
    CreatePictureBox WindowCount, "picRecess", 26, 39, 226, 22, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    ' Label
    CreateLabel WindowCount, "lblLoading", 6, 43, 266, , "Loading Game Data...", rockwell_15, , Alignment.alignCentre
End Sub

Public Sub CreateWindow_Dialogue()
' Create black background
    CreateWindow "winBlank", "", zOrder_Win, 0, 0, 800, 600, 0, , , , , , DesignTypes.desWin_Black, DesignTypes.desWin_Black, DesignTypes.desWin_Black, , , , , , , , , False, False
    ' Create dialogue window
    CreateWindow "winDialogue", "Warning", zOrder_Win, 0, 0, 348, 145, Tex_Item(38), , Fonts.rockwellDec_15, , 3, 5, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, , , , , , , , , , False
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnDialogue_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 335, 113, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Header
    CreatePictureBox WindowCount, "picShadow", 103, 44, 144, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblHeader", 103, 41, 144, , "Header", rockwellDec_15, White, Alignment.alignCentre
    ' Labels
    CreateLabel WindowCount, "lblBody_1", 15, 60, 314, , "Invalid username or password.", rockwell_15, , Alignment.alignCentre
    CreateLabel WindowCount, "lblBody_2", 15, 75, 314, , "Please try again.", rockwell_15, , Alignment.alignCentre
    ' Buttons
    CreateButton WindowCount, "btnYes", 104, 98, 68, 24, "Yes", rockwellDec_15, , , False, , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf Dialogue_Yes)
    CreateButton WindowCount, "btnNo", 180, 98, 68, 24, "No", rockwellDec_15, , , False, , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf Dialogue_No)
    CreateButton WindowCount, "btnOkay", 140, 98, 68, 24, "Okay", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf Dialogue_Okay)
    ' Input
    CreateTextbox WindowCount, "txtInput", 93, 75, 162, 18, , rockwell_15, White, Alignment.alignCentre, , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , , , , , True, 4, 2
    ' set active control
    'SetActiveControl WindowCount, GetControlIndex("winDialogue", "txtInput")
End Sub

Public Sub CreateWindow_Classes()
' Create window
    CreateWindow "winClasses", "Select Class", zOrder_Win, 0, 0, 364, 229, Tex_Item(17), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnClasses_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 352, 197, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment, , , , , , GetAddress(AddressOf Classes_DrawFace)
    ' Class Name
    CreatePictureBox WindowCount, "picShadow", 183, 42, 98, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblClassName", 183, 39, 98, , "Warrior", rockwellDec_15, White, Alignment.alignCentre
    ' Select Buttons
    CreateButton WindowCount, "btnLeft", 171, 40, 11, 13, , , , , , , Tex_GUI(12), Tex_GUI(14), Tex_GUI(16), , , , , , GetAddress(AddressOf btnClasses_Left)
    CreateButton WindowCount, "btnRight", 282, 40, 11, 13, , , , , , , Tex_GUI(13), Tex_GUI(15), Tex_GUI(17), , , , , , GetAddress(AddressOf btnClasses_Right)
    ' Accept Button
    CreateButton WindowCount, "btnAccept", 183, 185, 98, 22, "Accept", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnClasses_Accept)
    ' Text background
    CreatePictureBox WindowCount, "picBackground", 127, 55, 210, 124, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    ' Overlay
    CreatePictureBox WindowCount, "picOverlay", 6, 26, 0, 0, , , , , , , , , , , , , , , , GetAddress(AddressOf Classes_DrawText)
End Sub

Public Sub CreateWindow_NewChar()
' Create window
    CreateWindow "winNewChar", "Create Character", zOrder_Win, 0, 0, 291, 172, Tex_Item(17), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnNewChar_Cancel)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 278, 140, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Name
    CreatePictureBox WindowCount, "picShadow_1", 29, 42, 124, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblName", 29, 39, 124, , "Name", rockwellDec_15, White, Alignment.alignCentre
    ' Textbox
    CreateTextbox WindowCount, "txtName", 29, 55, 124, 19, , Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , True, 5, 3
    ' Gender
    CreatePictureBox WindowCount, "picShadow_2", 29, 85, 124, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblGender", 29, 82, 124, , "Gender", rockwellDec_15, White, Alignment.alignCentre
    ' Checkboxes
    CreateCheckbox WindowCount, "chkMale", 29, 103, 55, , 1, "Male", rockwell_15, , Alignment.alignCentre, , , DesignTypes.desChkNorm, , , GetAddress(AddressOf chkNewChar_Male), , , 1
    CreateCheckbox WindowCount, "chkFemale", 90, 103, 62, , 0, "Female", rockwell_15, , Alignment.alignCentre, , , DesignTypes.desChkNorm, , , GetAddress(AddressOf chkNewChar_Female), , , 1
    ' Buttons
    CreateButton WindowCount, "btnAccept", 29, 127, 60, 24, "Accept", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnNewChar_Accept)
    CreateButton WindowCount, "btnCancel", 93, 127, 60, 24, "Cancel", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnNewChar_Cancel)
    ' Sprite
    CreatePictureBox WindowCount, "picShadow_3", 175, 42, 76, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblSprite", 175, 39, 76, , "Sprite", rockwellDec_15, White, Alignment.alignCentre
    ' Scene
    CreatePictureBox WindowCount, "picScene", 165, 55, 96, 96, , , , , Tex_GUI(11), Tex_GUI(11), Tex_GUI(11), , , , , , , , , GetAddress(AddressOf NewChar_OnDraw)
    ' Buttons
    CreateButton WindowCount, "btnLeft", 163, 40, 11, 13, , , , , , , Tex_GUI(12), Tex_GUI(14), Tex_GUI(16), , , , , , GetAddress(AddressOf btnNewChar_Left)
    CreateButton WindowCount, "btnRight", 252, 40, 11, 13, , , , , , , Tex_GUI(13), Tex_GUI(15), Tex_GUI(17), , , , , , GetAddress(AddressOf btnNewChar_Right)

    ' Set the active control
    'SetActiveControl GetWindowIndex("winNewChar"), GetControlIndex("winNewChar", "txtName")
End Sub

Public Sub CreateWindow_EscMenu()
' Create window
    CreateWindow "winEscMenu", "", zOrder_Win, 0, 0, 210, 156, 0, , , , , , DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, , , , , , , , , False, False
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 6, 198, 144, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Buttons
    CreateButton WindowCount, "btnReturn", 16, 16, 178, 28, "Return to Game(" & KeycodeChar(Options.Options) & ")", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnEscMenu_Return)
    CreateButton WindowCount, "btnOptions", 16, 48, 178, 28, "Options", rockwellDec_15, , , , , , , , DesignTypes.desOrange, DesignTypes.desOrange_Hover, DesignTypes.desOrange_Click, , , GetAddress(AddressOf btnEscMenu_Options)
    CreateButton WindowCount, "btnMainMenu", 16, 80, 178, 28, "Back to Main Menu", rockwellDec_15, , , , , , , , DesignTypes.desBlue, DesignTypes.desBlue_Hover, DesignTypes.desBlue_Click, , , GetAddress(AddressOf btnEscMenu_MainMenu)
    CreateButton WindowCount, "btnExit", 16, 112, 178, 28, "Exit the Game", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnEscMenu_Exit)
End Sub

Public Sub CreateWindow_Bars()
' Create window
    CreateWindow "winBars", "", zOrder_Win, 10, 10, 239, 77, 0, , , , , , DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, , , , , , , , , False, False

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 6, 227, 65, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Blank Bars
    CreatePictureBox WindowCount, "picHP_Blank", 15, 15, 209, 13, , , , , Tex_GUI(24), Tex_GUI(24), Tex_GUI(24)
    CreatePictureBox WindowCount, "picSP_Blank", 15, 32, 209, 13, , , , , Tex_GUI(25), Tex_GUI(25), Tex_GUI(25)
    CreatePictureBox WindowCount, "picEXP_Blank", 15, 49, 209, 13, , , , , Tex_GUI(26), Tex_GUI(26), Tex_GUI(26)
    ' Draw the bars
    CreatePictureBox WindowCount, "picBlank", 0, 0, 0, 0, , , , , , , , , , , , , , , , GetAddress(AddressOf Bars_OnDraw)
    ' Bar Labels
    CreatePictureBox WindowCount, "picHealth", 16, 11, 44, 14, , , , , Tex_GUI(21), Tex_GUI(21), Tex_GUI(21)
    CreatePictureBox WindowCount, "picSpirit", 16, 28, 44, 14, , , , , Tex_GUI(22), Tex_GUI(22), Tex_GUI(22)
    CreatePictureBox WindowCount, "picExperience", 16, 45, 74, 14, , , , , Tex_GUI(23), Tex_GUI(23), Tex_GUI(23)
    ' Labels
    CreateLabel WindowCount, "lblHP", 15, 14, 209, 12, "999/999", rockwellDec_10, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblMP", 15, 31, 209, 12, "999/999", rockwellDec_10, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblEXP", 15, 48, 209, 12, "999/999", rockwellDec_10, White, Alignment.alignCentre
End Sub

Public Sub CreateWindow_Menu()
' Create window
    CreateWindow "winMenu", "", zOrder_Win, 564, 563, 229, 31, 0, , , , , , , , , , , , , , , , , False, False

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Wood part
    CreatePictureBox WindowCount, "picWood", 0, 5, 228, 21, , , , , , , , DesignTypes.desWood, DesignTypes.desWood, DesignTypes.desWood
    ' Buttons
    CreateButton WindowCount, "btnChar", 8, 1, 29, 29, , , Yellow, Tex_Item(108), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Char), , , -1, -2, , , "Character"
    CreateButton WindowCount, "btnInv", 44, 1, 29, 29, , , Yellow, Tex_Item(1), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Inv), , , -1, -2, , , "Inventory"
    CreateButton WindowCount, "btnSkills", 82, 1, 29, 29, , , Yellow, Tex_Item(109), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Skills), , , -1, -2, , , "Skills"
    CreateButton WindowCount, "btnGuild", 155, 1, 29, 29, , , Yellow, Tex_Item(107), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Guild), , , -1, -1, , , "Guild"
    CreateButton WindowCount, "btnQuest", 191, 1, 29, 29, , , Yellow, Tex_Item(23), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Quest), , , -1, -2, , , "Missoes"
    CreateButton WindowCount, "btnSerial", 119, 1, 29, 29, , , Yellow, Tex_Item(106), , , , , , DesignTypes.desGrey, DesignTypes.desGrey, DesignTypes.desGrey, , , , , , -1, -2, , , "Menu Disponivel"
End Sub

Public Sub CreateWindow_Hotbar()
' Create window
    CreateWindow "winHotbar", "", zOrder_Win, 372, 10, 418, 36, 0, , , , , , , , , , , , , GetAddress(AddressOf Hotbar_MouseMove), GetAddress(AddressOf Hotbar_MouseDown), GetAddress(AddressOf Hotbar_MouseMove), GetAddress(AddressOf Hotbar_DblClick), False, False, GetAddress(AddressOf DrawHotbar)
End Sub

Public Sub CreateWindow_Inventory()
' Create window
    CreateWindow "winInventory", "Inventory", zOrder_Win, 0, 0, 202, 319, Tex_Item(1), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , GetAddress(AddressOf Inventory_MouseMove), GetAddress(AddressOf Inventory_MouseDown), GetAddress(AddressOf Inventory_MouseMove), GetAddress(AddressOf Inventory_DblClick), , , GetAddress(AddressOf DrawInventory)
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Inv)
    ' Gold amount
    CreatePictureBox WindowCount, "picBlank", 8, 293, 186, 18, , , , , Tex_GUI(67), Tex_GUI(67), Tex_GUI(67)
    CreateLabel WindowCount, "lblGold", 42, 296, 100, 14, "0 $", verdanaBold_12, Yellow, , , , , , , GetAddress(AddressOf SendTradeGold)
End Sub

Public Sub CreateWindow_Description()
' Create window
    CreateWindow "winDescription", "", zOrder_Win, 0, 0, 193, 142, 0, , , , , , DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, , , , , , , , , False

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Name
    CreateLabel WindowCount, "lblName", 8, 12, 177, , "(SB) Flame Sword", rockwellDec_15, BrightBlue, Alignment.alignCentre
    ' Sprite box
    CreatePictureBox WindowCount, "picSprite", 18, 32, 68, 68, , , , , , , , DesignTypes.desDescPic, DesignTypes.desDescPic, DesignTypes.desDescPic, , , , , , GetAddress(AddressOf Description_OnDraw)
    ' Sep
    CreatePictureBox WindowCount, "picSep", 96, 28, 1, 92, , , , , Tex_GUI(44), Tex_GUI(44), Tex_GUI(44)
    ' Requirements
    CreateLabel WindowCount, "lblClass", 5, 102, 92, , "Warrior", verdana_12, LightGreen, Alignment.alignCentre
    CreateLabel WindowCount, "lblLevel", 5, 114, 92, , "Level 20", verdana_12, BrightRed, Alignment.alignCentre
    ' Bar
    CreatePictureBox WindowCount, "picBar", 19, 114, 66, 12, False, , , , Tex_GUI(45), Tex_GUI(45), Tex_GUI(45)
End Sub

Public Sub CreateWindow_DragBox()
' Create window
    CreateWindow "winDragBox", "", zOrder_Win, 0, 0, 32, 32, 0, , , , , , , , , , , , GetAddress(AddressOf DragBox_Check), , , , , , , GetAddress(AddressOf DragBox_OnDraw)
    ' Need to set up unique mouseup event
    Windows(WindowCount).Window.EntCallBack(entStates.MouseUp) = GetAddress(AddressOf DragBox_Check)
End Sub

Public Sub CreateWindow_Skills()
' Create window
    CreateWindow "winSkills", "Skills", zOrder_Win, 0, 0, 202, 297, Tex_Item(109), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , GetAddress(AddressOf Skills_MouseMove), GetAddress(AddressOf Skills_MouseDown), GetAddress(AddressOf Skills_MouseMove), GetAddress(AddressOf Skills_DblClick), , , GetAddress(AddressOf DrawSkills)
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Skills)
End Sub

Public Sub CreateWindow_Chat()
' Create window
    CreateWindow "winChat", "", zOrder_Win, 8, 422, 352, 152, 0, False, , , , , , , , , , , , , , , , False

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Channel boxes
    CreateCheckbox WindowCount, "chkGame", 10, 2, 49, 23, 1, "Game", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Game)
    CreateCheckbox WindowCount, "chkMap", 60, 2, 49, 23, 1, "Map", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Map)
    CreateCheckbox WindowCount, "chkGlobal", 110, 2, 49, 23, 1, "Global", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Global)
    CreateCheckbox WindowCount, "chkParty", 160, 2, 49, 23, 1, "Party", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Party)
    CreateCheckbox WindowCount, "chkGuild", 210, 2, 49, 23, 1, "Guild", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Guild)
    CreateCheckbox WindowCount, "chkPrivate", 260, 2, 49, 23, 1, "Private", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Private)
    CreateCheckbox WindowCount, "chkQuest", 310, 2, 49, 23, 1, "Quest", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Quest)
    CreateCheckbox WindowCount, "chkEvent", 310, 2, 49, 23, 1, "Event", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Event)
    ' Blank picturebox - ondraw wrapper
    CreatePictureBox WindowCount, "picNull", 0, 0, 0, 0, , , , , , , , , , , , , , , , GetAddress(AddressOf OnDraw_Chat)
    ' Chat button
    CreateButton WindowCount, "btnChat", 296, 124 + 16, 48, 20, "Say", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnSay_Click)
    ' Chat Textbox
    CreateTextbox WindowCount, "txtChat", 12, 127 + 16, 286, 25, , Fonts.verdana_12, , , , , , , , , , , , , , , , True, , , , , CHAT_LENGTH
    ' buttons
    CreateButton WindowCount, "btnUp", 328, 28, 11, 13, , , , , , , Tex_GUI(4), Tex_GUI(52), Tex_GUI(4), , , , , , GetAddress(AddressOf btnChat_Up)
    CreateButton WindowCount, "btnDown", 327, 122, 11, 13, , , , , , , Tex_GUI(5), Tex_GUI(53), Tex_GUI(5), , , , , , GetAddress(AddressOf btnChat_Down)
    ' Scroll
    CreateButton WindowCount, "btnScroll", 330, 50, 15, 78, , , , , False, , Tex_GUI(78), Tex_GUI(78), Tex_GUI(78), , , , , , GetAddress(AddressOf ChatScroll_MouseDown), GetAddress(AddressOf ChatScroll_MouseMove)

    ' Custom Handlers for mouse up
    Windows(WindowCount).Controls(GetControlIndex("winChat", "btnUp")).EntCallBack(entStates.MouseUp) = GetAddress(AddressOf btnChat_Up_MouseUp)
    Windows(WindowCount).Controls(GetControlIndex("winChat", "btnDown")).EntCallBack(entStates.MouseUp) = GetAddress(AddressOf btnChat_Down_MouseUp)

    ' Set the active control
    'SetActiveControl GetWindowIndex("winChat"), GetControlIndex("winChat", "txtChat")

    ' sort out the tabs
    With Windows(GetWindowIndex("winChat"))
        .Controls(GetControlIndex("winChat", "chkGame")).Value = Options.channelState(ChatChannel.chGame)
        .Controls(GetControlIndex("winChat", "chkMap")).Value = Options.channelState(ChatChannel.chMap)
        .Controls(GetControlIndex("winChat", "chkGlobal")).Value = Options.channelState(ChatChannel.chGlobal)
        .Controls(GetControlIndex("winChat", "chkParty")).Value = Options.channelState(ChatChannel.chParty)
        .Controls(GetControlIndex("winChat", "chkGuild")).Value = Options.channelState(ChatChannel.chGuild)
        .Controls(GetControlIndex("winChat", "chkPrivate")).Value = Options.channelState(ChatChannel.chPrivate)
        .Controls(GetControlIndex("winChat", "chkQuest")).Value = Options.channelState(ChatChannel.chQuest)
        .Controls(GetControlIndex("winChat", "chkEvent")).Value = Options.channelState(ChatChannel.chEvent)
    End With
End Sub

Public Sub CreateWindow_ChatSmall()
' Create window
    CreateWindow "winChatSmall", "", zOrder_Win, 8, 438, 0, 0, 0, False, , , , , , , , , , , , , , , , False, , GetAddress(AddressOf OnDraw_ChatSmall), , True

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Chat Label
    CreateLabel WindowCount, "lblMsg", 12, 127, 286, 25, "Press 'Enter' to open chatbox.", verdana_12, Grey
End Sub

Public Sub CreateWindow_Options()
' Create window
    CreateWindow "winOptions", "", zOrder_Win, 0, 0, 210, 262, 0, , , , , , DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, , , , , , , , , False, False
    ' Centralise it
    CentraliseWindow WindowCount

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 6, 198, 250, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' General
    CreatePictureBox WindowCount, "picBlank", 35, 25, 140, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblBlank", 35, 22, 140, , "General Options", rockwellDec_15, White, Alignment.alignCentre
    ' Check boxes
    CreateCheckbox WindowCount, "chkMusic", 35, 40, 80, , , "Music", rockwellDec_10, , , , , DesignTypes.desChkNorm
    CreateCheckbox WindowCount, "chkSound", 115, 40, 80, , , "Sound", rockwellDec_10, , , , , DesignTypes.desChkNorm
    CreateCheckbox WindowCount, "chkAutotiles", 35, 60, 80, , , "Autotiles", rockwellDec_10, , , , , DesignTypes.desChkNorm
    CreateCheckbox WindowCount, "chkFullscreen", 115, 60, 80, , , "Fullscreen", rockwellDec_10, , , , , DesignTypes.desChkNorm
    CreateCheckbox WindowCount, "chkReconnect", 35, 80, 80, , , "Reconect", rockwellDec_10, , , , , DesignTypes.desChkNorm
    CreateCheckbox WindowCount, "chkItemName", 115, 80, 80, , , "Item Name", rockwellDec_10, , , , , DesignTypes.desChkNorm
    CreateCheckbox WindowCount, "chkItemAnimation", 115, 100, 80, , , "Item Animation", rockwellDec_10, , , , , DesignTypes.desChkNorm
    CreateCheckbox WindowCount, "chkFPSConection", 35, 100, 80, , , "Fps/Ping", rockwellDec_10, , , , , DesignTypes.desChkNorm
    ' Resolution
    CreatePictureBox WindowCount, "picBlank", 35, 115, 140, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblBlank", 35, 112, 140, , "Select Resolution", rockwellDec_15, White, Alignment.alignCentre
    ' combobox
    CreateComboBox WindowCount, "cmbRes", 30, 130, 150, 18, DesignTypes.desComboNorm, verdana_12
    ' Renderer
    CreatePictureBox WindowCount, "picBlank", 35, 155, 140, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblBlank", 35, 152, 140, , "DirectX Mode", rockwellDec_15, White, Alignment.alignCentre
    ' Check boxes
    CreateComboBox WindowCount, "cmbRender", 30, 170, 150, 18, DesignTypes.desComboNorm, verdana_12
    ' Button
    CreateButton WindowCount, "btnChangeControls", 45, 200, 120, 15, "Change Controls", rockwellDec_15, , , , , , , , DesignTypes.desBlue, DesignTypes.desBlue_Hover, DesignTypes.desBlue_Click, , , GetAddress(AddressOf btnChangeControls_Open)
    ' Button
    CreateButton WindowCount, "btnConfirm", (Windows(WindowCount).Window.Width / 2) - (80 / 2), (Windows(WindowCount).Window.Height) - 40, 80, 22, "Confirm", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnOptions_Confirm)

    ' Populate the options screen
    SetOptionsScreen
End Sub

Public Sub CreateWindow_Shop()
' Create window
    CreateWindow "winShop", "Shop", zOrder_Win, 0, 0, 278, 293, Tex_Item(17), False, Fonts.rockwellDec_15, , 2, 5, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , GetAddress(AddressOf Shop_MouseMove), GetAddress(AddressOf Shop_MouseDown), GetAddress(AddressOf Shop_MouseMove), GetAddress(AddressOf Shop_MouseMove), , , GetAddress(AddressOf DrawShopBackground)
    ' additional mouse event
    Windows(WindowCount).Window.EntCallBack(entStates.MouseUp) = GetAddress(AddressOf Shop_MouseMove)
    ' Centralise it
    CentraliseWindow WindowCount

    zOrder_Con = 1

    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnShop_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 215, 266, 50, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment, , , , , , GetAddress(AddressOf DrawShop)
    ' Picture Box
    CreatePictureBox WindowCount, "picItemBG", 13, 222, 36, 36, , , , , Tex_GUI(54), Tex_GUI(54), Tex_GUI(54)
    CreatePictureBox WindowCount, "picItem", 15, 224, 32, 32
    ' Buttons
    CreateButton WindowCount, "btnBuy", 190, 228, 70, 24, "Buy", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnShopBuy)
    CreateButton WindowCount, "btnSell", 190, 228, 70, 24, "Sell", rockwellDec_15, White, , False, , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnShopSell)
    ' Gold
    CreatePictureBox WindowCount, "picBlank", 9, 266, 162, 18, , , , , Tex_GUI(55), Tex_GUI(55), Tex_GUI(55)
    ' Buying/Selling
    CreateCheckbox WindowCount, "chkBuying", 173, 265, 49, 20, 1, , , , , , , DesignTypes.desChkCustom_Buying, , , GetAddress(AddressOf chkShopBuying)
    CreateCheckbox WindowCount, "chkSelling", 222, 265, 49, 20, 0, , , , , , , DesignTypes.desChkCustom_Selling, , , GetAddress(AddressOf chkShopSelling)
    ' Labels
    CreateLabel WindowCount, "lblName", 56, 226, 300, , "Test Item", verdanaBold_12, Black, Alignment.alignLeft
    CreateLabel WindowCount, "lblCost", 56, 240, 300, , "1000 $", verdana_12, Black, Alignment.alignLeft
    ' Gold
    CreateLabel WindowCount, "lblGold", 44, 269, 300, , "0 $", verdana_12
End Sub

Public Sub CreateWindow_NpcChat()
' Create window
    CreateWindow "winNpcChat", "Conversation with [Name]", zOrder_Win, 0, 0, 480, 228, Tex_Item(111), False, Fonts.rockwellDec_15, , 2, 11, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount

    zOrder_Con = 1

    ' Close Button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnNpcChat_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 468, 198, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Face background
    CreatePictureBox WindowCount, "picFaceBG", 20, 40, 102, 102, , , , , Tex_GUI(60), Tex_GUI(60), Tex_GUI(60)
    ' Actual Face
    CreatePictureBox WindowCount, "picFace", 23, 43, 96, 96, , , , , Tex_Face(1), Tex_Face(1), Tex_Face(1)
    ' Chat BG
    CreatePictureBox WindowCount, "picChatBG", 128, 39, 334, 104, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    ' Chat
    CreateLabel WindowCount, "lblChat", 136, 44, 318, 102, "[Text]", rockwellDec_15, White, Alignment.alignCentre
    ' Reply buttons
    CreateButton WindowCount, "btnOpt4", 69, 145, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt4), , , , , DarkGrey
    CreateButton WindowCount, "btnOpt3", 69, 162, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt3), , , , , DarkGrey
    CreateButton WindowCount, "btnOpt2", 69, 179, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt2), , , , , DarkGrey
    CreateButton WindowCount, "btnOpt1", 69, 196, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt1), , , , , DarkGrey

    ' Cache positions
    optPos(1) = 196
    optPos(2) = 179
    optPos(3) = 162
    optPos(4) = 145
    optHeight = 228
End Sub

Public Sub CreateWindow_Message()
' Create window
    CreateWindow "winMessage", "Mensagem!", zOrder_Win, 0, 0, 358, 169, Tex_Item(111), False, Fonts.rockwellDec_15, , 2, 11, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount

    zOrder_Con = 1

    ' Close Button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMessage_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 346, 130, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Chat BG
    CreatePictureBox WindowCount, "picChatBG", 12, 39, 334, 104, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    ' Chat
    CreateLabel WindowCount, "lblChat", 20, 44, 318, 102, "[Text]", rockwellDec_15, White, Alignment.alignCentre
End Sub

Public Sub CreateWindow_RightClick()
' Create window
    CreateWindow "winRightClickBG", "", zOrder_Win, 0, 0, 800, 600, 0, , , , , , , , , , , , , , GetAddress(AddressOf RightClick_Close), , , False
    ' Centralise it
    CentraliseWindow WindowCount
End Sub

Public Sub CreateWindow_PlayerMenu()
' Create window
    CreateWindow "winPlayerMenu", "", zOrder_Win, 0, 0, 110, 106, 0, , , , , , DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, , , , , , GetAddress(AddressOf RightClick_Close), , , False
    ' Centralise it
    CentraliseWindow WindowCount

    zOrder_Con = 1

    ' Name
    CreateButton WindowCount, "btnName", 8, 8, 94, 18, "[Name]", verdanaBold_12, White, , , , , , , DesignTypes.desMenuHeader, DesignTypes.desMenuHeader, DesignTypes.desMenuHeader, , , GetAddress(AddressOf RightClick_Close)
    ' Options
    CreateButton WindowCount, "btnParty", 8, 26, 94, 18, "Invite to Party", verdana_12, White, , , , , , , , DesignTypes.desMenuOption, , , , GetAddress(AddressOf PlayerMenu_Party)
    CreateButton WindowCount, "btnTrade", 8, 44, 94, 18, "Request Trade", verdana_12, White, , , , , , , , DesignTypes.desMenuOption, , , , GetAddress(AddressOf PlayerMenu_Trade)
    CreateButton WindowCount, "btnGuild", 8, 62, 94, 18, "Invite to Guild", verdana_12, White, , , , , , , , DesignTypes.desMenuOption, , , , GetAddress(AddressOf PlayerMenu_Guild)
    CreateButton WindowCount, "btnPM", 8, 80, 94, 18, "Private Message", verdana_12, White, , , , , , , , DesignTypes.desMenuOption, , , , GetAddress(AddressOf PlayerMenu_PM)
End Sub

Public Sub CreateWindow_Party()
' Create window
    CreateWindow "winParty", "", zOrder_Win, 4, 78, 252, 158, 0, , , , , , DesignTypes.desWin_Party, DesignTypes.desWin_Party, DesignTypes.desWin_Party, , , , , , , , , False

    zOrder_Con = 1

    ' Name labels
    CreateLabel WindowCount, "lblName1", 60, 20, 173, , "Richard - Level 10", rockwellDec_10
    CreateLabel WindowCount, "lblName2", 60, 60, 173, , "Anna - Level 18", rockwellDec_10
    CreateLabel WindowCount, "lblName3", 60, 100, 173, , "Doleo - Level 25", rockwellDec_10
    ' Empty Bars - HP
    CreatePictureBox WindowCount, "picEmptyBar_HP1", 58, 34, 173, 9, , , , , Tex_GUI(62), Tex_GUI(62), Tex_GUI(62)
    CreatePictureBox WindowCount, "picEmptyBar_HP2", 58, 74, 173, 9, , , , , Tex_GUI(62), Tex_GUI(62), Tex_GUI(62)
    CreatePictureBox WindowCount, "picEmptyBar_HP3", 58, 114, 173, 9, , , , , Tex_GUI(62), Tex_GUI(62), Tex_GUI(62)
    ' Empty Bars - SP
    CreatePictureBox WindowCount, "picEmptyBar_SP1", 58, 44, 173, 9, , , , , Tex_GUI(63), Tex_GUI(63), Tex_GUI(63)
    CreatePictureBox WindowCount, "picEmptyBar_SP2", 58, 84, 173, 9, , , , , Tex_GUI(63), Tex_GUI(63), Tex_GUI(63)
    CreatePictureBox WindowCount, "picEmptyBar_SP3", 58, 124, 173, 9, , , , , Tex_GUI(63), Tex_GUI(63), Tex_GUI(63)
    ' Filled bars - HP
    CreatePictureBox WindowCount, "picBar_HP1", 58, 34, 173, 9, , , , , Tex_GUI(64), Tex_GUI(64), Tex_GUI(64)
    CreatePictureBox WindowCount, "picBar_HP2", 58, 74, 173, 9, , , , , Tex_GUI(64), Tex_GUI(64), Tex_GUI(64)
    CreatePictureBox WindowCount, "picBar_HP3", 58, 114, 173, 9, , , , , Tex_GUI(64), Tex_GUI(64), Tex_GUI(64)
    ' Filled bars - SP
    CreatePictureBox WindowCount, "picBar_SP1", 58, 44, 173, 9, , , , , Tex_GUI(65), Tex_GUI(65), Tex_GUI(65)
    CreatePictureBox WindowCount, "picBar_SP2", 58, 84, 173, 9, , , , , Tex_GUI(65), Tex_GUI(65), Tex_GUI(65)
    CreatePictureBox WindowCount, "picBar_SP3", 58, 124, 173, 9, , , , , Tex_GUI(65), Tex_GUI(65), Tex_GUI(65)
    ' Shadows
    CreatePictureBox WindowCount, "picShadow1", 20, 24, 32, 32, , , , , Tex_Shadow, Tex_Shadow, Tex_Shadow
    CreatePictureBox WindowCount, "picShadow2", 20, 64, 32, 32, , , , , Tex_Shadow, Tex_Shadow, Tex_Shadow
    CreatePictureBox WindowCount, "picShadow3", 20, 104, 32, 32, , , , , Tex_Shadow, Tex_Shadow, Tex_Shadow
    ' Characters
    CreatePictureBox WindowCount, "picChar1", 20, 20, 32, 32, , , , , Tex_Char(1), Tex_Char(1), Tex_Char(1)
    CreatePictureBox WindowCount, "picChar2", 20, 60, 32, 32, , , , , Tex_Char(1), Tex_Char(1), Tex_Char(1)
    CreatePictureBox WindowCount, "picChar3", 20, 100, 32, 32, , , , , Tex_Char(1), Tex_Char(1), Tex_Char(1)
End Sub

Public Sub CreateWindow_Invitations()
' Create window
    CreateWindow "winInvite_Party", "", zOrder_Win, screenWidth - 234, screenHeight - 80, 223, 37, 0, , , , , , DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, , , , , , , , , False

    zOrder_Con = 1

    ' Button
    CreateButton WindowCount, "btnInvite", 11, 12, 201, 14, ColourChar & White & "Richard " & ColourChar & "-1" & "has invited you to a party.", verdana_12, Grey, , , , , , , , , , , , GetAddress(AddressOf btnInvite_Party), , , , , Green

    ' Create window
    CreateWindow "winInvite_Trade", "", zOrder_Win, screenWidth - 234, screenHeight - 80, 223, 37, 0, , , , , , DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, , , , , , , , , False
    ' Button
    CreateButton WindowCount, "btnInvite", 11, 12, 201, 14, ColourChar & White & "Richard " & ColourChar & "-1" & "has invited you to a party.", verdana_12, Grey, , , , , , , , , , , , GetAddress(AddressOf btnInvite_Trade), , , , , Green

    ' Create window
    CreateWindow "winInvite_Guild", "", zOrder_Win, screenWidth - 234, screenHeight - 80, 223, 37, 0, , , , , , DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, , , , , , , , , False
    ' Button
    CreateButton WindowCount, "btnInvite", 11, 12, 201, 14, ColourChar & White & "Richard " & ColourChar & "-1" & "has invited you to a Guild.", verdana_12, Grey, , , , , , , , , , , , GetAddress(AddressOf btnInvite_Guild), , , , , Green
End Sub

Public Sub CreateWindow_Trade()
' Create window
    CreateWindow "winTrade", "Trading with [Name]", zOrder_Win, 0, 0, 412, 386, Tex_Item(112), False, Fonts.rockwellDec_15, , 2, 5, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , , , , , , , GetAddress(AddressOf DrawTrade)
    ' Centralise it
    CentraliseWindow WindowCount

    zOrder_Con = 1

    ' Close Button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnTrade_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 10, 312, 392, 66, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Labels
    CreatePictureBox WindowCount, "picShadow", 36, 30, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblYourTrade", 36, 27, 142, 9, "Robin's Offer", rockwellDec_15, White, Alignment.alignCentre
    CreatePictureBox WindowCount, "picShadow", 36 + 200, 30, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblTheirTrade", 36 + 200, 27, 142, 9, "Richard's Offer", rockwellDec_15, White, Alignment.alignCentre
    ' Buttons
    CreateButton WindowCount, "btnAccept", 134, 340, 68, 24, "Accept", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnTrade_Accept)
    CreateButton WindowCount, "btnDecline", 210, 340, 68, 24, "Decline", rockwellDec_15, White, , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnTrade_Close)
    ' Labels
    CreateLabel WindowCount, "lblStatus", 114, 322, 184, , "", verdanaBold_12, Red, Alignment.alignCentre
    ' Amounts
    CreateLabel WindowCount, "lblBlank", 25, 330, 100, , "Total Value", verdanaBold_12, Black, Alignment.alignCentre
    CreateLabel WindowCount, "lblBlank", 285, 330, 100, , "Total Value", verdanaBold_12, Black, Alignment.alignCentre
    CreateLabel WindowCount, "lblYourValue", 25, 344, 100, , "52,812g", verdana_12, Black, Alignment.alignCentre
    CreateLabel WindowCount, "lblTheirValue", 285, 344, 100, , "12,531g", verdana_12, Black, Alignment.alignCentre
    ' Item Containers
    CreatePictureBox WindowCount, "picYour", 14, 46, 184, 260, , , , , , , , , , , , GetAddress(AddressOf TradeMouseMove_Your), GetAddress(AddressOf TradeMouseDown_Your), GetAddress(AddressOf TradeMouseMove_Your), , GetAddress(AddressOf DrawYourTrade)
    CreatePictureBox WindowCount, "picTheir", 214, 46, 184, 260, , , , , , , , , , , , GetAddress(AddressOf TradeMouseMove_Their), GetAddress(AddressOf TradeMouseMove_Their), GetAddress(AddressOf TradeMouseMove_Their), , GetAddress(AddressOf DrawTheirTrade)
End Sub

Public Sub CreateWindow_Combobox()
' background window
    CreateWindow "winComboMenuBG", "ComboMenuBG", zOrder_Win, 0, 0, 800, 600, 0, , , , , , , , , , , , , , GetAddress(AddressOf CloseComboMenu), , , False, False

    zOrder_Con = 1

    ' window
    CreateWindow "winComboMenu", "ComboMenu", zOrder_Win, 0, 0, 100, 100, 0, , Fonts.verdana_12, , , , DesignTypes.desComboMenuNorm, , , , , , , , , , , False, False
    ' centralise it
    CentraliseWindow WindowCount
End Sub

' Rendering & Initialisation
Public Sub InitGUI()

' Starter values
    zOrder_Win = 1
    zOrder_Con = 1

    ' Menu
    CreateWindow_Login
    CreateWindow_Loading
    CreateWindow_Dialogue
    CreateWindow_Classes
    CreateWindow_NewChar
    CreateWindow_Register
    CreateWindow_Reconnect
    CreateWindow_Serial

    ' Game
    CreateWindow_Combobox
    CreateWindow_EscMenu
    CreateWindow_Bars
    CreateWindow_Menu
    CreateWindow_Hotbar
    CreateWindow_Inventory
    CreateWindow_Character
    CreateWindow_Description
    CreateWindow_DragBox
    CreateWindow_Skills
    CreateWindow_Chat
    CreateWindow_ChatSmall
    CreateWindow_Options
    CreateWindow_ChangeControls
    CreateWindow_Shop
    CreateWindow_NpcChat
    CreateWindow_Party
    CreateWindow_EnemyBars
    CreateWindow_Invitations
    CreateWindow_Trade
    CreateWindow_Guild
    CreateWindow_GuildMaker
    CreateWindow_GuildMenu
    CreateWindow_Bank
    CreateWindow_Quest
    CreateWindow_Message
    CreateWindow_CheckIn
    CreateWindow_Lottery

    ' Menus
    CreateWindow_Clipboard
    CreateWindow_RightClick
    CreateWindow_PlayerMenu
End Sub
