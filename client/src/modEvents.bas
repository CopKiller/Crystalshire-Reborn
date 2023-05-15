Attribute VB_Name = "modEvents"
Option Explicit

Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

' temporary event
Public cpEvent As EventRec

Sub CopyEvent_Map(X As Long, Y As Long)
    Dim Count As Long, i As Long
    Count = Map.TileData.EventCount
    If Count = 0 Then Exit Sub

    For i = 1 To Count
        If Map.TileData.Events(i).X = X And Map.TileData.Events(i).Y = Y Then
            ' copy it
            CopyMemory ByVal VarPtr(cpEvent), ByVal VarPtr(Map.TileData.Events(i)), LenB(Map.TileData.Events(i))
            ' exit
            Exit Sub
        End If
    Next
End Sub

Sub PasteEvent_Map(X As Long, Y As Long)
    Dim Count As Long, i As Long, eventNum As Long
    Count = Map.TileData.EventCount

    If Count > 0 Then
        For i = 1 To Count
            If Map.TileData.Events(i).X = X And Map.TileData.Events(i).Y = Y Then
                ' already an event - paste over it
                eventNum = i
            End If
        Next
    End If

    ' couldn't find one - create one
    If eventNum = 0 Then
        ' increment count
        AddEvent X, Y, True
        eventNum = Count + 1
    End If

    ' copy it
    CopyMemory ByVal VarPtr(Map.TileData.Events(eventNum)), ByVal VarPtr(cpEvent), LenB(cpEvent)

    ' set position
    Map.TileData.Events(eventNum).X = X
    Map.TileData.Events(eventNum).Y = Y
End Sub

Sub AddEvent(X As Long, Y As Long, Optional ByVal cancelLoad As Boolean = False)
    Dim Count As Long, pageCount As Long, i As Long
    Count = Map.TileData.EventCount + 1
    ' make sure there's not already an event
    If Count - 1 > 0 Then
        For i = 1 To Count - 1
            If Map.TileData.Events(i).X = X And Map.TileData.Events(i).Y = Y Then
                ' already an event - edit it
                If Not cancelLoad Then EventEditorInit i
                Exit Sub
            End If
        Next
    End If
    ' increment count
    Map.TileData.EventCount = Count
    ReDim Preserve Map.TileData.Events(1 To Count)
    ' set the new event
    Map.TileData.Events(Count).X = X
    Map.TileData.Events(Count).Y = Y
    ' give it a new page
    pageCount = Map.TileData.Events(Count).pageCount + 1
    Map.TileData.Events(Count).pageCount = pageCount
    ReDim Preserve Map.TileData.Events(Count).EventPage(1 To pageCount)
    ' load the editor
    If Not cancelLoad Then EventEditorInit Count
End Sub

Sub DeleteEvent(X As Long, Y As Long)
    Dim Count As Long, i As Long, lowIndex As Long
    If Not InMapEditor Then Exit Sub

    Count = Map.TileData.EventCount
    For i = 1 To Count
        If Map.TileData.Events(i).X = X And Map.TileData.Events(i).Y = Y Then
            ' delete it
            ClearEvent i
            lowIndex = i
            Exit For
        End If
    Next

    ' not found anything
    If lowIndex = 0 Then Exit Sub

    ' move everything down an index
    For i = lowIndex To Count - 1
        CopyEvent i + 1, i
    Next
    ' delete the last index
    ClearEvent Count
    ' set the new count
    Map.TileData.EventCount = Count - 1
End Sub

Sub ClearEvent(eventNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map.TileData.Events(eventNum)), LenB(Map.TileData.Events(eventNum)))
End Sub

Sub CopyEvent(original As Long, newone As Long)
    CopyMemory ByVal VarPtr(Map.TileData.Events(newone)), ByVal VarPtr(Map.TileData.Events(original)), LenB(Map.TileData.Events(original))
End Sub

Sub EventEditorInit(eventNum As Long)
    Dim i As Long
    EditorEvent = eventNum
    ' copy the event data to the temp event
    CopyMemory ByVal VarPtr(tmpEvent), ByVal VarPtr(Map.TileData.Events(eventNum)), LenB(Map.TileData.Events(eventNum))
    ' populate form
    With frmEditor_Events
        ' set the tabs
        .tabPages.Tabs.Clear
        For i = 1 To tmpEvent.pageCount
            .tabPages.Tabs.Add , , Str(i)
        Next
        ' items
        .cmbHasItem.Clear
        .cmbHasItem.AddItem "None"
        For i = 1 To MAX_ITEMS
            .cmbHasItem.AddItem i & ": " & Trim$(Item(i).Name)
        Next
        ' variables
        .cmbPlayerVar.Clear
        .cmbPlayerVar.AddItem "None"
        For i = 1 To MAX_BYTE
            .cmbPlayerVar.AddItem i
        Next
        ' name
        .txtName.text = tmpEvent.Name
        ' enable delete button
        If tmpEvent.pageCount > 1 Then
            .cmdDeletePage.enabled = True
        Else
            .cmdDeletePage.enabled = False
        End If
        .cmdPastePage.enabled = False
        ' set the commands frame
        .fraCommands.Width = 417
        .fraCommands.Height = 497
        ' set the dialogue frame
        .fraDialogue.Width = 417
        .fraDialogue.Height = 497
        ' Load page 1 to start off with
        curPageNum = 1
        EventEditorLoadPage curPageNum
    End With
    ' show the editor
    frmEditor_Events.Show
End Sub

Sub AddCommand(theType As EventType)
    Dim Count As Long
    ' update the array
    With tmpEvent.EventPage(curPageNum)
        Count = .CommandCount + 1
        ReDim Preserve .Commands(1 To Count)
        .CommandCount = Count
        ' set the shit
        Select Case theType
        Case EventType.evAddText
            ' set the values
            .Commands(Count).Type = EventType.evAddText
            .Commands(Count).text = frmEditor_Events.txtAddText_Text.text
            .Commands(Count).Colour = frmEditor_Events.scrlAddText_Colour.Value
            If frmEditor_Events.optAddText_Game.Value Then
                .Commands(Count).channel = 0
            ElseIf frmEditor_Events.optAddText_Map.Value Then
                .Commands(Count).channel = 1
            ElseIf frmEditor_Events.optAddText_Global.Value Then
                .Commands(Count).channel = 2
            End If
        Case EventType.evShowChatBubble
            .Commands(Count).Type = EventType.evShowChatBubble
            .Commands(Count).text = frmEditor_Events.txtChatBubble.text
            .Commands(Count).Colour = frmEditor_Events.scrlChatBubble.Value
            .Commands(Count).TargetType = frmEditor_Events.cmbChatBubbleType.ListIndex
            .Commands(Count).Target = frmEditor_Events.cmbChatBubble.ListIndex
        Case EventType.evPlayerVar
            .Commands(Count).Type = EventType.evPlayerVar
            .Commands(Count).Target = frmEditor_Events.cmbVariable.ListIndex
            .Commands(Count).Colour = Val(frmEditor_Events.txtVariable.text)
        Case EventType.evWarpPlayer
            .Commands(Count).Type = EventType.evWarpPlayer
            .Commands(Count).X = frmEditor_Events.scrlWPX.Value
            .Commands(Count).Y = frmEditor_Events.scrlWPY.Value
            .Commands(Count).Target = frmEditor_Events.scrlWPMap.Value
        End Select
    End With
    ' re-list the commands
    EventListCommands
End Sub

Sub EditCommand()
    With tmpEvent.EventPage(curPageNum).Commands(curCommand)
        Select Case .Type
        Case EventType.evAddText
            .text = frmEditor_Events.txtAddText_Text.text
            .Colour = frmEditor_Events.scrlAddText_Colour.Value
            If frmEditor_Events.optAddText_Game.Value Then
                .channel = 0
            ElseIf frmEditor_Events.optAddText_Map.Value Then
                .channel = 1
            ElseIf frmEditor_Events.optAddText_Global.Value Then
                .channel = 2
            End If
        Case EventType.evShowChatBubble
            .text = frmEditor_Events.txtChatBubble.text
            .Colour = frmEditor_Events.scrlChatBubble.Value
            .TargetType = frmEditor_Events.cmbChatBubbleType.ListIndex
            .Target = frmEditor_Events.cmbChatBubble.ListIndex
        Case EventType.evPlayerVar
            .Target = frmEditor_Events.cmbVariable.ListIndex
            .Colour = Val(frmEditor_Events.txtVariable.text)
        Case EventType.evWarpPlayer
            .X = frmEditor_Events.scrlWPX.Value
            .Y = frmEditor_Events.scrlWPY.Value
        End Select
    End With
    ' re-list the commands
    EventListCommands
End Sub

Sub EventListCommands()
    Dim i As Long, Count As Long
    frmEditor_Events.lstCommands.Clear
    ' check if there are any
    Count = tmpEvent.EventPage(curPageNum).CommandCount
    If Count > 0 Then
        ' list them
        For i = 1 To Count
            With tmpEvent.EventPage(curPageNum).Commands(i)
                Select Case .Type
                Case EventType.evAddText
                    ListCommandAdd "@>Add Text: " & .text & " - Colour: " & GetColourString(.Colour) & " - Channel: " & .channel
                Case EventType.evShowChatBubble
                    ListCommandAdd "@>Show Chat Bubble: " & .text & " - Colour: " & GetColourString(.Colour) & " - Target Type: " & .TargetType & " - Target: " & .Target
                Case EventType.evPlayerVar
                    ListCommandAdd "@>Change variable #" & .Target & " to " & .Colour
                Case EventType.evWarpPlayer
                    ListCommandAdd "@>Warp Player to Map #" & .Target & ", X: " & .X & ", Y: " & .Y
                Case Else
                    ListCommandAdd "@>Unknown"
                End Select
            End With
        Next
    Else
        frmEditor_Events.lstCommands.AddItem "@>"
    End If
    frmEditor_Events.lstCommands.ListIndex = 0
    curCommand = 1
End Sub

Sub ListCommandAdd(s As String)
    Static X As Long
    frmEditor_Events.lstCommands.AddItem s
    ' scrollbar
    If X < frmEditor_Events.TextWidth(s & "  ") Then
        X = frmEditor_Events.TextWidth(s & "  ")
        If frmEditor_Events.ScaleMode = vbTwips Then X = X / Screen.TwipsPerPixelX    ' if twips change to pixels
        SendMessageByNum frmEditor_Events.lstCommands.hwnd, LB_SETHORIZONTALEXTENT, X, 0
    End If
End Sub

Sub EventEditorLoadPage(pageNum As Long)
' populate form
    With tmpEvent.EventPage(pageNum)
        GraphicSelX = .GraphicX
        GraphicSelY = .GraphicY
        frmEditor_Events.cmbGraphic.ListIndex = .GraphicType
        frmEditor_Events.cmbHasItem.ListIndex = .HasItemNum
        frmEditor_Events.cmbMoveFreq.ListIndex = .MoveFreq
        frmEditor_Events.cmbMoveSpeed.ListIndex = .MoveSpeed
        frmEditor_Events.cmbMoveType.ListIndex = .MoveType
        frmEditor_Events.cmbPlayerVar.ListIndex = .PlayerVarNum
        frmEditor_Events.cmbPriority.ListIndex = .Priority
        frmEditor_Events.cmbSelfSwitch.ListIndex = .SelfSwitchNum
        frmEditor_Events.cmbTrigger.ListIndex = .Trigger
        frmEditor_Events.chkDirFix.Value = .DirFix
        frmEditor_Events.chkHasItem.Value = .chkHasItem
        frmEditor_Events.chkPlayerVar.Value = .chkPlayerVar
        frmEditor_Events.chkSelfSwitch.Value = .chkSelfSwitch
        frmEditor_Events.chkStepAnim.Value = .StepAnim
        frmEditor_Events.chkWalkAnim.Value = .WalkAnim
        frmEditor_Events.chkWalkThrough.Value = .WalkThrough
        frmEditor_Events.txtPlayerVariable = .PlayerVariable
        frmEditor_Events.scrlGraphic.Value = .Graphic
        If .chkHasItem = 0 Then frmEditor_Events.cmbHasItem.enabled = False Else frmEditor_Events.cmbHasItem.enabled = True
        If .chkSelfSwitch = 0 Then frmEditor_Events.cmbSelfSwitch.enabled = False Else frmEditor_Events.cmbSelfSwitch.enabled = True
        If .chkPlayerVar = 0 Then
            frmEditor_Events.cmbPlayerVar.enabled = False
            frmEditor_Events.txtPlayerVariable.enabled = False
        Else
            frmEditor_Events.cmbPlayerVar.enabled = True
            frmEditor_Events.txtPlayerVariable.enabled = True
        End If
        ' show the commands
        EventListCommands
    End With
End Sub

Sub EventEditorOK()
' copy the event data from the temp event
    CopyMemory ByVal VarPtr(Map.TileData.Events(EditorEvent)), ByVal VarPtr(tmpEvent), LenB(tmpEvent)
    ' unload the form
    Unload frmEditor_Events
End Sub
