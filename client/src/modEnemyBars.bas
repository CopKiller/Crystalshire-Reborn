Attribute VB_Name = "modEnemyBars"
Option Explicit

Public Sub CreateWindow_EnemyBars()
' Create window
    CreateWindow "winEnemyBars", "", zOrder_Win, (screenWidth - 252), 78, 252, 158, 0, , , , , , DesignTypes.desWin_Party, DesignTypes.desWin_Party, DesignTypes.desWin_Party, , , , , , , , , False

    ' Set the index for spawning controls
    zOrder_Con = 1

    ' Close Button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.Width - 30, 15, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf Close_EnemyBar)
    ' Name labels
    CreateLabel WindowCount, "lblName", 60, 20, 173, , "Enemy - Level 10", rockwellDec_10
    ' Empty Bars - HP
    CreatePictureBox WindowCount, "picEmptyBar_HP", 58, 34, 173, 9, , , , , Tex_GUI(62), Tex_GUI(62), Tex_GUI(62)
    ' Empty Bars - SP
    CreatePictureBox WindowCount, "picEmptyBar_SP", 58, 44, 173, 9, , , , , Tex_GUI(63), Tex_GUI(63), Tex_GUI(63)
    ' Filled bars - HP
    CreatePictureBox WindowCount, "picBar_HP", 58, 34, 173, 9, , , , , Tex_GUI(64), Tex_GUI(64), Tex_GUI(64)
    ' Filled bars - SP
    CreatePictureBox WindowCount, "picBar_SP", 58, 44, 173, 9, , , , , Tex_GUI(65), Tex_GUI(65), Tex_GUI(65)
    ' Shadows
    CreatePictureBox WindowCount, "picShadow", 20, 24, 32, 32, , , , , Tex_Shadow, Tex_Shadow, Tex_Shadow
    ' Characters
    CreatePictureBox WindowCount, "picChar", 20, 20, 32, 32
End Sub

Private Sub Close_EnemyBar()
    HideWindow GetWindowIndex("winEnemyBars")
End Sub

Sub UpdateEnemyInterface()
    Dim pIndex As Long

    ' unload it if we're not in target
    If myTargetType = 0 Or myTarget = 0 Then
        HideWindow GetWindowIndex("winEnemyBars")
        Exit Sub
    End If

    ' load the window
    ShowWindow GetWindowIndex("winEnemyBars")

    ' fill the controls
    With Windows(GetWindowIndex("winEnemyBars"))

        ' clear controls first
        .Controls(GetControlIndex("winEnemyBars", "lblName")).text = vbNullString
        .Controls(GetControlIndex("winEnemyBars", "picEmptyBar_HP")).visible = False
        .Controls(GetControlIndex("winEnemyBars", "picEmptyBar_SP")).visible = False
        .Controls(GetControlIndex("winEnemyBars", "picBar_HP")).visible = False
        .Controls(GetControlIndex("winEnemyBars", "picBar_SP")).visible = False
        .Controls(GetControlIndex("winEnemyBars", "picShadow")).visible = False
        .Controls(GetControlIndex("winEnemyBars", "picChar")).visible = False
        .Controls(GetControlIndex("winEnemyBars", "picChar")).Value = 0
        Windows(GetWindowIndex("winEnemyBars")).Window.origLeft = (screenWidth - 252)

        ' labels
        pIndex = myTarget
        If pIndex > 0 Then
            ' If pIndex <> MyIndex Then


            If myTargetType = TARGET_TYPE_PLAYER Then
                If IsPlaying(pIndex) Then
                    ' name and level
                    .Controls(GetControlIndex("winEnemyBars", "lblName")).visible = True
                    .Controls(GetControlIndex("winEnemyBars", "lblName")).text = Trim$(GetPlayerName(pIndex)) & " - " & GetPlayerLevel(pIndex)
                    ' picture
                    .Controls(GetControlIndex("winEnemyBars", "picShadow")).visible = True
                    .Controls(GetControlIndex("winEnemyBars", "picChar")).visible = True
                    ' store the player's index as a value for later use
                    .Controls(GetControlIndex("winEnemyBars", "picChar")).Value = pIndex
                    .Controls(GetControlIndex("winEnemyBars", "picChar")).image(0) = Tex_Char(GetPlayerSprite(pIndex))
                    .Controls(GetControlIndex("winEnemyBars", "picChar")).image(1) = Tex_Char(GetPlayerSprite(pIndex))
                    .Controls(GetControlIndex("winEnemyBars", "picChar")).image(2) = Tex_Char(GetPlayerSprite(pIndex))
                    ' bars
                    .Controls(GetControlIndex("winEnemyBars", "picEmptyBar_HP")).visible = True
                    .Controls(GetControlIndex("winEnemyBars", "picEmptyBar_SP")).visible = True
                    .Controls(GetControlIndex("winEnemyBars", "picBar_HP")).visible = True
                    .Controls(GetControlIndex("winEnemyBars", "picBar_SP")).visible = True
                End If
            ElseIf myTargetType = TARGET_TYPE_NPC Then
                ' name and level
                .Controls(GetControlIndex("winEnemyBars", "lblName")).visible = True
                .Controls(GetControlIndex("winEnemyBars", "lblName")).text = Trim$(NPC(MapNpc(pIndex).Num).Name) & " - " & NPC(MapNpc(pIndex).Num).Level
                ' picture
                .Controls(GetControlIndex("winEnemyBars", "picShadow")).visible = True
                .Controls(GetControlIndex("winEnemyBars", "picChar")).visible = True
                ' store the player's index as a value for later use
                .Controls(GetControlIndex("winEnemyBars", "picChar")).Value = pIndex
                .Controls(GetControlIndex("winEnemyBars", "picChar")).image(0) = Tex_Char(NPC(MapNpc(pIndex).Num).Sprite)
                .Controls(GetControlIndex("winEnemyBars", "picChar")).image(1) = Tex_Char(NPC(MapNpc(pIndex).Num).Sprite)
                .Controls(GetControlIndex("winEnemyBars", "picChar")).image(2) = Tex_Char(NPC(MapNpc(pIndex).Num).Sprite)
                ' bars
                .Controls(GetControlIndex("winEnemyBars", "picEmptyBar_HP")).visible = True
                .Controls(GetControlIndex("winEnemyBars", "picEmptyBar_SP")).visible = True
                .Controls(GetControlIndex("winEnemyBars", "picBar_HP")).visible = True
                .Controls(GetControlIndex("winEnemyBars", "picBar_SP")).visible = True
            End If
        End If

        .Window.Height = 78

        ' update the bars
        UpdateEnemyBars
    End With
End Sub

Sub UpdateEnemyBars()
    Dim i As Long, pIndex As Long, barWidth As Long, Width As Long

    ' unload it if we're not in target
    If myTargetType = 0 Or myTarget = 0 Then
        HideWindow GetWindowIndex("winEnemyBars")
        Exit Sub
    End If

    ' max bar width
    barWidth = 173

    ' make sure we're in a party
    With Windows(GetWindowIndex("winEnemyBars"))
        ' get the pIndex from the control
        If .Controls(GetControlIndex("winEnemyBars", "picChar")).visible = True Then
            pIndex = .Controls(GetControlIndex("winEnemyBars", "picChar")).Value
            ' make sure they exist
            If pIndex > 0 Then
                If myTargetType = TARGET_TYPE_PLAYER Then
                    If IsPlaying(pIndex) Then
                        ' get playername and level atualization
                        If .Controls(GetControlIndex("winEnemyBars", "lblName")).text <> Trim$(GetPlayerName(pIndex)) & " - " & GetPlayerLevel(pIndex) Then
                            .Controls(GetControlIndex("winEnemyBars", "lblName")).text = Trim$(GetPlayerName(pIndex)) & " - " & GetPlayerLevel(pIndex)
                        End If
                        ' get their health
                        If GetPlayerVital(pIndex, HP) > 0 And GetPlayerMaxVital(pIndex, HP) > 0 Then
                            Width = ((GetPlayerVital(pIndex, Vitals.HP) / barWidth) / (GetPlayerMaxVital(pIndex, Vitals.HP) / barWidth)) * barWidth
                            .Controls(GetControlIndex("winEnemyBars", "picBar_HP")).Width = Width
                        Else
                            .Controls(GetControlIndex("winEnemyBars", "picBar_HP")).Width = 0
                        End If
                        ' get their spirit
                        If GetPlayerVital(pIndex, MP) > 0 And GetPlayerMaxVital(pIndex, MP) > 0 Then
                            Width = ((GetPlayerVital(pIndex, Vitals.MP) / barWidth) / (GetPlayerMaxVital(pIndex, Vitals.MP) / barWidth)) * barWidth
                            .Controls(GetControlIndex("winEnemyBars", "picBar_SP")).Width = Width
                        Else
                            .Controls(GetControlIndex("winEnemyBars", "picBar_SP")).Width = 0
                        End If
                    End If
                ElseIf myTargetType = TARGET_TYPE_NPC Then
                    ' get playername and level atualization

                    If .Controls(GetControlIndex("winEnemyBars", "lblName")).text <> Trim$(NPC(MapNpc(pIndex).Num).Name) & " - " & NPC(MapNpc(pIndex).Num).Level Then
                        .Controls(GetControlIndex("winEnemyBars", "lblName")).text = Trim$(NPC(MapNpc(pIndex).Num).Name) & " - " & NPC(MapNpc(pIndex).Num).Level
                    End If

                    ' get their health
                    If GetNpcVitals(pIndex, HP) > 0 And GetNpcMaxVitals(pIndex, HP) > 0 Then
                        Width = ((GetNpcVitals(pIndex, Vitals.HP) / barWidth) / (GetNpcMaxVitals(pIndex, Vitals.HP) / barWidth)) * barWidth
                        .Controls(GetControlIndex("winEnemyBars", "picBar_HP")).Width = Width
                    Else
                        .Controls(GetControlIndex("winEnemyBars", "picBar_HP")).Width = 0
                    End If
                    ' get their spirit
                    If GetNpcVitals(pIndex, MP) > 0 And GetNpcMaxVitals(pIndex, MP) > 0 Then
                        Width = ((GetNpcVitals(pIndex, Vitals.MP) / barWidth) / (GetNpcMaxVitals(pIndex, Vitals.MP) / barWidth)) * barWidth
                        .Controls(GetControlIndex("winEnemyBars", "picBar_SP")).Width = Width
                    Else
                        .Controls(GetControlIndex("winEnemyBars", "picBar_SP")).Width = 0
                    End If

                    If MapNpc(pIndex).Dead = YES Then
                        .Controls(GetControlIndex("winEnemyBars", "lblName")).text = .Controls(GetControlIndex("winEnemyBars", "lblName")).text & " (Dead)"
                        .Controls(GetControlIndex("winEnemyBars", "lblName")).textColour = BrightRed
                    Else
                        .Controls(GetControlIndex("winEnemyBars", "lblName")).textColour = White
                    End If

                End If
            End If
        End If
    End With
End Sub
