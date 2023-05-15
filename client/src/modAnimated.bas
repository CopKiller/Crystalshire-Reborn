Attribute VB_Name = "modAnimated"
Option Explicit

'item animated
Private StepItem As Byte
Private ItemTmr As Currency

'quest objetives animated
Private StepQuestObj As Byte
Private QuestObjTmr As Currency

Public Enum Animated
    TextureItem = 1
    TextureQuestObj
End Enum

Public Sub RenderTexture_Animated(Texture As Long, ByVal X As Long, ByVal Y As Long, ByVal sX As Single, ByVal sY As Single, _
                                  ByVal w As Long, ByVal h As Long, ByVal sW As Single, ByVal sH As Single, ByVal AnimType As Animated, _
                                  Optional ByVal Colour As Long = -1, Optional ByVal offset As Boolean = False, _
                                  Optional ByVal degrees As Single = 0, Optional ByVal Shadow As Byte = 0)

    If AnimType = TextureItem Then
        If ItemTmr <= getTime Then
            If StepItem < 4 Then
                StepItem = StepItem + 1
                ItemTmr = getTime + 200
            Else
                StepItem = 0
            End If
        End If

        If StepItem = 1 Then
            Y = Y - 2
        ElseIf StepItem = 2 Then
            Y = Y - 4
        ElseIf StepItem = 3 Then
            Y = Y - 2
        End If
    End If

    If AnimType = TextureQuestObj Then
        If QuestObjTmr <= getTime Then
            If StepQuestObj < 4 Then
                StepQuestObj = StepQuestObj + 1
                QuestObjTmr = getTime + 100
            Else
                StepQuestObj = 0
            End If
        End If

        If StepQuestObj = 1 Then
            Y = Y - 2
        ElseIf StepQuestObj = 2 Then
            Y = Y - 4
        ElseIf StepQuestObj = 3 Then
            Y = Y - 6
        End If
    End If

    RenderTexture Texture, X, Y, sX, sY, w, h, sW, sH, Colour, offset, degrees, Shadow
End Sub

Public Function VerifyWindowsIsInCur() As Boolean
    Dim i As Integer
    For i = 1 To WindowCount
        With Windows(i)
            '.Window.state = entStates.Normal
            If .Window.visible Then
                If Not .Window.clickThrough Then
                    If GlobalX >= .Window.Left And GlobalX <= .Window.Left + .Window.Width Then
                        If GlobalY >= .Window.top And GlobalY <= .Window.top + .Window.Height Then
                            VerifyWindowsIsInCur = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End With
    Next i
End Function
