Attribute VB_Name = "modWeather"
Option Explicit

'Weather Type Constants
Public Const WEATHER_TYPE_NONE As Byte = 0
Public Const WEATHER_TYPE_RAIN As Byte = 1
Public Const WEATHER_TYPE_SNOW As Byte = 2
Public Const WEATHER_TYPE_HAIL As Byte = 3
Public Const WEATHER_TYPE_SANDSTORM As Byte = 4
Public Const WEATHER_TYPE_STORM As Byte = 5

Public Const MAX_WEATHER_PARTICLES As Long = 250

Public DrawThunder As Long

Public WeatherParticle(1 To MAX_WEATHER_PARTICLES) As WeatherParticleRec

Public WeatherImpact(1 To MAX_WEATHER_PARTICLES) As WeatherGroundRec

Private Type WeatherGroundRec
    Impact As Boolean
    tmr As Currency
    step As Byte
    X As Long
    Y As Long
End Type

Private Type WeatherParticleRec
    Type As Long
    X As Long
    Y As Long
    RangeX As Long
    RangeY As Long
    Velocity As Long
    InUse As Long
End Type

Public Sub ProcessWeather()
    Dim i As Long, X As Integer, Y As Integer
    If Map.MapData.Weather > 0 Then
        i = Rand(1, 101 - Map.MapData.WeatherIntensity)
        If i = 1 Then
            'Add a new particle
            For i = 1 To MAX_WEATHER_PARTICLES
                If WeatherParticle(i).InUse = False Then
                        WeatherParticle(i).InUse = True
                        WeatherParticle(i).Type = Map.MapData.Weather
                        WeatherParticle(i).Velocity = Rand(5, 20)
                        WeatherParticle(i).X = Rand(0, Map.MapData.MaxX * 32)   '(TileView.Left * 32) + Rand(32, frmMain.ScaleWidth)
                        WeatherParticle(i).Y = TileView.top * 32
                    Exit For
                End If
            Next
        End If
    End If

    If Map.MapData.Weather = WEATHER_TYPE_STORM Then
        i = Rand(1, 400 - Map.MapData.WeatherIntensity)
        If i = 1 Then
            'Draw Thunder
            DrawThunder = Rand(15, 22)
            Play_Sound "Thunder.wav", -1, -1
        End If
    End If

    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If (WeatherParticle(i).Y / 32) > Map.MapData.MaxY Then
                WeatherParticle(i).InUse = False
                        WeatherImpact(i).Impact = True
                        WeatherImpact(i).X = WeatherParticle(i).X
                        WeatherImpact(i).Y = WeatherParticle(i).Y
            ElseIf Map.MapData.Weather = WEATHER_TYPE_STORM Or Map.MapData.Weather = WEATHER_TYPE_RAIN Then
                If IsValidMapPoint(WeatherParticle(i).X / 32, WeatherParticle(i).Y / 32) Then
                    If Rand(1, 400 - Map.MapData.WeatherIntensity) <= 10 Then
                        WeatherParticle(i).InUse = False
                        WeatherImpact(i).Impact = True
                        WeatherImpact(i).X = WeatherParticle(i).X
                        WeatherImpact(i).Y = WeatherParticle(i).Y
                    End If
                End If
            ElseIf Map.MapData.Weather = WEATHER_TYPE_HAIL Then
                If IsValidMapPoint(WeatherParticle(i).X / 32, WeatherParticle(i).Y / 32) Then
                    If Rand(1, 400 - Map.MapData.WeatherIntensity) <= 10 Then
                        WeatherParticle(i).InUse = False
                    End If
                End If
            End If
            WeatherParticle(i).Y = WeatherParticle(i).Y + WeatherParticle(i).Velocity
        End If
        
        ' Animação ao tocar o chao
        With WeatherImpact(i)
            If .Impact Then
                If .step = 0 Then
                    If .tmr <= getTime Then
                        .step = 1
                        .tmr = getTime + 150
                    End If
                ElseIf .step = 1 Then
                    If .tmr <= getTime Then
                        .step = 2
                        .tmr = getTime + 150
                    End If
                ElseIf .step = 2 Then
                    If .tmr <= getTime Then
                        .step = 3
                        .tmr = getTime + 150
                    End If
                ElseIf .step = 3 Then
                    If .tmr <= getTime Then
                        .step = 0
                        .tmr = 0
                        .Impact = False
                    End If
                End If
            End If
        End With
    Next
End Sub

Public Sub DrawWeather()
    Dim Color As Long, i As Long, SpriteLeft As Long
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).Type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(i).Type - 1
            End If

            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(i).X), ConvertMapY(WeatherParticle(i).Y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
        End If
    Next
End Sub

Public Sub DrawWeather_Impact(ByVal WeatherID As Byte)
' Animação ao tocar o chao
    RenderTexture Tex_Weather, ConvertMapX(WeatherImpact(WeatherID).X), ConvertMapY(WeatherImpact(WeatherID).Y), WeatherImpact(WeatherID).step * 32, 32, 32, 32, 32, 32, -1
End Sub
