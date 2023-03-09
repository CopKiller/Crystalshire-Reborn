Attribute VB_Name = "modDirectX8"
Option Explicit

' Texture wrapper
Public Tex_Anim() As Long, Tex_Char() As Long, Tex_Face() As Long, Tex_Item() As Long, Tex_Paperdoll() As Long, Tex_Resource() As Long
Attribute Tex_Char.VB_VarUserMemId = 1073741824
Attribute Tex_Face.VB_VarUserMemId = 1073741824
Attribute Tex_Item.VB_VarUserMemId = 1073741824
Attribute Tex_Paperdoll.VB_VarUserMemId = 1073741824
Attribute Tex_Resource.VB_VarUserMemId = 1073741824
Public Tex_Spellicon() As Long, Tex_Tileset() As Long, Tex_Fog() As Long, Tex_GUI() As Long, Tex_Design() As Long, Tex_Gradient() As Long, Tex_Surface() As Long
Attribute Tex_Spellicon.VB_VarUserMemId = 1073741830
Public Tex_Bars As Long, Tex_Blood As Long, Tex_Direction As Long, Tex_Misc As Long, Tex_Target As Long, Tex_Shadow As Long, Tex_Weather As Long
Attribute Tex_Bars.VB_VarUserMemId = 1073741837
Public Tex_Fader As Long, Tex_Blank As Long, Tex_Event As Long, Tex_Light As Long, Tex_LightMap As Long
Attribute Tex_Fader.VB_VarUserMemId = 1073741843
Public Tex_Captcha() As Long, Tex_Panoramas() As Long, Tex_Flags() As Long, Tex_Status() As Long, Tex_Sun() As Long

' Texture count
Public Count_Anim As Long, Count_Char As Long, Count_Face As Long, Count_GUI As Long, Count_Design As Long, Count_Gradient As Long
Attribute Count_Anim.VB_VarUserMemId = 1073741847
Attribute Count_Char.VB_VarUserMemId = 1073741847
Attribute Count_Face.VB_VarUserMemId = 1073741847
Attribute Count_GUI.VB_VarUserMemId = 1073741847
Attribute Count_Design.VB_VarUserMemId = 1073741847
Attribute Count_Gradient.VB_VarUserMemId = 1073741847
Public Count_Item As Long, Count_Paperdoll As Long, Count_Resource As Long, Count_Spellicon As Long, Count_Tileset As Long, Count_Fog As Long, Count_Surface As Long
Attribute Count_Item.VB_VarUserMemId = 1073741853
Attribute Count_Paperdoll.VB_VarUserMemId = 1073741853
Attribute Count_Resource.VB_VarUserMemId = 1073741853
Attribute Count_Spellicon.VB_VarUserMemId = 1073741853
Attribute Count_Tileset.VB_VarUserMemId = 1073741853
Attribute Count_Fog.VB_VarUserMemId = 1073741853
Attribute Count_Surface.VB_VarUserMemId = 1073741853
Public Count_Captcha As Long, Count_Panoramas As Long, Count_Flags As Long, Count_Status As Long, Count_Sun As Long
Attribute Count_Captcha.VB_VarUserMemId = 1073741860

' Menu BackGround Randomics
Public MenuBG As Byte

' Variables
Public DX8 As DirectX8
Attribute DX8.VB_VarUserMemId = 1073741861
Public D3D As Direct3D8
Attribute D3D.VB_VarUserMemId = 1073741862
Public D3DX As D3DX8
Attribute D3DX.VB_VarUserMemId = 1073741863
Public D3DDevice As Direct3DDevice8
Attribute D3DDevice.VB_VarUserMemId = 1073741864
Public DXVB As Direct3DVertexBuffer8
Attribute DXVB.VB_VarUserMemId = 1073741865
Public D3DWindow As D3DPRESENT_PARAMETERS
Attribute D3DWindow.VB_VarUserMemId = 1073741866
Public mhWnd As Long
Attribute mhWnd.VB_VarUserMemId = 1073741867
Public BackBuffer As Direct3DSurface8
Attribute BackBuffer.VB_VarUserMemId = 1073741868

Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE    'Or D3DFVF_SPECULAR

Public Type TextureStruct
    Texture As Direct3DTexture8
    data() As Byte
    w As Long
    h As Long
End Type

Public Type TextureDataStruct
    data() As Byte
End Type

Public Type Vertex
    X As Single
    Y As Single
    z As Single
    RHW As Single
    Colour As Long
    tu As Single
    tv As Single
End Type

Public mClip As RECT
Attribute mClip.VB_VarUserMemId = 1073741869
Public Box(0 To 3) As Vertex
Attribute Box.VB_VarUserMemId = 1073741870
Public mTexture() As TextureStruct
Attribute mTexture.VB_VarUserMemId = 1073741871
Public mTextures As Long
Attribute mTextures.VB_VarUserMemId = 1073741872
Public CurrentTexture As Long
Attribute CurrentTexture.VB_VarUserMemId = 1073741873

Public ScreenWidth As Long, ScreenHeight As Long
Attribute ScreenWidth.VB_VarUserMemId = 1073741874
Attribute ScreenHeight.VB_VarUserMemId = 1073741874
Public TileWidth As Long, TileHeight As Long
Attribute TileWidth.VB_VarUserMemId = 1073741876
Attribute TileHeight.VB_VarUserMemId = 1073741876
Public ScreenX As Long, ScreenY As Long
Attribute ScreenX.VB_VarUserMemId = 1073741878
Attribute ScreenY.VB_VarUserMemId = 1073741878
Public curResolution As Byte, isFullscreen As Boolean
Attribute curResolution.VB_VarUserMemId = 1073741880
Attribute isFullscreen.VB_VarUserMemId = 1073741880

Public Const DegreeToRadian As Single = 0.0174532919296
Public Const RadianToDegree As Single = 57.2958300962816

Public Sub InitDX8(ByVal hWnd As Long)
    Dim DispMode As D3DDISPLAYMODE, Width As Long, Height As Long

    mhWnd = hWnd

    Set DX8 = New DirectX8
    Set D3D = DX8.Direct3DCreate
    Set D3DX = New D3DX8

    ' set size
    GetResolutionSize curResolution, Width, Height
    ScreenWidth = Width
    ScreenHeight = Height
    TileWidth = (Width / 32) - 1
    TileHeight = (Height / 32) - 1
    ScreenX = (TileWidth) * PIC_X
    ScreenY = (TileHeight) * PIC_Y

    ' set up window
    Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    DispMode.Format = D3DFMT_A8R8G8B8

    If Options.Fullscreen = 0 Then
        isFullscreen = False
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        D3DWindow.hDeviceWindow = hWnd
        D3DWindow.BackBufferFormat = DispMode.Format
        D3DWindow.Windowed = 1
    Else
        isFullscreen = True
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
        D3DWindow.BackBufferCount = 1
        D3DWindow.BackBufferFormat = DispMode.Format
        D3DWindow.BackBufferWidth = ScreenWidth
        D3DWindow.BackBufferHeight = ScreenHeight
    End If

    Select Case Options.Render
    Case 1    ' hardware
        If LoadDirectX(D3DCREATE_HARDWARE_VERTEXPROCESSING, hWnd) <> 0 Then
            Options.Fullscreen = 0
            Options.Resolution = 0
            Options.Render = 0
            SaveOptions
            Call MsgBox("Could not initialize DirectX with hardware vertex processing.", vbCritical)
            Call DestroyGame
        End If
    Case 2    ' mixed
        If LoadDirectX(D3DCREATE_MIXED_VERTEXPROCESSING, hWnd) <> 0 Then
            Options.Fullscreen = 0
            Options.Resolution = 0
            Options.Render = 0
            SaveOptions
            Call MsgBox("Could not initialize DirectX with mixed vertex processing.", vbCritical)
            Call DestroyGame
        End If
    Case 3    ' software
        If LoadDirectX(D3DCREATE_SOFTWARE_VERTEXPROCESSING, hWnd) <> 0 Then
            Options.Fullscreen = 0
            Options.Resolution = 0
            Options.Render = 0
            SaveOptions
            Call MsgBox("Could not initialize DirectX with software vertex processing.", vbCritical)
            Call DestroyGame
        End If
    Case Else    ' auto
        If LoadDirectX(D3DCREATE_HARDWARE_VERTEXPROCESSING, hWnd) <> 0 Then
            If LoadDirectX(D3DCREATE_MIXED_VERTEXPROCESSING, hWnd) <> 0 Then
                If LoadDirectX(D3DCREATE_SOFTWARE_VERTEXPROCESSING, hWnd) <> 0 Then
                    Options.Fullscreen = 0
                    Options.Resolution = 0
                    Options.Render = 0
                    SaveOptions
                    Call MsgBox("Could not initialize DirectX.  DX8VB.dll may not be registered.", vbCritical)
                    Call DestroyGame
                End If
            End If
        End If
    End Select

    ' Render states
    Call D3DDevice.SetVertexShader(FVF)
    Call D3DDevice.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
    Call D3DDevice.SetRenderState(D3DRS_LIGHTING, False)
    Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
    Call D3DDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
    Call D3DDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, 2)
    Call D3DDevice.SetStreamSource(0, DXVB, Len(Box(0)))
End Sub

Public Function LoadDirectX(ByVal BehaviourFlags As CONST_D3DCREATEFLAGS, ByVal hWnd As Long)
    On Error GoTo ErrorInit

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, BehaviourFlags, D3DWindow)
    Exit Function

ErrorInit:
    LoadDirectX = 1
End Function

Sub DestroyDX8()
    Dim i As Long
    'For i = 1 To mTextures
    '    mTexture(i).data
    'Next
    If Not DX8 Is Nothing Then Set DX8 = Nothing
    If Not D3D Is Nothing Then Set D3D = Nothing
    If Not D3DX Is Nothing Then Set D3DX = Nothing
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
End Sub

Public Sub LoadTextures()
    Dim i As Long
    ' Arrays
    Tex_Flags = LoadTextureFiles(Count_Flags, App.path & Path_Flags)
    Tex_Panoramas = LoadTextureFiles(Count_Panoramas, App.path & Path_Panoramas)
    Tex_Captcha = LoadTextureFiles(Count_Captcha, App.path & Path_Captcha)
    Tex_Tileset = LoadTextureFiles(Count_Tileset, App.path & Path_Tileset)
    Tex_Anim = LoadTextureFiles(Count_Anim, App.path & Path_Anim)
    Tex_Char = LoadTextureFiles(Count_Char, App.path & Path_Char)
    Tex_Face = LoadTextureFiles(Count_Face, App.path & Path_Face)
    Tex_Item = LoadTextureFiles(Count_Item, App.path & Path_Item)
    Tex_Paperdoll = LoadTextureFiles(Count_Paperdoll, App.path & Path_Paperdoll)
    Tex_Resource = LoadTextureFiles(Count_Resource, App.path & Path_Resource)
    Tex_Spellicon = LoadTextureFiles(Count_Spellicon, App.path & Path_Spellicon)
    Tex_GUI = LoadTextureFiles(Count_GUI, App.path & Path_GUI)
    Tex_Design = LoadTextureFiles(Count_Design, App.path & Path_Design)
    Tex_Gradient = LoadTextureFiles(Count_Gradient, App.path & Path_Gradient)
    Tex_Surface = LoadTextureFiles(Count_Surface, App.path & Path_Surface)
    Tex_Status = LoadTextureFiles(Count_Status, App.path & Path_Status)
    Tex_Fog = LoadTextureFiles(Count_Fog, App.path & Path_Fog)
    Tex_Sun = LoadTextureFiles(Count_Sun, App.path & Path_Sun)
    ' Singles
    Tex_Bars = LoadTextureFile(App.path & Path_Misc & "bars.png")
    Tex_Blood = LoadTextureFile(App.path & Path_Misc & "blood.png")
    Tex_Misc = LoadTextureFile(App.path & Path_Misc & "misc.png")
    Tex_Direction = LoadTextureFile(App.path & Path_Misc & "direction.png")
    Tex_Target = LoadTextureFile(App.path & Path_Misc & "target.png")
    Tex_Shadow = LoadTextureFile(App.path & Path_Misc & "shadow.png")
    Tex_Fader = LoadTextureFile(App.path & Path_Misc & "fader.png")
    Tex_Blank = LoadTextureFile(App.path & Path_Misc & "blank.png")
    Tex_Event = LoadTextureFile(App.path & Path_Misc & "event.png")
    Tex_Weather = LoadTextureFile(App.path & Path_Misc & "weather.png")
    Tex_Light = LoadTextureFile(App.path & Path_Misc & "light.png")
    Tex_LightMap = LoadTextureFile(App.path & Path_Misc & "lightmap.png")
End Sub

Public Function LoadTextureFiles(ByRef Counter As Long, ByVal path As String) As Long()
    Dim Texture() As Long
    Dim i As Long

    Counter = 1

    Do While dir$(path & Counter + 1 & ".png") <> vbNullString
        Counter = Counter + 1
    Loop

    ReDim Texture(0 To Counter)

    For i = 1 To Counter
        Texture(i) = LoadTextureFile(path & i & ".png")
        If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
    Next

    LoadTextureFiles = Texture
End Function

Public Function LoadTextureFile(ByVal path As String, Optional ByVal DontReuse As Boolean) As Long
    Dim data() As Byte
    Dim f As Long

    If dir$(path) = vbNullString Then
        Call MsgBox("""" & path & """ could not be found.")
        End
    End If

    f = FreeFile
    Open path For Binary As #f
    ReDim data(0 To LOF(f) - 1)
    Get #f, , data
    Close #f

    LoadTextureFile = LoadTexture(data, DontReuse)
End Function

Public Function LoadTexture(ByRef data() As Byte, Optional ByVal DontReuse As Boolean) As Long
    Dim i As Long, LeftF As Byte

    If AryCount(data) = 0 Then
        Exit Function
    End If

    mTextures = mTextures + 1
    LoadTexture = mTextures
    ReDim Preserve mTexture(1 To mTextures) As TextureStruct
    mTexture(mTextures).w = ByteToInt(data(18), data(19))
    mTexture(mTextures).h = ByteToInt(data(22), data(23))
    mTexture(mTextures).data = data
    Set mTexture(mTextures).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, data(0), AryCount(data), mTexture(mTextures).w, mTexture(mTextures).h, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
End Function

Public Sub CheckGFX()
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then
        Do While D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST
            If PeekMessage(M, 0, 0, 0, PM_NOREMOVE) Then DoEvents
        Loop
        Call ResetGFX
    End If
End Sub

Public Sub ResetGFX()
    Dim Temp() As TextureDataStruct
    Dim i As Long, n As Long

    n = mTextures
    ReDim Temp(1 To n)
    For i = 1 To n
        Set mTexture(i).Texture = Nothing
        Temp(i).data = mTexture(i).data
    Next

    Erase mTexture
    mTextures = 0

    Call D3DDevice.Reset(D3DWindow)
    Call D3DDevice.SetVertexShader(FVF)
    Call D3DDevice.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
    Call D3DDevice.SetRenderState(D3DRS_LIGHTING, False)
    Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
    Call D3DDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
    Call D3DDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, 2)

    For i = 1 To n
        Call LoadTexture(Temp(i).data)
    Next
End Sub

Public Sub SetTexture(ByVal TextureNum As Long)
    If TextureNum > 0 Then
        Call D3DDevice.SetTexture(0, mTexture(TextureNum).Texture)
        CurrentTexture = TextureNum
    Else
        Call D3DDevice.SetTexture(0, Nothing)
        CurrentTexture = 0
    End If
End Sub

Public Sub RenderTexture(Texture As Long, ByVal X As Long, ByVal Y As Long, ByVal sX As Single, ByVal sY As Single, ByVal w As Long, ByVal h As Long, ByVal sW As Single, ByVal sH As Single, Optional ByVal Colour As Long = -1, Optional ByVal offset As Boolean = False, Optional ByVal degrees As Single = 0, Optional ByVal Shadow As Byte = 0)
    SetTexture Texture
    RenderGeom X, Y, sX, sY, w, h, sW, sH, Colour, offset, degrees, Shadow
End Sub

Public Sub RenderGeom(ByVal X As Long, ByVal Y As Long, ByVal sX As Single, ByVal sY As Single, ByVal w As Long, ByVal h As Long, ByVal sW As Single, ByVal sH As Single, Optional ByVal Colour As Long = -1, Optional ByVal offset As Boolean = False, Optional ByVal degress As Single = 0, Optional ByVal Shadow As Byte = 0)
Dim i As Long

    If CurrentTexture = 0 Then Exit Sub
    If w = 0 Then Exit Sub
    If h = 0 Then Exit Sub
    If sW = 0 Then Exit Sub
    If sH = 0 Then Exit Sub
    
    If mClip.Right <> 0 Then
        If mClip.top <> 0 Then
            If mClip.Left > X Then
                sX = sX + (mClip.Left - X) / (w / sW)
                sW = sW - (mClip.Left - X) / (w / sW)
                w = w - (mClip.Left - X)
                X = mClip.Left
            End If
            
            If mClip.top > Y Then
                sY = sY + (mClip.top - Y) / (h / sH)
                sH = sH - (mClip.top - Y) / (h / sH)
                h = h - (mClip.top - Y)
                Y = mClip.top
            End If
            
            If mClip.Right < X + w Then
                sW = sW - (X + w - mClip.Right) / (w / sW)
                w = -X + mClip.Right
            End If
            
            If mClip.bottom < Y + h Then
                sH = sH - (Y + h - mClip.bottom) / (h / sH)
                h = -Y + mClip.bottom
            End If
            
            If w <= 0 Then Exit Sub
            If h <= 0 Then Exit Sub
            If sW <= 0 Then Exit Sub
            If sH <= 0 Then Exit Sub
        End If
    End If
    
    Call GeomCalc(CurrentTexture, X, Y, w, h, sX, sY, sW, sH, Colour, degress, Shadow)
    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
End Sub

Public Sub GeomCalc(ByVal TextureNum As Long, ByVal X As Single, ByVal Y As Single, ByVal w As Integer, ByVal h As Integer, ByVal sX As Single, ByVal sY As Single, ByVal sW As Single, ByVal sH As Single, ByVal Color As Long, Optional ByVal degrees As Single = 0, Optional ByVal Shadow As Byte = 0)
    Dim RadAngle As Single ' The angle in Radians
    Dim CenterX As Single, CenterY As Single
    Dim NewX As Single, NewY As Single
    Dim SinRad As Single, CosRad As Single, i As Long

    sW = (sW + sX) / mTexture(TextureNum).w + 0.000003
    sH = (sH + sY) / mTexture(TextureNum).h + 0.000003
    sX = sX / mTexture(TextureNum).w + 0.000003
    sY = sY / mTexture(TextureNum).h + 0.000003
    
    GeomSetBox X, Y, w, h, Color, sX, sY, sW, sH

    ' Check if a rotation is required
    If degrees <> 0 And degrees <> 360 Then
        ' Converts the angle to rotate by into radians
        RadAngle = degrees * DegreeToRadian
        ' Set the CenterX and CenterY values
        CenterX = X + (w * 0.5)
        CenterY = Y + (h * 0.5)
        ' Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)
        ' Loops through the passed vertex buffer
        For i = 0 To 3
            ' Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (Box(i).X - CenterX) * CosRad - (Box(i).Y - CenterY) * SinRad
            NewY = CenterY + (Box(i).Y - CenterY) * CosRad + (Box(i).X - CenterX) * SinRad
            ' Applies the new co-ordinates to the buffer
            Box(i).X = NewX
            Box(i).Y = NewY
        Next
    End If
    
    If Shadow > 0 Then
     'Efeito VbGore Sombra
     '* 0.3
        Box(0).X = X + w
        Box(0).Y = Y + h
        Box(1).X = Box(0).X - w
        Box(1).Y = Box(0).Y
    End If
End Sub

'Private Function CalculateShadowPosition() As Integer
'    If GameHours >= 7 And GameHours <= 12 Then
'        CalculateShadowPosition = (-11 * 10)
'    Else
'
'    End If
'End Function

Private Sub GeomSetBox(ByVal X As Single, ByVal Y As Single, ByVal w As Integer, ByVal h As Integer, ByVal Color As Long, ByVal sX As Single, ByVal sY As Single, ByVal sW As Single, ByVal sH As Single)
    Box(0) = MakeVertex(X, Y, 0, 1, Color, 1, sX, sY)
    Box(1) = MakeVertex(X + w, Y, 0, 1, Color, 0, sW, sY)
    Box(2) = MakeVertex(X, Y + h, 0, 1, Color, 0, sX, sH)
    Box(3) = MakeVertex(X + w, Y + h, 0, 1, Color, 0, sW, sH)
End Sub

Private Function MakeVertex(X As Single, Y As Single, z As Single, RHW As Single, Color As Long, Specular As Long, tu As Single, tv As Single) As Vertex
    MakeVertex.X = X
    MakeVertex.Y = Y
    MakeVertex.z = z
    MakeVertex.RHW = RHW
    MakeVertex.Colour = Color
    MakeVertex.tu = tu
    MakeVertex.tv = tv
End Function

' GDI rendering
Public Sub GDIRenderAnimation()
    Dim i As Long, Animationnum As Long, ShouldRender As Boolean, Width As Long, Height As Long, looptime As Long, FrameCount As Long
    Dim sX As Long, sY As Long, sRECT As RECT
    sRECT.top = 0
    sRECT.bottom = 192
    sRECT.Left = 0
    sRECT.Right = 192

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).Value

        If Animationnum <= 0 Or Animationnum > Count_Anim Then
            ' don't render lol
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)

            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            ShouldRender = False

            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= getTime Then

                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If

                AnimEditorTimer(i) = getTime
                ShouldRender = True
            End If

            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then
                    ' total width divided by frame count
                    Width = 192
                    Height = 192
                    sY = (Height * ((AnimEditorFrame(i) - 1) \ AnimColumns))
                    sX = (Width * (((AnimEditorFrame(i) - 1) Mod AnimColumns)))
                    ' Start Rendering
                    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call D3DDevice.BeginScene
                    'EngineRenderRectangle Tex_Anim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    RenderTexture Tex_Anim(Animationnum), 0, 0, sX, sY, Width, Height, Width, Height
                    ' Finish Rendering
                    Call D3DDevice.EndScene
                    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(i).hWnd, ByVal 0)
                End If
            End If
        End If

    Next

End Sub

' GDI rendering
Public Sub GDIRenderResource(ByRef picBox As PictureBox, ByVal Sprite As Long)
    Dim Width As Long, Height As Long, sRECT As RECT
    
    ' exit out if doesn't exist
    If Sprite <= 0 Or Sprite > Count_Resource Then
        picBox.Cls
        Exit Sub
    End If
    
    Height = mTexture(Tex_Resource(Sprite)).h
    Width = mTexture(Tex_Resource(Sprite)).w
    
    If Height = 0 Or Width = 0 Then
        Height = 1
        Width = 1
    End If
    
    sRECT.top = 0
    sRECT.bottom = Height
    sRECT.Left = 0
    sRECT.Right = Width
    
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    
    RenderTexture Tex_Resource(Sprite), 0, 0, 0, 0, Width, Height, Width, Height
    ' Finish Rendering
    
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderChar(ByRef picBox As PictureBox, ByVal Sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If Sprite <= 0 Or Sprite > Count_Char Then Exit Sub
    Height = 32
    Width = 32
    sRECT.top = 0
    sRECT.bottom = sRECT.top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    RenderTexture Tex_Char(Sprite), 0, 0, 0, 0, Width, Height, Width, Height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderLight(ByRef picBox As PictureBox)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    Height = (frmEditor_Map.scrlTamanho.Value * 32)
    Width = (frmEditor_Map.scrlTamanho.Value * 32)
    sRECT.top = 0
    sRECT.bottom = sRECT.top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    RenderTexture Tex_Light, 0, 0, 0, 0, Width, Height, 128, 128, D3DColorARGB(frmEditor_Map.scrlA, frmEditor_Map.scrlR, frmEditor_Map.scrlG, frmEditor_Map.scrlB)
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderFace(ByRef picBox As PictureBox, ByVal Sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If Sprite <= 0 Or Sprite > Count_Face Then Exit Sub
    Height = mTexture(Tex_Face(Sprite)).h
    Width = mTexture(Tex_Face(Sprite)).w

    If Height = 0 Or Width = 0 Then
        Height = 1
        Width = 1
    End If

    sRECT.top = 0
    sRECT.bottom = sRECT.top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle Tex_Face(sprite), 0, 0, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Face(Sprite), 0, 0, 0, 0, Width, Height, Width, Height
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Sub GDIRenderEventGraphic()
    Dim Height As Long, Width As Long, GraphicType As Long, graphicNum As Long, sX As Long, sY As Long, texNum As Long
    Dim sRECT As RECT, Graphic As Long

    If Not frmEditor_Events.visible Then Exit Sub
    If curPageNum = 0 Then Exit Sub

    GraphicType = tmpEvent.EventPage(curPageNum).GraphicType
    Graphic = tmpEvent.EventPage(curPageNum).Graphic
    sX = tmpEvent.EventPage(curPageNum).GraphicX
    sY = tmpEvent.EventPage(curPageNum).GraphicY

    If GraphicType = 0 Then Exit Sub
    If Graphic = 0 Then Exit Sub

    Height = 32
    Width = 32

    Select Case GraphicType
    Case 0    ' nothing
        texNum = 0
    Case 1    ' Character
        If Graphic <= Count_Char Then texNum = Tex_Char(Graphic) Else texNum = 0
    Case 2    ' Tileset
        If Graphic <= Count_Tileset Then texNum = Tex_Tileset(Graphic) Else texNum = 0
    End Select

    If texNum = 0 Then
        frmEditor_Events.picGraphic.Cls
        Exit Sub
    End If

    sRECT.top = 0
    sRECT.bottom = sRECT.top + frmEditor_Events.picGraphic.ScaleHeight
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + frmEditor_Events.picGraphic.ScaleWidth

    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, DX8Colour(White, 255), 1#, 0)
    Call D3DDevice.BeginScene

    RenderTexture texNum, (frmEditor_Events.picGraphic.ScaleWidth / 2) - 16, (frmEditor_Events.picGraphic.ScaleHeight / 2) - 16, sX * 32, sY * 32, Width, Height, Width, Height

    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Events.picGraphic.hWnd, ByVal 0)
End Sub

Sub GDIRenderEventGraphicSel()
    Dim Height As Long, Width As Long, GraphicType As Long, graphicNum As Long, sX As Long, sY As Long, texNum As Long
    Dim sRECT As RECT, Graphic As Long

    If Not frmEditor_Events.visible Then Exit Sub
    If Not frmEditor_Events.fraGraphic.visible Then Exit Sub
    If curPageNum = 0 Then Exit Sub

    GraphicType = tmpEvent.EventPage(curPageNum).GraphicType
    Graphic = tmpEvent.EventPage(curPageNum).Graphic

    If GraphicType = 0 Then Exit Sub
    If Graphic = 0 Then Exit Sub

    Select Case GraphicType
    Case 0    ' nothing
        texNum = 0
    Case 1    ' Character
        If Graphic <= Count_Char Then texNum = Tex_Char(Graphic) Else texNum = 0
    Case 2    ' Tileset
        If Graphic <= Count_Tileset Then texNum = Tex_Tileset(Graphic) Else texNum = 0
    End Select

    If texNum = 0 Then
        frmEditor_Events.picGraphicSel.Cls
        Exit Sub
    End If

    Width = mTexture(texNum).w
    Height = mTexture(texNum).h

    sRECT.top = 0
    sRECT.bottom = sRECT.top + frmEditor_Events.picGraphicSel.ScaleHeight
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + frmEditor_Events.picGraphicSel.ScaleWidth

    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, DX8Colour(White, 255), 1#, 0)
    Call D3DDevice.BeginScene

    RenderTexture texNum, 0, 0, 0, 0, Width, Height, Width, Height
    RenderDesign DesignTypes.desTileBox, GraphicSelX * 32, GraphicSelY * 32, 32, 32

    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Events.picGraphicSel.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderTileset()
    Dim Height As Long, Width As Long, tileSet As Byte, sRECT As RECT
    ' find tileset number
    tileSet = frmEditor_Map.scrlTileSet.Value

    ' exit out if doesn't exist
    If tileSet <= 0 Or tileSet > Count_Tileset Then Exit Sub
    Height = mTexture(Tex_Tileset(tileSet)).h
    Width = mTexture(Tex_Tileset(tileSet)).w

    If Height = 0 Or Width = 0 Then
        Height = 1
        Width = 1
    End If

    frmEditor_Map.picBackSelect.Width = Width
    frmEditor_Map.picBackSelect.Height = Height
    sRECT.top = 0
    sRECT.bottom = Height
    sRECT.Left = 0
    sRECT.Right = Width

    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.Value > 0 Then

        Select Case frmEditor_Map.scrlAutotile.Value

        Case 1    ' autotile
            shpSelectedWidth = 64
            shpSelectedHeight = 96

        Case 2    ' fake autotile
            shpSelectedWidth = 32
            shpSelectedHeight = 32

        Case 3    ' animated
            shpSelectedWidth = 192
            shpSelectedHeight = 96

        Case 4    ' cliff
            shpSelectedWidth = 64
            shpSelectedHeight = 64

        Case 5    ' waterfall
            shpSelectedWidth = 64
            shpSelectedHeight = 96
        End Select

    End If

    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, DX8Colour(White, 255), 1#, 0)
    Call D3DDevice.BeginScene

    'EngineRenderRectangle Tex_Tileset(Tileset), 0, 0, 0, 0, width, height, width, height, width, height
    If Tex_Tileset(tileSet) <= 0 Then Exit Sub
    RenderTexture Tex_Tileset(tileSet), 0, 0, 0, 0, Width, Height, Width, Height
    ' draw selection boxes
    RenderDesign DesignTypes.desTileBox, shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, frmEditor_Map.picBackSelect.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderItem(ByRef picBox As PictureBox, ByVal Sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If Sprite <= 0 Or Sprite > Count_Item Then Exit Sub
    Height = mTexture(Tex_Item(Sprite)).h
    Width = mTexture(Tex_Item(Sprite)).w
    sRECT.top = 0
    sRECT.bottom = 32
    sRECT.Left = 0
    sRECT.Right = 32
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle Tex_Item(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(Sprite), 0, 0, 0, 0, 32, 32, 32, 32
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderSpell(ByRef picBox As PictureBox, ByVal Sprite As Long)
    Dim Height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If Sprite <= 0 Or Sprite > Count_Spellicon Then Exit Sub
    Height = mTexture(Tex_Spellicon(Sprite)).h
    Width = mTexture(Tex_Spellicon(Sprite)).w

    If Height = 0 Or Width = 0 Then
        Height = 1
        Width = 1
    End If

    sRECT.top = 0
    sRECT.bottom = Height
    sRECT.Left = 0
    sRECT.Right = Width
    ' Start Rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
    'EngineRenderRectangle Tex_Spellicon(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Spellicon(Sprite), 0, 0, 0, 0, 32, 32, 32, 32
    ' Finish Rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

' Directional blocking
Public Sub DrawDirection(ByVal X As Long, ByVal Y As Long)
    Dim i As Long, top As Long, Left As Long
    ' render grid
    top = 24
    Left = 0
    'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), left, top, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Direction, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), Left, top, 32, 32, 32, 32

    ' render dir blobs
    For i = 1 To 4
        Left = (i - 1) * 8

        ' find out whether render blocked or not
        If Not isDirBlocked(Map.TileData.Tile(X, Y).DirBlock, CByte(i)) Then
            top = 8
        Else
            top = 16
        End If

        'render!
        'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), left, top, 8, 8, 8, 8, 8, 8
        RenderTexture Tex_Direction, ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), Left, top, 8, 8, 8, 8
    Next

End Sub

Public Sub DrawFade()
    RenderTexture Tex_Blank, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, DX8Colour(White, fadeAlpha)
End Sub

Private Sub DrawFog()
    Dim fogNum As Long, Color As Long, X As Long, Y As Long
    Dim fogWidth As Integer, fogHeight As Integer

    fogNum = Map.MapData.Fog
    If fogNum <= 0 Or fogNum > Count_Fog Then Exit Sub

    Color = D3DColorRGBA(255, 255, 255, 255 - Map.MapData.FogOpacity)
    fogWidth = mTexture(Tex_Fog(fogNum)).w
    fogHeight = mTexture(Tex_Fog(fogNum)).h
    
    ' reset the position
    If fogOffsetX < (256 * -1) Then fogOffsetX = 0
    If fogOffsetY < (256 * -1) Then fogOffsetY = 0
    
    For X = 0 To (ScreenWidth / fogWidth) + 1
        For Y = 0 To ((ScreenHeight) / fogHeight) + 1
            RenderTexture Tex_Fog(fogNum), (X * fogWidth) + fogOffsetX, (Y * fogHeight) + fogOffsetY, 0, 0, fogWidth, fogHeight, fogWidth, fogHeight, Color
        Next Y
    Next X

    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
End Sub

Private Sub DrawTint()
    Dim Color As Long

    ' Tem Visibilidade?
    If Map.MapData.Alpha = 0 Then Exit Sub

    ' Tem alguma cor pra renderizar?
    If Map.MapData.Red > 0 Or Map.MapData.Green > 0 Or Map.MapData.Blue > 0 Or Map.MapData.Alpha > 0 Then
        Color = D3DColorRGBA(Map.MapData.Red, Map.MapData.Green, Map.MapData.Blue, Map.MapData.Alpha)

        RenderTexture Tex_Blank, ConvertMapX(0), ConvertMapY(0), 0, 0, ((Map.MapData.MaxX + 1) * PIC_X), ((Map.MapData.MaxY + 1) * PIC_Y), 32, 32, Color

        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Private Sub DrawSun()
    Dim Width As Integer, Height As Integer

    If Not IsDay Then Exit Sub
    If Map.MapData.Sun = 0 Then Exit Sub
    
        Width = mTexture(Tex_Sun(Map.MapData.Sun)).w
        Height = mTexture(Tex_Sun(Map.MapData.Sun)).h

        RenderTexture Tex_Sun(Map.MapData.Sun), 0, 0, 0, 0, ScreenWidth, ScreenHeight, Width, Height, D3DColorRGBA(255, 255, 255, 100)

        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
End Sub

Public Sub DrawAutoTile(ByVal layernum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim YOffSet As Long, XOffSet As Long

    ' calculate the offset
    Select Case Map.TileData.Tile(X, Y).Autotile(layernum)

    Case AUTOTILE_WATERFALL
        YOffSet = (waterfallFrame - 1) * 32

    Case AUTOTILE_ANIM
        XOffSet = autoTileFrame * 64

    Case AUTOTILE_CLIFF
        YOffSet = -32
    End Select

    ' Draw the quarter
    RenderTexture Tex_Tileset(Map.TileData.Tile(X, Y).Layer(layernum).tileSet), destX, destY, Autotile(X, Y).Layer(layernum).srcX(quarterNum) + XOffSet, Autotile(X, Y).Layer(layernum).srcY(quarterNum) + YOffSet, 16, 16, 16, 16
End Sub

Sub DrawTileSelection()
    If frmEditor_Map.optEvents.Value Then
        RenderDesign DesignTypes.desTileBox, ConvertMapX(selTileX * PIC_X), ConvertMapY(selTileY * PIC_Y), 32, 32
    Else
        If frmEditor_Map.scrlAutotile > 0 Then
            RenderDesign DesignTypes.desTileBox, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), 32, 32
        Else
            RenderDesign DesignTypes.desTileBox, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), shpSelectedWidth, shpSelectedHeight
        End If
    End If
End Sub

' Rendering Procedures
Public Sub DrawMapTile(ByVal X As Long, ByVal Y As Long)
    Dim i As Long, tileSet As Long, sX As Long, sY As Long

    With Map.TileData.Tile(X, Y)
        ' draw the map
        For i = MapLayer.Ground To MapLayer.Mask2
            ' skip tile if tileset isn't set
            If Autotile(X, Y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).tileSet), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32
            ElseIf Autotile(X, Y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            ElseIf Autotile(X, Y).Layer(i).renderState = RENDER_STATE_APPEAR Then
                ' check if it's fading
                If TempTile(X, Y).fadeAlpha(i) > 0 Then
                    ' render it
                    tileSet = Map.TileData.Tile(X, Y).Layer(i).tileSet
                    sX = Map.TileData.Tile(X, Y).Layer(i).X
                    sY = Map.TileData.Tile(X, Y).Layer(i).Y
                    RenderTexture Tex_Tileset(tileSet), ConvertMapX(X * 32), ConvertMapY(Y * 32), sX * 32, sY * 32, 32, 32, 32, 32, DX8Colour(White, TempTile(X, Y).fadeAlpha(i))
                End If
            End If
        Next
    End With
End Sub

Public Sub DrawMapFringeTile(ByVal X As Long, ByVal Y As Long)
    Dim i As Long

    With Map.TileData.Tile(X, Y)
        ' draw the map
        For i = MapLayer.Fringe To MapLayer.Fringe2

            ' skip tile if tileset isn't set
            If Autotile(X, Y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).tileSet), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32
            ElseIf Autotile(X, Y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            End If
        Next
    End With
End Sub

Public Sub DrawHotbar()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, i As Long, T As Long, SS As String, CooldownTime As Long

    xO = Windows(GetWindowIndex("winHotbar")).Window.Left
    yO = Windows(GetWindowIndex("winHotbar")).Window.top

    ' render start + end wood
    RenderTexture Tex_GUI(31), xO - 1, yO + 3, 0, 0, 11, 26, 11, 26
    RenderTexture Tex_GUI(31), xO + 407, yO + 3, 0, 0, 11, 26, 11, 26

    For i = 1 To MAX_HOTBAR
        xO = Windows(GetWindowIndex("winHotbar")).Window.Left + HotbarLeft + ((i - 1) * HotbarOffsetX)
        yO = Windows(GetWindowIndex("winHotbar")).Window.top + HotbarTop
        Width = 36
        Height = 36
        ' don't render last one
        If i <> 10 Then
            ' render wood
            RenderTexture Tex_GUI(32), xO + 30, yO + 3, 0, 0, 13, 26, 13, 26
        End If
        ' render box
        RenderTexture Tex_GUI(30), xO - 2, yO - 2, 0, 0, Width, Height, Width, Height
        ' render icon
        If Not (DragBox.Origin = origin_Hotbar And DragBox.Slot = i) Then
            Select Case Hotbar(i).sType
            Case 1    ' inventory
                If Len(Item(Hotbar(i).Slot).Name) > 0 And Item(Hotbar(i).Slot).Pic > 0 Then
                    RenderTexture Tex_Item(Item(Hotbar(i).Slot).Pic), xO, yO, 0, 0, 32, 32, 32, 32
                End If
            Case 2    ' spell
                If Len(Spell(Hotbar(i).Slot).Name) > 0 And Spell(Hotbar(i).Slot).Icon > 0 Then
                    RenderTexture Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), xO, yO, 0, 0, 32, 32, 32, 32
                    For T = 1 To MAX_PLAYER_SPELLS
                        If PlayerSpells(T).Spell > 0 Then
                            If PlayerSpells(T).Spell = Hotbar(i).Slot And SpellCD(T) > 0 Then
                                CooldownTime = SpellCD(T)
                                SS = SecondsToHMS(CooldownTime)
                                RenderTexture Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), xO, yO, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                                RenderText font(Fonts.georgia_16), SS, xO + 15 - (TextWidth(font(Fonts.georgia_16), SS) / 2), yO - 6, BrightRed
                            End If
                        End If
                    Next
                End If
            End Select
        End If
        ' draw the numbers
        SS = KeycodeChar(Options.Hotbar(i))
        RenderText font(Fonts.rockwellDec_15), SS, xO + 4, yO + 19, White
    Next
End Sub

Public Sub RenderAppearTileFade()
    Dim X As Long, Y As Long, tileSet As Long, sX As Long, sY As Long, layernum As Long

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            For layernum = MapLayer.Ground To MapLayer.Mask
                ' check if it's fading
                If TempTile(X, Y).fadeAlpha(layernum) > 0 Then
                    ' render it
                    tileSet = Map.TileData.Tile(X, Y).Layer(layernum).tileSet
                    sX = Map.TileData.Tile(X, Y).Layer(layernum).X
                    sY = Map.TileData.Tile(X, Y).Layer(layernum).Y
                    RenderTexture Tex_Tileset(tileSet), ConvertMapX(X * 32), ConvertMapY(Y * 32), sX * 32, sY * 32, 32, 32, 32, 32, DX8Colour(White, TempTile(X, Y).fadeAlpha(layernum))
                End If
            Next
        Next
    Next
End Sub

Public Sub DrawSkills()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, i As Long, Y As Long, spellnum As Long, spellPic As Long, X As Long, top As Long, Left As Long, SS As String, CooldownTime As Long

    xO = Windows(GetWindowIndex("winSkills")).Window.Left
    yO = Windows(GetWindowIndex("winSkills")).Window.top

    Width = Windows(GetWindowIndex("winSkills")).Window.Width
    Height = Windows(GetWindowIndex("winSkills")).Window.Height

    ' render green
    RenderTexture Tex_GUI(34), xO + 4, yO + 23, 0, 0, Width - 8, Height - 27, 4, 4

    Width = 76
    Height = 76

    Y = yO + 23
    ' render grid - row
    For i = 1 To 4
        If i = 4 Then Height = 42
        RenderTexture Tex_GUI(35), xO + 4, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next

    ' actually draw the icons
    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i).Spell
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            ' not dragging?
            If Not (DragBox.Origin = origin_Spells And DragBox.Slot = i) Then
                spellPic = Spell(spellnum).Icon

                If spellPic > 0 And spellPic <= Count_Spellicon Then
                    top = yO + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = xO + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))

                    If SpellCD(i) > 0 Then
                        CooldownTime = SpellCD(i)
                        SS = SecondsToHMS(CooldownTime)
                        RenderTexture Tex_Spellicon(spellPic), Left, top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                        RenderText font(Fonts.georgia_16), SS, Left + 15 - (TextWidth(font(Fonts.georgia_16), SS) / 2), top - 6, BrightRed
                    Else
                        RenderTexture Tex_Spellicon(spellPic), Left, top, 0, 0, 32, 32, 32, 32
                    End If
                End If
            End If
        End If
    Next
End Sub

Public Sub RenderMapName()
    Dim zonetype As String, Colour As Long

    If Map.MapData.Moral = 0 Then
        zonetype = "PK Zone"
        Colour = Red
    ElseIf Map.MapData.Moral = 1 Then
        zonetype = "Safe Zone"
        Colour = White
    ElseIf Map.MapData.Moral = 2 Then
        zonetype = "Boss Chamber"
        Colour = Grey
    End If

    RenderText font(Fonts.rockwellDec_10), Trim$(Map.MapData.Name) & " - " & zonetype, ScreenWidth - 15 - TextWidth(font(Fonts.rockwellDec_10), Trim$(Map.MapData.Name) & " - " & zonetype), 45, Colour, 255
End Sub

Public Sub DrawShopBackground()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, i As Long, Y As Long

    xO = Windows(GetWindowIndex("winShop")).Window.Left
    yO = Windows(GetWindowIndex("winShop")).Window.top
    Width = Windows(GetWindowIndex("winShop")).Window.Width
    Height = Windows(GetWindowIndex("winShop")).Window.Height

    ' render green
    RenderTexture Tex_GUI(34), xO + 4, yO + 23, 0, 0, Width - 8, Height - 27, 4, 4

    Width = 76
    Height = 76

    Y = yO + 23
    ' render grid - row
    For i = 1 To 3
        If i = 3 Then Height = 42
        RenderTexture Tex_GUI(35), xO + 4, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 232, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
    ' render bottom wood
    RenderTexture Tex_GUI(1), xO + 4, Y - 34, 0, 0, 270, 72, 270, 72
End Sub

Public Sub DrawShop()
    Dim xO As Long, yO As Long, ItemPic As Long, itemNum As Long, Amount As Long, i As Long, top As Long, Left As Long, Y As Long, X As Long, Colour As Long
    Dim rec As RECT

    If InShop = 0 Then Exit Sub

    xO = Windows(GetWindowIndex("winShop")).Window.Left
    yO = Windows(GetWindowIndex("winShop")).Window.top

    If Not shopIsSelling Then
        ' render the shop items
        For i = 1 To MAX_TRADES
            itemNum = Shop(InShop).TradeItem(i).Item

            ' draw early
            top = yO + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            Left = xO + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            ' draw selected square
            If shopSelectedSlot = i Then RenderTexture Tex_GUI(61), Left, top, 0, 0, 32, 32, 32, 32

            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= Count_Item Then
                    ' draw item
                    If Options.ItemAnimation = YES Then
                        rec.top = 0
                        rec.Left = Shop(InShop).TradeItem(i).Frame * PIC_X
                    End If
                    RenderTexture Tex_Item(ItemPic), Left, top, rec.Left, rec.top, 32, 32, 32, 32
                End If
            End If
        Next
    Else
        ' render the shop items
        For i = 1 To MAX_TRADES
            itemNum = GetPlayerInvItemNum(MyIndex, i)

            ' draw early
            top = yO + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            Left = xO + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            ' draw selected square
            If shopSelectedSlot = i Then RenderTexture Tex_GUI(61), Left, top, 0, 0, 32, 32, 32, 32

            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= Count_Item Then

                    ' draw item
                    RenderTexture Tex_Item(ItemPic), Left, top, 0, 0, 32, 32, 32, 32

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = top + 21
                        X = Left + 1
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, i))

                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Colour = BrightGreen
                        End If

                        RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), X, Y, Colour
                    End If
                End If
            End If
        Next
    End If
End Sub

Sub DrawTrade()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, i As Long, Y As Long, X As Long

    xO = Windows(GetWindowIndex("winTrade")).Window.Left
    yO = Windows(GetWindowIndex("winTrade")).Window.top
    Width = Windows(GetWindowIndex("winTrade")).Window.Width
    Height = Windows(GetWindowIndex("winTrade")).Window.Height

    ' render green
    RenderTexture Tex_GUI(34), xO + 4, yO + 23, 0, 0, Width - 8, Height - 27, 4, 4

    ' top wood
    RenderTexture Tex_GUI(1), xO + 4, yO + 23, 100, 100, Width - 8, 18, Width - 8, 18
    ' left wood
    RenderTexture Tex_GUI(1), xO + 4, yO + 41, 350, 0, 5, Height - 45, 5, Height - 45
    ' right wood
    RenderTexture Tex_GUI(1), xO + Width - 9, yO + 41, 350, 0, 5, Height - 45, 5, Height - 45
    ' centre wood
    RenderTexture Tex_GUI(1), xO + 203, yO + 41, 350, 0, 6, Height - 45, 6, Height - 45
    ' bottom wood
    RenderTexture Tex_GUI(1), xO + 4, yO + 307, 100, 100, Width - 8, 75, Width - 8, 75

    ' left
    Width = 76
    Height = 76
    Y = yO + 41
    For i = 1 To 4
        If i = 4 Then Height = 38
        RenderTexture Tex_GUI(35), xO + 4 + 5, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80 + 5, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156 + 5, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next

    ' right
    Width = 76
    Height = 76
    Y = yO + 41
    For i = 1 To 4
        If i = 4 Then Height = 38
        RenderTexture Tex_GUI(35), xO + 4 + 205, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80 + 205, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156 + 205, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
End Sub

Sub DrawYourTrade()
    Dim i As Long, itemNum As Long, ItemPic As Long, top As Long, Left As Long, Colour As Long, Amount As String, X As Long, Y As Long
    Dim xO As Long, yO As Long, rec As RECT

    xO = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).Left
    yO = Windows(GetWindowIndex("winTrade")).Window.top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picYour")).top

    ' your items
    For i = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                top = yO + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                Left = xO + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))

                ' draw icon
                If Options.ItemAnimation = YES Then
                    rec.top = 0
                    rec.Left = TradeYourOffer(i).Frame * PIC_X
                End If
                RenderTexture Tex_Item(ItemPic), Left, top, rec.Left, rec.top, 32, 32, 32, 32

                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).Value > 1 Then
                    Y = top + 21
                    X = Left + 1
                    Amount = CStr(TradeYourOffer(i).Value)

                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If

                    RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), X, Y, Colour
                End If
            End If
        End If
    Next
End Sub

Sub DrawTheirTrade()
    Dim i As Long, itemNum As Long, ItemPic As Long, top As Long, Left As Long, Colour As Long, Amount As String, X As Long, Y As Long
    Dim xO As Long, yO As Long, rec As RECT

    xO = Windows(GetWindowIndex("winTrade")).Window.Left + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).Left
    yO = Windows(GetWindowIndex("winTrade")).Window.top + Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "picTheir")).top

    ' their items
    For i = 1 To MAX_INV
        itemNum = TradeTheirOffer(i).num
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                top = yO + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                Left = xO + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))

                ' draw icon
                If Options.ItemAnimation = YES Then
                    rec.top = 0
                    rec.Left = TradeTheirOffer(i).Frame * PIC_X
                End If
                RenderTexture Tex_Item(ItemPic), Left, top, rec.Left, rec.top, 32, 32, 32, 32

                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).Value > 1 Then
                    Y = top + 21
                    X = Left + 1
                    Amount = CStr(TradeTheirOffer(i).Value)

                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If

                    RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), X, Y, Colour
                End If
            End If
        End If
    Next
End Sub

Public Sub DrawInventory()
    Dim xO As Long, yO As Long, Width As Long, Height As Long, i As Long, Y As Long, itemNum As Long, ItemPic As Long, X As Long, top As Long, Left As Long, Amount As String
    Dim Colour As Long, skipItem As Boolean, amountModifier As Long, tmpItem As Long, rec As RECT

    xO = Windows(GetWindowIndex("winInventory")).Window.Left
    yO = Windows(GetWindowIndex("winInventory")).Window.top
    Width = Windows(GetWindowIndex("winInventory")).Window.Width
    Height = Windows(GetWindowIndex("winInventory")).Window.Height

    ' render green
    RenderTexture Tex_GUI(34), xO + 4, yO + 23, 0, 0, Width - 8, Height - 27, 4, 4

    Width = 76
    Height = 76

    Y = yO + 23
    ' render grid - row
    For i = 1 To 4
        If i = 4 Then Height = 38
        RenderTexture Tex_GUI(35), xO + 4, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 80, Y, 0, 0, Width, Height, Width, Height
        RenderTexture Tex_GUI(35), xO + 156, Y, 0, 0, 42, Height, 42, Height
        Y = Y + 76
    Next
    ' render bottom wood
    RenderTexture Tex_GUI(1), xO + 4, yO + 289, 100, 100, 194, 26, 194, 26

    ' actually draw the icons
    For i = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, i)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ' not dragging?
            If Not (DragBox.Origin = origin_Inventory And DragBox.Slot = i) Then
                ItemPic = Item(itemNum).Pic

                ' exit out if we're offering item in a trade.
                amountModifier = 0
                If InTrade > 0 Then
                    For X = 1 To MAX_INV
                        tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).num)
                        If TradeYourOffer(X).num = i Then
                            ' check if currency
                            If Not Item(tmpItem).Stackable > 0 Then
                                ' normal item, exit out
                                skipItem = True
                            Else
                                ' if amount = all currency, remove from inventory
                                If TradeYourOffer(X).Value = GetPlayerInvItemValue(MyIndex, i) Then
                                    skipItem = True
                                Else
                                    ' not all, change modifier to show change in currency count
                                    amountModifier = TradeYourOffer(X).Value
                                End If
                            End If
                        End If
                    Next
                End If

                If Not skipItem Then
                    If ItemPic > 0 And ItemPic <= Count_Item Then
                        top = yO + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        Left = xO + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))

                        ' draw icon
                        If Options.ItemAnimation = YES Then
                            rec.top = 0
                            rec.Left = PlayerInv(i).Frame * PIC_X
                        End If
                        RenderTexture Tex_Item(ItemPic), Left, top, rec.Left, rec.top, 32, 32, 32, 32

                        ' If item is a stack - draw the amount you have
                        If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                            Y = top + 21
                            X = Left + 1
                            Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier

                            ' Draw currency but with k, m, b etc. using a convertion function
                            If CLng(Amount) < 1000000 Then
                                Colour = White
                            ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                                Colour = Yellow
                            ElseIf CLng(Amount) > 10000000 Then
                                Colour = BrightGreen
                            End If

                            RenderText font(Fonts.verdana_12), ConvertCurrency(Amount), X, Y, Colour
                        End If
                    End If
                End If
                ' reset
                skipItem = False
            End If
        End If
    Next
End Sub

Public Sub DrawChatBubble(ByVal Index As Long)
    Dim theArray() As String, X As Long, Y As Long, i As Long, MaxWidth As Long, X2 As Long, Y2 As Long, Colour As Long, tmpNum As Long

    With chatBubble(Index)
        ' exit out early
        If .Target = 0 Then Exit Sub
        ' calculate position
        Select Case .TargetType
        Case TARGET_TYPE_PLAYER
            ' it's a player
            If Not GetPlayerMap(.Target) = GetPlayerMap(MyIndex) Then Exit Sub
            ' change the colour depending on access
            Colour = DarkBrown
            ' it's on our map - get co-ords
            X = ConvertMapX((Player(.Target).X * 32) + Player(.Target).XOffSet) + 16
            Y = ConvertMapY((Player(.Target).Y * 32) + Player(.Target).YOffSet) - 32
        Case TARGET_TYPE_EVENT
            Colour = .Colour
            X = ConvertMapX(Map.TileData.Events(.Target).X * 32) + 16
            Y = ConvertMapY(Map.TileData.Events(.Target).Y * 32) - 16
        Case Else
            Exit Sub
        End Select

        ' word wrap
        WordWrap_Array .Msg, ChatBubbleWidth, theArray
        ' find max width
        tmpNum = UBound(theArray)

        For i = 1 To tmpNum
            If TextWidth(font(Fonts.georgiaDec_16), theArray(i)) > MaxWidth Then MaxWidth = TextWidth(font(Fonts.georgiaDec_16), theArray(i))
        Next

        ' calculate the new position
        X2 = X - (MaxWidth \ 2)
        Y2 = Y - (UBound(theArray) * 12)
        ' render bubble - top left
        RenderTexture Tex_GUI(33), X2 - 9, Y2 - 5, 0, 0, 9, 5, 9, 5
        ' top right
        RenderTexture Tex_GUI(33), X2 + MaxWidth, Y2 - 5, 119, 0, 9, 5, 9, 5
        ' top
        RenderTexture Tex_GUI(33), X2, Y2 - 5, 9, 0, MaxWidth, 5, 5, 5
        ' bottom left
        RenderTexture Tex_GUI(33), X2 - 9, Y, 0, 19, 9, 6, 9, 6
        ' bottom right
        RenderTexture Tex_GUI(33), X2 + MaxWidth, Y, 119, 19, 9, 6, 9, 6
        ' bottom - left half
        RenderTexture Tex_GUI(33), X2, Y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' bottom - right half
        RenderTexture Tex_GUI(33), X2 + (MaxWidth \ 2) + 6, Y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' left
        RenderTexture Tex_GUI(33), X2 - 9, Y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
        ' right
        RenderTexture Tex_GUI(33), X2 + MaxWidth, Y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1
        ' center
        RenderTexture Tex_GUI(33), X2, Y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
        ' little pointy bit
        RenderTexture Tex_GUI(33), X - 5, Y, 58, 19, 11, 11, 11, 11
        ' render each line centralised
        tmpNum = UBound(theArray)

        For i = 1 To tmpNum
            RenderText font(Fonts.georgia_16), theArray(i), X - (TextWidth(font(Fonts.georgiaDec_16), theArray(i)) / 2), Y2, Colour
            Y2 = Y2 + 12
        Next

        ' check if it's timed out - close it if so
        If .Timer + 5000 < getTime Then
            .Active = False
        End If
    End With
End Sub

Public Function isConstAnimated(ByVal Sprite As Long) As Boolean
    isConstAnimated = False

    Select Case Sprite

    Case 16, 21, 22, 26, 28
        isConstAnimated = True
    End Select

End Function

Public Function hasSpriteShadow(ByVal Sprite As Long) As Boolean
    hasSpriteShadow = True

    Select Case Sprite

    Case 25, 26
        hasSpriteShadow = False
    End Select

End Function

Public Sub DrawPlayer(ByVal Index As Long)
    Dim Anim As Byte
    Dim X As Long
    Dim Y As Long
    Dim Sprite As Long, spritetop As Long
    Dim rec As RECT
    Dim attackspeed As Long

    ' pre-load sprite for calculations
    Sprite = GetPlayerSprite(Index)

    'SetTexture Tex_Char(Sprite)
    If Sprite < 1 Or Sprite > Count_Char Then Exit Sub

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).Speed
    Else
        attackspeed = 1000
    End If

    If Not isConstAnimated(GetPlayerSprite(Index)) Then
        ' Reset frame
        Anim = 1

        ' Check for attacking animation
        If Player(Index).AttackTimer + (attackspeed / 2) > getTime Then
            If Player(Index).Attacking = 1 Then
                Anim = 2
            End If

        Else

            ' If not attacking, walk normally
            Select Case GetPlayerDir(Index)

            Case DIR_UP

                If (Player(Index).YOffSet > 8) Then Anim = Player(Index).step

            Case DIR_DOWN

                If (Player(Index).YOffSet < -8) Then Anim = Player(Index).step

            Case DIR_LEFT

                If (Player(Index).XOffSet > 8) Then Anim = Player(Index).step

            Case DIR_RIGHT

                If (Player(Index).XOffSet < -8) Then Anim = Player(Index).step
            End Select

        End If

    Else

        If Player(Index).AnimTimer + 100 <= getTime Then
            Player(Index).Anim = Player(Index).Anim + 1

            If Player(Index).Anim >= 3 Then Player(Index).Anim = 0
            Player(Index).AnimTimer = getTime
        End If

        Anim = Player(Index).Anim
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If Player(Index).Attacking <> 0 Then
            If .AttackTimer + attackspeed < getTime Then
                .Attacking = 0
                .AttackTimer = 0
            End If
        End If

    End With

    ' Set the left
    Select Case GetPlayerDir(Index)

    Case DIR_UP
        spritetop = 3

    Case DIR_RIGHT
        spritetop = 2

    Case DIR_DOWN
        spritetop = 0

    Case DIR_LEFT
        spritetop = 1
    End Select

   ' With rec
   '     .top = spritetop * (mTexture(Tex_Char(Sprite)).h / 4)
    '    .Height = (mTexture(Tex_Char(Sprite)).h / 4)
    '    .Left = Anim * (mTexture(Tex_Char(Sprite)).w / 4)
    '    .Width = (mTexture(Tex_Char(Sprite)).w / 4)
   ' End With
    
    With rec
        .top = (mTexture(Tex_Char(Sprite)).h / 4) * spritetop
        .bottom = .top + mTexture(Tex_Char(Sprite)).h / 4
        .Left = Anim * (mTexture(Tex_Char(Sprite)).w / 4)
        .Right = .Left + (mTexture(Tex_Char(Sprite)).w / 4)
    End With

    ' Calculate the X
    X = GetPlayerX(Index) * PIC_X + Player(Index).XOffSet - ((mTexture(Tex_Char(Sprite)).w / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (mTexture(Tex_Char(Sprite)).h) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffSet - ((mTexture(Tex_Char(Sprite)).h / 4) - 32) - 4
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffSet - 4
    End If

    If IsDay Then
        Dim Height As Long
        'Height = (rec.Height - rec.top)
        'RenderTexture Tex_Char(Sprite), ConvertMapX(X), ConvertMapY(Y) + (Height * 1.5) + 8, rec.Left, rec.top, rec.Width - rec.Left, rec.top - rec.Height, rec.Width - rec.Left, rec.top - rec.top, D3DColorRGBA(255, 255, 255, 100)
        'RenderTexture Tex_Char(Sprite), ConvertMapX(X), ConvertMapY(Y + 32), rec.Left, rec.top, rec.Width, rec.Height, rec.Width, rec.Height, D3DColorRGBA(0, 0, 0, 100), , 180, 3
        DrawShadow Sprite, X, Y + 5, rec, 50, 0, 0, 0
    End If
    RenderTexture Tex_Char(Sprite), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.top, rec.Right - rec.Left, rec.bottom - rec.top, rec.Right - rec.Left, rec.bottom - rec.top

    ' Mostra o player animado na janela do enemybars!
    'If myTarget > 0 Then
    '    If myTargetType = TARGET_TYPE_PLAYER Then
    '        If Index = myTarget Then
    '            With Windows(GetWindowIndex("winEnemyBars"))
    '                If .Window.visible Then
    '                    Sprite = GetPlayerSprite(.Controls(GetControlIndex("winEnemyBars", "picChar")).Value)
    '                    If Sprite > 0 And Sprite <= Count_Char Then
    '                        RenderTexture Tex_Char(Sprite), .Window.Left + .Controls(GetControlIndex("winEnemyBars", "picChar")).Left, .Window.top + .Controls(GetControlIndex("winEnemyBars", "picChar")).top, rec.Left, rec.top, rec.Width, rec.Height, rec.Width, rec.Height
    '                    End If
    '                End If
    '            End With
    '        End If
    '    End If
    'End If
End Sub

Private Sub DrawShadow(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, rec As RECT, Optional a As Byte = 255, Optional r As Byte = 255, Optional g As Byte = 255, Optional B As Byte = 255)
    Dim X As Long
    Dim Y As Long
    Dim Width As Long
    Dim Height As Long
    Dim SombraSize As Long

    If Sprite < 1 Or Sprite > Count_Char Then Exit Sub
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.bottom - rec.top)

    SombraSize = Height

    RenderTexture Tex_Char(Sprite), X, Y + (Height * 1.5) + 8, rec.Left, rec.top, rec.Right - rec.Left, rec.top - rec.bottom, rec.Right - rec.Left, rec.bottom - rec.top, D3DColorRGBA(r, g, B, a)
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim X As Long
    Dim Y As Long
    Dim Sprite As Long, spritetop As Long
    Dim rec As RECT
    Dim attackspeed As Long

    ' pre-load texture for calculations
    Sprite = NPC(MapNpc(MapNpcNum).num).Sprite

    'SetTexture Tex_Char(Sprite)
    If Sprite < 1 Or Sprite > Count_Char Then Exit Sub
    attackspeed = 1000

    If Not isConstAnimated(NPC(MapNpc(MapNpcNum).num).Sprite) Then
        ' Reset frame
        Anim = 1

        ' Check for attacking animation
        If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > getTime Then
            If MapNpc(MapNpcNum).Attacking = 1 Then
                Anim = 2
            End If

        Else

            ' If not attacking, walk normally
            Select Case MapNpc(MapNpcNum).dir

            Case DIR_UP

                If (MapNpc(MapNpcNum).YOffSet > 8) Then Anim = MapNpc(MapNpcNum).step

            Case DIR_DOWN

                If (MapNpc(MapNpcNum).YOffSet < -8) Then Anim = MapNpc(MapNpcNum).step

            Case DIR_LEFT

                If (MapNpc(MapNpcNum).XOffSet > 8) Then Anim = MapNpc(MapNpcNum).step

            Case DIR_RIGHT

                If (MapNpc(MapNpcNum).XOffSet < -8) Then Anim = MapNpc(MapNpcNum).step
            End Select

        End If

    Else

        With MapNpc(MapNpcNum)

            If .AnimTimer + 100 <= getTime Then
                .Anim = .Anim + 1

                If .Anim >= 3 Then .Anim = 0
                .AnimTimer = getTime
            End If

            Anim = .Anim
        End With

    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)

        If .AttackTimer + attackspeed < getTime Then
            .Attacking = 0
            .AttackTimer = 0
        End If
        
    If .Dead > 0 Then
        Anim = 3
    End If
    
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).dir

    Case DIR_UP
        spritetop = 3

    Case DIR_RIGHT
        spritetop = 2

    Case DIR_DOWN
        spritetop = 0

    Case DIR_LEFT
        spritetop = 1
    End Select

   ' With rec
    '    .top = (mTexture(Tex_Char(Sprite)).h / 4) * spritetop
    '    .Height = mTexture(Tex_Char(Sprite)).h / 4
   '     .Left = Anim * (mTexture(Tex_Char(Sprite)).w / 4)
   '     .Width = (mTexture(Tex_Char(Sprite)).w / 4)
   ' End With
   
   With rec
        .top = (mTexture(Tex_Char(Sprite)).h / 4) * spritetop
        .bottom = .top + mTexture(Tex_Char(Sprite)).h / 4
        .Left = Anim * (mTexture(Tex_Char(Sprite)).w / 4)
        .Right = .Left + (mTexture(Tex_Char(Sprite)).w / 4)
    End With

    ' Calculate the X
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffSet - ((mTexture(Tex_Char(Sprite)).w / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (mTexture(Tex_Char(Sprite)).h / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffSet - ((mTexture(Tex_Char(Sprite)).h / 4) - 32) - 4
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffSet - 4
    End If

    ' Sombra do npc
    If NPC(MapNpc(MapNpcNum).num).Shadow > 0 And IsDay And MapNpc(MapNpcNum).Dead = NO Then
        DrawShadow Sprite, X, Y + 5, rec, 50, 0, 0, 0
    End If

    RenderTexture Tex_Char(Sprite), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.top, rec.Right - rec.Left, rec.bottom - rec.top, rec.Right - rec.Left, rec.bottom - rec.top
    
    ' Mostra o npc animado na janela do enemybars!
    'If myTarget > 0 Then
    '    If myTargetType = TARGET_TYPE_NPC Then
    '        If MapNpcNum = myTarget Then
    '            With Windows(GetWindowIndex("winEnemyBars"))
    '                If .Window.visible Then
    '                    Sprite = NPC(MapNpc(.Controls(GetControlIndex("winEnemyBars", "picChar")).Value).num).Sprite
    '                    If Sprite > 0 And Sprite <= Count_Char Then
    '                        RenderTexture Tex_Char(Sprite), .Window.Left + .Controls(GetControlIndex("winEnemyBars", "picChar")).Left, .Window.top + .Controls(GetControlIndex("winEnemyBars", "picChar")).top, rec.Left, rec.top, rec.Width, rec.Height, rec.Width, rec.Height
    '                    End If
    '                End If
    '            End With
    '        End If
    '    End If
    'End If
        
End Sub

Sub DrawEvent(eventNum As Long, pageNum As Long)
    Dim texNum As Long, X As Long, Y As Long

    ' render it
    With Map.TileData.Events(eventNum).EventPage(pageNum)
        If .GraphicType > 0 Then
            If .Graphic > 0 Then
                Select Case .GraphicType
                Case 1    ' character
                    If .Graphic < Count_Char Then
                        texNum = Tex_Char(.Graphic)
                    End If
                Case 2    ' tileset
                    If .Graphic < Count_Tileset Then
                        texNum = Tex_Tileset(.Graphic)
                    End If
                End Select
                If texNum > 0 Then
                    X = ConvertMapX(Map.TileData.Events(eventNum).X * 32)
                    Y = ConvertMapY(Map.TileData.Events(eventNum).Y * 32)
                    RenderTexture texNum, X, Y, .GraphicX * 32, .GraphicY * 32, 32, 32, 32, 32
                End If
            End If
        End If
    End With
End Sub

Sub DrawLowerEvents()
    Dim i As Long, X As Long

    If Map.TileData.EventCount = 0 Then Exit Sub
    For i = 1 To Map.TileData.EventCount
        ' find the active page
        If Map.TileData.Events(i).pageCount > 0 Then
            X = ActiveEventPage(i)
            If X > 0 Then
                ' make sure it's lower
                If Map.TileData.Events(i).EventPage(X).Priority <> 2 Then
                    ' render event
                    DrawEvent i, X
                End If
            End If
        End If
    Next
End Sub

Sub DrawUpperEvents()
    Dim i As Long, X As Long

    If Map.TileData.EventCount = 0 Then Exit Sub
    For i = 1 To Map.TileData.EventCount
        ' find the active page
        If Map.TileData.Events(i).pageCount > 0 Then
            X = ActiveEventPage(i)
            If X > 0 Then
                ' make sure it's lower
                If Map.TileData.Events(i).EventPage(X).Priority = 2 Then
                    ' render event
                    DrawEvent i, X
                End If
            End If
        End If
    Next
End Sub

Public Sub DrawTarget(ByVal X As Long, ByVal Y As Long)
    Dim Width As Long, Height As Long
    ' calculations
    Width = mTexture(Tex_Target).w / 2
    Height = mTexture(Tex_Target).h
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2) + 16
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    'EngineRenderRectangle Tex_Target, x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Target, X, Y, 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawTargetHover()
    Dim i As Long, X As Long, Y As Long, Width As Long, Height As Long

    If diaIndex > 0 Then Exit Sub
    Width = mTexture(Tex_Target).w / 2
    Height = mTexture(Tex_Target).h

    If Width <= 0 Then Width = 1
    If Height <= 0 Then Height = 1

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            X = (Player(i).X * 32) + Player(i).XOffSet + 32
            Y = (Player(i).Y * 32) + Player(i).YOffSet + 32

            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    X = ConvertMapX(X)
                    Y = ConvertMapY(Y)
                    RenderTexture Tex_Target, X - 16 - (Width / 2), Y - 16 - (Height / 2), Width, 0, Width, Height, Width, Height
                End If
            End If
        End If

    Next

    For i = 1 To MAX_MAP_NPCS

        If MapNpc(i).num > 0 Then
            X = (MapNpc(i).X * 32) + MapNpc(i).XOffSet + 32
            Y = (MapNpc(i).Y * 32) + MapNpc(i).YOffSet + 32

            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    X = ConvertMapX(X)
                    Y = ConvertMapY(Y)
                    RenderTexture Tex_Target, X - 16 - (Width / 2), Y - 16 - (Height / 2), Width, 0, Width, Height, Width, Height
                End If
            End If
        End If

    Next

End Sub

Public Sub DrawResource(ByVal Resource_num As Long)
    Dim Resource_master As Long
    Dim Resource_state As Long
    Dim Resource_sprite As Long
    Dim rec As RECT
    Dim X As Long, Y As Long
    Dim Width As Long, Height As Long, i As Long, Alpha As Byte
    Dim sString As String
    
    X = MapResource(Resource_num).X
    Y = MapResource(Resource_num).Y

    If X < 0 Or X > Map.MapData.MaxX Then Exit Sub
    If Y < 0 Or Y > Map.MapData.MaxY Then Exit Sub
    ' Get the Resource type
    Resource_master = Map.TileData.Tile(X, Y).Data1

    If Resource_master = 0 Then Exit Sub
    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then    ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then    ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' pre-load texture for calculations
    'SetTexture Tex_Resource(Resource_sprite)
    ' src rect
    With rec
        .top = 0
        .bottom = mTexture(Tex_Resource(Resource_sprite)).h
        .Left = 0
        .Right = mTexture(Tex_Resource(Resource_sprite)).w
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (mTexture(Tex_Resource(Resource_sprite)).w / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - mTexture(Tex_Resource(Resource_sprite)).h + 32
    Width = rec.Right - rec.Left
    Height = rec.bottom - rec.top

    If ConvertMapY(GetPlayerY(MyIndex)) < ConvertMapY(MapResource(Resource_num).Y) And ConvertMapY(GetPlayerY(MyIndex)) > ConvertMapY(MapResource(Resource_num).Y) - ((mTexture(Tex_Resource(Resource_sprite)).h) / 32) Then
        If ConvertMapX(GetPlayerX(MyIndex)) > ConvertMapX(MapResource(Resource_num).X) - ((mTexture(Tex_Resource(Resource_sprite)).w / 2 + 16) / 32) And ConvertMapX(GetPlayerX(MyIndex)) <= ConvertMapX(MapResource(Resource_num).X) + ((mTexture(Tex_Resource(Resource_sprite)).w / 2) / 32) Then
            Alpha = 100
        Else
            Alpha = 255
        End If
    Else
        Alpha = 255
    End If
    
    If Resource(Resource_master).Shadow > 0 Then
        If Alpha <> 100 Then
            ', rec.Left, rec.top, rec.Right - rec.Left, rec.bottom - rec.top, rec.Right - rec.Left, rec.bottom - rec.top
            'RenderTexture Tex_Resource(Resource_sprite), ConvertMapX(X), ConvertMapY(Y + Height - 20), 0, 0, Width, Height, Width, Height, D3DColorARGB(100, 0, 0, 0), , 180, 3
            DrawResourceShadow Resource_sprite, X, Y + 5, rec, 50, 0, 0, 0
        Else
            'RenderTexture Tex_Resource(Resource_sprite), ConvertMapX(X), ConvertMapY(Y + Height - 20), 0, 0, Width, Height, Width, Height, D3DColorARGB(50, 0, 0, 0), , 180, 3
            DrawResourceShadow Resource_sprite, X, Y + 5, rec, 15, 0, 0, 0
        End If
    End If

    RenderTexture Tex_Resource(Resource_sprite), ConvertMapX(X), ConvertMapY(Y), 0, 0, Width, Height, Width, Height, D3DColorARGB(Alpha, 255, 255, 255)

    For i = 1 To MAX_QUESTS
        'check if the npc is the next task to any quest: [?] symbol
        If Trim$(Quest(i).Name) <> "" Then
            If Player(MyIndex).PlayerQuest(i).Status = QUEST_STARTED Then
                If Quest(i).Task(Player(MyIndex).PlayerQuest(i).ActualTask).Resource = Resource_master Then
                    X = ConvertMapX(MapResource(Resource_num).X * PIC_X) + (mTexture(Tex_GUI(4)).w / 2)
                    Y = ConvertMapY(MapResource(Resource_num).Y * PIC_Y) + 32
                    RenderTexture_Animated Tex_GUI(4), X, Y, 0, 0, 13, 13, 13, 13, TextureQuestObj, D3DColorARGB(255, 255, 255, 0)
                    
                    If GlobalX >= X And GlobalX <= X + 13 Then
                        If GlobalY >= Y And GlobalY <= Y + 13 Then
                            sString = "Objetivo de Missao!"
                            Call RenderEntity_Square(Tex_Design(6), GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)) - 5, GlobalY - 35, TextWidth(font(Fonts.georgiaBold_16), sString) + 10, 20, 5, 200)
                            Call RenderText(font(Fonts.georgiaBold_16), sString, GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)), GlobalY - 32, Yellow)
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Private Sub DrawResourceShadow(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, rec As RECT, Optional a As Byte = 255, Optional r As Byte = 255, Optional g As Byte = 255, Optional B As Byte = 255)
    Dim X As Long
    Dim Y As Long
    Dim Width As Long
    Dim Height As Long
    Dim SombraSize As Long

    If Sprite < 1 Or Sprite > Count_Resource Then Exit Sub
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.bottom - rec.top)

    SombraSize = Height

    RenderTexture Tex_Resource(Sprite), X, Y + (Height * 1.5) + 8, rec.Left, rec.top, rec.Right - rec.Left, rec.top - rec.bottom, rec.Right - rec.Left, rec.bottom - rec.top, D3DColorRGBA(r, g, B, a)
End Sub

Public Sub DrawItem(ByVal itemNum As Long)
    Dim PicNum As Integer, dontRender As Boolean, i As Long, tmpIndex As Long, Colour As Byte, textX As Long, textY As Long
    Dim sString As String, ItemSizeMouse As Long, rec As RECT
    PicNum = Item(MapItem(itemNum).num).Pic
    
    ' Default item size
    ItemSizeMouse = 32

    If PicNum < 1 Or PicNum > Count_Item Then Exit Sub
    
    ' Animação ao dropar
    If MapItem(itemNum).Gravity < 0 Then
        MapItem(itemNum).Gravity = MapItem(itemNum).Gravity + 1
        MapItem(itemNum).YOffSet = MapItem(itemNum).YOffSet - 3
    ElseIf MapItem(itemNum).Gravity < 11 Then
        MapItem(itemNum).Gravity = MapItem(itemNum).Gravity + 1
        MapItem(itemNum).YOffSet = MapItem(itemNum).YOffSet + 3
        
        If MapItem(itemNum).Gravity = 11 Then
            MapItem(itemNum).YOffSet = 0
        End If
    End If

    ' if it's not us then don't render
    If MapItem(itemNum).playerName <> vbNullString Then
        If Trim$(MapItem(itemNum).playerName) <> Trim$(GetPlayerName(MyIndex)) Then

            dontRender = True
        End If

        ' make sure it's not a party drop
        If Party.Leader > 0 Then

            For i = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party.Member(i)

                If tmpIndex > 0 Then
                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(itemNum).playerName) Then
                        If MapItem(itemNum).bound = 0 Then

                            dontRender = False
                        End If
                    End If
                End If

            Next

        End If
    End If

    'If Not dontRender Then EngineRenderRectangle Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32, 32, 32
    If Not dontRender Then
    
        
        With rec
            rec.top = 0
            rec.Left = MapItem(itemNum).Frame * PIC_X
        End With
        
        
        ' Recicles variables to use in Centralize Item on mousepoint
        textX = MapItem(itemNum).X * PIC_X
        'textY = MapItem(itemNum).Y * PIC_Y
        
        textY = (MapItem(itemNum).Y * PIC_Y) + MapItem(itemNum).YOffSet
    
        If GlobalX >= ConvertMapX(MapItem(itemNum).X * PIC_X) And GlobalX <= ConvertMapX(MapItem(itemNum).X * PIC_X) + PIC_X Then
            If GlobalY >= ConvertMapY(MapItem(itemNum).Y * PIC_Y) And GlobalY <= ConvertMapY(MapItem(itemNum).Y * PIC_Y) + PIC_Y Then
                ItemSizeMouse = (PIC_X + (PIC_X / 2))
                textX = textX - ((ItemSizeMouse - PIC_X) / 2)
                textY = textY - ((ItemSizeMouse - PIC_Y) / 2)
                Call GroundItem_MouseMove(GlobalX, GlobalY, MapItem(itemNum).num, MapItem(itemNum).bound)
            End If
        End If
        
        If Options.ItemAnimation = YES Then
            RenderTexture_Animated Tex_Item(PicNum), ConvertMapX(textX), ConvertMapY(textY), rec.Left, rec.top, ItemSizeMouse, ItemSizeMouse, PIC_X, PIC_Y, TextureItem
        Else
            RenderTexture Tex_Item(PicNum), ConvertMapX(textX), ConvertMapY(textY), 0, 0, ItemSizeMouse, ItemSizeMouse, PIC_X, PIC_Y
        End If
        
        Colour = GetItemNameColour(Item(MapItem(itemNum).num).Rarity)
        If Options.ItemName = YES Or CurX = MapItem(itemNum).X And CurY = MapItem(itemNum).Y Then
            RenderText font(Fonts.rockwell_15), Trim$(Item(MapItem(itemNum).num).Name), 16 + ConvertMapX(MapItem(itemNum).X * PIC_X) - (TextWidth(font(Fonts.rockwell_15), Trim$(Item(MapItem(itemNum).num).Name)) / 2), ConvertMapY(MapItem(itemNum).Y * PIC_Y) - 10, Colour
        End If
    End If

    For i = 1 To MAX_QUESTS
        'check if the npc is the next task to any quest: [?] symbol
        If Trim$(Quest(i).Name) <> "" Then
            If Player(MyIndex).PlayerQuest(i).Status = QUEST_STARTED Then
                If Quest(i).Task(Player(MyIndex).PlayerQuest(i).ActualTask).Item = MapItem(itemNum).num Then
                    textX = 16 + ConvertMapX(MapItem(itemNum).X * PIC_X) - (mTexture(Tex_GUI(5)).w / 2)
                    textY = ConvertMapY(MapItem(itemNum).Y * PIC_Y) - 20
                    RenderTexture_Animated Tex_GUI(5), textX, textY, 0, 0, 13, 13, 13, 13, TextureQuestObj, D3DColorARGB(255, 255, 255, 0)

                    If GlobalX >= textX And GlobalX <= textX + 13 Then
                        If GlobalY >= textY And GlobalY <= textY + 13 Then
                            sString = "Objetivo de Missao!"
                            Call RenderEntity_Square(Tex_Design(6), GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)) - 5, GlobalY - 35, TextWidth(font(Fonts.georgiaBold_16), sString) + 10, 20, 5, 200)
                            Call RenderText(font(Fonts.georgiaBold_16), sString, GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)), GlobalY - 32, Yellow)
                        End If
                    End If

                End If
            End If
        End If
    Next


End Sub

Private Sub GroundItem_MouseMove(ByVal X As Long, ByVal Y As Long, ByVal itemNum As Long, ByVal SoulBound As Byte)
    Dim i As Long
    Dim IsBound As Boolean

    ' exit out early if dragging
    If DragBox.Type <> part_None Then Exit Sub

        ' exit out if we're offering that item
        ' make sure we're not dragging the item
        If DragBox.Type = Part_Item And DragBox.Value = itemNum Then Exit Sub
        ' calc position
        X = GlobalX - Windows(GetWindowIndex("winDescription")).Window.Width
        Y = GlobalY - Windows(GetWindowIndex("winDescription")).Window.Height
        ' offscreen?
        If X < 0 Then
            ' switch to right
            X = GlobalX
        End If
        
        If Y < 0 Then
            ' switch to right
            Y = GlobalY
        End If
        ' go go go
        
        If SoulBound > 0 Then IsBound = True
        
        ShowItemDesc X, Y, itemNum, IsBound
End Sub

Public Sub DrawBars()
    Dim Left As Long, top As Long, Width As Long, Height As Long
    Dim tmpX As Long, tmpY As Long, barWidth As Long, i As Long, NpcNum As Long
    Dim MaxHP As Long, HP As Long
    Dim MaxMP As Long, MP As Long
    Dim partyIndex As Long

    ' dynamic bar calculations
    Width = mTexture(Tex_Bars).w
    Height = mTexture(Tex_Bars).h / 4

    ' render npc health bars
    For i = 1 To MAX_MAP_NPCS
        NpcNum = MapNpc(i).num
        ' exists?
        If NpcNum > 0 Then
            ' alive?

            HP = GetNpcVitals(i, Vitals.HP)
            MP = GetNpcVitals(i, Vitals.MP)
            MaxHP = GetNpcMaxVitals(i, Vitals.HP)
            MaxMP = GetNpcMaxVitals(i, Vitals.MP)

            If HP > 0 And HP < MaxHP Then
                ' lock to npc
                tmpX = MapNpc(i).X * PIC_X + MapNpc(i).XOffSet + 16 - (Width / 2)
                tmpY = MapNpc(i).Y * PIC_Y + MapNpc(i).YOffSet + 35

                ' calculate the width to fill
                If Width > 0 Then BarWidth_NpcHP_Max(i) = ((HP / Width) / (MaxHP / Width)) * Width

                ' draw bar background
                top = Height * 1    ' HP bar background
                Left = 0
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, top, Width, Height, Width, Height

                ' draw the bar proper
                top = 0    ' HP bar
                Left = 0
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, top, BarWidth_NpcHP(i), Height, BarWidth_NpcHP(i), Height
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer).Spell).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffSet + 16 - (Width / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffSet + 35 + Height + 1

            ' calculate the width to fill
            If Width > 0 Then barWidth = (getTime - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer).Spell).CastTime * 1000)) * Width

            ' draw bar background
            top = Height * 3    ' cooldown bar background
            Left = 0
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, top, Width, Height, Width, Height

            ' draw the bar proper
            top = Height * 2    ' cooldown bar
            Left = 0
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, top, barWidth, Height, barWidth, Height
        End If
    End If

    HP = GetPlayerVital(MyIndex, Vitals.HP)
    MP = GetPlayerVital(MyIndex, Vitals.MP)
    MaxHP = GetPlayerMaxVital(MyIndex, Vitals.HP)
    MaxMP = GetPlayerMaxVital(MyIndex, Vitals.MP)
    ' draw own health bar
    If HP > 0 And HP < MaxHP Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffSet + 16 - (Width / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).YOffSet + 35

        ' calculate the width to fill
        If Width > 0 Then BarWidth_PlayerHP_Max(MyIndex) = ((HP / Width) / (MaxHP / Width)) * Width

        ' draw bar background
        top = Height * 1    ' HP bar background
        Left = 0
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, top, Width, Height, Width, Height

        ' draw the bar proper
        top = 0    ' HP bar
        Left = 0
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, top, BarWidth_PlayerHP(MyIndex), Height, BarWidth_PlayerHP(MyIndex), Height
    End If
End Sub

Sub DrawMenuBG()

'renderiza a parallax do menu
    RenderTexture Tex_Panoramas(MenuBG), ParallaxX, 0, 0, 0, ScreenWidth, ScreenHeight, 800, 600
    RenderTexture Tex_Panoramas(MenuBG), ParallaxX + ScreenWidth, 0, 0, 0, ScreenWidth, ScreenHeight, 800, 600
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
    Dim Sprite As Integer, sRECT As GeomRec, Width As Long, Height As Long, FrameCount As Long
    Dim X As Long, Y As Long, lockindex As Long

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If

    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)

    If Sprite < 1 Or Sprite > Count_Anim Then Exit Sub
    ' pre-load texture for calculations
    'SetTexture Tex_Anim(Sprite)
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    ' total width divided by frame count
    Width = 192    'mTexture(Tex_Anim(Sprite)).width / frameCount
    Height = 192    'mTexture(Tex_Anim(Sprite)).height

    With sRECT
        .top = (Height * ((AnimInstance(Index).frameIndex(Layer) - 1) \ AnimColumns))
        .Height = Height
        .Left = (Width * (((AnimInstance(Index).frameIndex(Layer) - 1) Mod AnimColumns)))
        .Width = Width
    End With

    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then    ' if <> none

        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex

            ' check if is ingame
            If IsPlaying(lockindex) Then

                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).XOffSet
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).YOffSet
                End If
            End If

        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex

            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then

                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).XOffSet
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).YOffSet
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If

            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If

    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    'EngineRenderRectangle Tex_Anim(sprite), x, y, sRECT.left, sRECT.top, sRECT.width, sRECT.height, sRECT.width, sRECT.height, sRECT.width, sRECT.height
    RenderTexture Tex_Anim(Sprite), X, Y, sRECT.Left, sRECT.top, sRECT.Width, sRECT.Height, sRECT.Width, sRECT.Height
End Sub

Public Sub DrawGDI()

    If frmEditor_Animation.visible Then
        GDIRenderAnimation
    ElseIf frmEditor_Item.visible Then
        GDIRenderItem frmEditor_Item.picItem, frmEditor_Item.scrlPic.Value
    ElseIf frmEditor_Map.visible Then
        GDIRenderTileset
        
        If frmEditor_Map.fraLight.visible Then GDIRenderLight frmEditor_Map.picLight
        
        If frmEditor_Events.visible Then
            GDIRenderEventGraphic
            GDIRenderEventGraphicSel
        End If
    ElseIf frmEditor_NPC.visible Then
        GDIRenderChar frmEditor_NPC.picSprite, frmEditor_NPC.scrlSprite.Value
    ElseIf frmEditor_Resource.visible Then
        GDIRenderResource frmEditor_Resource.picNormalPic, frmEditor_Resource.scrlNormalPic.Value
        GDIRenderResource frmEditor_Resource.picExhaustedPic, frmEditor_Resource.scrlExhaustedPic.Value
    ElseIf frmEditor_Spell.visible Then
        GDIRenderSpell frmEditor_Spell.picSprite, frmEditor_Spell.scrlIcon.Value
    End If

End Sub

' Main Loop
Public Sub Render_Graphics()
    Dim X As Long, Y As Long, i As Long, bgColour As Long, RenderY As Long

    On Error GoTo Retry
Retry:

    ' fuck off if we're not doing anything
    If GettingMap Then Exit Sub

    ' update the camera
    UpdateCamera

    ' check graphics
    CheckGFX

    ' Start rendering
    If Not InMapEditor Then
        bgColour = 0
    Else
        bgColour = DX8Colour(Red, 255)
    End If

    ' Bg
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, bgColour, 1#, 0)
    Call D3DDevice.BeginScene

    ' render black if map
    If InMapEditor Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                    RenderTexture Tex_Fader, ConvertMapX(X * 32), ConvertMapY(Y * 32), 0, 0, 32, 32, 32, 32
                End If
            Next
        Next
    End If

    If Map.MapData.Panorama > 0 And Map.MapData.Panorama <= Count_Panoramas Then
        RenderTexture Tex_Panoramas(Map.MapData.Panorama), ParallaxX, 0, 0, 0, ScreenWidth, ScreenHeight, 800, 600
        RenderTexture Tex_Panoramas(Map.MapData.Panorama), ParallaxX + ScreenWidth, 0, 0, 0, ScreenWidth, ScreenHeight, 800, 600
    End If

    ' Render appear tile fades
    RenderAppearTileFade

    ' render lower tiles
    If Count_Tileset > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                    Call DrawMapTile(X, Y)
                End If
            Next
        Next
    End If

    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).num > 0 Then    ' no npc set
            If MapNpc(i).Dead = YES Then
                Call DrawNpc(i)
            End If
        End If
    Next

    ' render the decals
    For i = 1 To Blood_HighIndex
        Call DrawBlood(i)
    Next

    ' render the items
    If Count_Item > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call DrawItem(i)
            End If
        Next
    End If

    ' draw animations
    If Count_Anim > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                DrawAnimation i, 0
            End If
        Next
    End If

    ' draw events
    DrawLowerEvents

    For i = 1 To MAX_WEATHER_PARTICLES
        With WeatherImpact(i)
            If .Impact Then
                DrawWeather_Impact i
            End If
        End With
    Next i

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For RenderY = 0 To Map.MapData.MaxY

        If Count_Char > 0 Then
            ' Players
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).Y = RenderY Then
                        Call DrawPlayer(i)
                    End If
                End If
            Next

            ' Npcs
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then    ' no npc set
                    If MapNpc(i).Dead = NO Then
                        If MapNpc(i).Y = RenderY Then
                            Call DrawNpc(i)
                        End If
                    End If
                End If
            Next
        End If

        ' Resources
        If Count_Resource > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = RenderY Then
                            Call DrawResource(i)
                        End If
                    Next

                End If
            End If
        End If

    Next RenderY

    ' render out upper tiles
    If Count_Tileset > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                    Call DrawMapFringeTile(X, Y)
                End If
            Next
        Next
    End If

    ' render animations
    If Count_Anim > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                DrawAnimation i, 1
            End If
        Next
    End If

    DrawFog

    DrawNight

    DrawTileOutline

    DrawWeather

    ' blit out lights
    For X = TileView.Left To TileView.Right
        For Y = TileView.top To TileView.bottom
            If IsValidMapPoint(X, Y) Then
                If Map.TileData.Tile(X, Y).Type = TILE_TYPE_LIGHT Then
                    If Not IsDay Then
                        Call DrawLight(X * 32, Y * 32, Map.TileData.Tile(X, Y).Data1, Map.TileData.Tile(X, Y).Data2, Map.TileData.Tile(X, Y).Data3, Map.TileData.Tile(X, Y).Data4, Map.TileData.Tile(X, Y).Data5)
                    End If
                End If
            End If
        Next
    Next

    DrawSun

    ' draw events
    DrawUpperEvents

    DrawTint

    ' render target
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
            DrawTarget (Player(myTarget).X * 32) + Player(myTarget).XOffSet, (Player(myTarget).Y * 32) + Player(myTarget).YOffSet
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            DrawTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).XOffSet, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).YOffSet
        End If
    End If

    ' blt the hover icon
    DrawTargetHover

    ' draw the bars
    DrawBars

    ' draw attributes
    If InMapEditor Then
        DrawMapAttributes
        DrawMapEvents
    End If

    ' draw player names
    If Not screenshotMode Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                    DrawPlayerStatus i
                    DrawPlayerName i
                    DrawGuild i
                Else
                    If CurX = GetPlayerX(i) And CurY = GetPlayerY(i) Then
                        DrawPlayerStatus i
                        DrawPlayerName i
                        DrawGuild i
                    Else
                        DrawPlayerName i
                        DrawGuild i
                        DrawPlayerStatus i
                    End If
                End If
            End If
        Next
    End If

    ' draw npc names
    If Not screenshotMode Then
        For i = 1 To MAX_MAP_NPCS
            With MapNpc(i)
                If .num > 0 Then
                    If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                        Call DrawNpcStatus(i)
                        Call DrawNpcName(i)
                    Else
                        If CurX = .X And CurY = .Y Then
                            Call DrawNpcStatus(i)
                            Call DrawNpcName(i)
                        Else
                            Call DrawNpcName(i)
                            Call DrawNpcStatus(i)
                        End If
                    End If
                End If
            End With
        Next
    End If

    ' draw action msg
    For i = 1 To MAX_BYTE
        DrawActionMsg i
    Next

    If InMapEditor Then
        If frmEditor_Map.optBlock.Value = True Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.top To TileView.bottom
                    If IsValidMapPoint(X, Y) Then
                        Call DrawDirection(X, Y)
                    End If
                Next
            Next
        End If
    End If

    ' draw the messages
    For i = 1 To MAX_BYTE
        If chatBubble(i).Active Then
            DrawChatBubble i
        End If
    Next

    ' Not ScreenShot Mode
    If Not screenshotMode Then
        RenderTexture Tex_GUI(43), 0, 0, 0, 0, ScreenWidth, 64, 1, 64
        RenderTexture Tex_GUI(42), 0, ScreenHeight - 64, 0, 0, ScreenWidth, 64, 1, 64

        If Options.FPSConection = YES Then
            RenderText font(Fonts.rockwell_15), "FPS: " & GameFPS & Space(24) & "Ping: " & Ping, 1, 1, White
        End If

        ' Not Hide Gui, Not ScreenShot Mode - Game Hour
        If Not hideGUI Then
            RenderText font(Fonts.rockwellDec_10), KeepTwoDigit(GameHours) & ":" & KeepTwoDigit(GameMinutes) & ":" & KeepTwoDigit(GameSeconds), ScreenWidth - 15 - TextWidth(font(Fonts.rockwellDec_10), KeepTwoDigit(GameHours) & ":" & KeepTwoDigit(GameMinutes) & ":" & KeepTwoDigit(GameSeconds)), 65, Yellow

            If QuestTimeToFinish <> vbNullString And QuestNameToFinish <> vbNullString Then
                If QuestSelect > 0 Then
                    RenderText font(Fonts.rockwell_15), QuestTimeToFinish, ConvertMapX(GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffSet + (PIC_X \ 2) - (TextWidth(font(Fonts.rockwell_15), QuestTimeToFinish) \ 2)), ConvertMapY(GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).YOffSet - 52), Yellow
                    RenderText font(Fonts.rockwell_15), QuestNameToFinish, ConvertMapX(GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffSet + (PIC_X \ 2) - (TextWidth(font(Fonts.rockwell_15), QuestNameToFinish) \ 2)), ConvertMapY(GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).YOffSet - 68), Yellow
                End If
            End If
            ' Not Map Editor, Not Hide Gui, Not ScreenShot Mode
            If Not InMapEditor Then
                If DrawThunder > 0 Then
                    RenderTexture Tex_Blank, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160)
                    DrawThunder = DrawThunder - 1
                End If

                ' draw map name
                RenderMapName

                RenderEntities
            Else
                ' draw loc
                If BLoc Then
                    RenderText font(Fonts.georgiaDec_16), Trim$("cur x: " & CurX & " y: " & CurY), 260, 6, Yellow
                    RenderText font(Fonts.georgiaDec_16), Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 260, 22, Yellow
                    RenderText font(Fonts.georgiaDec_16), Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 260, 38, Yellow
                End If
                DrawTileSelection
            End If
        End If
    End If

    ' End the rendering
    Call D3DDevice.EndScene

    If D3DDevice.TestCooperativeLevel = D3D_OK And Not D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST And Not D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    End If

    ' GDI Rendering
    DrawGDI
End Sub

Public Sub Render_Menu()
' check graphics
    CheckGFX
    ' Start rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, &HFFFFFF, 1#, 0)
    Call D3DDevice.BeginScene
    ' Render menu background
    DrawMenuBG
    ' Render entities
    RenderEntities
    ' render white fade
    DrawFade

    If Options.FPSConection = YES Then
        RenderText font(Fonts.rockwell_15), "FPS: " & GameFPS, 1, 1, White
    End If
    ' End the rendering
    Call D3DDevice.EndScene

    If D3DDevice.TestCooperativeLevel = D3D_OK And Not D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST And Not D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    End If
End Sub

Public Sub DrawBlood(ByVal Index As Long)

    BloodCount = mTexture(Tex_Blood).w / 32
    
    With Blood(Index)
        RenderTexture Tex_Blood, ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), Index * 32, 0, 32, 32, 32, 32, D3DColorARGB(Blood(Index).Alpha, 255, 255, 255)
    End With

End Sub

Private Sub DrawNpcStatus(ByVal MapNpcNum As Long)
    Dim X As Long, Y As Long, XX As Long, YY As Long, rec As RECT, sString As String

    With MapNpc(MapNpcNum)
        X = (.X * PIC_X) + .XOffSet
        Y = (.Y * PIC_Y) - 38 + .YOffSet
        X = ConvertMapX(X)
        Y = ConvertMapY(Y)

        'draw npc stun balão
        If .StunDuration > 0 Or .Dead > 0 Then
            rec.top = 0
            rec.Left = .StatusFrame * PIC_X
            RenderTexture Tex_Status(Status.Confused), X, Y, rec.Left, rec.top, 25, 25, 32, 32

            sString = "Npc está confuso!"
            If GlobalX >= X And GlobalX <= X + PIC_X Then
                If GlobalY >= Y And GlobalY <= Y + PIC_Y Then
                    Call RenderEntity_Square(Tex_Design(6), GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)) - 5, GlobalY - 35, TextWidth(font(Fonts.georgiaBold_16), sString) + 10, 20, 5, 200)
                    Call RenderText(font(Fonts.georgiaBold_16), sString, GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)), GlobalY - 32, Red)
                End If
            End If
            
            Exit Sub
        End If

        If CheckNpcHaveQuest(.num) Then
            rec.top = 0
            rec.Left = .StatusFrame * PIC_X
            RenderTexture Tex_Status(Status.Question), X, Y, rec.Left, rec.top, 25, 25, 32, 32

            sString = "Missao Disponivel!"
            If GlobalX >= X And GlobalX <= X + PIC_X Then
                If GlobalY >= Y And GlobalY <= Y + PIC_Y Then

                    ' calc position
                    XX = GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)) - 5
                    YY = GlobalY - 35
                    ' offscreen?
                    If XX < 0 Then
                        ' switch to right
                        XX = GlobalX
                    End If

                    If YY < 0 Then
                        ' switch to right
                        YY = GlobalY
                    End If
                    Call RenderEntity_Square(Tex_Design(6), XX, YY, TextWidth(font(Fonts.georgiaBold_16), sString) + 10, 20, 5, 200)
                    Call RenderText(font(Fonts.georgiaBold_16), sString, XX + 5, YY + 3, Green)
                End If
            End If
            
            Exit Sub
        End If

        If CheckNpcQuestProgress(.num) Then
            rec.top = 0
            rec.Left = .StatusFrame * PIC_X
            RenderTexture Tex_Status(Status.Important), X, Y, rec.Left, rec.top, 25, 25, 32, 32

            sString = "Missao em Progresso!"
            If GlobalX >= X And GlobalX <= X + PIC_X Then
                If GlobalY >= Y And GlobalY <= Y + PIC_Y Then
                    Call RenderEntity_Square(Tex_Design(6), GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)) - 5, GlobalY - 35, TextWidth(font(Fonts.georgiaBold_16), sString) + 10, 20, 5, 200)
                    Call RenderText(font(Fonts.georgiaBold_16), sString, GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)), GlobalY - 32, Yellow)
                End If
            End If
            
            Exit Sub
        End If
        
        If NPC(.num).Balao > 0 Then
            rec.top = 0
            rec.Left = .StatusFrame * PIC_X
            RenderTexture Tex_Status(NPC(.num).Balao), X, Y, rec.Left, rec.top, 25, 25, 32, 32

            Select Case NPC(.num).Balao
            Case Status.typing
                sString = "Npc Conversa..."
            Case Status.Afk
                sString = "Npc Dormindo..."
            Case Status.Confused
                sString = "Npc Confuso..."
            Case Else
                sString = "Status de Npc..."
            End Select

            If GlobalX >= X And GlobalX <= X + PIC_X Then
                If GlobalY >= Y And GlobalY <= Y + PIC_Y Then
                    Call RenderEntity_Square(Tex_Design(6), GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)) - 5, GlobalY - 35, TextWidth(font(Fonts.georgiaBold_16), sString) + 10, 20, 5, 200)
                    Call RenderText(font(Fonts.georgiaBold_16), sString, GlobalX - ((TextWidth(font(Fonts.georgiaBold_16), sString) / 2)), GlobalY - 32, White)
                End If
            End If
        End If

    End With

End Sub

Public Function CheckNpcHaveQuest(ByVal NpcNum As Integer) As Boolean
    Dim i As Byte

    CheckNpcHaveQuest = False

    ' Se o npc não tiver nenhuma conversa, então não tem quest
    If NPC(NpcNum).Conv <= 0 Then Exit Function

    With Conv(NPC(NpcNum).Conv)

        ' Se a conversa nao tem quantidade de abas pra verificação do for
        If .chatCount <= 0 Then Exit Function

        For i = 1 To .chatCount
            If .Conv(i).Event = 5 Then    ' É quest?
                If .Conv(i).Data1 > 0 Then
                    If Not QuestInProgress(.Conv(i).Data1) Then
                        If Player(MyIndex).PlayerQuest(.Conv(i).Data1).Status = QUEST_NOT_STARTED Or Player(MyIndex).PlayerQuest(.Conv(i).Data1).Status = QUEST_COMPLETED_BUT Then
                            CheckNpcHaveQuest = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next i

    End With

End Function

Public Function CheckNpcQuestProgress(ByVal NpcNum As Integer) As Boolean
    Dim i As Byte

    CheckNpcQuestProgress = False

    ' Se o npc não tiver nenhuma conversa, então não tem quest
    If NPC(NpcNum).Conv <= 0 Then Exit Function

    With Conv(NPC(NpcNum).Conv)

        ' Se a conversa nao tem quantidade de abas pra verificação do for
        If .chatCount <= 0 Then Exit Function

        For i = 1 To .chatCount
            If .Conv(i).Event = 5 Then    ' É quest?
                If .Conv(i).Data1 > 0 Then
                    If Player(MyIndex).PlayerQuest(.Conv(i).Data1).Status = QUEST_STARTED Then
                        CheckNpcQuestProgress = True
                        Exit Function
                    End If
                End If
            End If
        Next i

    End With

End Function

Public Sub DrawTileOutline()
    If IsValidMapPoint(CurX, CurY) Then
        Dim rec As RECT

        With rec
            .top = 0
            .bottom = .top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        RenderTexture Tex_Misc, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), rec.Left, rec.top, rec.Right - rec.Left, rec.bottom - rec.top, rec.Right - rec.Left, rec.bottom - rec.top, D3DColorRGBA(255, 255, 255, 255)
    End If
End Sub

Private Sub DrawLight(ByVal X As Long, ByVal Y As Long, ByVal a As Long, ByVal r As Long, ByVal g As Long, ByVal B As Long, ByVal LightSize As Byte)
    RenderTexture Tex_Light, ConvertMapX(X) - (((LightSize * 32) - 32) / 2), ConvertMapY(Y) - (((LightSize * 32) - 32) / 2), 0, 0, LightSize * 32, LightSize * 32, 128, 128, D3DColorARGB(Abs(Int(a) - Rand(0, 25)), Int(r), Int(g), Int(B))
End Sub

Sub DrawNight()
    Dim Alpha As Integer, X As Long, Y As Long

    If IsDay Then Exit Sub

    'Night/Day
    If Map.MapData.DayNight = 1 Then
        Alpha = 150
    ElseIf GameHours >= 18 And GameHours <= 24 Then
        Alpha = 150
    ElseIf GameHours >= 0 And GameHours < 4 Then
        Alpha = 150
    ElseIf GameHours = 4 Then
        Alpha = 120
    ElseIf GameHours = 5 Then
        Alpha = 90
    ElseIf GameHours = 6 Then
        Alpha = 50
    ElseIf GameHours >= 7 And GameHours < 18 Then
            Alpha = 0
    End If
    
    Alpha = 250

    RenderTexture Tex_LightMap, ConvertMapX(GetPlayerX(MyIndex) * 32) + Player(MyIndex).XOffSet + 16 - 1300, ConvertMapY(GetPlayerY(MyIndex) * 32) + Player(MyIndex).YOffSet - 812.5, 0, 0, 2600, 1625, 2600, 1625, D3DColorARGB(Alpha, 0, 0, 0)
End Sub
