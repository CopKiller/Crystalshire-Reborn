Attribute VB_Name = "modDx8toFile"
Option Explicit

' This method allows you to convert your Direct3DTexture8 or Direct3DSurface8 texture into an image file

Public Enum ImageFormats    ' All the image formats supported by GDIp
    BMP
    PNG
    JPEG
End Enum

Public Function ConvertSurfaceToArray(Surface As Direct3DSurface8, ByRef Pixels() As Byte) As Boolean
    Dim SurfDesc As D3DSURFACE_DESC, LockedRect As D3DLOCKED_RECT, rc As RECT
    Dim TempSurf As Direct3DSurface8

    ' This methods converts a Direct3DSurface8 using a format ARGB with 8 bits
    'per component into an array of bytes

    ' You can get the ARGB value in the form of a long by copying 4 bytes (starting from the first) using CopyMemory
    ' i.e CopyMemory Long, byte(count), 4

    ' Each individual component of the ARGB can be extracted directly from the array
    ' For eg. starting from 0, x = 0 -> A = Arr(x + 3), R = Arr(x + 2), G = Arr(x + 1), B = Arr(x)

    ' Currently this method only provides support for 32bpp images. Though you can add support for 24bpp or 16bpp or 8bpp.
    ' Contact me if you want more help with that.

    Call Surface.GetDesc(SurfDesc)

    ' Check if it's a format that we can convert to a byte array
    If SurfDesc.Format = D3DFMT_A8R8G8B8 Or D3DFMT_X8R8G8B8 Then
        ' Set up the rect so that DX8 knows which parts of the surface we want to copy from
        With rc    ' This will select the entire surface
            .Left = 0
            .Right = SurfDesc.Width
            .top = 0
            .bottom = SurfDesc.Height
        End With

        If SurfDesc.Pool = D3DPOOL_DEFAULT Then
            ' If the Texture is in the D3DPool_Default pool then we can't lock the rect
            ' We have make a copy of the Surface elsewhere before we attemptt lock it
            Set TempSurf = D3DDevice.CreateImageSurface(SurfDesc.Width, SurfDesc.Height, SurfDesc.Format)
            ' Copy the rect to the TempSurf
            Call D3DDevice.CopyRects(Surface, rc, 1, TempSurf, rc)
        Else
            ' If not in the Default pool then just set the TempSurf to the original Surf.
            Set TempSurf = Surface
        End If

        ' Lock the surface to get the pixel data.
        Call TempSurf.LockRect(LockedRect, ByVal 0, 0)
        ReDim Pixels((LockedRect.Pitch * SurfDesc.Height) - 1)    ' Pitch = Bytes Per Pixel * width
        If Not DXCopyMemory(Pixels(0), ByVal LockedRect.pBits, LockedRect.Pitch * SurfDesc.Height) = D3D_OK Then
            ' We weren't able to copy the data
            ConvertSurfaceToArray = False
            Exit Function
        End If
        ' If you want, you can tweak the Pixels array. To put it back into the DX8 texture copy back into pbits.
        ' Just reverse the Dest and Source parameters in DXCopyMemory. Make sure UnLockRect is called. Else it the texture won't update

        Call TempSurf.UnlockRect    ' This call is essential. If LockRect is called and this isn't called then it can lead to memory leaks

        ConvertSurfaceToArray = True
    End If
End Function

Public Function ConvertTextureToArray(Texture As Direct3DTexture8, Pixels() As Byte) As Boolean
    ConvertTextureToArray = ConvertSurfaceToArray(Texture.GetSurfaceLevel(0), Pixels)
End Function

Public Sub SaveDirectX8SurfToMemory(data() As Byte, Surf As Direct3DSurface8, Optional ByVal ImageType As ImageFormats = PNG)
    Dim GDIToken As cGDIpToken, renderer As cGDIpRenderer, image As cGDIpImage, i As Long
    Dim PixelFormat As ImageColorFormatConstants, Width As Long, Height As Long, SurfDesc As D3DSURFACE_DESC
    Dim ImageStride As Long, DataPointer As Long, Pixels() As Byte

    If Not ConvertSurfaceToArray(Surf, Pixels) Then Exit Sub
    Call Surf.GetDesc(SurfDesc)

    Set GDIToken = New cGDIpToken
    Set renderer = New cGDIpRenderer
    Set image = New cGDIpImage

    Call renderer.AttachTokenClass(GDIToken)
    i = renderer.CreateHGraphicsFromHWND(frmMain.hWnd)    ' afaik the hWnd can be the hWnd of any form in the project

    If image.LoadPicture_FromNothing(SurfDesc.Width, SurfDesc.Height, i, GDIToken) Then
        ' Set up the values for locking the image
        'Image.ColorFormat
        PixelFormat = PixelFormat32bppPARGB

        Width = image.Width: Height = image.Height

        ' lock the GDIp Image
        DataPointer = image.LockImageBits(ImageLockModeWrite, PixelFormat, 0, 0, Width, Height, ImageStride)
        If DataPointer > 0 Then
            ' Copy the data that we had in DirectX8 Texture to the GDIp image
            Call CopyMemory(ByVal DataPointer, Pixels(0), ImageStride * Height)
            ' Unlock the Gdip image so that the image updates, pass back the values that we used in LockImageBits
            Call image.UnLockImageBits(PixelFormat, Width, Height, ImageStride, DataPointer)
        End If

        If ImageType = ImageFormats.PNG Then
            Call image.SaveAsPNG(data)
        ElseIf ImageType = ImageFormats.BMP Then
            Call image.SaveAsBMP(data)
        ElseIf ImageType = ImageFormats.JPEG Then
            Call image.SaveAsJPG(data)
        Else
            Call image.SaveAsPNG(data)    ' Just save as a PNG
        End If
    End If
    Set image = Nothing
    Set renderer = Nothing
    Set GDIToken = Nothing
End Sub

Public Sub SaveDirectX8SurfToFile(ByVal path As String, Surf As Direct3DSurface8, Optional ByVal ImageType As ImageFormats = PNG)
    Dim data() As Byte, f As Long

    ' Convert the texture into a readable file
    Call SaveDirectX8SurfToMemory(data, Surf, ImageType)

    ' Dump the data we got into the file at the path

    If FileExist(path) Then Kill path

    f = FreeFile
    Open path For Binary As #f
    Put #f, , data
    Close #f

    ' The file should now be a valid image
End Sub

Public Sub SaveDirectX8TextureToMemory(data() As Byte, Texture As Direct3DTexture8, Optional ByVal ImageType As ImageFormats = ImageFormats.PNG)
    Call SaveDirectX8SurfToMemory(data, Texture.GetSurfaceLevel(0), ImageType)
End Sub

Public Sub SaveDirectX8TextureToFile(ByVal path As String, Texture As Direct3DTexture8, Optional ByVal ImageType As ImageFormats = ImageFormats.PNG)
    Call SaveDirectX8SurfToFile(path, Texture.GetSurfaceLevel(0), ImageType)
End Sub

Public Sub ScreenShotMap(ByVal Ground As Boolean, ByVal Fringe As Boolean, ByVal Resources As Boolean)
    Dim X As Integer, Y As Integer, i As Integer, NamePaste As String

    NamePaste = "Vazio"

    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene

    If Ground = True Then

        NamePaste = "Ground"
        ' render lower tiles
        If Count_Tileset > 0 Then
            For X = 0 To map.MapData.MaxX
                For Y = 0 To map.MapData.MaxY
                    Call DrawMapTile(X, Y)
                Next
            Next
        End If

    End If

    ' Resources
    If Resources = True Then
        NamePaste = "Resources"
        If Count_Resource > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        Call DrawResource(i)
                    Next
                End If
            End If
        End If
    End If

    ' render out upper tiles
    If Fringe = True Then
        NamePaste = "Fringe"
        If Count_Tileset > 0 Then
            For X = 0 To map.MapData.MaxX
                For Y = 0 To map.MapData.MaxY
                    Call DrawMapFringeTile(X, Y)
                Next
            Next
        End If
    End If

    If Ground And Resources And Fringe Then
        NamePaste = "MapaCompleto"
    End If

    ' End the rendering
    Call D3DDevice.EndScene

    If D3DDevice.TestCooperativeLevel = D3D_OK And Not D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST And Not D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Call D3DDevice.Present(ByVal 0, ByVal 0, ByVal 0, ByVal 0)
        Call SaveDirectX8SurfToFile(App.path & "\Mapas Salvos\" & NamePaste & GetPlayerMap(MyIndex) & ".png", D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), ImageFormats.PNG)
    End If
End Sub
