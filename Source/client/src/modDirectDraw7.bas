Attribute VB_Name = "modGraphics"
Option Explicit
' **********************
' ** Renders graphics **
' **********************
' DirectX8 Object
Private DirectX8 As DirectX8 'The master DirectX object.
Private Direct3D As Direct3D8 'Controls all things 3D.
Public Direct3D_Device As Direct3DDevice8 'Represents the hardware rendering.
Private Direct3DX As D3DX8

'The 2D (Transformed and Lit) vertex format.
Public Const FVF_TLVERTEX As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

'The 2D (Transformed and Lit) vertex format type.
Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    RHW As Single
    color As Long
    TU As Single
    TV As Single
End Type

Public BBWidth As Long
Public BBHeight As Long

Private Vertex_List(3) As TLVERTEX '4 vertices will make a square.

'Some color depth constants to help make the DX constants more readable.
Private Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Private Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Private Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8

Public RenderingMode As Long

Private Direct3D_Window As D3DPRESENT_PARAMETERS 'Backbuffer and viewport description.
Private Display_Mode As D3DDISPLAYMODE

'Graphic Textures
Public Tex_Item() As DX8TextureRec ' arrays
Public Tex_Character() As DX8TextureRec
Public Tex_Paperdoll() As DX8TextureRec
Public Tex_Tileset() As DX8TextureRec
Public Tex_Resource() As DX8TextureRec
Public Tex_Animation() As DX8TextureRec
Public Tex_SpellIcon() As DX8TextureRec
Public Tex_Face() As DX8TextureRec
Public Tex_Fog() As DX8TextureRec
Public Tex_GUI() As DX8TextureRec
Public Tex_Furniture() As DX8TextureRec
Public Tex_Door As DX8TextureRec ' singes
Public Tex_Blood As DX8TextureRec
Public Tex_Misc As DX8TextureRec
Public Tex_Direction As DX8TextureRec
Public Tex_Target As DX8TextureRec
Public Tex_Bars As DX8TextureRec
Public Tex_Selection As DX8TextureRec
Public Tex_White As DX8TextureRec
Public Tex_Weather As DX8TextureRec
Public Tex_ChatBubble As DX8TextureRec
Public Tex_Fade As DX8TextureRec
Public Tex_Arrows As DX8TextureRec
Public Tex_Pic() As DX8TextureRec
Public Tex_Projectiles() As DX8TextureRec

' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public numitems As Long
Public NumFurniture As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Public NumFogs As Long
Public NumGUI As Long
Public NumPics As Long
Public NumProjectiles As Long

'Hair
'Body
'Shirt
'Legs
'Shoes

Public Tex_CharHair() As DX8TextureRec
Public Tex_CharBodies() As DX8TextureRec
Public Tex_CharShirts() As DX8TextureRec
Public Tex_MaleLegs() As DX8TextureRec
Public Tex_FemaleLegs() As DX8TextureRec
Public Tex_MaleShoes() As DX8TextureRec
Public Tex_FemaleShoes() As DX8TextureRec
Public NumCharHair As Long
Public NumCharBodies As Long
Public NumCharShirts As Long
Public NumMaleLegs As Long
Public NumMaleShoes As Long
Public NumFemaleLegs As Long
Public NumFemaleShoes As Long

Public Tex_FHair() As DX8TextureRec
Public Tex_FHairB() As DX8TextureRec
Public Tex_FHeads() As DX8TextureRec
Public Tex_FEyes() As DX8TextureRec
Public Tex_FEyebrows() As DX8TextureRec
Public Tex_FEars() As DX8TextureRec
Public Tex_FMouth() As DX8TextureRec
Public Tex_FNose() As DX8TextureRec
Public Tex_FShirts() As DX8TextureRec
Public Tex_FEtc() As DX8TextureRec
Public NumFaceHair As Long
Public NumFaceHeads As Long
Public NumFaceEyes As Long
Public NumFaceEyebrows As Long
Public NumFaceEars As Long
Public NumFaceMouths As Long
Public NumFaceNoses As Long
Public NumFaceShirts As Long
Public NumFaceEtc As Long

Public MapCache As MapCacheRec

Private Type MapTexRec
    MapTexRec() As Direct3DTexture8
    MapTexSurf() As Direct3DSurface8 'handle to the surface data of the texture
End Type

Private Type LayerFrameRec
    Frame(1 To 3) As MapTexRec
End Type

Private Type MapCacheRec
    Layers(2) As LayerFrameRec
End Type

Public BackBufferSurf As Direct3DSurface8

Public Type DX8TextureRec
    Texture As Long
    Width As Long
    Height As Long
    filepath As String
    TexWidth As Long
    TexHeight As Long
End Type

Public Type GlobalTextureRec
    Texture As Direct3DTexture8
    TexWidth As Long
    TexHeight As Long
    TextureTimer As Long
    IsLoaded As Boolean
    ImageData() As Byte
    HasData As Boolean
End Type

Public Type rect
    Top As Long
    Left As Long
    Bottom As Long
    Right As Long
End Type

Public gTexture() As GlobalTextureRec
Public LastTexture As Long
Public TextureSaved As Long
Public NumTextures As Long

Public MapLayerImage As MapLayerCacheRec

Public Type MapLayerCacheRec
    Texture As Direct3DTexture8
    TexWidth As Long
    TexHeight As Long
    TextureTimer As Long
    IsLoaded As Boolean
    ImageData() As Byte
    HasData As Boolean
End Type

' ********************
' ** Initialization **
' ********************
Public Function InitDX8() As Boolean


   On Error GoTo errorhandler

    Set DirectX8 = New DirectX8 'Creates the DirectX object.
    Set Direct3D = DirectX8.Direct3DCreate() 'Creates the Direct3D object using the DirectX object.
    Set Direct3DX = New D3DX8
    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_DISCARD 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'DispMode.Format = D3DFMT_X8R8G8B8
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = BBWidth ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = BBHeight 'frmMain.ScaleHeight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hwnd 'Use frmMain as the device window.
    'we've already setup for Direct3D_Window.
    If TryCreateDirectX8Device = False Then
        MsgBox "Unable to initialize DirectX8. You may be missing dx8vb.dll or have incompatible hardware to use DirectX8."
        DestroyGame
    End If

    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
            .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
            .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
    ' Initialise the surfaces
    LoadTextures
    
    ' We're done
    InitDX8 = True



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "InitDX8", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Function
Public Function LoadDirectX(ByVal BehaviourFlags As CONST_D3DCREATEFLAGS, ByVal hwnd As Long)
On Error GoTo ErrorInit

    Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, BehaviourFlags, Direct3D_Window)
    Exit Function

ErrorInit:
    LoadDirectX = 1
End Function

Function TryCreateDirectX8Device() As Boolean
Dim i As Long

   On Error GoTo errorhandler

       TryCreateDirectX8Device = False
       Select Case Options.Render
        Case 1 ' hardware
            If LoadDirectX(D3DCREATE_HARDWARE_VERTEXPROCESSING, frmMain.hwnd) <> 0 Then
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with hardware vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case 2 ' mixed
            If LoadDirectX(D3DCREATE_MIXED_VERTEXPROCESSING, frmMain.hwnd) <> 0 Then
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with mixed vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case 3 ' software
            If LoadDirectX(D3DCREATE_SOFTWARE_VERTEXPROCESSING, frmMain.hwnd) <> 0 Then
                Options.Render = 0
                SaveOptions
                Call MsgBox("Could not initialize DirectX with software vertex processing.", vbCritical)
                Call DestroyGame
            End If
        Case Else ' auto
            If LoadDirectX(D3DCREATE_HARDWARE_VERTEXPROCESSING, frmMain.hwnd) <> 0 Then
                If LoadDirectX(D3DCREATE_MIXED_VERTEXPROCESSING, frmMain.hwnd) <> 0 Then
                    If LoadDirectX(D3DCREATE_SOFTWARE_VERTEXPROCESSING, frmMain.hwnd) <> 0 Then
                        Options.Render = 0
                        SaveOptions
                        Call MsgBox("Could not initialize DirectX.  DX8VB.dll may not be registered.", vbCritical)
                        Call DestroyGame
                    End If
                End If
            End If
    End Select
    TryCreateDirectX8Device = True


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "TryCreateDirectX8Device", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Function GetNearestPOT(Value As Long) As Long
Dim i As Long




   On Error GoTo errorhandler

    Do While 2 ^ i < Value
        i = i + 1
    Loop
    GetNearestPOT = 2 ^ i


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "GetNearestPOT", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
        
End Function
Public Sub LoadTexture1(ByRef TextureRec As DX8TextureRec)
Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, i As Long
Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long

   On Error GoTo errorhandler
    If TextureRec.filepath = "" Then
        Exit Sub
    End If
    If gTexture(TextureRec.Texture).HasData = False Then
        Set GDIToken = New cGDIpToken
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        Set SourceBitmap = New cGDIpImage
        Call SourceBitmap.LoadPicture_FileName(TextureRec.filepath, GDIToken)
            TextureRec.Width = SourceBitmap.Width
        TextureRec.Height = SourceBitmap.Height
            newWidth = GetNearestPOT(TextureRec.Width)
        newHeight = GetNearestPOT(TextureRec.Height)
        If newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height Then
            Set ConvertedBitmap = New cGDIpImage
            Set GDIGraphics = New cGDIpRenderer
            i = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
            If MapLayerImage.IsLoaded = False Then
                Call ConvertedBitmap.LoadPicture_FromNothing(256, 256, i, GDIToken) 'I HAVE NO IDEA why this is backwards but it works.
                Call ConvertedBitmap.SaveAsPNG(ImageData)
                MapLayerImage.ImageData = ImageData
                MapLayerImage.IsLoaded = True
                ConvertedBitmap.Clear
                LoadTexture1 TextureRec
                Exit Sub
            Else
                Call ConvertedBitmap.LoadPicture_FromNothing(newHeight, newWidth, i, GDIToken) 'I HAVE NO IDEA why this is backwards but it works.
            End If
            Call GDIGraphics.DestroyHGraphics(i)
            i = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
            Call GDIGraphics.AttachTokenClass(GDIToken)
            Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, i)
            Call ConvertedBitmap.SaveAsPNG(ImageData)
            GDIGraphics.DestroyHGraphics (i)
            gTexture(TextureRec.Texture).ImageData = ImageData
            ConvertedBitmap.Clear
            SourceBitmap.Clear
            Set ConvertedBitmap = Nothing
            Set GDIGraphics = Nothing
            Set SourceBitmap = Nothing
        Else
            Call SourceBitmap.SaveAsPNG(ImageData)
            gTexture(TextureRec.Texture).ImageData = ImageData
            Set SourceBitmap = Nothing
        End If
    Else
        ImageData = gTexture(TextureRec.Texture).ImageData
    End If
    Set gTexture(TextureRec.Texture).Texture = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, _
                                                    ImageData(0), _
                                                    UBound(ImageData) + 1, _
                                                    newWidth, _
                                                    newHeight, _
                                                    D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
    gTexture(TextureRec.Texture).TexWidth = newWidth
    gTexture(TextureRec.Texture).TexHeight = newHeight
    gTexture(TextureRec.Texture).TextureTimer = GetTickCount + 100000
    gTexture(TextureRec.Texture).IsLoaded = True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadTexture1", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub LoadTextures()
Dim i As Long

   On Error GoTo errorhandler

    NumTextures = 0

    Call CheckCharacters
    Call CheckGUIs
    Call CheckBodies
    Call CheckPaperdolls
    Call CheckTilesets
    Call CheckAnimations
    Call CheckItems
    Call CheckPics
    Call CheckFurniture
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    Call CheckCharFaces
    Call CheckFogs
    Call CheckProjectiles
    NumTextures = NumTextures + 12
    ReDim Preserve gTexture(NumTextures)
    Tex_Fade.filepath = App.path & GFX_PATH & "misc\fader.png"
    Tex_Fade.Texture = NumTextures - 11
    Tex_ChatBubble.filepath = App.path & GFX_PATH & "misc\chatbubble.png"
    Tex_ChatBubble.Texture = NumTextures - 10
    Tex_Weather.filepath = App.path & GFX_PATH & "misc\weather.png"
    Tex_Weather.Texture = NumTextures - 9
    Tex_White.filepath = App.path & GFX_PATH & "misc\white.png"
    Tex_White.Texture = NumTextures - 8
    Tex_Door.filepath = App.path & GFX_PATH & "misc\door.png"
    Tex_Door.Texture = NumTextures - 7
    Tex_Direction.filepath = App.path & GFX_PATH & "misc\direction.png"
    Tex_Direction.Texture = NumTextures - 6
    Tex_Target.filepath = App.path & GFX_PATH & "misc\target.png"
    Tex_Target.Texture = NumTextures - 5
    Tex_Misc.filepath = App.path & GFX_PATH & "misc\misc.png"
    Tex_Misc.Texture = NumTextures - 4
    Tex_Blood.filepath = App.path & GFX_PATH & "misc\blood.png"
    Tex_Blood.Texture = NumTextures - 3
    Tex_Bars.filepath = App.path & GFX_PATH & "misc\bars.png"
    Tex_Bars.Texture = NumTextures - 2
    Tex_Selection.filepath = App.path & GFX_PATH & "misc\select.png"
    Tex_Selection.Texture = NumTextures - 1
    Tex_Arrows.filepath = App.path & GFX_PATH & "misc\arrows.png"
    Tex_Arrows.Texture = NumTextures
    EngineInitFontTextures




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadTextures", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub UnloadTextures()
Dim i As Long
    On Error Resume Next
    For i = 1 To UBound(gTexture)
        Set gTexture(i).Texture = Nothing
        ZeroMemory ByVal VarPtr(gTexture(i)), LenB(gTexture(i))
    Next
    ReDim gTexture(1)

    For i = 1 To NumTileSets
        Tex_Tileset(i).Texture = 0
    Next

    For i = 1 To numitems
        Tex_Item(i).Texture = 0
    Next

    For i = 1 To NumCharacters
        Tex_Character(i).Texture = 0
    Next
    For i = 1 To NumPaperdolls
        Tex_Paperdoll(i).Texture = 0
    Next
    For i = 1 To NumResources
        Tex_Resource(i).Texture = 0
    Next
    For i = 1 To NumAnimations
        Tex_Animation(i).Texture = 0
    Next
    For i = 1 To NumSpellIcons
        Tex_SpellIcon(i).Texture = 0
    Next
    For i = 1 To NumFaces
        Tex_Face(i).Texture = 0
    Next
    Tex_Misc.Texture = 0
    Tex_Blood.Texture = 0
    Tex_Door.Texture = 0
    Tex_Direction.Texture = 0
    Tex_Target.Texture = 0
    Tex_Selection.Texture = 0
    UnloadFontTextures
    ClearMapCache



End Sub

' **************
' ** Drawing **
' **************
Public Sub RenderTexture(ByRef TextureRec As DX8TextureRec, ByVal dX As Single, ByVal dY As Single, ByVal sx As Single, ByVal sy As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional color As Long = -1, Optional GameScreen As Boolean = False, Optional StrictSDim As Boolean = False)
    Dim TextureNum As Long
    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single, rec As rect

   On Error GoTo errorhandler

    TextureNum = TextureRec.Texture
    If gTexture(TextureRec.Texture).IsLoaded = False Then
        LoadTexture1 TextureRec
    Else
        gTexture(TextureRec.Texture).TextureTimer = GetTickCount + 10000
    End If
    textureWidth = gTexture(TextureNum).TexWidth
    textureHeight = gTexture(TextureNum).TexHeight
    Dim X1 As Long, Y1 As Long, sW As Long, sH As Long, dw As Long, dH As Long, srcX As Long, srcY As Long
    Dim tOffsetX As Long, tOffsetY As Long
    If GameScreen Then
        'Lets do some trimming
        X1 = dX
        Y1 = dY
        srcX = sx
        srcY = sy
        dw = dWidth
        dH = dHeight
        sW = sWidth
        sH = sHeight
            If tOffsetX = 0 Then
            'Compares to game screen beginning
            If dX + dw < GameScreenBounds.Left Then
                dw = 0
            Else
                'Trimming
                If dX < GameScreenBounds.Left Then
                    tOffsetX = (GameScreenBounds.Left - dX)
                    srcX = srcX + tOffsetX
                    sW = sW - tOffsetX
                    dw = dw - tOffsetX
                End If
            End If
        End If
              If tOffsetY = 0 Then
            If dY + dH < GameScreenBounds.Top Then
                dH = 0
            Else
                'Trimming
                If dY < GameScreenBounds.Top Then
                    tOffsetY = (GameScreenBounds.Top - dY)
                    srcY = srcY + tOffsetY
                    sH = sH - tOffsetY
                    dH = dH - tOffsetY
                End If
            End If
        End If
            'Compares to Screen End
        If dX > GameScreenBounds.Right Then
            sW = 0
            dw = 0
        Else
            If dX + dw > GameScreenBounds.Right Then
                dw = dw + (GameScreenBounds.Right - (dX + dw))
                sW = sW + (GameScreenBounds.Right - (dX + sW))
            End If
        End If
            If dY > GameScreenBounds.Bottom Then
            sH = 0
            dH = 0
        Else
            If dY + dH > GameScreenBounds.Bottom Then
                dH = dH + (GameScreenBounds.Bottom - (dY + dH))
                sH = sH + (GameScreenBounds.Bottom - (dY + sH))
            End If
        End If
            'Do Stuff
        Y1 = dY + tOffsetY
        X1 = dX + tOffsetX
        tOffsetX = 0
        tOffsetY = 0
        dX = X1
        dY = Y1
                If tOffsetX = 0 Then
            'Compares to mapbeginning
            If dX + dw < ConvertMapX(TileView.Left * 32) Then
                dw = 0
            Else
                'Trimming
                If dX < ConvertMapX(TileView.Left * 32) Then
                    tOffsetX = (ConvertMapX(TileView.Left * 32) - dX)
                    srcX = srcX + tOffsetX
                    sW = sW - tOffsetX
                    dw = dw - tOffsetX
                End If
            End If
        End If
                If tOffsetY = 0 Then
            If dY + dH < ConvertMapY(TileView.Top * 32) Then
                dH = 0
            Else
                'Trimming
                If dY < ConvertMapY(TileView.Top * 32) Then
                    tOffsetY = (ConvertMapY(TileView.Top * 32) - dY)
                    srcY = srcY + tOffsetY
                    sH = sH - tOffsetY
                    dH = dH - tOffsetY
                End If
            End If
        End If
            dY = dY + tOffsetY
        dX = dX + tOffsetX
            'Compares to Map end
        If dX > ConvertMapX((TileView.Right + 1) * 32) Then
            sW = 0
            dw = 0
        Else
            If dX + dw > ConvertMapX((TileView.Right + 1) * 32) Then
                dw = dw + (ConvertMapX((TileView.Right + 1) * 32) - (dX + dw))
                sW = sW + (ConvertMapX((TileView.Right + 1) * 32) - (dX + sW))
            End If
        End If
            If dY > ConvertMapY((TileView.Bottom + 1) * 32) Then
            sH = 0
            dH = 0
        Else
            If dY + dH > ConvertMapY((TileView.Bottom + 1) * 32) Then
                dH = dH + (ConvertMapY((TileView.Bottom + 1) * 32) - (dY + dH))
                sH = sH + (ConvertMapY((TileView.Bottom + 1) * 32) - (dY + sH))
            End If
        End If
        
        Y1 = dY
        X1 = dX

            dX = X1
        dY = Y1
        dWidth = dw
        dHeight = dH
        If StrictSDim = False Then
            sx = srcX
            sy = srcY
            sWidth = sW
            sHeight = sH
        End If
    End If
    If sy + sHeight > textureHeight Then Exit Sub
    If sx + sWidth > textureWidth Then Exit Sub
    If sx < 0 Then Exit Sub
    If sy < 0 Then Exit Sub

    sx = sx - 0.5
    sy = sy - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sx / textureWidth)
    sourceY = (sy / textureHeight)
    sourceWidth = ((sx + sWidth) / textureWidth)
    sourceHeight = ((sy + sHeight) / textureHeight)
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    Direct3D_Device.SetTexture 0, gTexture(TextureNum).Texture
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "RenderTexture", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub RenderMapChunk(tex As Direct3DTexture8, ByVal dX As Single, ByVal dY As Single, ByVal sx As Single, ByVal sy As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional color As Long = -1, Optional GameScreen As Boolean = False, Optional StrictSDim As Boolean = False)
    Dim TextureNum As Long
    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single, rec As rect

   On Error GoTo errorhandler


    textureWidth = 256
    textureHeight = 256
    Dim X1 As Long, Y1 As Long, sW As Long, sH As Long, dw As Long, dH As Long, srcX As Long, srcY As Long
    Dim tOffsetX As Long, tOffsetY As Long
    If GameScreen Then
        'Lets do some trimming
        X1 = dX
        Y1 = dY
        srcX = sx
        srcY = sy
        dw = dWidth
        dH = dHeight
        sW = sWidth
        sH = sHeight
            If tOffsetX = 0 Then
            'Compares to game screen beginning
            If dX + dw < GameScreenBounds.Left Then
                dw = 0
            Else
                'Trimming
                If dX < GameScreenBounds.Left Then
                    tOffsetX = (GameScreenBounds.Left - dX)
                    srcX = srcX + tOffsetX
                    sW = sW - tOffsetX
                    dw = dw - tOffsetX
                End If
            End If
        End If
              If tOffsetY = 0 Then
            If dY + dH < GameScreenBounds.Top Then
                dH = 0
            Else
                'Trimming
                If dY < GameScreenBounds.Top Then
                    tOffsetY = (GameScreenBounds.Top - dY)
                    srcY = srcY + tOffsetY
                    sH = sH - tOffsetY
                    dH = dH - tOffsetY
                End If
            End If
        End If
            'Compares to Screen End
        If dX > GameScreenBounds.Right Then
            sW = 0
            dw = 0
        Else
            If dX + dw > GameScreenBounds.Right Then
                dw = dw + (GameScreenBounds.Right - (dX + dw))
                sW = sW + (GameScreenBounds.Right - (dX + sW))
            End If
        End If
            If dY > GameScreenBounds.Bottom Then
            sH = 0
            dH = 0
        Else
            If dY + dH > GameScreenBounds.Bottom Then
                dH = dH + (GameScreenBounds.Bottom - (dY + dH))
                sH = sH + (GameScreenBounds.Bottom - (dY + sH))
            End If
        End If
            'Do Stuff
        Y1 = dY + tOffsetY
        X1 = dX + tOffsetX
        tOffsetX = 0
        tOffsetY = 0
        dX = X1
        dY = Y1
                If tOffsetX = 0 Then
            'Compares to mapbeginning
            If dX + dw < ConvertMapX(TileView.Left * 32) Then
                dw = 0
            Else
                'Trimming
                If dX < ConvertMapX(TileView.Left * 32) Then
                    tOffsetX = (ConvertMapX(TileView.Left * 32) - dX)
                    srcX = srcX + tOffsetX
                    sW = sW - tOffsetX
                    dw = dw - tOffsetX
                End If
            End If
        End If
                If tOffsetY = 0 Then
            If dY + dH < ConvertMapY(TileView.Top * 32) Then
                dH = 0
            Else
                'Trimming
                If dY < ConvertMapY(TileView.Top * 32) Then
                    tOffsetY = (ConvertMapY(TileView.Top * 32) - dY)
                    srcY = srcY + tOffsetY
                    sH = sH - tOffsetY
                    dH = dH - tOffsetY
                End If
            End If
        End If
            dY = dY + tOffsetY
        dX = dX + tOffsetX
            'Compares to Map end
        If dX > ConvertMapX((TileView.Right + 1) * 32) Then
            sW = 0
            dw = 0
        Else
            If dX + dw > ConvertMapX((TileView.Right + 1) * 32) Then
                dw = dw + (ConvertMapX((TileView.Right + 1) * 32) - (dX + dw))
                sW = sW + (ConvertMapX((TileView.Right + 1) * 32) - (dX + sW))
            End If
        End If
            If dY > ConvertMapY((TileView.Bottom + 1) * 32) Then
            sH = 0
            dH = 0
        Else
            If dY + dH > ConvertMapY((TileView.Bottom + 1) * 32) Then
                dH = dH + (ConvertMapY((TileView.Bottom + 1) * 32) - (dY + dH))
                sH = sH + (ConvertMapY((TileView.Bottom + 1) * 32) - (dY + sH))
            End If
        End If
        
        Y1 = dY
        X1 = dX

            dX = X1
        dY = Y1
        dWidth = dw
        dHeight = dH
        If StrictSDim = False Then
            sx = srcX
            sy = srcY
            sWidth = sW
            sHeight = sH
        End If
    End If
    If sy + sHeight > textureHeight Then Exit Sub
    If sx + sWidth > textureWidth Then Exit Sub
    If sx < 0 Then Exit Sub
    If sy < 0 Then Exit Sub

    sx = sx - 0.5
    sy = sy - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sx / textureWidth)
    sourceY = (sy / textureHeight)
    sourceWidth = ((sx + sWidth) / textureWidth)
    sourceHeight = ((sy + sHeight) / textureHeight)
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    Direct3D_Device.SetTexture 0, tex
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "RenderTexture", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub RenderTextureByRects(TextureRec As DX8TextureRec, sRect As rect, dRect As rect)


   On Error GoTo errorhandler

    RenderTexture TextureRec, dRect.Left, dRect.Top, sRect.Left, sRect.Top, dRect.Right - dRect.Left, dRect.Bottom - dRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)





   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "RenderTextureByRects", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawDirection(ByVal X As Long, ByVal Y As Long)
Dim rec As rect
Dim i As Long

    ' render grid

   On Error GoTo errorhandler

    rec.Top = 24
    rec.Left = 0
    rec.Right = rec.Left + 32
    rec.Bottom = rec.Top + 32
    RenderTexture Tex_Direction, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255), True
    ' render dir blobs
    For i = 1 To 4
        rec.Left = (i - 1) * 8
        rec.Right = rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(X, Y).DirBlock, CByte(i)) Then
            rec.Top = 8
        Else
            rec.Top = 16
        End If
        rec.Bottom = rec.Top + 8
        'render!
        RenderTexture Tex_Direction, ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Next




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawDirection", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawTarget(ByVal X As Long, ByVal Y As Long)
Dim sRect As rect
Dim Width As Long, Height As Long


   On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRect
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    ' clipping
    If Y < 0 Then
        With sRect
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRect
            .Left = .Left - X
        End With
        X = 0
    End If
    RenderTexture Tex_Target, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawTarget", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawHover(ByVal tType As Long, ByVal target As Long, ByVal X As Long, ByVal Y As Long)
Dim sRect As rect
Dim Width As Long, Height As Long


   On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRect
        .Top = 0
        .Bottom = Height
        .Left = Width
        .Right = .Left + Width
    End With
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    ' clipping
    If Y < 0 Then
        With sRect
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRect
            .Left = .Left - X
        End With
        X = 0
    End If
    RenderTexture Tex_Target, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawHover", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawMapTile(ByVal X As Long, ByVal Y As Long)
Dim rec As rect
Dim i As Long
Dim tOffsetX As Long, tOffsetY As Long, sx As Long, sy As Long, w As Long, h As Long, X1 As Long, Y1 As Long


   On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Ground To MapLayer.Mask2
            If .Layer(i).Tileset > 0 Then
                If Autotile(X, Y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                    'RenderTexture Tex_Tileset(.Layer(i).Tileset), x1, y1, .Layer(i).x * 32 + sx, .Layer(i).y * 32 + sy, w, h, w, h, -1
                    RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32, -1, True
                ElseIf Autotile(X, Y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                    ' Draw autotiles
                    DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                    DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                    DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                    DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
                End If
            End If
        Next
    End With
    
    With Map.exTile(X, Y)
        For i = ExMapLayer.Mask3 To ExMapLayer.Mask5
            If .Layer(i).Tileset > 0 Then
                If Autotile(X, Y).ExLayer(i).renderState = RENDER_STATE_NORMAL Then
                    'RenderTexture Tex_Tileset(.Layer(i).Tileset), x1, y1, .Layer(i).x * 32 + sx, .Layer(i).y * 32 + sy, w, h, w, h, -1
                    RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32, -1, True
                ElseIf Autotile(X, Y).ExLayer(i).renderState = RENDER_STATE_AUTOTILE Then
                    ' Draw autotiles
                    DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y, 0, True, True
                    DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y, 0, True, True
                    DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y, 0, True, True
                    DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y, 0, True, True
                End If
            End If
        Next
    End With



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawMapTile", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawMapFringeTile(ByVal X As Long, ByVal Y As Long)
Dim rec As rect
Dim i As Long



   On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            If .Layer(i).Tileset > 0 Then
                If Autotile(X, Y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                    ' Draw normally
                    RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32, -1, True
                ElseIf Autotile(X, Y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                    ' Draw autotiles
                    DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                    DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                    DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                    DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
                End If
            End If
        Next
    End With
    
    With Map.exTile(X, Y)
        For i = ExMapLayer.Fringe3 To ExMapLayer.Fringe5
            If .Layer(i).Tileset > 0 Then
                If Autotile(X, Y).ExLayer(i).renderState = RENDER_STATE_NORMAL Then
                    'RenderTexture Tex_Tileset(.Layer(i).Tileset), x1, y1, .Layer(i).x * 32 + sx, .Layer(i).y * 32 + sy, w, h, w, h, -1
                    RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32, -1, True
                ElseIf Autotile(X, Y).ExLayer(i).renderState = RENDER_STATE_AUTOTILE Then
                    ' Draw autotiles
                    DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y, 0, True, True
                    DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y, 0, True, True
                    DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y, 0, True, True
                    DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y, 0, True, True
                End If
            End If
        Next
    End With




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawMapFringeTile", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawDoor(ByVal X As Long, ByVal Y As Long)
Dim rec As rect
Dim X2 As Long, Y2 As Long

    ' sort out animation

   On Error GoTo errorhandler

    With TempTile(X, Y)
        If .DoorAnimate = 1 Then ' opening
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame < 4 Then
                    .DoorFrame = .DoorFrame + 1
                Else
                    .DoorAnimate = 2 ' set to closing
                End If
                .DoorTimer = GetTickCount
            End If
        ElseIf .DoorAnimate = 2 Then ' closing
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame > 1 Then
                    .DoorFrame = .DoorFrame - 1
                Else
                    .DoorAnimate = 0 ' end animation
                End If
                .DoorTimer = GetTickCount
            End If
        End If
            If .DoorFrame = 0 Then .DoorFrame = 1
    End With

    With rec
        .Top = 0
        .Bottom = Tex_Door.Height
        .Left = ((TempTile(X, Y).DoorFrame - 1) * (Tex_Door.Width / 4))
        .Right = .Left + (Tex_Door.Width / 4)
    End With

    X2 = (X * PIC_X)
    Y2 = (Y * PIC_Y) - (Tex_Door.Height / 2) + 4
    RenderTexture Tex_Door, ConvertMapX(X2), ConvertMapY(Y2), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawDoor", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawBlood(ByVal Index As Long)
Dim rec As rect
    'load blood then

   On Error GoTo errorhandler

    BloodCount = Tex_Blood.Width / 32
    With Blood(Index)
        ' check if we should be seeing it
        If .Timer + 20000 < GetTickCount Then Exit Sub
            rec.Top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        RenderTexture Tex_Blood, ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255), True
    End With




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawBlood", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
Dim Sprite As Long
Dim sRect As rect
Dim dRect As rect
Dim i As Long
Dim Width As Long, Height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim X As Long, Y As Long
Dim lockindex As Long


   On Error GoTo errorhandler

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    ' total width divided by frame count
    Width = Tex_Animation(Sprite).Width / FrameCount
    Height = Tex_Animation(Sprite).Height
    sRect.Top = 0
    sRect.Bottom = Height
    sRect.Left = (AnimInstance(Index).frameIndex(Layer) - 1) * Width
    sRect.Right = sRect.Left + Width
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).XOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).YOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).Num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).XOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).YOffset
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
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_ZONENPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If ZoneNPC(AnimInstance(Index).LockZone).Npc(lockindex).Num > 0 Then
                ' check if alive
                If ZoneNPC(AnimInstance(Index).LockZone).Npc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (ZoneNPC(AnimInstance(Index).LockZone).Npc(lockindex).X * PIC_X) + 16 - (Width / 2) + ZoneNPC(AnimInstance(Index).LockZone).Npc(lockindex).XOffset
                    Y = (ZoneNPC(AnimInstance(Index).LockZone).Npc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + ZoneNPC(AnimInstance(Index).LockZone).Npc(lockindex).YOffset
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
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_PET Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) And Player(lockindex).Pet.Alive = True Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (Player(lockindex).Pet.X * PIC_X) + 16 - (Width / 2) + Player(lockindex).Pet.XOffset
                    Y = (Player(lockindex).Pet.Y * PIC_Y) + 16 - (Height / 2) + Player(lockindex).Pet.YOffset
                End If
            End If
        End If
    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    ' Clip to screen
    If Y < 0 Then

        With sRect
            .Top = .Top - Y
        End With

        Y = 0
    End If

    If X < 0 Then

        With sRect
            .Left = .Left - X
        End With

        X = 0
    End If
    RenderTexture Tex_Animation(Sprite), X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawAnimation", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawItem(ByVal ItemNum As Long)
Dim PicNum As Long
Dim rec As rect
Dim MaxFrames As Byte

    ' if it's not us then don't render

   On Error GoTo errorhandler

    If MapItem(ItemNum).playerName <> vbNullString Then
        If MapItem(ItemNum).playerName <> Trim$(GetPlayerName(MyIndex)) Then Exit Sub
    End If
    ' get the picture
    PicNum = Item(MapItem(ItemNum).Num).pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

    If Tex_Item(PicNum).Width > 64 Then ' has more than 1 frame
        With rec
            .Top = 0
            .Bottom = 32
            .Left = (MapItem(ItemNum).Frame * 32)
            .Right = .Left + 32
        End With
    Else
        With rec
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If
    RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(ItemNum).X * PIC_X), ConvertMapY(MapItem(ItemNum).Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255), True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawItem", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawFurniture(ByVal Index As Long, Layer As Long)
Dim i As Long, ItemNum As Long
Dim rec As rect
Dim X As Long, Y As Long, Width As Long, Height As Long, X1 As Long, Y1 As Long

   On Error GoTo errorhandler

    ItemNum = Furniture(Index).ItemNum
    If Item(ItemNum).type <> ITEM_TYPE_FURNITURE Then Exit Sub
    i = Item(ItemNum).data2
    Width = Item(ItemNum).FurnitureWidth
    Height = Item(ItemNum).FurnitureHeight
    If Width > 4 Then Width = 4
    If Height > 4 Then Height = 4
    If i <= 0 Or i > NumFurniture Then Exit Sub

    ' make sure it's not out of map
    If Furniture(Index).X > Map.MaxX Then Exit Sub
    If Furniture(Index).Y > Map.MaxY Then Exit Sub
    For X1 = 0 To Width - 1
        For Y1 = 0 To Height
            If Item(Furniture(Index).ItemNum).FurnitureFringe(X1, Y1) = Layer Then
                ' Set base x + y, then the offset due to size
                X = (Furniture(Index).X * 32) + (X1 * 32)
                Y = (Furniture(Index).Y * 32 - (Height * 32)) + (Y1 * 32)
                            X = ConvertMapX(X)
                Y = ConvertMapY(Y)
                            RenderTexture Tex_Furniture(i), X, Y, 32 * X1, Y1 * 32, 32, 32, 32, 32, -1, True
            End If
        Next
    Next




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawFurniture", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawMapResource(ByVal Resource_num As Long, Optional ByVal screenShot As Boolean = False)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As rect
Dim X As Long, Y As Long

    ' make sure it's not out of map

   On Error GoTo errorhandler

    If MapResource(Resource_num).X > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.MaxY Then Exit Sub
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).Data1
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' src rect
    With rec
        .Top = 0
        .Bottom = Tex_Resource(Resource_sprite).Height
        .Left = 0
        .Right = Tex_Resource(Resource_sprite).Width
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (Tex_Resource(Resource_sprite).Width / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - Tex_Resource(Resource_sprite).Height + 32
    ' render it
    If Not screenShot Then
        Call DrawResource(Resource_sprite, X, Y, rec)
    Else
        Call ScreenshotResource(Resource_sprite, X, Y, rec)
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawMapResource", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub DrawResource(ByVal Resource As Long, ByVal dX As Long, dY As Long, rec As rect)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long
Dim destRect As rect



   On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub

    X = ConvertMapX(dX)
    Y = ConvertMapY(dY)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    RenderTexture Tex_Resource(Resource), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255), True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawResource", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub ScreenshotResource(ByVal Resource As Long, ByVal X As Long, Y As Long, rec As rect)
Dim Width As Long
Dim Height As Long
Dim destRect As rect



   On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If
    RenderTexture Tex_Resource(Resource), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ScreenshotResource", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub DrawBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRect As rect
Dim barWidth As Long
Dim i As Long, npcNum As Long, partyIndex As Long, X As Long


   On Error GoTo errorhandler

    If gTexture(Tex_Bars.Texture).IsLoaded = False Then LoadTexture1 Tex_Bars
    ' dynamic bar calculations
    sWidth = Tex_Bars.Width
    sHeight = Tex_Bars.Height / 4
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(i).Num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(npcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).X * PIC_X + MapNpc(i).XOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(i).Y * PIC_Y + MapNpc(i).YOffset + 35
                            ' calculate the width to fill
                barWidth = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (Npc(npcNum).HP / sWidth)) * sWidth
                            ' draw bar background
                With sRect
                    .Top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                            ' draw the bar proper
                With sRect
                    .Top = 0 ' HP bar
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
            End If
        End If
    Next
    ' render health bars
    For i = 1 To MAX_ZONES
        For X = 1 To MAX_MAP_NPCS * 2
            npcNum = ZoneNPC(i).Npc(X).Num
            ' exists?
            If npcNum > 0 Then
                ' alive?
                If ZoneNPC(i).Npc(X).Vital(Vitals.HP) > 0 And ZoneNPC(i).Npc(X).Vital(Vitals.HP) < Npc(npcNum).HP Then
                    ' lock to npc
                    tmpX = ZoneNPC(i).Npc(X).X * PIC_X + ZoneNPC(i).Npc(X).XOffset + 16 - (sWidth / 2)
                    tmpY = ZoneNPC(i).Npc(X).Y * PIC_Y + ZoneNPC(i).Npc(X).YOffset + 35
                                    ' calculate the width to fill
                    barWidth = ((ZoneNPC(i).Npc(X).Vital(Vitals.HP) / sWidth) / (Npc(npcNum).HP / sWidth)) * sWidth
                                    ' draw bar background
                    With sRect
                        .Top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                                    ' draw the bar proper
                    With sRect
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                End If
            End If
        Next
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffset + 35 + sHeight + 1
                    ' calculate the width to fill
            barWidth = (GetTickCount - SpellBufferTimer) / ((spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth
                    ' draw bar background
            With sRect
                .Top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                    ' draw the bar proper
            With sRect
                .Top = sHeight * 2 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
        End If
    End If
   ' check for pet casting time bar
    If PetSpellBuffer > 0 Then
        If spell(Pet(Player(MyIndex).Pet.Num).spell(PetSpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = Player(MyIndex).Pet.X * PIC_X + Player(MyIndex).Pet.XOffset + 16 - (sWidth / 2)
            tmpY = Player(MyIndex).Pet.Y * PIC_Y + Player(MyIndex).Pet.YOffset + 35 + sHeight + 1
            
            ' calculate the width to fill
            barWidth = (GetTickCount - PetSpellBufferTimer) / ((spell(Pet(Player(MyIndex).Pet.Num).spell(PetSpellBuffer)).CastTime * 1000)) * sWidth
            
            ' draw bar background
            With sRect
                .Top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
            
            ' draw the bar proper
            With sRect
                .Top = sHeight * 2 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
        End If
    End If
    For i = 1 To MAX_PLAYERS
        If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            ' draw own health bar
            If GetPlayerVital(i, Vitals.HP) > 0 And GetPlayerVital(i, Vitals.HP) < GetPlayerMaxVital(i, Vitals.HP) Then
                ' lock to Player
                tmpX = GetPlayerX(i) * PIC_X + Player(i).XOffset + 16 - (sWidth / 2)
                tmpY = GetPlayerY(i) * PIC_X + Player(i).YOffset + 35
                   ' calculate the width to fill
                barWidth = ((GetPlayerVital(i, Vitals.HP) / sWidth) / (GetPlayerMaxVital(i, Vitals.HP) / sWidth)) * sWidth
                   ' draw bar background
                With sRect
                    .Top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                   ' draw the bar proper
                With sRect
                    .Top = 0 ' HP bar
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
            End If
            If Player(i).Pet.Alive Then
                ' draw own health bar
                If Player(i).Pet.Health > 0 And Player(i).Pet.Health < Player(i).Pet.MaxHp Then
                    ' lock to Player
                    tmpX = Player(i).Pet.X * PIC_X + Player(i).Pet.XOffset + 16 - (sWidth / 2)
                    tmpY = Player(i).Pet.Y * PIC_X + Player(i).Pet.YOffset + 35
                       ' calculate the width to fill
                    barWidth = ((Player(i).Pet.Health / sWidth) / (Player(i).Pet.MaxHp / sWidth)) * sWidth
                       ' draw bar background
                    With sRect
                        .Top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                       ' draw the bar proper
                    With sRect
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                End If
            End If
        End If
    Next
    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).XOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).YOffset + 35
                                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                                    ' draw bar background
                    With sRect
                        .Top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                                    ' draw the bar proper
                    With sRect
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                End If
            End If
        Next
    End If
                



   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawBars", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawPlayer(ByVal Index As Long)
Dim Anim As Byte, i As Long, X As Long, Y As Long, z As Long, X1 As Long, Y1 As Long
Dim Sprite As Long, spritetop As Long
Dim rec As rect
Dim attackspeed As Long, TextureNum As DX8TextureRec, drawnshield As Boolean

   On Error GoTo errorhandler

    If Player(MyIndex).InHouse > 0 Then
        If Player(Index).InHouse <> Player(MyIndex).InHouse Then Exit Sub
    End If
    If CharMode = 1 Then
        For i = 1 To SpriteEnum.Sprite_Count - 1
            Sprite = Player(Index).Sprite(i)
            Select Case Player(Index).Sex
                Case SEX_MALE
                    Select Case i
                        Case SpriteEnum.Body
                            TextureNum = Tex_CharBodies(Player(Index).Sprite(SpriteEnum.Body))
                        Case SpriteEnum.Hair
                            TextureNum = Tex_CharHair(Player(Index).Sprite(SpriteEnum.Hair))
                        Case SpriteEnum.Pants
                            TextureNum = Tex_MaleLegs(Player(Index).Sprite(SpriteEnum.Pants))
                        Case SpriteEnum.Shirt
                            TextureNum = Tex_CharShirts(Player(Index).Sprite(SpriteEnum.Shirt))
                        Case SpriteEnum.Shoes
                            TextureNum = Tex_MaleShoes(Player(Index).Sprite(SpriteEnum.Shoes))
                    End Select
                    If TextureNum.Texture > 0 Then z = 1
                Case SEX_FEMALE
                    Select Case i
                        Case SpriteEnum.Body
                            TextureNum = Tex_CharBodies(Player(Index).Sprite(SpriteEnum.Body))
                        Case SpriteEnum.Hair
                            TextureNum = Tex_CharHair(Player(Index).Sprite(SpriteEnum.Hair))
                        Case SpriteEnum.Pants
                            TextureNum = Tex_FemaleLegs(Player(Index).Sprite(SpriteEnum.Pants))
                        Case SpriteEnum.Shirt
                            TextureNum = Tex_CharShirts(Player(Index).Sprite(SpriteEnum.Shirt))
                        Case SpriteEnum.Shoes
                            TextureNum = Tex_FemaleShoes(Player(Index).Sprite(SpriteEnum.Shoes))
                    End Select
                                If TextureNum.Texture > 0 Then z = 1
            End Select
                If i = SpriteEnum.Hair Then
                If GetPlayerEquipment(Index, Helmet) > 0 Then
                    If Item(Player(Index).Equipment(Equipment.Helmet)).Paperdoll > 0 Then
                        z = 0
                    End If
                End If
            End If
                If z > 0 Then
                        ' speed from weapon
                If GetPlayerEquipment(Index, Weapon) > 0 Then
                    attackspeed = Item(GetPlayerEquipment(Index, Weapon)).speed
                Else
                    attackspeed = 1000
                End If
                    ' Reset frame
                Anim = 0
                ' Check for attacking animation
                If Player(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
                    If Player(Index).Attacking = 1 Then
                        Anim = 3
                    End If
                Else
                    ' If not attacking, walk normally
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            If (Player(Index).YOffset > 8) Then Anim = Player(Index).Step
                        Case DIR_DOWN
                            If (Player(Index).YOffset < -8) Then Anim = Player(Index).Step
                        Case DIR_LEFT
                            If (Player(Index).XOffset > 8) Then Anim = Player(Index).Step
                        Case DIR_RIGHT
                            If (Player(Index).XOffset < -8) Then Anim = Player(Index).Step
                    End Select
                End If
                    ' Check to see if we want to stop making him attack
                With Player(Index)
                    If .AttackTimer + attackspeed < GetTickCount Then
                        .Attacking = 0
                        .AttackTimer = 0
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
                    With rec
                    .Top = spritetop * (TextureNum.Height / 4)
                    .Bottom = (TextureNum.Height / 4)
                    .Left = Anim * (TextureNum.Width / 4)
                    .Right = (TextureNum.Width / 4)
                End With
                    ' Calculate the X
                X = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((TextureNum.Width / 4 - 32) / 2)
                    ' Is the player's height more than 32..?
                If (TextureNum.Height) > 32 Then
                    ' Create a 32 pixel offset for larger sprites
                    Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((TextureNum.Height / 4) - 32) - 4
                Else
                    ' Proceed as normal
                    Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 4
                End If
                        If GetPlayerEquipment(Index, Equipment.Shield) > 0 Then
                    If Item(GetPlayerEquipment(Index, Equipment.Shield)).Paperdoll > 0 Then
                        If Player(Index).dir = DIR_UP Or Player(Index).dir = DIR_LEFT Then
                            If drawnshield = False Then
                                Call DrawPaperdoll(ConvertMapX(X), ConvertMapY(Y), Item(GetPlayerEquipment(Index, Equipment.Shield)).Paperdoll, Anim, spritetop)
                                drawnshield = True
                            End If
                        End If
                    End If
                End If
                        Select Case Player(Index).Sex
                    Case SEX_MALE
                        Select Case i
                            Case SpriteEnum.Body
                                RenderTexture Tex_CharBodies(Player(Index).Sprite(SpriteEnum.Body)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
                            Case SpriteEnum.Hair
                                RenderTexture Tex_CharHair(Player(Index).Sprite(SpriteEnum.Hair)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
                            Case SpriteEnum.Pants
                                RenderTexture Tex_MaleLegs(Player(Index).Sprite(SpriteEnum.Pants)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
                            Case SpriteEnum.Shirt
                                RenderTexture Tex_CharShirts(Player(Index).Sprite(SpriteEnum.Shirt)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
                            Case SpriteEnum.Shoes
                                RenderTexture Tex_MaleShoes(Player(Index).Sprite(SpriteEnum.Shoes)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
                        End Select
                        If TextureNum.Texture > 0 Then z = 1
                    Case SEX_FEMALE
                        Select Case i
                            Case SpriteEnum.Body
                                RenderTexture Tex_CharBodies(Player(Index).Sprite(SpriteEnum.Body)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
                            Case SpriteEnum.Hair
                                RenderTexture Tex_CharHair(Player(Index).Sprite(SpriteEnum.Hair)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
                            Case SpriteEnum.Pants
                                RenderTexture Tex_FemaleLegs(Player(Index).Sprite(SpriteEnum.Pants)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
                            Case SpriteEnum.Shirt
                                RenderTexture Tex_CharShirts(Player(Index).Sprite(SpriteEnum.Shirt)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
                            Case SpriteEnum.Shoes
                                RenderTexture Tex_FemaleShoes(Player(Index).Sprite(SpriteEnum.Shoes)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
                        End Select
                        If TextureNum.Texture > 0 Then z = 1
                End Select
            End If
            z = 0
        Next
    Else
        Sprite = Player(Index).Face(FaceEnum.Head)
        If Player(Index).Face(FaceEnum.Head) > NumCharacters Or Player(Index).Face(FaceEnum.Head) <= 0 Then TextureNum.Texture = 0 Else TextureNum = Tex_Character(Player(Index).Face(FaceEnum.Head))
        If TextureNum.Texture > 0 Then z = 1
        If z > 0 Then
            ' speed from weapon
            If GetPlayerEquipment(Index, Weapon) > 0 Then
                attackspeed = Item(GetPlayerEquipment(Index, Weapon)).speed
            Else
                attackspeed = 1000
            End If
                    
            ' Reset frame
            Anim = 0
            ' Check for attacking animation
            If Player(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
                If Player(Index).Attacking = 1 Then
                    Anim = 3
                End If
            Else
                ' If not attacking, walk normally
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        If (Player(Index).YOffset > 8) Then Anim = Player(Index).Step
                    Case DIR_DOWN
                        If (Player(Index).YOffset < -8) Then Anim = Player(Index).Step
                    Case DIR_LEFT
                        If (Player(Index).XOffset > 8) Then Anim = Player(Index).Step
                    Case DIR_RIGHT
                        If (Player(Index).XOffset < -8) Then Anim = Player(Index).Step
                End Select
            End If
            
            ' Check to see if we want to stop making him attack
            With Player(Index)
                If .AttackTimer + attackspeed < GetTickCount Then
                    .Attacking = 0
                    .AttackTimer = 0
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
                    
            With rec
                .Top = spritetop * (TextureNum.Height / 4)
                .Bottom = (TextureNum.Height / 4)
                .Left = Anim * (TextureNum.Width / 4)
                .Right = (TextureNum.Width / 4)
            End With
            
            ' Calculate the X
            X = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((TextureNum.Width / 4 - 32) / 2)
                ' Is the player's height more than 32..?
            If (TextureNum.Height) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((TextureNum.Height / 4) - 32) - 4
            Else
                ' Proceed as normal
                Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 4
            End If
            
            If GetPlayerEquipment(Index, Equipment.Shield) > 0 Then
                If Item(GetPlayerEquipment(Index, Equipment.Shield)).Paperdoll > 0 Then
                    If Player(Index).dir = DIR_UP Or Player(Index).dir = DIR_LEFT Then
                        If drawnshield = False Then
                            Call DrawPaperdoll(ConvertMapX(X), ConvertMapY(Y), Item(GetPlayerEquipment(Index, Equipment.Shield)).Paperdoll, Anim, spritetop)
                            drawnshield = True
                        End If
                    End If
                End If
            End If
                
            RenderTexture Tex_Character(Player(Index).Face(FaceEnum.Head)), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Right, rec.Bottom, rec.Right, rec.Bottom, -1, True
        End If
        z = 0
    End If
    
    ' check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll <= NumPaperdolls Then
                    TextureNum = Tex_Paperdoll(Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll)
                    ' Calculate the X
                    X = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((TextureNum.Width / 4 - 32) / 2)
                    ' Is the player's height more than 32..?
                    If (TextureNum.Height) > 32 Then
                        ' Create a 32 pixel offset for larger sprites
                        Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((TextureNum.Height / 4) - 32) - 4
                    Else
                        ' Proceed as normal
                        Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 4
                    End If
                    If PaperdollOrder(i) = Equipment.Shield Then
                        If Player(Index).dir = DIR_DOWN Or Player(Index).dir = DIR_RIGHT Then
                            Call DrawPaperdoll(ConvertMapX(X), ConvertMapY(Y), Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
                        End If
                    Else
                        If Anim = 3 Then
                            If Player(Index).Attacking = 0 Then
                                If PaperdollOrder(i) = Equipment.Weapon Then
                                    Call DrawPaperdoll(ConvertMapX(X), ConvertMapY(Y), Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, 1, spritetop)
                                Else
                                    Call DrawPaperdoll(ConvertMapX(X), ConvertMapY(Y), Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
                                End If
                            Else
                                Call DrawPaperdoll(ConvertMapX(X), ConvertMapY(Y), Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
                            End If
                        Else
                            Call DrawPaperdoll(ConvertMapX(X), ConvertMapY(Y), Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
                        End If
                    End If
                End If
            End If
        End If
    Next

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawPlayer", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte, i As Long, X As Long, Y As Long, Sprite As Long, spritetop As Long
Dim rec As rect
Dim attackspeed As Long


   On Error GoTo errorhandler

    If MapNpc(MapNpcNum).Num = 0 Then Exit Sub ' no npc set
    Sprite = Npc(MapNpc(MapNpcNum).Num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    attackspeed = 1000

    ' Reset frame
    Anim = 0
    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If MapNpc(MapNpcNum).Attacking = 1 Then
            Anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
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

    With rec
        .Top = (Tex_Character(Sprite).Height / 4) * spritetop
        .Bottom = .Top + Tex_Character(Sprite).Height / 4
        .Left = Anim * (Tex_Character(Sprite).Width / 4)
        .Right = .Left + (Tex_Character(Sprite).Width / 4)
    End With

    ' Calculate the X
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset
    End If

    Call DrawSprite(Sprite, X, Y, rec)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawNpc", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawZoneNpc(zonenum As Long, npcNum As Long)
Dim Anim As Byte, i As Long, X As Long, Y As Long, Sprite As Long, spritetop As Long
Dim rec As rect
Dim attackspeed As Long


   On Error GoTo errorhandler

    If ZoneNPC(zonenum).Npc(npcNum).Num = 0 Then Exit Sub ' no npc set
    If ZoneNPC(zonenum).Npc(npcNum).Map <> GetPlayerMap(MyIndex) Then Exit Sub
    Sprite = Npc(ZoneNPC(zonenum).Npc(npcNum).Num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    attackspeed = 1000

    ' Reset frame
    Anim = 0
    ' Check for attacking animation
    If ZoneNPC(zonenum).Npc(npcNum).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If ZoneNPC(zonenum).Npc(npcNum).Attacking = 1 Then
            Anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case ZoneNPC(zonenum).Npc(npcNum).dir
            Case DIR_UP
                If (ZoneNPC(zonenum).Npc(npcNum).YOffset > 8) Then Anim = ZoneNPC(zonenum).Npc(npcNum).Step
            Case DIR_DOWN
                If (ZoneNPC(zonenum).Npc(npcNum).YOffset < -8) Then Anim = ZoneNPC(zonenum).Npc(npcNum).Step
            Case DIR_LEFT
                If (ZoneNPC(zonenum).Npc(npcNum).XOffset > 8) Then Anim = ZoneNPC(zonenum).Npc(npcNum).Step
            Case DIR_RIGHT
                If (ZoneNPC(zonenum).Npc(npcNum).XOffset < -8) Then Anim = ZoneNPC(zonenum).Npc(npcNum).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With ZoneNPC(zonenum).Npc(npcNum)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case ZoneNPC(zonenum).Npc(npcNum).dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (Tex_Character(Sprite).Height / 4) * spritetop
        .Bottom = .Top + Tex_Character(Sprite).Height / 4
        .Left = Anim * (Tex_Character(Sprite).Width / 4)
        .Right = .Left + (Tex_Character(Sprite).Width / 4)
    End With

    ' Calculate the X
    X = ZoneNPC(zonenum).Npc(npcNum).X * PIC_X + ZoneNPC(zonenum).Npc(npcNum).XOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = ZoneNPC(zonenum).Npc(npcNum).Y * PIC_Y + ZoneNPC(zonenum).Npc(npcNum).YOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = ZoneNPC(zonenum).Npc(npcNum).Y * PIC_Y + ZoneNPC(zonenum).Npc(npcNum).YOffset
    End If

    Call DrawSprite(Sprite, X, Y, rec)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawZoneNpc", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawPaperdoll(ByVal X2 As Long, ByVal Y2 As Long, ByVal Sprite As Long, ByVal Anim As Long, ByVal spritetop As Long)
Dim rec As rect
Dim X As Long, Y As Long
Dim Width As Long, Height As Long


   On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    With rec
        .Top = spritetop * (Tex_Paperdoll(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Paperdoll(Sprite).Height / 4)
        .Left = Anim * (Tex_Paperdoll(Sprite).Width / 4)
        .Right = .Left + (Tex_Paperdoll(Sprite).Width / 4)
    End With
    ' clipping
    'x = ConvertMapX(x2)
    'y = ConvertMapY(y2)
    X = X2
    Y = Y2
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' Clip to screen
    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If
    RenderTexture Tex_Paperdoll(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255), True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Private Sub DrawSprite(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, rec As rect)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long



   On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    RenderTexture Tex_Character(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255), True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawSprite", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawFog()
Dim fogNum As Long, color As Long, X As Long, Y As Long, renderState As Long


   On Error GoTo errorhandler

    fogNum = CurrentFog
    If fogNum <= 0 Or fogNum > NumFogs Then Exit Sub
    color = D3DColorRGBA(255, 255, 255, 255 - CurrentFogOpacity)

    renderState = 0
    ' render state
    Select Case renderState
        Case 1 ' Additive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    For X = 0 To ((Map.MaxX * 32) / 256) + 1
        For Y = 0 To ((Map.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), ConvertMapX((X * 256) + fogOffsetX), ConvertMapY((Y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, color, True
        Next
    Next
    ' reset render state
    If renderState > 0 Then
        Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawFog", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawTint()
Dim color As Long

   On Error GoTo errorhandler

    color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    RenderTexture Tex_White, ConvertMapX(TileView.Left * 32), ConvertMapY(TileView.Top * 32), 0, 0, (TileView.Right - TileView.Left + 1) * 32, (TileView.Bottom - TileView.Top + 1) * 32, 32, 32, color, True, True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawTint", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawWeather()
Dim color As Long, i As Long, SpriteLeft As Long


   On Error GoTo errorhandler

    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(i).type - 1
            End If
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(i).X), ConvertMapY(WeatherParticle(i).Y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1, True
        End If
    Next




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawWeather", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_DrawTileset()
Dim Height As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Width As Long
Dim Tileset As Long
Dim sRect As rect
Dim dRect As rect, scrlX As Long, scrlY As Long

    ' find tileset number

   On Error GoTo errorhandler

    Tileset = frmEditor_Map.scrlTileSet.Value
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    scrlX = frmEditor_Map.scrlPictureX.Value * PIC_X
    scrlY = frmEditor_Map.scrlPictureY.Value * PIC_Y
    Height = Tex_Tileset(Tileset).Height - scrlY
    Width = Tex_Tileset(Tileset).Width - scrlX
    sRect.Left = frmEditor_Map.scrlPictureX.Value * PIC_X
    sRect.Top = frmEditor_Map.scrlPictureY.Value * PIC_Y
    sRect.Right = sRect.Left + Width
    sRect.Bottom = sRect.Top + Height
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    RenderTextureByRects Tex_Tileset(Tileset), sRect, dRect
    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.Value > 0 Then
        Select Case frmEditor_Map.scrlAutotile.Value
            Case 1 ' autotile
                EditorTileWidth = 2
                EditorTileHeight = 3
            Case 2 ' fake autotile
                EditorTileWidth = 1
                EditorTileHeight = 1
            Case 3 ' animated
                EditorTileWidth = 6
                EditorTileHeight = 3
            Case 4 ' cliff
                EditorTileWidth = 2
                EditorTileHeight = 2
            Case 5 ' waterfall
                EditorTileWidth = 2
                EditorTileHeight = 3
        End Select
    End If
    With destRect
        .X1 = (EditorTileX * 32) - sRect.Left
        .X2 = (EditorTileWidth * 32) + .X1
        .Y1 = (EditorTileY * 32) - sRect.Top
        .Y2 = (EditorTileHeight * 32) + .Y1
    End With
    DrawSelectionBox destRect
        With srcRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    With destRect
        .X1 = 0
        .X2 = frmEditor_Map.picBack.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Map.picBack.ScaleHeight
    End With
    'Now render the selection tiles and we are done!
                    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picBack.hwnd, ByVal (0)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorMap_DrawTileset", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Sub DrawSelectionBox(dRect As D3DRECT)
Dim Width As Long, Height As Long, X As Long, Y As Long


   On Error GoTo errorhandler

    Width = dRect.X2 - dRect.X1
    Height = dRect.Y2 - dRect.Y1
    X = dRect.X1
    Y = dRect.Y1
    If Width > 6 And Height > 6 Then
        'Draw Box 32 by 32 at graphicselx and graphicsely
        RenderTexture Tex_Selection, X, Y, 1, 1, 2, 2, 2, 2, -1 'top left corner
        RenderTexture Tex_Selection, X + 2, Y, 3, 1, Width - 4, 2, 32 - 6, 2, -1 'top line
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y, 29, 1, 2, 2, 2, 2, -1 'top right corner
        RenderTexture Tex_Selection, X, Y + 2, 1, 3, 2, Height - 4, 2, 32 - 6, -1 'Left Line
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2, 32 - 3, 3, 2, Height - 4, 2, 32 - 6, -1 'right line
        RenderTexture Tex_Selection, X, Y + 2 + (Height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 'bottom left corner
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2 + (Height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 'bottom right corner
        RenderTexture Tex_Selection, X + 2, Y + 2 + (Height - 4), 3, 32 - 3, Width - 4, 2, 32 - 6, 2, -1 'bottom line
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawSelectionBox", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawTileOutline()
Dim rec As rect


   On Error GoTo errorhandler

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    RenderTexture Tex_Misc, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255), True




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawTileOutline", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub NewCharacterDrawSprite()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRect As rect
Dim dRect As rect
Dim Width As Long, Height As Long, TextureNum As DX8TextureRec, i As Long
    'Blank Because this will become faces

   On Error GoTo errorhandler

    dRect = NCPreviewBounds
    If CharMode = 1 Then
        If NewCharHair > 0 And NewCharHair <= NumFaceHair Then
            If Tex_FHairB(NewCharHair).Texture > 0 Then
                RenderTexture Tex_FHairB(NewCharHair), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FHairB(NewCharHair).Width, Tex_FHairB(NewCharHair).Height
            End If
        End If
        If NewCharHead > 0 And NewCharHead <= NumFaceHeads Then
            If Tex_FHeads(NewCharHead).Texture > 0 Then
                RenderTexture Tex_FHeads(NewCharHead), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FHeads(NewCharHead).Width, Tex_FHeads(NewCharHead).Height
            End If
        End If
        If NewCharEye > 0 And NewCharEye <= NumFaceEyes Then
            If Tex_FEyes(NewCharEye).Texture > 0 Then
                RenderTexture Tex_FEyes(NewCharEye), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FEyes(NewCharEye).Width, Tex_FEyes(NewCharEye).Height
            End If
        End If
        If NewCharEyebrow > 0 And NewCharEyebrow <= NumFaceEyebrows Then
            If Tex_FEyebrows(NewCharEyebrow).Texture > 0 Then
                RenderTexture Tex_FEyebrows(NewCharEyebrow), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FEyebrows(NewCharEyebrow).Width, Tex_FEyebrows(NewCharEyebrow).Height
            End If
        End If
        If NewCharNose > 0 And NewCharNose <= NumFaceNoses Then
            If Tex_FNose(NewCharNose).Texture > 0 Then
                RenderTexture Tex_FNose(NewCharNose), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FNose(NewCharNose).Width, Tex_FNose(NewCharNose).Height
            End If
        End If
        If NewCharMouth > 0 And NewCharMouth <= NumFaceMouths Then
            If Tex_FMouth(NewCharMouth).Texture > 0 Then
                RenderTexture Tex_FMouth(NewCharMouth), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FMouth(NewCharMouth).Width, Tex_FMouth(NewCharMouth).Height
            End If
        End If
        If NewCharEar > 0 And NewCharEar <= NumFaceEars Then
            If Tex_FEars(NewCharEar).Texture > 0 Then
                RenderTexture Tex_FEars(NewCharEar), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FEars(NewCharEar).Width, Tex_FEars(NewCharEar).Height
            End If
        End If
        If NewCharEtc > 0 And NewCharEtc <= NumFaceEtc Then
            If Tex_FEtc(NewCharEtc).Texture > 0 Then
                RenderTexture Tex_FEtc(NewCharEtc), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FEtc(NewCharEtc).Width, Tex_FEtc(NewCharEtc).Height
            End If
        End If
        If NewCharHair > 0 And NewCharHair <= NumFaceHair Then
            If Tex_FHair(NewCharHair).Texture > 0 Then
                RenderTexture Tex_FHair(NewCharHair), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FHair(NewCharHair).Width, Tex_FHair(NewCharHair).Height
            End If
        End If
        If NewCharShirt > 0 And NewCharShirt <= NumFaceShirts Then
            If Tex_FShirts(NewCharShirt).Texture > 0 Then
                RenderTexture Tex_FShirts(NewCharShirt), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FShirts(NewCharShirt).Width, Tex_FShirts(NewCharShirt).Height
            End If
        End If
    Else
        If NewCharHead > 0 And NewCharHead <= NumFaces Then
            If Tex_Face(NewCharHead).Texture > 0 Then
                RenderTexture Tex_Face(NewCharHead), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_Face(NewCharHead).Width, Tex_Face(NewCharHead).Height
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "NewCharacterDrawSprite", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub SelCharDrawSprite()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRect As rect
Dim dRect As rect
Dim Width As Long, Height As Long, TextureNum As DX8TextureRec, i As Long

   On Error GoTo errorhandler

    dRect = CharPreviewBounds
    If SelectedChar <= 0 Or SelectedChar > UBound(CharSelection) Then Exit Sub
    If Trim$(CharSelection(SelectedChar).Name) = "Free Character Slot" Then Exit Sub
    If CharMode = 1 Then
        If CharSelection(SelectedChar).Face(FaceEnum.Hair) > 0 And CharSelection(SelectedChar).Face(FaceEnum.Hair) <= NumFaceHair Then
            If Tex_FHairB(CharSelection(SelectedChar).Face(FaceEnum.Hair)).Texture > 0 Then
                RenderTexture Tex_FHairB(CharSelection(SelectedChar).Face(FaceEnum.Hair)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FHairB(CharSelection(SelectedChar).Face(FaceEnum.Hair)).Width, Tex_FHairB(CharSelection(SelectedChar).Face(FaceEnum.Hair)).Height
            End If
        End If
        If CharSelection(SelectedChar).Face(FaceEnum.Head) > 0 And CharSelection(SelectedChar).Face(FaceEnum.Head) <= NumFaceHeads Then
            If Tex_FHeads(CharSelection(SelectedChar).Face(FaceEnum.Head)).Texture > 0 Then
                RenderTexture Tex_FHeads(CharSelection(SelectedChar).Face(FaceEnum.Head)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FHeads(CharSelection(SelectedChar).Face(FaceEnum.Head)).Width, Tex_FHeads(CharSelection(SelectedChar).Face(FaceEnum.Head)).Height
            End If
        End If
        If CharSelection(SelectedChar).Face(FaceEnum.Eyes) > 0 And CharSelection(SelectedChar).Face(FaceEnum.Eyes) <= NumFaceEyes Then
            If Tex_FEyes(CharSelection(SelectedChar).Face(FaceEnum.Eyes)).Texture > 0 Then
                RenderTexture Tex_FEyes(CharSelection(SelectedChar).Face(FaceEnum.Eyes)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FEyes(CharSelection(SelectedChar).Face(FaceEnum.Eyes)).Width, Tex_FEyes(CharSelection(SelectedChar).Face(FaceEnum.Eyes)).Height
            End If
        End If
        If CharSelection(SelectedChar).Face(FaceEnum.EyeBrows) > 0 And CharSelection(SelectedChar).Face(FaceEnum.EyeBrows) <= NumFaceEyebrows Then
            If Tex_FEyebrows(CharSelection(SelectedChar).Face(FaceEnum.EyeBrows)).Texture > 0 Then
                RenderTexture Tex_FEyebrows(CharSelection(SelectedChar).Face(FaceEnum.EyeBrows)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FEyebrows(CharSelection(SelectedChar).Face(FaceEnum.EyeBrows)).Width, Tex_FEyebrows(CharSelection(SelectedChar).Face(FaceEnum.EyeBrows)).Height
            End If
        End If
        If CharSelection(SelectedChar).Face(FaceEnum.Nose) > 0 And CharSelection(SelectedChar).Face(FaceEnum.Nose) <= NumFaceNoses Then
            If Tex_FNose(CharSelection(SelectedChar).Face(FaceEnum.Nose)).Texture > 0 Then
                RenderTexture Tex_FNose(CharSelection(SelectedChar).Face(FaceEnum.Nose)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FNose(CharSelection(SelectedChar).Face(FaceEnum.Nose)).Width, Tex_FNose(CharSelection(SelectedChar).Face(FaceEnum.Nose)).Height
            End If
        End If
        If CharSelection(SelectedChar).Face(FaceEnum.Mouth) > 0 And CharSelection(SelectedChar).Face(FaceEnum.Mouth) <= NumFaceMouths Then
            If Tex_FMouth(CharSelection(SelectedChar).Face(FaceEnum.Mouth)).Texture > 0 Then
                RenderTexture Tex_FMouth(CharSelection(SelectedChar).Face(FaceEnum.Mouth)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FMouth(CharSelection(SelectedChar).Face(FaceEnum.Mouth)).Width, Tex_FMouth(CharSelection(SelectedChar).Face(FaceEnum.Mouth)).Height
            End If
        End If
        If CharSelection(SelectedChar).Face(FaceEnum.Ears) > 0 And CharSelection(SelectedChar).Face(FaceEnum.Ears) <= NumFaceEars Then
            If Tex_FEars(CharSelection(SelectedChar).Face(FaceEnum.Ears)).Texture > 0 Then
                RenderTexture Tex_FEars(CharSelection(SelectedChar).Face(FaceEnum.Ears)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FEars(CharSelection(SelectedChar).Face(FaceEnum.Ears)).Width, Tex_FEars(CharSelection(SelectedChar).Face(FaceEnum.Ears)).Height
            End If
        End If
        If CharSelection(SelectedChar).Face(FaceEnum.Etc) > 0 And CharSelection(SelectedChar).Face(FaceEnum.Etc) <= NumFaceEtc Then
            If Tex_FEtc(CharSelection(SelectedChar).Face(FaceEnum.Etc)).Texture > 0 Then
                RenderTexture Tex_FEtc(CharSelection(SelectedChar).Face(FaceEnum.Etc)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FEtc(CharSelection(SelectedChar).Face(FaceEnum.Etc)).Width, Tex_FEtc(CharSelection(SelectedChar).Face(FaceEnum.Etc)).Height
            End If
        End If
        If CharSelection(SelectedChar).Face(FaceEnum.Hair) > 0 And CharSelection(SelectedChar).Face(FaceEnum.Hair) <= NumFaceHair Then
            If Tex_FHair(CharSelection(SelectedChar).Face(FaceEnum.Hair)).Texture > 0 Then
                RenderTexture Tex_FHair(CharSelection(SelectedChar).Face(FaceEnum.Hair)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FHair(CharSelection(SelectedChar).Face(FaceEnum.Hair)).Width, Tex_FHair(CharSelection(SelectedChar).Face(FaceEnum.Hair)).Height
            End If
        End If
        If CharSelection(SelectedChar).Face(FaceEnum.Shirt) > 0 And CharSelection(SelectedChar).Face(FaceEnum.Shirt) <= NumFaceShirts Then
            If Tex_FShirts(CharSelection(SelectedChar).Face(FaceEnum.Shirt)).Texture > 0 Then
                RenderTexture Tex_FShirts(CharSelection(SelectedChar).Face(FaceEnum.Shirt)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_FShirts(CharSelection(SelectedChar).Face(FaceEnum.Shirt)).Width, Tex_FShirts(CharSelection(SelectedChar).Face(FaceEnum.Shirt)).Height
            End If
        End If
    Else
        If CharSelection(SelectedChar).Face(FaceEnum.Head) > 0 And CharSelection(SelectedChar).Face(FaceEnum.Head) <= NumFaces Then
            If Tex_Face(CharSelection(SelectedChar).Face(FaceEnum.Head)).Texture > 0 Then
                RenderTexture Tex_Face(CharSelection(SelectedChar).Face(FaceEnum.Head)), dRect.Left, dRect.Top, 0, 0, dRect.Right, dRect.Bottom, Tex_Face(CharSelection(SelectedChar).Face(FaceEnum.Head)).Width, Tex_Face(CharSelection(SelectedChar).Face(FaceEnum.Head)).Height
            End If
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "SelCharDrawSprite", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorMap_DrawMapItem()
Dim ItemNum As Long
Dim sRect As rect, destRect As D3DRECT
Dim dRect As rect


   On Error GoTo errorhandler

    ItemNum = Item(frmEditor_Map.scrlMapItem.Value).pic

    If ItemNum < 1 Or ItemNum > numitems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(ItemNum), sRect, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picMapItem.hwnd, ByVal (0)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorMap_DrawMapItem", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorMap_DrawKey()
Dim ItemNum As Long
Dim sRect As rect, destRect As D3DRECT
Dim dRect As rect



   On Error GoTo errorhandler

    ItemNum = Item(frmEditor_Map.scrlMapKey.Value).pic

    If ItemNum < 1 Or ItemNum > numitems Then
        frmEditor_Map.picMapKey.Cls
        Exit Sub
    End If
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    RenderTextureByRects Tex_Item(ItemNum), sRect, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picMapKey.hwnd, ByVal (0)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorMap_DrawKey", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorItem_DrawItem()
Dim ItemNum As Long
Dim sRect As rect, destRect As D3DRECT
Dim dRect As rect


   On Error GoTo errorhandler

    ItemNum = frmEditor_Item.scrlPic.Value

    If ItemNum < 1 Or ItemNum > numitems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If


    ' rect for source
    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    ' same for destination as source
    dRect = sRect
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(ItemNum), sRect, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picItem.hwnd, ByVal (0)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorItem_DrawItem", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorPet_DrawPet()
Dim SpriteNum As Long
Dim sRect As rect, destRect As D3DRECT
Dim dRect As rect


   On Error GoTo errorhandler

    SpriteNum = frmEditor_Pet.scrlSprite.Value

    If SpriteNum < 1 Or SpriteNum > NumCharacters Then
        frmEditor_Pet.picSprite.Cls
        Exit Sub
    End If


    ' rect for source
    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    ' same for destination as source
    dRect = sRect
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(SpriteNum), sRect, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Pet.picSprite.hwnd, ByVal (0)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorPet_DrawPet", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorItem_DrawFurniture()
Dim ItemNum As Long
Dim sRect As rect, destRect As D3DRECT
Dim dRect As rect, X As Long, Y As Long

   On Error GoTo errorhandler

    If frmEditor_Item.fraFurniture.Visible = False Then Exit Sub
    ItemNum = frmEditor_Item.scrlFurniture.Value

    If ItemNum < 1 Or ItemNum > NumFurniture Then
        frmEditor_Item.picFurniture.Cls
        Exit Sub
    End If
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTexture Tex_Furniture(ItemNum), 0, 0, 0, 0, Tex_Furniture(ItemNum).Width, Tex_Furniture(ItemNum).Height, Tex_Furniture(ItemNum).Width, Tex_Furniture(ItemNum).Height
    If frmEditor_Item.optSetBlocks.Value = True Then
        For X = 0 To 3
            For Y = 0 To 3
                If X <= (Tex_Furniture(ItemNum).Width / 32) - 1 Then
                    If Y <= (Tex_Furniture(ItemNum).Height / 32) - 1 Then
                        If Item(EditorIndex).FurnitureBlocks(X, Y) = 1 Then
                             RenderText Font_Default, "X", X * 32 + 8, Y * 32 + 8, BrightRed
                        Else
                             RenderText Font_Default, "O", X * 32 + 8, Y * 32 + 8, Blue
                        End If
                    End If
                End If
            Next
        Next
    ElseIf frmEditor_Item.optSetFringe.Value = True Then
        For X = 0 To 3
            For Y = 0 To 3
                If X <= Item(EditorIndex).FurnitureWidth - 1 Then
                    If Y <= Item(EditorIndex).FurnitureHeight Then
                        If Item(EditorIndex).FurnitureFringe(X, Y) = 1 Then
                             RenderText Font_Default, "O", X * 32 + 8, Y * 32 + 8, Blue
                        End If
                    End If
                End If
            Next
        Next
    End If
    With destRect
        .X1 = 0
        .X2 = frmEditor_Item.picFurniture.Width
        .Y1 = 0
        .Y2 = frmEditor_Item.picFurniture.Height
    End With
                    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picFurniture.hwnd, ByVal (0)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorItem_DrawFurniture", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorItem_DrawPaperdoll()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRect As rect
Dim dRect As rect

   On Error GoTo errorhandler

    Sprite = frmEditor_Item.scrlPaperdoll.Value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    ' rect for source
    sRect.Top = 0
    sRect.Bottom = Tex_Paperdoll(Sprite).Height / 4
    sRect.Left = 0
    sRect.Right = Tex_Paperdoll(Sprite).Width / 4
    ' same for destination as source
    dRect = sRect
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Paperdoll(Sprite), sRect, dRect
                    With destRect
        .X1 = 0
        .X2 = Tex_Paperdoll(Sprite).Width / 4
        .Y1 = 0
        .Y2 = Tex_Paperdoll(Sprite).Height / 4
    End With
                    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picPaperdoll.hwnd, ByVal (0)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorItem_DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorSpell_DrawIcon()
Dim iconnum As Long, destRect As D3DRECT
Dim sRect As rect
Dim dRect As rect



   On Error GoTo errorhandler

    iconnum = frmEditor_Spell.scrlIcon.Value
    If iconnum < 1 Or iconnum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_SpellIcon(iconnum), sRect, dRect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Spell.picSprite.hwnd, ByVal (0)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorSpell_DrawIcon", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorProjectile_DrawProjectile()
Dim iconnum As Long, destRect As D3DRECT
Dim sRect As rect
Dim dRect As rect

   On Error GoTo errorhandler

    iconnum = frmEditor_Projectile.scrlPic.Value
    If iconnum < 1 Or iconnum > NumProjectiles Then
        frmEditor_Projectile.picProjectile.Cls
        Exit Sub
    End If
    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X * 4
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X * 4
    With destRect
        .X1 = 0
        .X2 = PIC_X * 4
        .Y1 = 0
        .Y2 = PIC_Y
    End With
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Projectiles(iconnum), sRect, dRect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Projectile.picProjectile.hwnd, ByVal (0)

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorProjectile_DrawProjectile", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorAnim_DrawAnim()
Dim Animationnum As Long
Dim sRect As rect
Dim dRect As rect
Dim i As Long
Dim Width As Long, Height As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim looptime As Long
Dim FrameCount As Long
Dim ShouldRender As Boolean


   On Error GoTo errorhandler

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).Value
            If Animationnum < 1 Or Animationnum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
                    ShouldRender = False
                    ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If
                If ShouldRender Then
                'frmEditor_Animation.picSprite(i).Cls
                            If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    ' total width divided by frame count
                    Width = Tex_Animation(Animationnum).Width / frmEditor_Animation.scrlFrameCount(i).Value
                    Height = Tex_Animation(Animationnum).Height
                                    sRect.Top = 0
                    sRect.Bottom = Height
                    sRect.Left = (AnimEditorFrame(i) - 1) * Width
                    sRect.Right = sRect.Left + Width
                                    dRect.Top = 0
                    dRect.Bottom = Height
                    dRect.Left = 0
                    dRect.Right = Width
                                    RenderTextureByRects Tex_Animation(Animationnum), sRect, dRect
                                    With srcRect
                        .X1 = 0
                        .X2 = frmEditor_Animation.picSprite(i).Width
                        .Y1 = 0
                        .Y2 = frmEditor_Animation.picSprite(i).Height
                    End With
                                                With destRect
                        .X1 = 0
                        .X2 = frmEditor_Animation.picSprite(i).Width
                        .Y1 = 0
                        .Y2 = frmEditor_Animation.picSprite(i).Height
                    End With
                                                Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Animation.picSprite(i).hwnd, ByVal (0)
                End If
            End If
        End If
    Next




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorAnim_DrawAnim", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorNpc_DrawSprite()
Dim Sprite As Long, destRect As D3DRECT
Dim sRect As rect
Dim dRect As rect



   On Error GoTo errorhandler

    Sprite = frmEditor_NPC.scrlSprite.Value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    sRect.Top = 0
    sRect.Bottom = SIZE_Y
    sRect.Left = PIC_X * 3 ' facing down
    sRect.Right = sRect.Left + SIZE_X
    dRect.Top = 0
    dRect.Bottom = SIZE_Y
    dRect.Left = 0
    dRect.Right = SIZE_X
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(Sprite), sRect, dRect
    With destRect
        .X1 = 0
        .X2 = SIZE_X
        .Y1 = 0
        .Y2 = SIZE_Y
    End With
                    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_NPC.picSprite.hwnd, ByVal (0)




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorNpc_DrawSprite", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorResource_DrawSprite()
Dim Sprite As Long
Dim sRect As rect, destRect As D3DRECT, srcRect As D3DRECT
Dim dRect As rect

    ' normal sprite

   On Error GoTo errorhandler

    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRect.Top = 0
        sRect.Bottom = Tex_Resource(Sprite).Height
        sRect.Left = 0
        sRect.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRect, dRect
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(Sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(Sprite).Height
        End With
            With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picNormalPic.hwnd, ByVal (0)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRect.Top = 0
        sRect.Bottom = Tex_Resource(Sprite).Height
        sRect.Left = 0
        sRect.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRect, dRect
            With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
            With srcRect
            .X1 = 0
            .X2 = Tex_Resource(Sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(Sprite).Height
        End With
                        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picExhaustedPic.hwnd, ByVal (0)
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorResource_DrawSprite", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub Render_Menu()
Dim X As Long, z As Long
Dim Y As Long
Dim i As Long
Dim rec As rect
Dim rec_pos As rect, srcRect As D3DRECT, destRect As D3DRECT
   On Error GoTo errorhandler
    'Check for device lost.
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    For i = 1 To NumTextures
        If gTexture(i).IsLoaded Then
            If gTexture(i).TextureTimer < GetTickCount Then
                If i <> Font_Default.Texture.Texture And i <> Font_Georgia.Texture.Texture Then
                    Set gTexture(i).Texture = Nothing
                    ZeroMemory ByVal VarPtr(gTexture(i)), LenB(gTexture(i))
                End If
            End If
        End If
    Next
    With srcRect
        .X1 = 0
        .X2 = MenuWidth
        .Y1 = 0
        .Y2 = MenuHeight
    End With
            With destRect
        .X1 = 0
        .X2 = frmMain.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.ScaleHeight
    End With
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
            Direct3D_Device.BeginScene
            If InIntro = 1 Then
                If IntroType = 1 Then
                    RenderTexture IntroImages(IntroStep), 0, 0, 0, 0, MenuWidth, MenuHeight, IntroImages(IntroStep).Width, IntroImages(IntroStep).Height
                End If
            Else
                DrawMenu
            End If
            If FadeAmount > 0 Then RenderTexture Tex_Fade, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, FadeAmount)
            If BFPS Then
                RenderText Font_Default, "FPS: " & CStr(GameFPS), 2, 1, Yellow, 0
            End If
        Direct3D_Device.EndScene
        If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        Direct3D_Device.Present srcRect, destRect, 0, ByVal 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Err.Clear
        Exit Sub
    Else
        HandleError "Render_Menu", "modGraphics", Err.Number, Err.Description, Erl
        Err.Clear
    End If
End Sub

Public Sub Render_Graphics()
Dim X As Long, z As Long
Dim Y As Long
Dim i As Long
Dim rec As rect
Dim rec_pos As rect, srcRect As D3DRECT, destRect As D3DRECT
   On Error GoTo errorhandler
    'Check for device lost.
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub
    For i = 1 To NumTextures
        If gTexture(i).IsLoaded Then
            If gTexture(i).TextureTimer < GetTickCount Then
                If i <> Font_Default.Texture.Texture And i <> Font_Georgia.Texture.Texture Then
                    Set gTexture(i).Texture = Nothing
                    ZeroMemory ByVal VarPtr(gTexture(i)), LenB(gTexture(i))
                End If
            End If
        End If
    Next
    ' update the viewpoint
    UpdateCamera
    
    If Options.GfxMode = 0 Then
        If CacheMap = False Then
            CacheMapLayers
            CacheMap = True
        End If
    End If

   Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
            Direct3D_Device.BeginScene

            
            If NumTileSets > 0 Then
                If InMapEditor Or Options.GfxMode = 1 Then
                    For X = TileView.Left To TileView.Right
                        For Y = TileView.Top To TileView.Bottom
                            If IsValidMapPoint(X, Y) Then
                                Call DrawMapTile(X, Y)
                            End If
                        Next
                    Next
                Else
                    RenderLowerMap
                End If
            End If
            
            
            
                ' render the decals
            For i = 1 To MAX_BYTE
                Call DrawBlood(i)
            Next
                ' Blit out the items
            If numitems > 0 Then
                For i = 1 To MAX_MAP_ITEMS
                    If MapItem(i).Num > 0 Then
                        Call DrawItem(i)
                    End If
                Next
            End If
            ' Furniture
            If FurnitureHouse > 0 Then
                If FurnitureHouse = Player(MyIndex).InHouse Then
                    If FurnitureCount > 0 Then
                        For i = 1 To FurnitureCount
                            If Furniture(i).ItemNum > 0 Then
                                Call DrawFurniture(i, 0)
                            End If
                        Next
                    End If
                End If
            End If
            If Map.CurrentEvents > 0 Then
                For i = 1 To Map.CurrentEvents
                    If Map.MapEvents(i).Position = 0 Then
                        DrawEvent i
                    End If
                Next
            End If
                    ' draw animations
            If NumAnimations > 0 Then
                For i = 1 To MAX_BYTE
                    If AnimInstance(i).Used(0) Then
                        DrawAnimation i, 0
                    End If
                Next
            End If
                ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
            For Y = 0 To Map.MaxY
                If NumCharacters > 0 Then
                                
                    ' Players
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                            If Player(i).Y = Y Then
                                Call DrawPlayer(i)
                            End If
                            If Player(i).Pet.Alive = True Then
                                If Player(i).Pet.Y = Y Then
                                    DrawPet (i)
                                End If
                            End If
                        End If
                    Next
                    
                    If Map.CurrentEvents > 0 Then
                        For i = 1 To Map.CurrentEvents
                            If Map.MapEvents(i).Position = 1 Then
                                If Y = Map.MapEvents(i).Y Then
                                    DrawEvent i
                                End If
                            End If
                        Next
                    End If
                                                                ' Npcs
                    For i = 1 To Npc_HighIndex
                        If MapNpc(i).Y = Y Then
                            Call DrawNpc(i)
                        End If
                    Next
                                    For i = 1 To MAX_ZONES
                        For X = 1 To MAX_MAP_NPCS * 2
                            If ZoneNPC(i).Npc(X).Vital(Vitals.HP) > 0 Then
                                If ZoneNPC(i).Npc(X).Num > 0 Then
                                    If ZoneNPC(i).Npc(X).Y = Y Then
                                        'Draw Zone NPC
                                        DrawZoneNpc i, X
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If
                            ' Resources
                If NumResources > 0 Then
                    If Resources_Init Then
                        If Resource_Index > 0 Then
                            For i = 1 To Resource_Index
                                If MapResource(i).Y = Y Then
                                    Call DrawMapResource(i)
                                End If
                            Next
                        End If
                    End If
                End If
            
            Next
            
            'Projectiles
            If NumProjectiles > 0 Then
                For i = 1 To MAX_PROJECTILES
                    If MapProjectiles(i).ProjectileNum > 0 Then
                        DrawProjectile i
                    End If
                Next
            End If
            
                    ' animations
            If NumAnimations > 0 Then
                For i = 1 To MAX_BYTE
                    If AnimInstance(i).Used(1) Then
                        DrawAnimation i, 1
                    End If
                Next
            End If
                ' blit out upper tiles
            If NumTileSets > 0 Then
                If InMapEditor Or Options.GfxMode = 1 Then
                    For X = TileView.Left To TileView.Right
                        For Y = TileView.Top To TileView.Bottom
                            If IsValidMapPoint(X, Y) Then
                                Call DrawMapFringeTile(X, Y)
                            End If
                        Next
                    Next
                Else
                    RenderUpperMap
                End If
            End If
                    ' Furniture
            If FurnitureHouse > 0 Then
                If FurnitureHouse = Player(MyIndex).InHouse Then
                    If FurnitureCount > 0 Then
                        For i = 1 To FurnitureCount
                            If Furniture(i).ItemNum > 0 Then
                                Call DrawFurniture(i, 1)
                            End If
                        Next
                    End If
                End If
            End If
                    If Map.CurrentEvents > 0 Then
                For i = 1 To Map.CurrentEvents
                    If Map.MapEvents(i).Position = 2 Then
                        DrawEvent i
                    End If
                Next
            End If
                    DrawWeather
            DrawFog
            DrawTint
                    ' blit out a square at mouse cursor
            If InMapEditor Then
                If frmEditor_Map.optBlock.Value = True Then
                    For X = TileView.Left To TileView.Right
                        For Y = TileView.Top To TileView.Bottom
                            If IsValidMapPoint(X, Y) Then
                                If Map.Tile(X, Y).type = TILE_TYPE_BLOCKED Then
                                    RenderText Font_Default, "B", ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5), ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5), BrightRed, 0, True, True
                                End If
                            End If
                        Next
                    Next
                End If
                Call DrawTileOutline
            End If
                    If DragInvSlotNum > 0 Then
                If Player(MyIndex).InHouse = MyIndex Then
                    If Item(PlayerInv(DragInvSlotNum).Num).type = ITEM_TYPE_FURNITURE Then
                        If Item(PlayerInv(DragInvSlotNum).Num).data2 > 0 And Item(PlayerInv(DragInvSlotNum).Num).data2 <= NumFurniture Then
                            RenderTexture Tex_Furniture(Item(PlayerInv(DragInvSlotNum).Num).data2), ConvertMapX(CurX * 32), ConvertMapY(CurY * 32), 0, 0, Tex_Furniture(Item(PlayerInv(DragInvSlotNum).Num).data2).Width, Tex_Furniture(Item(PlayerInv(DragInvSlotNum).Num).data2).Height, Tex_Furniture(Item(PlayerInv(DragInvSlotNum).Num).data2).Width, Tex_Furniture(Item(PlayerInv(DragInvSlotNum).Num).data2).Height, -1, True
                        End If
                    End If
                End If
            End If
                    ' Render the bars
            DrawBars
                    ' Draw the target icon
            If myTarget > 0 Then
                If myTargetType = TARGET_TYPE_PLAYER Then
                    DrawTarget (Player(myTarget).X * 32) + Player(myTarget).XOffset, (Player(myTarget).Y * 32) + Player(myTarget).YOffset
                ElseIf myTargetType = TARGET_TYPE_NPC Then
                    DrawTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).XOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).YOffset
                ElseIf myTargetType = TARGET_TYPE_ZONENPC Then
                    If ZoneNPC(myTargetZone).Npc(myTarget).Map = GetPlayerMap(MyIndex) Then
                        DrawTarget (ZoneNPC(myTargetZone).Npc(myTarget).X * 32) + ZoneNPC(myTargetZone).Npc(myTarget).XOffset, (ZoneNPC(myTargetZone).Npc(myTarget).Y * 32) + ZoneNPC(myTargetZone).Npc(myTarget).YOffset
                    End If
                ElseIf myTargetType = TARGET_TYPE_PET Then
                    DrawTarget (Player(myTarget).Pet.X * 32) + Player(myTarget).Pet.XOffset, (Player(myTarget).Pet.Y * 32) + Player(myTarget).Pet.YOffset
                End If
            End If
                    ' Draw the hover icon
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If Player(i).Map = Player(MyIndex).Map Then
                        If Player(i).InHouse = Player(MyIndex).InHouse Then
                            If CurX = Player(i).X And CurY = Player(i).Y Then
                                If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                                    ' dont render lol
                                Else
                                    DrawHover TARGET_TYPE_PLAYER, i, (Player(i).X * 32) + Player(i).XOffset, (Player(i).Y * 32) + Player(i).YOffset
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Num > 0 Then
                    If CurX = MapNpc(i).X And CurY = MapNpc(i).Y Then
                        If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                            ' dont render lol
                        Else
                            DrawHover TARGET_TYPE_NPC, i, (MapNpc(i).X * 32) + MapNpc(i).XOffset, (MapNpc(i).Y * 32) + MapNpc(i).YOffset
                        End If
                    End If
                End If
            Next
                    For i = 1 To MAX_ZONES
                For z = 1 To MAX_MAP_NPCS * 2
                    If ZoneNPC(i).Npc(z).Num > 0 Then
                        If ZoneNPC(i).Npc(z).Map = GetPlayerMap(MyIndex) Then
                            If CurX = ZoneNPC(i).Npc(z).X And CurY = ZoneNPC(i).Npc(z).Y Then
                                If myTargetType = TARGET_TYPE_ZONENPC And myTarget = X And myTargetZone = i Then
                                    ' dont render lol
                                Else
                                    DrawHover TARGET_TYPE_ZONENPC, z, (ZoneNPC(i).Npc(z).X * 32) + ZoneNPC(i).Npc(z).XOffset, (ZoneNPC(i).Npc(z).Y * 32) + ZoneNPC(i).Npc(z).YOffset
                                End If
                            End If
                        End If
                    End If
                Next
            Next
                    If DrawThunder > 0 Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160), True, True: DrawThunder = DrawThunder - 1
                    DrawWalkArrow
                    ' Get rec
            With rec
                .Top = Camera.Top
                .Bottom = .Top + ScreenY
                .Left = Camera.Left
                .Right = .Left + ScreenX
            End With
                        ' rec_pos
            With rec_pos
                .Bottom = ScreenY
                .Right = ScreenX
            End With
            'With srcRect
            '    .X1 = 0
            '    .X2 = GameScreenWidth
            '    .Y1 = 0
            '    .Y2 = GameScreenHeight
            'End With
            'With destRect
            '    .X1 = 0
            '    .X2 = GameScreenWidth
            '    .Y1 = 0
            '    .Y2 = GameScreenHeight
            'End With
            
            With srcRect
                .X1 = 0
                .X2 = BBWidth
                .Y1 = 0
                .Y2 = BBHeight
            End With
                    With destRect
                .X1 = 0
                .X2 = BBWidth
                .Y1 = 0
                .Y2 = BBHeight
            End With
    
                    ' draw player names
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    Call DrawPlayerName(i)
                    If Player(i).Pet.Health > 0 And Trim$(Player(i).Pet.Num) > 0 Then
                        Call DrawPlayerPetName(i)
                    End If
                End If
            Next
                    For i = 1 To Map.CurrentEvents
                If Map.MapEvents(i).Visible = 1 Then
                    If Map.MapEvents(i).ShowName = 1 Then
                        DrawEventName (i)
                    End If
                End If
            Next
                    ' draw npc names
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Num > 0 Then
                    Call DrawNpcName(i)
                End If
            Next
                    For i = 1 To MAX_ZONES
                For X = 1 To MAX_MAP_NPCS * 2
                    If ZoneNPC(i).Npc(X).Num > 0 Then
                        If ZoneNPC(i).Npc(X).Map = GetPlayerMap(MyIndex) Then
                            DrawZoneNpcName i, X
                        End If
                    End If
                Next
            Next
                        ' draw the messages
            For i = 1 To MAX_BYTE
                If chatBubble(i).active Then
                    DrawChatBubble i
                End If
            Next
            If ActionMsgTick < GetTickCount Then
                For i = 1 To Action_HighIndex
                    Call DrawActionMsg(i, True)
                Next i
                ActionMsgTick = GetTickCount + 20
            Else
                For i = 1 To Action_HighIndex
                    Call DrawActionMsg(i, False)
                Next i
            End If
            If InMapEditor And frmEditor_Map.optEvent.Value = True Then DrawEvents
            If InMapEditor Then Call DrawMapAttributes
                    If FadeAmount > 0 Then RenderTexture Tex_Fade, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, FadeAmount)
            If FlashTimer > GetTickCount Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, -1, True, True
            If Not InMapEditor Then DrawPictures
            DrawGUI
            z = 90
            If BFPS Then
                RenderText Font_Default, "FPS: " & CStr(GameFPS), 10, z, Yellow, 0
                z = z + 14
            End If
                    ' draw cursor, player X and Y locations
            If BLoc Then
                RenderText Font_Default, Trim$("cur x: " & CurX & " y: " & CurY), 10, z, Yellow, 0
                RenderText Font_Default, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 10, z + 14, Yellow, 0
                RenderText Font_Default, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 10, z + 28, Yellow, 0
                z = z + 42
            End If
            If BPing Then
                RenderText Font_Default, Trim$("Ping: " & CStr(Ping)), 10, z, Yellow
            End If
        Direct3D_Device.EndScene
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        Direct3D_Device.Present srcRect, destRect, 0, ByVal 0
        DrawGDI
    End If

    ' Error handler
    Exit Sub
errorhandler:
    If Direct3D_Device Is Nothing Then Exit Sub
    
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Err.Clear
        Exit Sub
    Else
        If Options.Debug = 1 Then
            HandleError "Render_Graphics", "modGraphics", Err.Number, Err.Description, Erl
            Err.Clear
        End If
    End If
End Sub

Sub HandleDeviceLost()
'Do a loop while device is lost
   Do While Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST
       Exit Sub
   Loop
   
   UnloadTextures
   
   'Reset the device
   Direct3D_Device.Reset Direct3D_Window
   
   DirectX_ReInit
   LoadTextures
   LoadGUI
   



End Sub

Private Function DirectX_ReInit() As Boolean

    On Error GoTo Error_Handler

    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
        Direct3D_Window.Windowed = True 'The app will be in windowed mode.

    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'we've already setup for Direct3D_Window.
    'Creates the rendering device with some useful info, along with the info
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = BBWidth ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = BBHeight 'frmMain.ScaleHeight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hwnd 'Use frmMain as the device window.
    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
            .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
    DirectX_ReInit = True

    Exit Function
Error_Handler:
    MsgBox "An error occured while attempting to re-initialize DirectX", vbCritical
    DestroyGame
    DirectX_ReInit = False
End Function

Public Sub UpdateCamera()
Dim offsetX As Long
Dim offsetY As Long
Dim startX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    offsetX = Player(MyIndex).XOffset + PIC_X
    offsetY = Player(MyIndex).YOffset + PIC_Y
    If Map.MaxX = MAX_MAPX + 1 Then
        startX = GetPlayerX(MyIndex) - StartXValue
        If startX <= 0 Then
            offsetX = 0
            If startX = 0 Then
                If Player(MyIndex).XOffset > 0 Then
                    offsetX = Player(MyIndex).XOffset
                End If
            End If
            startX = 0
        End If
        EndX = startX + EndXValue
        If EndX > Map.MaxX Then
            offsetX = 32
            If EndX = Map.MaxX Then
                If Player(MyIndex).XOffset < 0 Then
                    offsetX = Player(MyIndex).XOffset + PIC_X
                End If
            End If
            EndX = Map.MaxX
            startX = EndX - MAX_MAPX - 1
        End If
    Else
        startX = GetPlayerX(MyIndex) - StartXValue
        If startX <= 0 Then
            offsetX = 0
            If startX = -1 Then
                If Player(MyIndex).XOffset > 0 Then
                    offsetX = Player(MyIndex).XOffset
                End If
            End If
            startX = 0
        End If
        
        If Map.MaxX <= EndXValue Then
            EndX = Map.MaxX
            startX = 0
            offsetX = 0
        Else
            EndX = startX + EndXValue + 2
        End If
        
        If EndX >= Map.MaxX Then
            offsetX = 32
            If EndX = Map.MaxX + 1 Then
                If Player(MyIndex).XOffset < 0 Then
                    offsetX = Player(MyIndex).XOffset + PIC_X
                End If
            End If
            EndX = Map.MaxX
            startX = EndX - MAX_MAPX - 1
        End If
    End If
    
    If Map.MaxY = MAX_MAPY + 1 Then
        StartY = GetPlayerY(MyIndex) - StartYValue
        If StartY <= 0 Then
            offsetY = 0
            If StartY = 0 Then
                If Player(MyIndex).YOffset > 0 Then
                    offsetY = Player(MyIndex).YOffset
                End If
            End If
            StartY = 0
        End If
        EndY = StartY + EndYValue
        If EndY > Map.MaxY Then
            offsetY = 32
            If EndY = Map.MaxY Then
                If Player(MyIndex).YOffset < 0 Then
                    offsetY = Player(MyIndex).YOffset + PIC_Y
                End If
            End If
            EndY = Map.MaxY
            StartY = EndY - MAX_MAPY - 1
        End If
    Else
        StartY = GetPlayerY(MyIndex) - StartYValue
        If StartY <= 0 Then
            offsetY = 0
            If StartY = -1 Then
                If Player(MyIndex).YOffset > 0 Then
                    offsetY = Player(MyIndex).YOffset
                End If
            End If
            StartY = 0
        End If
        If Map.MaxY < EndYValue Then
            EndY = Map.MaxY
            StartY = 0
            offsetY = 0
        Else
            EndY = StartY + EndYValue + 2
        End If
        If EndY >= Map.MaxY Then
            offsetY = 32
            If EndY = Map.MaxY + 1 Then
                If Player(MyIndex).YOffset < 0 Then
                    offsetY = Player(MyIndex).YOffset + PIC_Y
                End If
            End If
            EndY = Map.MaxY
            StartY = EndY - MAX_MAPY - 1
        End If
    End If
    
    If EndX - startX < EndXValue Then
        offsetX = offsetX - (((EndXValue - EndX) / 2) * 32)
    End If
    
    If EndY - StartY < EndYValue Then
        offsetY = offsetY - (((EndYValue - EndY) / 2) * 32)
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = startX
        .Right = EndX
    End With
    
    With Camera
        .Top = offsetY
        .Bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
    UpdateDrawMapName

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Function ConvertMapX(ByVal X As Long) As Long


   On Error GoTo errorhandler

    ConvertMapX = X - (TileView.Left * PIC_X) - Camera.Left + GameScreenX



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ConvertMapX", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Function ConvertMapY(ByVal Y As Long) As Long


   On Error GoTo errorhandler

    ConvertMapY = Y - (TileView.Top * PIC_Y) - Camera.Top + GameScreenY



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ConvertMapY", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean


   On Error GoTo errorhandler

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.Top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    InViewPort = True



   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "InViewPort", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean


   On Error GoTo errorhandler

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
    


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Function

Public Sub LoadTilesets()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim tilesetInUse() As Boolean


   On Error GoTo errorhandler

    ReDim tilesetInUse(0 To NumTileSets)
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If Map.Tile(X, Y).Layer(i).Tileset > 0 And Map.Tile(X, Y).Layer(i).Tileset <= NumTileSets Then
                    tilesetInUse(Map.Tile(X, Y).Layer(i).Tileset) = True
                End If
            Next
        Next
    Next




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawEvents()
Dim sRect As rect
Dim Width As Long, Height As Long, i As Long, X As Long, Y As Long

   On Error GoTo errorhandler

    If Map.EventCount <= 0 Then Exit Sub
    For i = 1 To Map.EventCount
        Width = 32
        Height = 32
        X = Map.Events(i).X * 32
        Y = Map.Events(i).Y * 32
        If Map.Events(i).pageCount <= 0 Then
                sRect.Top = 0
                sRect.Bottom = 32
                sRect.Left = 0
                sRect.Right = 32
                RenderTexture Tex_Selection, ConvertMapX(X), ConvertMapY(Y), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
            GoTo nextevent
        End If
        X = ConvertMapX(X)
        Y = ConvertMapY(Y)
            If i > Map.EventCount Then Exit Sub
        If 1 > Map.Events(i).pageCount Then Exit Sub
        Select Case Map.Events(i).Pages(1).GraphicType
            Case 0
                sRect.Top = 0
                sRect.Bottom = 32
                sRect.Left = 0
                sRect.Right = 32
                RenderTexture Tex_Selection, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
            Case 1
                If Map.Events(i).Pages(1).Graphic > 0 And Map.Events(i).Pages(1).Graphic <= NumCharacters Then
                                    sRect.Top = (Map.Events(i).Pages(1).GraphicY * (Tex_Character(Map.Events(i).Pages(1).Graphic).Height / 4))
                    sRect.Left = (Map.Events(i).Pages(1).GraphicX * (Tex_Character(Map.Events(i).Pages(1).Graphic).Width / 4))
                    sRect.Bottom = sRect.Top + 32
                    sRect.Right = sRect.Left + 32
                    RenderTexture Tex_Character(Map.Events(i).Pages(1).Graphic), X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                Else
                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                End If
            Case 2
                If Map.Events(i).Pages(1).Graphic > 0 And Map.Events(i).Pages(1).Graphic < NumTileSets Then
                    sRect.Top = Map.Events(i).Pages(1).GraphicY * 32
                    sRect.Left = Map.Events(i).Pages(1).GraphicX * 32
                    sRect.Bottom = sRect.Top + 32
                    sRect.Right = sRect.Left + 32
                    RenderTexture Tex_Tileset(Map.Events(i).Pages(1).Graphic), X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                Else
                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
                End If
        End Select
nextevent:
    Next




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawEvents", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub EditorEvent_DrawGraphic()
Dim sRect As rect, destRect As D3DRECT, srcRect As D3DRECT
Dim dRect As rect



   On Error GoTo errorhandler

    If frmEditor_Events.picGraphicSel.Visible Then
        Select Case frmEditor_Events.cmbGraphic.ListIndex
            Case 0
                'None
                frmEditor_Events.picGraphicSel.Cls
                Exit Sub
            Case 1
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumCharacters Then
                                    If Tex_Character(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
                        sRect.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRect.Right = sRect.Left + (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width - sRect.Left)
                    Else
                        sRect.Left = 0
                        sRect.Right = Tex_Character(frmEditor_Events.scrlGraphic.Value).Width
                    End If
                                    If Tex_Character(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
                        sRect.Top = frmEditor_Events.hScrlGraphicSel.Value
                        sRect.Bottom = sRect.Top + (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height - sRect.Top)
                    Else
                        sRect.Top = 0
                        sRect.Bottom = Tex_Character(frmEditor_Events.scrlGraphic.Value).Height
                    End If
                                    With dRect
                        .Top = 0
                        .Bottom = sRect.Bottom - sRect.Top
                        .Left = 0
                        .Right = sRect.Right - sRect.Left
                    End With
                                    With destRect
                        .X1 = dRect.Left
                        .X2 = dRect.Right
                        .Y1 = dRect.Top
                        .Y2 = dRect.Bottom
                    End With
                                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRect
                            .X1 = (GraphicSelX * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4)) - sRect.Left
                            .X2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4) + .X1
                            .Y1 = (GraphicSelY * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4)) - sRect.Top
                            .Y2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4) + .Y1
                        End With

                    Else
                        With destRect
                            .X1 = (GraphicSelX * 32) - sRect.Left
                            .X2 = ((GraphicSelX2 - GraphicSelX) * 32) + .X1
                            .Y1 = (GraphicSelY * 32) - sRect.Top
                            .Y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .Y1
                        End With
                    End If
                    DrawSelectionBox destRect
                                    With srcRect
                        .X1 = dRect.Left
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y1 = dRect.Top
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRect
                        .X1 = 0
                        .Y1 = 0
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picGraphicSel.hwnd, ByVal (0)
                                    If GraphicSelX <= 3 And GraphicSelY <= 3 Then
                    Else
                        GraphicSelX = 0
                        GraphicSelY = 0
                    End If
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
            Case 2
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumTileSets Then
                                    If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
                        sRect.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRect.Right = sRect.Left + 800
                    Else
                        sRect.Left = 0
                        sRect.Right = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width
                        sRect.Left = frmEditor_Events.hScrlGraphicSel.Value = 0
                    End If
                                    If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
                        sRect.Top = frmEditor_Events.vScrlGraphicSel.Value
                        sRect.Bottom = sRect.Top + 512
                    Else
                        sRect.Top = 0
                        sRect.Bottom = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height
                        frmEditor_Events.vScrlGraphicSel.Value = 0
                    End If
                                    If sRect.Left = -1 Then sRect.Left = 0
                    If sRect.Top = -1 Then sRect.Top = 0
                                    With dRect
                        .Top = 0
                        .Bottom = sRect.Bottom - sRect.Top
                        .Left = 0
                        .Right = sRect.Right - sRect.Left
                    End With
                                                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRect
                            .X1 = (GraphicSelX * 32) - sRect.Left
                            .X2 = PIC_X + .X1
                            .Y1 = (GraphicSelY * 32) - sRect.Top
                            .Y2 = PIC_Y + .Y1
                        End With

                    Else
                        With destRect
                            .X1 = (GraphicSelX * 32) - sRect.Left
                            .X2 = ((GraphicSelX2 - GraphicSelX) * 32) + .X1
                            .Y1 = (GraphicSelY * 32) - sRect.Top
                            .Y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .Y1
                        End With
                    End If
                                    DrawSelectionBox destRect
                                    With srcRect
                        .X1 = dRect.Left
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y1 = dRect.Top
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRect
                        .X1 = 0
                        .Y1 = 0
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picGraphicSel.hwnd, ByVal (0)
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
        End Select
    Else
        If tmpEvent.pageCount > 0 Then
            Select Case tmpEvent.Pages(curPageNum).GraphicType
                Case 0
                    frmEditor_Events.picGraphic.Cls
                Case 1
                    If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumCharacters Then
                        sRect.Top = tmpEvent.Pages(curPageNum).GraphicY * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                        sRect.Left = tmpEvent.Pages(curPageNum).GraphicX * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                        sRect.Bottom = sRect.Top + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                        sRect.Right = sRect.Left + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                        With dRect
                            dRect.Top = (193 / 2) - ((sRect.Bottom - sRect.Top) / 2)
                            dRect.Bottom = dRect.Top + (sRect.Bottom - sRect.Top)
                            dRect.Left = (121 / 2) - ((sRect.Right - sRect.Left) / 2)
                            dRect.Right = dRect.Left + (sRect.Right - sRect.Left)
                        End With
                        With destRect
                            .X1 = dRect.Left
                            .X2 = dRect.Right
                            .Y1 = dRect.Top
                            .Y2 = dRect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hwnd, ByVal (0)
                    End If
                Case 2
                    If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumTileSets Then
                        If tmpEvent.Pages(curPageNum).GraphicX2 = 0 Or tmpEvent.Pages(curPageNum).GraphicY2 = 0 Then
                            sRect.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                            sRect.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                            sRect.Bottom = sRect.Top + 32
                            sRect.Right = sRect.Left + 32
                            With dRect
                                dRect.Top = (193 / 2) - ((sRect.Bottom - sRect.Top) / 2)
                                dRect.Bottom = dRect.Top + (sRect.Bottom - sRect.Top)
                                dRect.Left = (120 / 2) - ((sRect.Right - sRect.Left) / 2)
                                dRect.Right = dRect.Left + (sRect.Right - sRect.Left)
                            End With
                            With destRect
                                .X1 = dRect.Left
                                .X2 = dRect.Right
                                .Y1 = dRect.Top
                                .Y2 = dRect.Bottom
                            End With
                            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                            Direct3D_Device.BeginScene
                            RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                            Direct3D_Device.EndScene
                            Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hwnd, ByVal (0)
                        Else
                            sRect.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                            sRect.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                            sRect.Bottom = sRect.Top + ((tmpEvent.Pages(curPageNum).GraphicY2 - tmpEvent.Pages(curPageNum).GraphicY) * 32)
                            sRect.Right = sRect.Left + ((tmpEvent.Pages(curPageNum).GraphicX2 - tmpEvent.Pages(curPageNum).GraphicX) * 32)
                            With dRect
                                dRect.Top = (193 / 2) - ((sRect.Bottom - sRect.Top) / 2)
                                dRect.Bottom = dRect.Top + (sRect.Bottom - sRect.Top)
                                dRect.Left = (120 / 2) - ((sRect.Right - sRect.Left) / 2)
                                dRect.Right = dRect.Left + (sRect.Right - sRect.Left)
                            End With
                            With destRect
                                .X1 = dRect.Left
                                .X2 = dRect.Right
                                .Y1 = dRect.Top
                                .Y2 = dRect.Bottom
                            End With
                            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                            Direct3D_Device.BeginScene
                            RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                            Direct3D_Device.EndScene
                            Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hwnd, ByVal (0)
                        End If
                    End If
            End Select
        End If
    End If




   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorEvent_DrawGraphic", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
Public Sub EditorEvent_DrawGFX()
Dim sRect As rect, destRect As D3DRECT, srcRect As D3DRECT
Dim dRect As rect



   On Error GoTo errorhandler
   
    If frmEditor_Events.fraDialogue.Visible Then
        If frmEditor_Events.fraCommand(0).Visible Then
            If frmEditor_Events.scrlShowTextFace.Value > 0 Then
                If frmEditor_Events.scrlShowTextFace.Value <= NumFaces Then
                    If Tex_Face(frmEditor_Events.scrlShowTextFace.Value).filepath <> "" Then
                        If Tex_Face(frmEditor_Events.scrlShowTextFace.Value).Texture = 0 Then
                            LoadTexture1 Tex_Face(frmEditor_Events.scrlShowTextFace.Value)
                        Else
                            If gTexture(Tex_Face(frmEditor_Events.scrlShowTextFace.Value).Texture).IsLoaded = False Then
                                LoadTexture1 Tex_Face(frmEditor_Events.scrlShowTextFace.Value)
                            End If
                        End If
                        
                        With srcRect
                            .X1 = 0
                            .X2 = Tex_Face(frmEditor_Events.scrlShowTextFace.Value).Width
                            .Y1 = 0
                            .Y2 = Tex_Face(frmEditor_Events.scrlShowTextFace.Value).Height
                        End With
                        With destRect
                            .X1 = 0
                            .X2 = frmEditor_Events.picShowTextFace.ScaleWidth
                            .Y1 = 0
                            .Y2 = frmEditor_Events.picShowTextFace.ScaleHeight
                        End With
                        With sRect
                            .Left = 0
                            .Right = Tex_Face(frmEditor_Events.scrlShowTextFace.Value).Width
                            .Top = 0
                            .Bottom = Tex_Face(frmEditor_Events.scrlShowTextFace.Value).Height
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Face(frmEditor_Events.scrlShowTextFace.Value), sRect, sRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picShowTextFace.hwnd, ByVal (0)
                    Else
                        frmEditor_Events.picShowTextFace.Cls
                    End If
                Else
                    frmEditor_Events.picShowTextFace.Cls
                End If
            Else
                frmEditor_Events.picShowTextFace.Cls
            End If
        ElseIf frmEditor_Events.fraCommand(1).Visible Then
            If frmEditor_Events.scrlShowChoicesFace.Value > 0 Then
                If frmEditor_Events.scrlShowChoicesFace.Value <= NumFaces Then
                    If Tex_Face(frmEditor_Events.scrlShowChoicesFace.Value).filepath <> "" Then
                        If Tex_Face(frmEditor_Events.scrlShowChoicesFace.Value).Texture = 0 Then
                            LoadTexture1 Tex_Face(frmEditor_Events.scrlShowChoicesFace.Value)
                        Else
                            If gTexture(Tex_Face(frmEditor_Events.scrlShowChoicesFace.Value).Texture).IsLoaded = False Then
                                LoadTexture1 Tex_Face(frmEditor_Events.scrlShowChoicesFace.Value)
                            End If
                        End If
                        
                        With srcRect
                            .X1 = 0
                            .X2 = Tex_Face(frmEditor_Events.scrlShowChoicesFace.Value).Width
                            .Y1 = 0
                            .Y2 = Tex_Face(frmEditor_Events.scrlShowChoicesFace.Value).Height
                        End With
                        With destRect
                            .X1 = 0
                            .X2 = frmEditor_Events.picShowChoicesFace.ScaleWidth
                            .Y1 = 0
                            .Y2 = frmEditor_Events.picShowChoicesFace.ScaleHeight
                        End With
                        With sRect
                            .Left = 0
                            .Right = Tex_Face(frmEditor_Events.scrlShowChoicesFace.Value).Width
                            .Top = 0
                            .Bottom = Tex_Face(frmEditor_Events.scrlShowChoicesFace.Value).Height
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Face(frmEditor_Events.scrlShowChoicesFace.Value), sRect, sRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picShowChoicesFace.hwnd, ByVal (0)
                    Else
                        frmEditor_Events.picShowChoicesFace.Cls
                    End If
                Else
                    frmEditor_Events.picShowChoicesFace.Cls
                End If
            Else
                frmEditor_Events.picShowChoicesFace.Cls
            End If
        ElseIf frmEditor_Events.fraCommand(33).Visible Then
            If NumPics > 0 Then
                If frmEditor_Events.scrlShowPicture.Value > 0 Then
                    If frmEditor_Events.scrlShowPicture.Value <= NumPics Then
                        If Tex_Pic(frmEditor_Events.scrlShowPicture.Value).filepath <> "" Then
                            If Tex_Pic(frmEditor_Events.scrlShowPicture.Value).Texture = 0 Then
                                LoadTexture1 Tex_Pic(frmEditor_Events.scrlShowPicture.Value)
                            Else
                                If gTexture(Tex_Pic(frmEditor_Events.scrlShowPicture.Value).Texture).IsLoaded = False Then
                                    LoadTexture1 Tex_Pic(frmEditor_Events.scrlShowPicture.Value)
                                End If
                            End If
                            
                            With srcRect
                                .X1 = 0
                                .X2 = Tex_Pic(frmEditor_Events.scrlShowPicture.Value).Width
                                .Y1 = 0
                                .Y2 = Tex_Pic(frmEditor_Events.scrlShowPicture.Value).Height
                            End With
                            With destRect
                                .X1 = 0
                                .X2 = frmEditor_Events.picShowPicture.ScaleWidth
                                .Y1 = 0
                                .Y2 = frmEditor_Events.picShowPicture.ScaleHeight
                            End With
                            With sRect
                                .Left = 0
                                .Right = Tex_Pic(frmEditor_Events.scrlShowPicture.Value).Width
                                .Top = 0
                                .Bottom = Tex_Pic(frmEditor_Events.scrlShowPicture.Value).Height
                            End With
                            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                            Direct3D_Device.BeginScene
                            RenderTextureByRects Tex_Pic(frmEditor_Events.scrlShowPicture.Value), sRect, sRect
                            Direct3D_Device.EndScene
                            Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picShowPicture.hwnd, ByVal (0)
                        Else
                            frmEditor_Events.picShowPicture.Cls
                        End If
                    Else
                        frmEditor_Events.picShowPicture.Cls
                    End If
                Else
                    frmEditor_Events.picShowPicture.Cls
                End If
            End If
        End If
    End If

   

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "EditorEvent_DrawFaces", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawEvent(id As Long)
    Dim X As Long, Y As Long, Width As Long, Height As Long, sRect As rect, dRect As rect, Anim As Long, spritetop As Long

   On Error GoTo errorhandler

    If Map.MapEvents(id).Visible = 0 Then Exit Sub
    If InMapEditor Then Exit Sub
    Select Case Map.MapEvents(id).GraphicType
        Case 0
            Exit Sub
                Case 1
            If Map.MapEvents(id).GraphicNum <= 0 Or Map.MapEvents(id).GraphicNum > NumCharacters Then Exit Sub
            Width = Tex_Character(Map.MapEvents(id).GraphicNum).Width / 4
            Height = Tex_Character(Map.MapEvents(id).GraphicNum).Height / 4
            ' Reset frame
            If Map.MapEvents(id).Step = 3 Then
                Anim = 0
            ElseIf Map.MapEvents(id).Step = 1 Then
                Anim = 2
            End If
                    Select Case Map.MapEvents(id).dir
                Case DIR_UP
                    If (Map.MapEvents(id).YOffset > 8) Then Anim = Map.MapEvents(id).Step
                Case DIR_DOWN
                    If (Map.MapEvents(id).YOffset < -8) Then Anim = Map.MapEvents(id).Step
                Case DIR_LEFT
                    If (Map.MapEvents(id).XOffset > 8) Then Anim = Map.MapEvents(id).Step
                Case DIR_RIGHT
                    If (Map.MapEvents(id).XOffset < -8) Then Anim = Map.MapEvents(id).Step
            End Select
                    ' Set the left
            Select Case Map.MapEvents(id).ShowDir
                Case DIR_UP
                    spritetop = 3
                Case DIR_RIGHT
                    spritetop = 2
                Case DIR_DOWN
                    spritetop = 0
                Case DIR_LEFT
                    spritetop = 1
            End Select
                    If Map.MapEvents(id).WalkAnim = 1 Then Anim = 0
                    If Map.MapEvents(id).Moving = 0 Then Anim = Map.MapEvents(id).GraphicX
                    With sRect
                .Top = spritetop * Height
                .Bottom = .Top + Height
                .Left = Anim * Width
                .Right = .Left + Width
            End With
                ' Calculate the X
            X = Map.MapEvents(id).X * PIC_X + Map.MapEvents(id).XOffset - ((Width - 32) / 2)
                ' Is the player's height more than 32..?
            If (Height * 4) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                Y = Map.MapEvents(id).Y * PIC_Y + Map.MapEvents(id).YOffset - ((Height) - 32)
            Else
                ' Proceed as normal
                Y = Map.MapEvents(id).Y * PIC_Y + Map.MapEvents(id).YOffset
            End If
                ' render the actual sprite
            Call DrawSprite(Map.MapEvents(id).GraphicNum, X, Y, sRect)
                Case 2
            If Map.MapEvents(id).GraphicNum < 1 Or Map.MapEvents(id).GraphicNum > NumTileSets Then Exit Sub
                    If Map.MapEvents(id).GraphicY2 > 0 Or Map.MapEvents(id).GraphicX2 > 0 Then
                With sRect
                    .Top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .Top + ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) * 32)
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + ((Map.MapEvents(id).GraphicX2 - Map.MapEvents(id).GraphicX) * 32)
                End With
            Else
                With sRect
                    .Top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .Top + 32
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + 32
                End With
            End If
                    X = Map.MapEvents(id).X * 32
            Y = Map.MapEvents(id).Y * 32
                    X = X - ((sRect.Right - sRect.Left) / 2)
            Y = Y - (sRect.Bottom - sRect.Top) + 32
                            If Map.MapEvents(id).GraphicY2 > 0 Then
                RenderTexture Tex_Tileset(Map.MapEvents(id).GraphicNum), ConvertMapX(Map.MapEvents(id).X * 32), ConvertMapY((Map.MapEvents(id).Y - ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) - 1)) * 32), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
            Else
                RenderTexture Tex_Tileset(Map.MapEvents(id).GraphicNum), ConvertMapX(Map.MapEvents(id).X * 32), ConvertMapY(Map.MapEvents(id).Y * 32), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255), True
            End If
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawEvent", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

'This function will make it much easier to setup the vertices with the info it needs.
Private Function Create_TLVertex(X As Single, Y As Single, z As Single, RHW As Single, color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX


   On Error GoTo errorhandler

    Create_TLVertex.X = X
    Create_TLVertex.Y = Y
    Create_TLVertex.z = z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.color = color
    Create_TLVertex.TU = TU
    Create_TLVertex.TV = TV


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "Create_TLVertex", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function Ceiling(dblValIn As Double, dblCeilIn As Double) As Double
' round it

   On Error GoTo errorhandler

Ceiling = Round(dblValIn / dblCeilIn, 0) * dblCeilIn
' if it rounded down, force it up
If Ceiling < dblValIn Then Ceiling = Ceiling + dblCeilIn


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "Ceiling", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Sub DestroyDX8()

   On Error GoTo errorhandler

    UnloadTextures
    Set Direct3DX = Nothing
    Set Direct3D_Device = Nothing
    Set Direct3D = Nothing
    Set DirectX8 = Nothing


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DestroyDX8", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub DrawGDI()

   On Error GoTo errorhandler

    If frmEditor_Animation.Visible Then
        EditorAnim_DrawAnim
    End If
    If frmEditor_Item.Visible Then
        EditorItem_DrawItem
        EditorItem_DrawPaperdoll
        EditorItem_DrawFurniture
    End If
    If frmEditor_Pet.Visible Then
        EditorPet_DrawPet
    End If
    If frmEditor_Map.Visible Then
        EditorMap_DrawTileset
        If frmEditor_Map.fraMapItem.Visible Then EditorMap_DrawMapItem
        If frmEditor_Map.fraMapKey.Visible Then EditorMap_DrawKey
    End If
    If frmEditor_NPC.Visible Then
        EditorNpc_DrawSprite
    End If
    If frmEditor_Resource.Visible Then
        EditorResource_DrawSprite
    End If
    If frmEditor_Spell.Visible Then
        EditorSpell_DrawIcon
    End If
    If frmEditor_Events.Visible Then
        EditorEvent_DrawGraphic
        EditorEvent_DrawGFX
    End If
    If frmEditor_Projectile.Visible Then
        EditorProjectile_DrawProjectile
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawGDI", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   All of this code is for auto tiles and the math behind generating them.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub placeAutotile(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)

   On Error GoTo errorhandler
    If layerNum > MapLayer.Layer_Count - 1 Then
        layerNum = layerNum - (MapLayer.Layer_Count - 1)
        With Autotile(X, Y).ExLayer(layerNum).QuarterTile(tileQuarter)
            Select Case autoTileLetter
                Case "a"
                    .X = autoInner(1).X
                    .Y = autoInner(1).Y
                Case "b"
                    .X = autoInner(2).X
                    .Y = autoInner(2).Y
                Case "c"
                    .X = autoInner(3).X
                    .Y = autoInner(3).Y
                Case "d"
                    .X = autoInner(4).X
                    .Y = autoInner(4).Y
                Case "e"
                    .X = autoNW(1).X
                    .Y = autoNW(1).Y
                Case "f"
                    .X = autoNW(2).X
                    .Y = autoNW(2).Y
                Case "g"
                    .X = autoNW(3).X
                    .Y = autoNW(3).Y
                Case "h"
                    .X = autoNW(4).X
                    .Y = autoNW(4).Y
                Case "i"
                    .X = autoNE(1).X
                    .Y = autoNE(1).Y
                Case "j"
                    .X = autoNE(2).X
                    .Y = autoNE(2).Y
                Case "k"
                    .X = autoNE(3).X
                    .Y = autoNE(3).Y
                Case "l"
                    .X = autoNE(4).X
                    .Y = autoNE(4).Y
                Case "m"
                    .X = autoSW(1).X
                    .Y = autoSW(1).Y
                Case "n"
                    .X = autoSW(2).X
                    .Y = autoSW(2).Y
                Case "o"
                    .X = autoSW(3).X
                    .Y = autoSW(3).Y
                Case "p"
                    .X = autoSW(4).X
                    .Y = autoSW(4).Y
                Case "q"
                    .X = autoSE(1).X
                    .Y = autoSE(1).Y
                Case "r"
                    .X = autoSE(2).X
                    .Y = autoSE(2).Y
                Case "s"
                    .X = autoSE(3).X
                    .Y = autoSE(3).Y
                Case "t"
                    .X = autoSE(4).X
                    .Y = autoSE(4).Y
            End Select
        End With
    Else
        With Autotile(X, Y).Layer(layerNum).QuarterTile(tileQuarter)
            Select Case autoTileLetter
                Case "a"
                    .X = autoInner(1).X
                    .Y = autoInner(1).Y
                Case "b"
                    .X = autoInner(2).X
                    .Y = autoInner(2).Y
                Case "c"
                    .X = autoInner(3).X
                    .Y = autoInner(3).Y
                Case "d"
                    .X = autoInner(4).X
                    .Y = autoInner(4).Y
                Case "e"
                    .X = autoNW(1).X
                    .Y = autoNW(1).Y
                Case "f"
                    .X = autoNW(2).X
                    .Y = autoNW(2).Y
                Case "g"
                    .X = autoNW(3).X
                    .Y = autoNW(3).Y
                Case "h"
                    .X = autoNW(4).X
                    .Y = autoNW(4).Y
                Case "i"
                    .X = autoNE(1).X
                    .Y = autoNE(1).Y
                Case "j"
                    .X = autoNE(2).X
                    .Y = autoNE(2).Y
                Case "k"
                    .X = autoNE(3).X
                    .Y = autoNE(3).Y
                Case "l"
                    .X = autoNE(4).X
                    .Y = autoNE(4).Y
                Case "m"
                    .X = autoSW(1).X
                    .Y = autoSW(1).Y
                Case "n"
                    .X = autoSW(2).X
                    .Y = autoSW(2).Y
                Case "o"
                    .X = autoSW(3).X
                    .Y = autoSW(3).Y
                Case "p"
                    .X = autoSW(4).X
                    .Y = autoSW(4).Y
                Case "q"
                    .X = autoSE(1).X
                    .Y = autoSE(1).Y
                Case "r"
                    .X = autoSE(2).X
                    .Y = autoSE(2).Y
                Case "s"
                    .X = autoSE(3).X
                    .Y = autoSE(3).Y
                Case "t"
                    .X = autoSE(4).X
                    .Y = autoSE(4).Y
            End Select
        End With
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "placeAutotile", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub initAutotiles()
Dim X As Long, Y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    ' First, we need to re-size the array

   On Error GoTo errorhandler

    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).X = 32
    autoInner(1).Y = 0
    ' NE - b
    autoInner(2).X = 48
    autoInner(2).Y = 0
    ' SW - c
    autoInner(3).X = 32
    autoInner(3).Y = 16
    ' SE - d
    autoInner(4).X = 48
    autoInner(4).Y = 16
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).X = 0
    autoNW(1).Y = 32
    ' NE - f
    autoNW(2).X = 16
    autoNW(2).Y = 32
    ' SW - g
    autoNW(3).X = 0
    autoNW(3).Y = 48
    ' SE - h
    autoNW(4).X = 16
    autoNW(4).Y = 48
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).X = 32
    autoNE(1).Y = 32
    ' NE - g
    autoNE(2).X = 48
    autoNE(2).Y = 32
    ' SW - k
    autoNE(3).X = 32
    autoNE(3).Y = 48
    ' SE - l
    autoNE(4).X = 48
    autoNE(4).Y = 48
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).X = 0
    autoSW(1).Y = 64
    ' NE - n
    autoSW(2).X = 16
    autoSW(2).Y = 64
    ' SW - o
    autoSW(3).X = 0
    autoSW(3).Y = 80
    ' SE - p
    autoSW(4).X = 16
    autoSW(4).Y = 80
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).X = 32
    autoSE(1).Y = 64
    ' NE - r
    autoSE(2).X = 48
    autoSE(2).Y = 64
    ' SW - s
    autoSE(3).X = 32
    autoSE(3).Y = 80
    ' SE - t
    autoSE(4).X = 48
    autoSE(4).Y = 80
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                CalculateAutotile X, Y, layerNum
                ' cache the rendering state of the tiles and set them
                CacheRenderState X, Y, layerNum
            Next
            For layerNum = 1 To ExMapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                CalculateAutotile X, Y, layerNum + (MapLayer.Layer_Count - 1)
                ' cache the rendering state of the tiles and set them
                CacheRenderState X, Y, layerNum + (MapLayer.Layer_Count - 1)
            Next
        Next
    Next


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "initAutotiles", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CacheRenderState(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early

   On Error GoTo errorhandler

    If X < 0 Or X > Map.MaxX Or Y < 0 Or Y > Map.MaxY Then Exit Sub
    
    If layerNum > MapLayer.Layer_Count - 1 Then
        layerNum = layerNum - (MapLayer.Layer_Count - 1)
        With Map.exTile(X, Y)
            ' check if the tile can be rendered
            If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > NumTileSets Then
                Autotile(X, Y).ExLayer(layerNum).renderState = RENDER_STATE_NONE
                Exit Sub
            End If
                ' check if it needs to be rendered as an autotile
            If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Then
                ' default to... default
                Autotile(X, Y).ExLayer(layerNum).renderState = RENDER_STATE_NORMAL
            Else
                Autotile(X, Y).ExLayer(layerNum).renderState = RENDER_STATE_AUTOTILE
                ' cache tileset positioning
                For quarterNum = 1 To 4
                    Autotile(X, Y).ExLayer(layerNum).srcX(quarterNum) = (Map.exTile(X, Y).Layer(layerNum).X * 32) + Autotile(X, Y).ExLayer(layerNum).QuarterTile(quarterNum).X
                    Autotile(X, Y).ExLayer(layerNum).srcY(quarterNum) = (Map.exTile(X, Y).Layer(layerNum).Y * 32) + Autotile(X, Y).ExLayer(layerNum).QuarterTile(quarterNum).Y
                Next
            End If
        End With
    Else
        With Map.Tile(X, Y)
            ' check if the tile can be rendered
            If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > NumTileSets Then
                Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_NONE
                Exit Sub
            End If
                ' check if it's a key - hide mask if key is closed
            If layerNum = MapLayer.Mask Then
                If .type = TILE_TYPE_KEY Then
                    If TempTile(X, Y).DoorOpen = NO Then
                        Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_NONE
                        Exit Sub
                    Else
                        Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_NORMAL
                        Exit Sub
                    End If
                End If
            End If
                ' check if it needs to be rendered as an autotile
            If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Then
                ' default to... default
                Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_NORMAL
            Else
                Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_AUTOTILE
                ' cache tileset positioning
                For quarterNum = 1 To 4
                    Autotile(X, Y).Layer(layerNum).srcX(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).X * 32) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).X
                    Autotile(X, Y).Layer(layerNum).srcY(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).Y * 32) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).Y
                Next
            End If
        End With
    End If

    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CacheRenderState", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CalculateAutotile(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    ' Exit out if we don't have an auatotile

   On Error GoTo errorhandler
    If layerNum > MapLayer.Layer_Count - 1 Then
        If Map.exTile(X, Y).Autotile(layerNum - (MapLayer.Layer_Count - 1)) = 0 Then Exit Sub
        ' Okay, we have autotiling but which one?
        Select Case Map.exTile(X, Y).Autotile(layerNum - (MapLayer.Layer_Count - 1))
            ' Normal or animated - same difference
            Case AUTOTILE_NORMAL, AUTOTILE_ANIM
                ' North West Quarter
                CalculateNW_Normal layerNum, X, Y
                        ' North East Quarter
                CalculateNE_Normal layerNum, X, Y
                        ' South West Quarter
                CalculateSW_Normal layerNum, X, Y
                        ' South East Quarter
                CalculateSE_Normal layerNum, X, Y
                    ' Cliff
            Case AUTOTILE_CLIFF
                ' North West Quarter
                CalculateNW_Cliff layerNum, X, Y
                        ' North East Quarter
                CalculateNE_Cliff layerNum, X, Y
                        ' South West Quarter
                CalculateSW_Cliff layerNum, X, Y
                        ' South East Quarter
                CalculateSE_Cliff layerNum, X, Y
                    ' Waterfalls
            Case AUTOTILE_WATERFALL
                ' North West Quarter
                CalculateNW_Waterfall layerNum, X, Y
                        ' North East Quarter
                CalculateNE_Waterfall layerNum, X, Y
                        ' South West Quarter
                CalculateSW_Waterfall layerNum, X, Y
                        ' South East Quarter
                CalculateSE_Waterfall layerNum, X, Y
                ' Anything else
            Case Else
                ' Don't need to render anything... it's fake or not an autotile
        End Select
    Else
        If Map.Tile(X, Y).Autotile(layerNum) = 0 Then Exit Sub
        ' Okay, we have autotiling but which one?
        Select Case Map.Tile(X, Y).Autotile(layerNum)
            ' Normal or animated - same difference
            Case AUTOTILE_NORMAL, AUTOTILE_ANIM
                ' North West Quarter
                CalculateNW_Normal layerNum, X, Y
                        ' North East Quarter
                CalculateNE_Normal layerNum, X, Y
                        ' South West Quarter
                CalculateSW_Normal layerNum, X, Y
                        ' South East Quarter
                CalculateSE_Normal layerNum, X, Y
                    ' Cliff
            Case AUTOTILE_CLIFF
                ' North West Quarter
                CalculateNW_Cliff layerNum, X, Y
                        ' North East Quarter
                CalculateNE_Cliff layerNum, X, Y
                        ' South West Quarter
                CalculateSW_Cliff layerNum, X, Y
                        ' South East Quarter
                CalculateSE_Cliff layerNum, X, Y
                    ' Waterfalls
            Case AUTOTILE_WATERFALL
                ' North West Quarter
                CalculateNW_Waterfall layerNum, X, Y
                        ' North East Quarter
                CalculateNE_Waterfall layerNum, X, Y
                        ' South West Quarter
                CalculateSW_Waterfall layerNum, X, Y
                        ' South East Quarter
                CalculateSE_Waterfall layerNum, X, Y
                ' Anything else
            Case Else
                ' Don't need to render anything... it's fake or not an autotile
        End Select
    End If
    

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateAutotile", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateNW_Normal", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateNE_Normal", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateSW_Normal", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateSE_Normal", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    ' West

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 1, "e"
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateNW_Waterfall", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    ' East

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 2, "j"
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateNE_Waterfall", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    ' West

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 3, "g"
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateSW_Waterfall", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    ' East

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 4, "l"
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateSE_Waterfall", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    situation = AUTO_FILL
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateNW_Cliff", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    situation = AUTO_FILL
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateNE_Cliff", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    situation = AUTO_FILL
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateSW_Cliff", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South

   On Error GoTo errorhandler

    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    situation = AUTO_FILL
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CalculateSE_Cliff", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
Dim exTile As Boolean
   On Error GoTo errorhandler
   
    If layerNum > MapLayer.Layer_Count - 1 Then exTile = True: layerNum = layerNum - (MapLayer.Layer_Count - 1)
    checkTileMatch = True
    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MaxX Or Y2 < 0 Or Y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    If exTile Then
        ' fakes ALWAYS return true
        If Map.exTile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
            checkTileMatch = True
            Exit Function
        End If
    Else
        ' fakes ALWAYS return true
        If Map.Tile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
            checkTileMatch = True
            Exit Function
        End If
    End If
    If exTile Then
        ' check neighbour is an autotile
        If Map.exTile(X2, Y2).Autotile(layerNum) = 0 Then
            checkTileMatch = False
            Exit Function
        End If
    Else
        ' check neighbour is an autotile
        If Map.Tile(X2, Y2).Autotile(layerNum) = 0 Then
            checkTileMatch = False
            Exit Function
        End If
    End If
    If exTile Then
        ' check we're a matching
        If Map.exTile(X1, Y1).Layer(layerNum).Tileset <> Map.exTile(X2, Y2).Layer(layerNum).Tileset Then
            checkTileMatch = False
            Exit Function
        End If
        ' check tiles match
        If Map.exTile(X1, Y1).Layer(layerNum).X <> Map.exTile(X2, Y2).Layer(layerNum).X Then
            checkTileMatch = False
            Exit Function
        End If
        If Map.exTile(X1, Y1).Layer(layerNum).Y <> Map.exTile(X2, Y2).Layer(layerNum).Y Then
            checkTileMatch = False
            Exit Function
        End If
    Else
        ' check we're a matching
        If Map.Tile(X1, Y1).Layer(layerNum).Tileset <> Map.Tile(X2, Y2).Layer(layerNum).Tileset Then
            checkTileMatch = False
            Exit Function
        End If
        ' check tiles match
        If Map.Tile(X1, Y1).Layer(layerNum).X <> Map.Tile(X2, Y2).Layer(layerNum).X Then
            checkTileMatch = False
            Exit Function
        End If
        If Map.Tile(X1, Y1).Layer(layerNum).Y <> Map.Tile(X2, Y2).Layer(layerNum).Y Then
            checkTileMatch = False
            Exit Function
        End If
    End If


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "checkTileMatch", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long, Optional forceFrame As Long = 0, Optional strict As Boolean = True, Optional ExLayer As Boolean = False)
Dim YOffset As Long, XOffset As Long

    ' calculate the offset
    If forceFrame > 0 Then
        Select Case forceFrame - 1
            Case 0
                waterfallFrame = 1
            Case 1
                waterfallFrame = 2
            Case 2
                waterfallFrame = 0
        End Select
        ' animate autotiles
        Select Case forceFrame - 1
            Case 0
                autoTileFrame = 1
            Case 1
                autoTileFrame = 2
            Case 2
                autoTileFrame = 0
        End Select
    End If

   On Error GoTo errorhandler

    If ExLayer Then
        Select Case Map.exTile(X, Y).Autotile(layerNum)
            Case AUTOTILE_WATERFALL
                YOffset = (waterfallFrame - 1) * 32
            Case AUTOTILE_ANIM
                XOffset = autoTileFrame * 64
            Case AUTOTILE_CLIFF
                YOffset = -32
        End Select
        ' Draw the quarter
        'EngineRenderRectangle Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, 16, 16
        RenderTexture Tex_Tileset(Map.exTile(X, Y).Layer(layerNum).Tileset), destX, destY, Autotile(X, Y).ExLayer(layerNum).srcX(quarterNum) + XOffset, Autotile(X, Y).ExLayer(layerNum).srcY(quarterNum) + YOffset, 16, 16, 16, 16, -1, strict
    
    Else
    
        Select Case Map.Tile(X, Y).Autotile(layerNum)
            Case AUTOTILE_WATERFALL
                YOffset = (waterfallFrame - 1) * 32
            Case AUTOTILE_ANIM
                XOffset = autoTileFrame * 64
            Case AUTOTILE_CLIFF
                YOffset = -32
        End Select
        ' Draw the quarter
        'EngineRenderRectangle Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, 16, 16
        RenderTexture Tex_Tileset(Map.Tile(X, Y).Layer(layerNum).Tileset), destX, destY, Autotile(X, Y).Layer(layerNum).srcX(quarterNum) + XOffset, Autotile(X, Y).Layer(layerNum).srcY(quarterNum) + YOffset, 16, 16, 16, 16, -1, strict
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawAutoTile", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub DrawGUI()


   On Error GoTo errorhandler

    Call DrawGameGUI(True)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawGUI", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub DrawWalkArrow()

   On Error GoTo errorhandler

    If Options.ClicktoWalk = 1 Then
        If Map.Up > 0 Then
            If CurY = 0 Then
                RenderTexture Tex_Arrows, ConvertMapX(CurX * 32), ConvertMapY(0), 0, 0, 32, 32, 32, 32, -1, True
            End If
        End If
        If Map.Right > 0 Then
            If Map.MaxX = CurX Then
                RenderTexture Tex_Arrows, ConvertMapX(Map.MaxX * 32), ConvertMapY(CurY * 32), 32, 0, 32, 32, 32, 32, -1, True
            End If
        End If
        If Map.Down > 0 Then
            If CurY = Map.MaxY Then
                RenderTexture Tex_Arrows, ConvertMapX(CurX * 32), ConvertMapY(Map.MaxY * 32), 64, 0, 32, 32, 32, 32, -1, True
            End If
        End If
        If Map.Left > 0 Then
            If CurX = 0 Then
                RenderTexture Tex_Arrows, ConvertMapX(0 * 32), ConvertMapY((CurY) * 32), 96, 0, 32, 32, 32, 32, -1, True
            End If
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawWalkArrow", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub DrawPictures()
Dim i As Long, X As Long, Y As Long

   On Error GoTo errorhandler

    If NumPics > 0 Then
        For i = 1 To 10
            If Pictures(i).pic > 0 And Pictures(i).pic <= NumPics Then
                If Tex_Pic(Pictures(i).pic).filepath <> "" Then
                    If Tex_Pic(Pictures(i).pic).Texture = 0 Then
                        LoadTexture1 Tex_Pic(Pictures(i).pic)
                    Else
                        If gTexture(Tex_Pic(Pictures(i).pic).Texture).IsLoaded = False Then
                            LoadTexture1 Tex_Pic(Pictures(i).pic)
                        End If
                    End If
                    'Lets do some calculation!
                    Select Case Pictures(i).type
                        Case 1
                            X = X + Pictures(i).XOffset
                            Y = Y + Pictures(i).YOffset
                            RenderTexture Tex_Pic(Pictures(i).pic), X, Y, 0, 0, Tex_Pic(Pictures(i).pic).Width, Tex_Pic(Pictures(i).pic).Height, Tex_Pic(Pictures(i).pic).Width, Tex_Pic(Pictures(i).pic).Height, -1, False
                        Case 2
                            X = X - Tex_Pic(Pictures(i).pic).Width / 2
                            Y = Y - Tex_Pic(Pictures(i).pic).Height / 2
                            X = X + GameScreenWidth / 2
                            Y = Y + GameScreenHeight / 2
                            X = X + Pictures(i).XOffset
                            Y = Y + Pictures(i).YOffset
                            RenderTexture Tex_Pic(Pictures(i).pic), X, Y, 0, 0, Tex_Pic(Pictures(i).pic).Width, Tex_Pic(Pictures(i).pic).Height, Tex_Pic(Pictures(i).pic).Width, Tex_Pic(Pictures(i).pic).Height, -1, False
                        Case 3
                            X = ConvertMapX((Player(MyIndex).X * 32) + Player(MyIndex).XOffset)
                            Y = ConvertMapY((Player(MyIndex).Y * 32) + Player(MyIndex).YOffset)
                            X = X - Tex_Pic(Pictures(i).pic).Width / 2
                            Y = Y - Tex_Pic(Pictures(i).pic).Height / 2
                            X = X + Pictures(i).XOffset
                            Y = Y + Pictures(i).YOffset
                            RenderTexture Tex_Pic(Pictures(i).pic), X, Y, 0, 0, Tex_Pic(Pictures(i).pic).Width, Tex_Pic(Pictures(i).pic).Height, Tex_Pic(Pictures(i).pic).Width, Tex_Pic(Pictures(i).pic).Height, -1, False
                    End Select
                End If
            End If
        Next
    End If

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawPictures", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub ClearMapCache()
Dim i As Long, j As Long, X As Long, Y As Long
   On Error GoTo errorhandler
    If MapCacheX <> -1 And MapCacheY <> -1 Then
        For i = 1 To 2
            For j = 1 To 3
                For X = 0 To MapCacheX
                    For Y = 0 To MapCacheY
                        Set MapCache.Layers(i).Frame(j).MapTexRec(X, Y) = Nothing
                        Set MapCache.Layers(i).Frame(j).MapTexSurf(X, Y) = Nothing
                    Next
                Next
            Next
        Next
    End If
    
    CacheMap = False

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "ClearMapCache", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub CacheMapLayers()
Dim X As Long, Y As Long, MaxX As Long, MaxY As Long, i As Long, j As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
Dim xChunks As Long, yChunks As Long, xChunk As Long, yChunk As Long, u As Long
   On Error GoTo errorhandler
    Set BackBufferSurf = Direct3D_Device.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
    
    ClearMapCache
    
    MaxX = CInt((Map.MaxX * 32) / 256)
    MaxY = CInt((Map.MaxY * 32) / 256)
    
    MapCacheX = MaxX
    MapCacheY = MaxY
    
    For i = 1 To 1
        For j = 1 To 3
            ReDim MapCache.Layers(i).Frame(j).MapTexSurf(MaxX, MaxY)
            ReDim MapCache.Layers(i).Frame(j).MapTexRec(MaxX, MaxY)
            For X = 0 To MaxX
                For Y = 0 To MaxY
                    'Set MapCache.Layers(i).Frame(j).MapTexRec(X, Y) = Direct3DX.CreateTexture(Direct3D_Device, 256, 256, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
                    Set MapCache.Layers(i).Frame(j).MapTexRec(X, Y) = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, _
                                                    MapLayerImage.ImageData(0), _
                                                    UBound(MapLayerImage.ImageData) + 1, _
                                                    256, _
                                                    256, _
                                                    0, D3DUSAGE_RENDERTARGET, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
                    Set MapCache.Layers(i).Frame(j).MapTexSurf(X, Y) = MapCache.Layers(i).Frame(j).MapTexRec(X, Y).GetSurfaceLevel(0)
                Next
            Next
            
            For xChunk = 0 To MaxX
                For yChunk = 0 To MaxY
                    Direct3D_Device.SetRenderTarget MapCache.Layers(i).Frame(j).MapTexSurf(xChunk, yChunk), Nothing, 0
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    For u = 1 To MapLayer.Mask2
                        For X = 0 To 7
                            For Y = 0 To 7
                                If X + (xChunk * 8) <= Map.MaxX Then
                                    If Y + (yChunk * 8) <= Map.MaxY Then
                                        X1 = X + (xChunk * 8)
                                        Y1 = Y + (yChunk * 8)
                                        With Map.Tile(X1, Y1)
                                            If .Layer(u).Tileset > 0 Then
                                                If Autotile(X1, Y1).Layer(u).renderState = RENDER_STATE_NORMAL Then
                                                    RenderTexture Tex_Tileset(.Layer(u).Tileset), (X * 32), (Y * 32), .Layer(u).X * 32, .Layer(u).Y * 32, 32, 32, 32, 32, -1, False
                                                ElseIf Autotile(X1, Y1).Layer(u).renderState = RENDER_STATE_AUTOTILE Then
                                                    ' Draw autotiles
                                                    DrawAutoTile u, (X * 32), (Y * 32), 1, X1, Y1, j, False
                                                    DrawAutoTile u, (X * 32) + 16, (Y * 32), 2, X1, Y1, j, False
                                                    DrawAutoTile u, (X * 32), (Y * 32) + 16, 3, X1, Y1, j, False
                                                    DrawAutoTile u, (X * 32) + 16, (Y * 32) + 16, 4, X1, Y1, j, False
                                                End If
                                            End If
                                        End With
                                    End If
                                End If
                                DoEvents
                            Next
                        Next
                    Next
                    For u = ExMapLayer.Mask3 To ExMapLayer.Mask5
                        For X = 0 To 7
                            For Y = 0 To 7
                                If X + (xChunk * 8) <= Map.MaxX Then
                                    If Y + (yChunk * 8) <= Map.MaxY Then
                                        X1 = X + (xChunk * 8)
                                        Y1 = Y + (yChunk * 8)
                                        With Map.exTile(X1, Y1)
                                            If .Layer(u).Tileset > 0 Then
                                                If Autotile(X1, Y1).ExLayer(u).renderState = RENDER_STATE_NORMAL Then
                                                    RenderTexture Tex_Tileset(.Layer(u).Tileset), (X * 32), (Y * 32), .Layer(u).X * 32, .Layer(u).Y * 32, 32, 32, 32, 32, -1, False
                                                ElseIf Autotile(X1, Y1).ExLayer(u).renderState = RENDER_STATE_AUTOTILE Then
                                                    ' Draw autotiles
                                                    DrawAutoTile u, (X * 32), (Y * 32), 1, X1, Y1, j, False, True
                                                    DrawAutoTile u, (X * 32) + 16, (Y * 32), 2, X1, Y1, j, False, True
                                                    DrawAutoTile u, (X * 32), (Y * 32) + 16, 3, X1, Y1, j, False, True
                                                    DrawAutoTile u, (X * 32) + 16, (Y * 32) + 16, 4, X1, Y1, j, False, True
                                                End If
                                            End If
                                        End With
                                    End If
                                End If
                                DoEvents
                            Next
                        Next
                    Next
                    Direct3D_Device.EndScene
                Next
            Next
        Next
    Next
    
    For i = 2 To 2
        For j = 1 To 3
            ReDim MapCache.Layers(i).Frame(j).MapTexSurf(MaxX, MaxY)
            ReDim MapCache.Layers(i).Frame(j).MapTexRec(MaxX, MaxY)
            For X = 0 To MaxX
                For Y = 0 To MaxY
                    'Set MapCache.Layers(i).Frame(j).MapTexRec(X, Y) = Direct3DX.CreateTexture(Direct3D_Device, 256, 256, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
                    Set MapCache.Layers(i).Frame(j).MapTexRec(X, Y) = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, _
                                                    MapLayerImage.ImageData(0), _
                                                    UBound(MapLayerImage.ImageData) + 1, _
                                                    256, _
                                                    256, _
                                                    0, D3DUSAGE_RENDERTARGET, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
                    Set MapCache.Layers(i).Frame(j).MapTexSurf(X, Y) = MapCache.Layers(i).Frame(j).MapTexRec(X, Y).GetSurfaceLevel(0)
                Next
            Next
            
            For xChunk = 0 To MaxX
                For yChunk = 0 To MaxY
                    Direct3D_Device.SetRenderTarget MapCache.Layers(i).Frame(j).MapTexSurf(xChunk, yChunk), Nothing, 0
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    For u = MapLayer.Fringe To MapLayer.Fringe2
                        For X = 0 To 7
                            For Y = 0 To 7
                                If X + (xChunk * 8) <= Map.MaxX Then
                                    If Y + (yChunk * 8) <= Map.MaxY Then
                                        X1 = X + (xChunk * 8)
                                        Y1 = Y + (yChunk * 8)
                                        With Map.Tile(X1, Y1)
                                            If .Layer(u).Tileset > 0 Then
                                                If Autotile(X1, Y1).Layer(u).renderState = RENDER_STATE_NORMAL Then
                                                    RenderTexture Tex_Tileset(.Layer(u).Tileset), (X * 32), (Y * 32), .Layer(u).X * 32, .Layer(u).Y * 32, 32, 32, 32, 32, -1, False
                                                ElseIf Autotile(X1, Y1).Layer(u).renderState = RENDER_STATE_AUTOTILE Then
                                                    ' Draw autotiles
                                                    DrawAutoTile u, (X * 32), (Y * 32), 1, X1, Y1, j, False
                                                    DrawAutoTile u, (X * 32) + 16, (Y * 32), 2, X1, Y1, j, False
                                                    DrawAutoTile u, (X * 32), (Y * 32) + 16, 3, X1, Y1, j, False
                                                    DrawAutoTile u, (X * 32) + 16, (Y * 32) + 16, 4, X1, Y1, j, False
                                                End If
                                            End If
                                        End With
                                    End If
                                End If
                                DoEvents
                            Next
                        Next
                    Next
                    For u = ExMapLayer.Fringe3 To ExMapLayer.Fringe5
                        For X = 0 To 7
                            For Y = 0 To 7
                                If X + (xChunk * 8) <= Map.MaxX Then
                                    If Y + (yChunk * 8) <= Map.MaxY Then
                                        X1 = X + (xChunk * 8)
                                        Y1 = Y + (yChunk * 8)
                                        With Map.exTile(X1, Y1)
                                            If .Layer(u).Tileset > 0 Then
                                                If Autotile(X1, Y1).ExLayer(u).renderState = RENDER_STATE_NORMAL Then
                                                    RenderTexture Tex_Tileset(.Layer(u).Tileset), (X * 32), (Y * 32), .Layer(u).X * 32, .Layer(u).Y * 32, 32, 32, 32, 32, -1, False
                                                ElseIf Autotile(X1, Y1).ExLayer(u).renderState = RENDER_STATE_AUTOTILE Then
                                                    ' Draw autotiles
                                                    DrawAutoTile u, (X * 32), (Y * 32), 1, X1, Y1, j, False, True
                                                    DrawAutoTile u, (X * 32) + 16, (Y * 32), 2, X1, Y1, j, False, True
                                                    DrawAutoTile u, (X * 32), (Y * 32) + 16, 3, X1, Y1, j, False, True
                                                    DrawAutoTile u, (X * 32) + 16, (Y * 32) + 16, 4, X1, Y1, j, False, True
                                                End If
                                            End If
                                        End With
                                    End If
                                End If
                                DoEvents
                            Next
                        Next
                    Next
                    Direct3D_Device.EndScene
                Next
            Next
        Next
    Next
            
    Call Direct3D_Device.SetRenderTarget(BackBufferSurf, Nothing, 0)
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CacheGroundLayer", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Sub RenderLowerMap()
Dim X As Long, Y As Long, X1 As Long, Y1 As Long
Dim xChunk As Long, yChunk As Long, x1Chunk As Long, y1Chunk As Long
Dim X2 As Long, Y2 As Long, x3 As Long, y3 As Long, i As Long, j As Long

   On Error GoTo errorhandler
    'rendermapchunk(
    xChunk = Fix(TileView.Left / 8)
    yChunk = Fix(TileView.Top / 8)
    x1Chunk = Fix(TileView.Right / 8)
    y1Chunk = Fix(TileView.Bottom / 8)
    
    Select Case autoTileFrame
        Case 0
            j = 3
        Case 1
            j = 1
        Case 2
            j = 2
    End Select
    
    For i = 1 To 1
        For xChunk = Fix(TileView.Left / 8) To x1Chunk
            For yChunk = Fix(TileView.Top / 8) To y1Chunk
                X1 = 0
                Y1 = 0
                If (xChunk * 8) < TileView.Left Then X1 = TileView.Left - (xChunk * 8)
                If (yChunk * 8) < TileView.Top Then Y1 = TileView.Top - (yChunk * 8)
                x3 = 7 - X1
                y3 = 7 - Y1
                If (xChunk * 8) + X1 > TileView.Right Then x3 = ((xChunk * 8) + X1) - TileView.Right
                If (yChunk * 8) + Y1 > TileView.Bottom Then y3 = ((yChunk * 8) + Y1) - TileView.Bottom
                y3 = y3 + 1
                x3 = x3 + 1
                RenderMapChunk MapCache.Layers(i).Frame(j).MapTexRec(xChunk, yChunk), ConvertMapX((xChunk * 256) + (X1 * 32)), ConvertMapY((yChunk * 256) + (Y1 * 32)), X1 * 32, Y1 * 32, x3 * 32, y3 * 32, x3 * 32, y3 * 32
            Next
        Next
    Next

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "RenderLowerMap", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
Sub RenderUpperMap()
Dim X As Long, Y As Long, X1 As Long, Y1 As Long
Dim xChunk As Long, yChunk As Long, x1Chunk As Long, y1Chunk As Long
Dim X2 As Long, Y2 As Long, x3 As Long, y3 As Long, i As Long, j As Long

   On Error GoTo errorhandler
    'rendermapchunk(
    xChunk = Fix(TileView.Left / 8)
    yChunk = Fix(TileView.Top / 8)
    x1Chunk = Fix(TileView.Right / 8)
    y1Chunk = Fix(TileView.Bottom / 8)
    
    Select Case autoTileFrame
        Case 0
            j = 3
        Case 1
            j = 1
        Case 2
            j = 2
    End Select
    
    For i = 2 To 2
        For xChunk = Fix(TileView.Left / 8) To x1Chunk
            For yChunk = Fix(TileView.Top / 8) To y1Chunk
                X1 = 0
                Y1 = 0
                If (xChunk * 8) < TileView.Left Then X1 = TileView.Left - (xChunk * 8)
                If (yChunk * 8) < TileView.Top Then Y1 = TileView.Top - (yChunk * 8)
                x3 = 7 - X1
                y3 = 7 - Y1
                If (xChunk * 8) + X1 > TileView.Right Then x3 = ((xChunk * 8) + X1) - TileView.Right
                If (yChunk * 8) + Y1 > TileView.Bottom Then y3 = ((yChunk * 8) + Y1) - TileView.Bottom
                y3 = y3 + 1
                x3 = x3 + 1
                RenderMapChunk MapCache.Layers(i).Frame(j).MapTexRec(xChunk, yChunk), ConvertMapX((xChunk * 256) + (X1 * 32)), ConvertMapY((yChunk * 256) + (Y1 * 32)), X1 * 32, Y1 * 32, x3 * 32, y3 * 32, x3 * 32, y3 * 32
            Next
        Next
    Next

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "RenderUpperMap", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub DrawPet(ByVal Index As Long)
    Dim Anim As Byte, i As Long, X As Long, Y As Long
    Dim Sprite As Long, spritetop As Long
    Dim rec As rect
    Dim attackspeed As Long
    


   On Error GoTo errorhandler

    Sprite = Pet(Player(Index).Pet.Num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    attackspeed = 1000

    ' Reset frame
    If Player(Index).Pet.Step = 3 Then
        Anim = 0
    ElseIf Player(Index).Pet.Step = 1 Then
        Anim = 2
    End If
    
    ' Check for attacking animation
    If Player(Index).Pet.AttackTimer + (attackspeed / 2) > GetTickCount Then
        If Player(Index).Pet.Attacking = 1 Then
            Anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).Pet.YOffset > 8) Then Anim = Player(Index).Pet.Step
            Case DIR_DOWN
                If (Player(Index).Pet.YOffset < -8) Then Anim = Player(Index).Pet.Step
            Case DIR_LEFT
                If (Player(Index).Pet.XOffset > 8) Then Anim = Player(Index).Pet.Step
            Case DIR_RIGHT
                If (Player(Index).Pet.XOffset < -8) Then Anim = Player(Index).Pet.Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index).Pet
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case Player(Index).Pet.dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = spritetop * (Tex_Character(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Character(Sprite).Height / 4)
        .Left = Anim * (Tex_Character(Sprite).Width / 4)
        .Right = .Left + (Tex_Character(Sprite).Width / 4)
    End With

    ' Calculate the X
    X = Player(Index).Pet.X * PIC_X + Player(Index).Pet.XOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = Player(Index).Pet.Y * PIC_Y + Player(Index).Pet.YOffset - ((Tex_Character(Sprite).Width / 4) - 32)
    Else
        ' Proceed as normal
        Y = Player(Index).Pet.Y * PIC_Y + Player(Index).Pet.YOffset
    End If

    ' render the actual sprite
    Call DrawSprite(Sprite, X, Y, rec)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawPet", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub

Public Sub DrawProjectile(ByVal ProjectileNum As Long)
Dim rec As rect
Dim CanClearProjectile As Boolean
Dim CollisionIndex As Long
Dim CollisionType As Byte
Dim CollisionZone As Long
Dim XOffset As Long, YOffset As Long
Dim X As Long, Y As Long
Dim i As Long, z As Long
Dim Sprite As Long

    ' make sure it's not out of map

   On Error GoTo errorhandler
   
    ' check to see if it's time to move the Projectile
    If GetTickCount > MapProjectiles(ProjectileNum).TravelTime Then
        Select Case MapProjectiles(ProjectileNum).dir
            Case DIR_UP
                MapProjectiles(ProjectileNum).Y = MapProjectiles(ProjectileNum).Y - 1
            Case DIR_DOWN
                MapProjectiles(ProjectileNum).Y = MapProjectiles(ProjectileNum).Y + 1
            Case DIR_LEFT
                MapProjectiles(ProjectileNum).X = MapProjectiles(ProjectileNum).X - 1
            Case DIR_RIGHT
                MapProjectiles(ProjectileNum).X = MapProjectiles(ProjectileNum).X + 1
        End Select
        MapProjectiles(ProjectileNum).TravelTime = GetTickCount + Projectiles(MapProjectiles(ProjectileNum).ProjectileNum).speed
        MapProjectiles(ProjectileNum).Range = MapProjectiles(ProjectileNum).Range + 1
    End If
   
    X = MapProjectiles(ProjectileNum).X
    Y = MapProjectiles(ProjectileNum).Y
    
    'Check if its been going for over 1 minute, if so clear.
    If MapProjectiles(ProjectileNum).Timer < GetTickCount Then CanClearProjectile = True

    If X > Map.MaxX Or X < 0 Then CanClearProjectile = True
    If Y > Map.MaxY Or Y < 0 Then CanClearProjectile = True
    
    'Check for blocked wall collision
    If CanClearProjectile = False Then 'Add a check to prevent crashing
        If Map.Tile(X, Y).type = TILE_TYPE_BLOCKED Then CanClearProjectile = True
    End If
    
    'Check for npc collision
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).X = X And MapNpc(i).Y = Y Then
            CanClearProjectile = True
            CollisionIndex = i
            CollisionType = TARGET_TYPE_NPC
            CollisionZone = -1
            Exit For
        End If
    Next
    
    For i = 1 To MAX_ZONES
        For z = 1 To MAX_MAP_NPCS * 2
            If ZoneNPC(i).Npc(z).Num > 0 Then
                If ZoneNPC(i).Npc(z).Map = GetPlayerMap(MyIndex) Then
                    If ZoneNPC(i).Npc(z).X = X And ZoneNPC(i).Npc(z).Y = Y Then
                        CanClearProjectile = True
                        CollisionIndex = z
                        CollisionType = TARGET_TYPE_ZONENPC
                        CollisionZone = i
                        Exit For
                    End If
                End If
            End If
        Next
        If CollisionZone > 0 Then
            Exit For
        End If
    Next i
    
    'Check for pet and player collision
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                CanClearProjectile = True
                CollisionIndex = i
                CollisionType = TARGET_TYPE_PLAYER
                CollisionZone = -1
                If MapProjectiles(ProjectileNum).OwnerType = TARGET_TYPE_PLAYER Or MapProjectiles(ProjectileNum).OwnerType = TARGET_TYPE_PET Then
                    If MapProjectiles(ProjectileNum).Owner = i Then CanClearProjectile = False ' Reset if its the owner of projectile
                End If
                Exit For
            End If
            
            If Player(i).Pet.Alive = True Then
                If Player(i).Pet.X = X And Player(i).Pet.Y = Y Then
                    CanClearProjectile = True
                    CollisionIndex = i
                    CollisionType = TARGET_TYPE_PET
                    CollisionZone = -1
                    If MapProjectiles(ProjectileNum).OwnerType = TARGET_TYPE_PLAYER Or MapProjectiles(ProjectileNum).OwnerType = TARGET_TYPE_PET Then
                        If MapProjectiles(ProjectileNum).Owner = i Then CanClearProjectile = False ' Reset if its the owner of projectile
                    End If
                    Exit For
                End If
            End If
        End If
    Next
    
    'Check if it has hit its maximum range
    If MapProjectiles(ProjectileNum).Range >= Projectiles(MapProjectiles(ProjectileNum).ProjectileNum).Range + 1 Then CanClearProjectile = True
    
    'Clear the projectile if possible
    If CanClearProjectile = True Then
        'Only send the clear to the server if you're the projectile caster or the one hit (only if owner is not a player)
        If (MapProjectiles(ProjectileNum).OwnerType = TARGET_TYPE_PLAYER And MapProjectiles(ProjectileNum).Owner = MyIndex) _
        Or (MapProjectiles(ProjectileNum).OwnerType = TARGET_TYPE_PET And MapProjectiles(ProjectileNum).Owner = MyIndex) Then
            SendClearProjectile ProjectileNum, CollisionIndex, CollisionType, CollisionZone
        End If
        
        ClearMapProjectile ProjectileNum
        Exit Sub
    End If
    
    Sprite = Projectiles(MapProjectiles(ProjectileNum).ProjectileNum).Sprite
    If Sprite < 1 Or Sprite > NumProjectiles Then Exit Sub

    ' src rect
    With rec
        .Top = 0
        .Bottom = Tex_Projectiles(Sprite).Height
        .Left = MapProjectiles(ProjectileNum).dir * PIC_X
        .Right = .Left + PIC_X
    End With
    
    'Find the offset
    Select Case MapProjectiles(ProjectileNum).dir
        Case DIR_UP
            YOffset = ((MapProjectiles(ProjectileNum).TravelTime - GetTickCount) / Projectiles(MapProjectiles(ProjectileNum).ProjectileNum).speed) * PIC_Y
        Case DIR_DOWN
            YOffset = -((MapProjectiles(ProjectileNum).TravelTime - GetTickCount) / Projectiles(MapProjectiles(ProjectileNum).ProjectileNum).speed) * PIC_Y
        Case DIR_LEFT
            XOffset = ((MapProjectiles(ProjectileNum).TravelTime - GetTickCount) / Projectiles(MapProjectiles(ProjectileNum).ProjectileNum).speed) * PIC_X
        Case DIR_RIGHT
            XOffset = -((MapProjectiles(ProjectileNum).TravelTime - GetTickCount) / Projectiles(MapProjectiles(ProjectileNum).ProjectileNum).speed) * PIC_X
    End Select

    X = ConvertMapX(X * PIC_X)
    Y = ConvertMapY(Y * PIC_Y)
    RenderTexture Tex_Projectiles(Sprite), X + XOffset, Y + YOffset, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255), True

   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "DrawProjectile", "modGraphics", Err.Number, Err.Description, Erl
    Err.Clear

End Sub
