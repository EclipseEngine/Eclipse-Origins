VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIpRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Credit for these GDI+ classes go to LaVolpe
'http://www.vbforums.com/showthread.php?t=598771
Option Explicit

' HOW TO USE THIS CLASS
' 1. Simply, instantiate a single copy for the life of your project
' 2. The cGDIpToken class must also be instantiated for any of these GDI+ classes
' 3. Before most functions can be called, you must first call the AttachTokenClass method
' EXCEPTION: The following class methods that use a cGDIpImage class
'   do not require the AttachTokenClass method to be called in advance nor at all
'   :: CreateGraphicsFromImageClass, PlgBltImageClassToDC
'   :: RenderImageClassToDC, RenderImageClassToHGraphics

' This class can be created and destroyed on demand if you prefer
' This class wraps some of the common GDI+ graphics calls.
' There are over a 100 API calls in GDI+, add additional ones as needed
' ToDo: Add sample Region & Path calls in future updates


Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Interpolation As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hwnd As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipDrawImagePointsRectI Lib "gdiplus" (ByVal graphics As Long, ByVal pImage As Long, ByRef pPoints As Any, ByVal count As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, Optional ByVal pcallback As Long = 0&, Optional ByVal callbackData As Long = 0&) As Long
Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal graphics As Long, ByVal pColor As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef graphics As Long) As Long
Private Declare Function GdipSetCompositingQuality Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal compositingQuality As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long) As Long

Private Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GdipDrawLinesI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal count As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal count As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByRef pPoints As Any, ByVal count As Long, ByVal FillMode As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipDrawRectanglesI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef rects As Any, ByVal count As Long) As Long
Private Declare Function GdipFillRectanglesI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByRef rects As Any, ByVal count As Long) As Long

Private Const UnitPixel As Long = &H2&

Private Type POINTL
    X As Long
    Y As Long
End Type

Public Enum FillModeConstants
    FillModeAlternate = 0
    FillModeWinding = 1
End Enum
Public Enum InterpolationConstants
    InterpolationModeDefault = 0
    InterpolationModeLowQuality = 1
    InterpolationModeHighQuality = 2
    InterpolationModeBilinear = 3
    InterpolationModeBicubic = 4
    InterpolationModeNearestNeighbor = 5
    InterpolationModeHighQualityBilinear = 6
    InterpolationModeHighQualityBicubic = 7
End Enum
Public Enum SmoothingConstants
    SmoothingModeDefault = 0
    SmoothingModeHighSpeed = 1
    SmoothingModeHighQuality = 2
    SmoothingModeNone = 3
    SmoothingModeAntiAlias = 4
End Enum
Public Enum CompositingModeConstants
    CompositingQualityDefault = 0
    CompositingQualityHighSpeed = 1
    CompositingQualityHighQuality = 2
    CompositingQualityGammaCorrected = 3
    CompositingQualityAssumeLinear = 4
End Enum
Public Enum PixelOffsetModeConstants
    PixelOffsetModeDefault = 0
    PixelOffsetModeHighSpeed = 1
    PixelOffsetModeHighQuality = 2
    PixelOffsetModeNone = 3
    PixelOffsetModeHalf = 4
End Enum

Private m_Token As cGDIpToken

' MUST BE CALLED BEFORE ANY CLASS FUNCTIONS ARE CALLED
' EXCEPTION: The following class methods that use a cGDIpImage class do not
'   require the AttachTokenClass method to be called in advance nor at all
'   :: CreateGraphicsFromImageClass, PlgBltImageClassToDC,
'   :: RenderImageClassToDC, RenderImageClassToHGraphics
Public Sub AttachTokenClass(TokenClass As cGDIpToken)
    If m_Token Is Nothing Then
        If TokenClass.Token Then Set m_Token = TokenClass
    End If
End Sub

Public Function CreateGraphicsFromImageClass(ImageClass As cGDIpImage) As Long
    If Not ImageClass Is Nothing Then
        If ImageClass.Handle Then Call GdipGetImageGraphicsContext(ImageClass.Handle, CreateGraphicsFromImageClass)
    End If
End Function

Public Function CreateHGraphicsFromHDC(ByVal hDC As Long) As Long
    If Not (hDC = 0& Or (m_Token Is Nothing)) Then
        If m_Token.Token Then Call GdipCreateFromHDC(hDC, CreateHGraphicsFromHDC)
    End If
End Function

Public Function CreateHGraphicsFromHWND(ByVal hwnd As Long) As Long
    If Not (hwnd = 0& Or (m_Token Is Nothing)) Then
        If m_Token.Token Then Call GdipCreateFromHWND(hwnd, CreateHGraphicsFromHWND)
    End If
End Function

Public Function DestroyHGraphics(ByVal hGraphics As Long) As Boolean
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If m_Token.Token Then DestroyHGraphics = (GdipDeleteGraphics(hGraphics) = 0&)
    End If
End Function

Public Function DrawEllipse(ByVal hGraphics As Long, _
                            ByVal X As Long, ByVal Y As Long, _
                            ByVal Width As Long, ByVal Height As Long, _
                            Optional OutlinePenHandle As Long = 0&, _
                            Optional ARGBFillColor As Long = 0&, Optional ByVal FillBrushHandle As Long = 0&) As Boolean
                            
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If m_Token.Token Then
            If FillBrushHandle Then
                DrawEllipse = GdipFillEllipseI(hGraphics, FillBrushHandle, X, Y, Width, Height)
            ElseIf ARGBFillColor Then
                Call GdipCreateSolidFill(ARGBFillColor, FillBrushHandle)
                If FillBrushHandle Then
                    DrawEllipse = GdipFillEllipseI(hGraphics, FillBrushHandle, X, Y, Width, Height)
                    GdipDeleteBrush FillBrushHandle
                End If
            End If
            If OutlinePenHandle Then DrawEllipse = (GdipDrawEllipseI(hGraphics, OutlinePenHandle, X, Y, Width, Height) = 0&)
        End If
    End If
End Function

Public Function DrawLine(ByVal hGraphics As Long, ByVal hPen As Long, _
                        ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
                        
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If Not (m_Token.Token = 0& Or hPen = 0&) Then
            DrawLine = (GdipDrawLineI(hGraphics, hPen, X1, Y1, X2, Y2) = 0&)
        End If
    End If
End Function

Public Function DrawLines(ByVal hGraphics As Long, ByVal hPen As Long, _
                            ByRef PointsVarPtr As Long, ByVal PointCount As Long) As Boolean
    ' draws multiple lines at once
    ' How to use PointsVarPtr...
    ' your uPoints() array should be a 1D array of POINTAPI structures or a 2D array of longs with X,Y in the lowerbound
    ' i.e., a simple rect, 4Pts:    Dim uPoints(0 to 4) As POINTAPI
    '                               ... X values              ... Y Values
    '                               uPoints(0) = Left:       uPoints(0) = Top         ' from Top Left
    '                               uPoints(1) = Width+Left: uPoints(1) = Top         ' to Top Right
    '                               uPoints(2) = Width+Left: uPoints(2) = Height+Top  ' to Bottom Right
    '                               uPoints(3) = Left:       uPoints(3) = Height+Top  ' to Bottom Left
    '                               uPoints(4) = Left:       uPoints(4) = Top         ' to Top Left
    ' To draw the rect, you would pass PointsVarPtr as VarPtr(uPoints(0)) & PointCount as 5
    ' To draw just the right & bottom edges of rect: PointsVarPtr as VarPtr(uPoints(1)) & PointCount as 3
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If Not (m_Token.Token = 0& Or hPen = 0&) Then
            If Not (PointsVarPtr = 0& Or PointCount = 0&) Then DrawLines = (GdipDrawLinesI(hGraphics, hPen, ByVal PointsVarPtr, PointCount) = 0&)
        End If
    End If
End Function

Public Function DrawPolygon(ByVal hGraphics As Long, _
                            ByRef PointsVarPtr As Long, ByVal PointCount As Long, _
                            Optional FillMode As FillModeConstants = FillModeWinding, _
                            Optional OutlinePenHandle As Long = 0&, _
                            Optional ARGBFillColor As Long = 0&, Optional FillBrushHandle As Long) As Boolean
    ' Draws a polygon. This is slightly different than using DrawLines (multiple lines)
    '   to draw a polygon because it will automatically close the polygon
    ' How to use PointsVarPtr...
    ' your uPoints() array should be a 1D array of POINTAPI structures or a 2D array of longs with X,Y in the lowerbound
    ' i.e., a simple rect, 4Pts:    Dim uPoints(0 to 3) As POINTAPI
    '                               ... X values              ... Y Values
    '                               uPoints(0) = Left:       uPoints(0) = Top         ' from Top Left
    '                               uPoints(1) = Width+Left: uPoints(1) = Top         ' to Top Right
    '                               uPoints(2) = Width+Left: uPoints(2) = Height+Top  ' to Bottom Right
    '                               uPoints(3) = Left:       uPoints(3) = Height+Top  ' to Bottom Left
    '                                                                                 ' automatically closes to Top Left
    ' To draw the rect, you would pass PointsVarPtr as VarPtr(uPoints(0)) & PointCount as 4
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If Not (PointsVarPtr = 0& Or PointCount = 0&) Then
            If m_Token.Token Then
                If FillBrushHandle Then
                    DrawPolygon = (GdipFillPolygonI(hGraphics, FillBrushHandle, ByVal PointsVarPtr, PointCount, FillMode) = 0&)
                ElseIf ARGBFillColor Then
                    Call GdipCreateSolidFill(ARGBFillColor, FillBrushHandle)
                    If FillBrushHandle Then
                        DrawPolygon = (GdipFillPolygonI(hGraphics, FillBrushHandle, ByVal PointsVarPtr, PointCount, FillMode) = 0&)
                        GdipDeleteBrush FillBrushHandle
                    End If
                End If
                If OutlinePenHandle Then DrawPolygon = (GdipDrawPolygonI(hGraphics, OutlinePenHandle, ByVal PointsVarPtr, PointCount) = 0&)
            End If
        End If
    End If
End Function

Public Function DrawRectangle(ByVal hGraphics As Long, _
                                ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, _
                                Optional ByVal OutlinePenHandle As Long = 0&, _
                                Optional ByVal ARGBFillColor As Long = 0&, Optional ByVal FillBrushHandle As Long = 0&) As Boolean
                                
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If m_Token.Token Then
            If FillBrushHandle Then
                DrawRectangle = (GdipFillRectangleI(hGraphics, FillBrushHandle, Left, Top, Width, Height) = 0&)
            ElseIf ARGBFillColor Then
                Call GdipCreateSolidFill(ARGBFillColor, FillBrushHandle)
                If FillBrushHandle Then
                    DrawRectangle = (GdipFillRectangleI(hGraphics, FillBrushHandle, Left, Top, Width, Height) = 0&)
                    GdipDeleteBrush FillBrushHandle
                End If
            End If
            If OutlinePenHandle Then DrawRectangle = (GdipDrawRectangleI(hGraphics, OutlinePenHandle, Left, Top, Width, Height) = 0&)
        End If
    End If
End Function

Public Function DrawRectangles(ByVal hGraphics As Long, _
                                ByRef PointsVarPtr As Long, ByVal PointCount As Long, _
                                Optional OutlinePenHandle As Long = 0&, _
                                Optional ARGBFillColor As Long = 0&, Optional FillBrushHandle As Long = 0&) As Boolean
    ' draws multiple rectangles at once
    ' How to use PointsVarPtr...
    ' your uRECT() array should be 1D array containing RECT structures: ReDim uRECT(0 to NrOfRects-1)
    ' To draw the rects, you would pass PointsVarPtr as VarPtr(uRect(0)) & PointCount as NrOfRects
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If Not (PointsVarPtr = 0& Or PointCount = 0&) Then
            If m_Token.Token Then
                If FillBrushHandle Then
                    DrawRectangles = (GdipFillRectanglesI(hGraphics, FillBrushHandle, ByVal PointsVarPtr, PointCount) = 0&)
                ElseIf ARGBFillColor Then
                    Call GdipCreateSolidFill(ARGBFillColor, FillBrushHandle)
                    If FillBrushHandle Then
                        DrawRectangles = (GdipDrawRectanglesI(hGraphics, FillBrushHandle, ByVal PointsVarPtr, PointCount) = 0&)
                        GdipDeleteBrush FillBrushHandle
                    End If
                End If
                If OutlinePenHandle Then DrawRectangles = (GdipDrawRectanglesI(hGraphics, OutlinePenHandle, ByVal PointsVarPtr, PointCount) = 0&)
            End If
        End If
    End If
End Function

Public Function EraseHGraphics(ByVal hGraphics As Long, ByVal ARGBColor As Long) As Boolean
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If m_Token.Token Then EraseHGraphics = (GdipGraphicsClear(hGraphics, ARGBColor) = 0&)
    End If
End Function

Public Function PlgBltImageClassToDC(ImageClass As cGDIpImage, ByVal hDC As Long, _
                        ByVal SkewedTopLeftX As Long, ByVal SkewedTopLeftY As Long, _
                        ByVal SkewedTopRightX As Long, ByVal SkewedTopRightY As Long, _
                        ByVal SkewedBottomLeftX As Long, ByVal SkewedBottomLeftY As Long, _
                        Optional ByVal srcX As Long = 0&, Optional ByVal srcY As Long = 0&, _
                        Optional ByVal srcWidth As Long = 0&, Optional ByVal srcHeight As Long = 0&, _
                        Optional ByVal InterpolationMode As InterpolationConstants = InterpolationModeDefault, _
                        Optional ByVal SmoothingMode As SmoothingConstants = SmoothingModeHighSpeed, _
                        Optional ByVal CompositeMode As CompositingModeConstants = CompositingQualityHighSpeed, _
                        Optional ByVal PixelOffsetMode As PixelOffsetModeConstants = PixelOffsetModeHighSpeed, _
                        Optional ByVal EraseBkgARGBColor As Long = 0&, _
                        Optional ByVal EraseBkgBrushHandle As Long = 0&) As Boolean

    ' Polygon BLT of image.
    ' Does not support Rotation and brushes are not skewed
    ' -- if these effects are desired, you should render your image to a DC containing a 32bpp DIB,
    '       then load the DIB to an cGDIpImage class and send that class to this function

    ' For the various Modes, recommend googling a bit for the differences. Here is one: http://www.geekpedia.com/tutorial52_Antialiasing-or-Smoothing.html
    ' Parameters
    '   :: ImageClass - instantiated cGDIpImage class with a loaded image
    '   :: hDC - the target DC to render to
    '   :: SkewedTopLeftX,SkewedTopLeftY - the target top/left coords to render at
    '   :: SkewedTopRightX,SkewedTopRightY - the target top/right coords to render at
    '   :: SkewedBottomLeftX,SkewedBottomLeftY - the target bottom/left coords to render at
    '   :: srcx,srcy - the image top/left coords to render from
    '   :: srcWidth,srcHeight - the size of the image to be rendered
    '   :: InterpolationMode - end quality when scaling/rotating during rendering
    '   :: SmoothingMode - end quality of the rendering itself
    '   :: CompositeMode - how colors are rendered with background colors
    '   :: PixelOffsetMode - how pixels will be shifted to produce higher quality images
    '   :: EraseBkgARGBColor - ignored if EraseBkgBrushHandle is not null; fill color of erasing background
    '   :: EraseBkgBrushHandle - the brush used to erase the background
    
    ' Note. Clarification of the Skewedxxx parameters.
    '   The parameters are where the skewed/rotated points would end up.
    '   For example. Let's say you were rotating a square image 180 degrees
    '       The SkewedTopLeft X,Y would be the image's bottom right corner
    '       The SkewedTopRight X,Y would be the images bottom left corner
    '       The SkewedBottomLeft X,Y would be the image's top right corner
    '       The SkewedBottomRight X,Y is calculated by the API; therefore is never passed

    Dim hGraphics As Long, hBrush As Long, plgPts(0 To 2) As POINTL
    Dim plgFillPts() As POINTL
    
    If hDC = 0& Or (ImageClass Is Nothing) Then Exit Function
    If ImageClass.Handle = 0& Then Exit Function

    GdipCreateFromHDC hDC, hGraphics
    If hGraphics = 0& Then Exit Function

    If srcWidth = 0& Then srcWidth = ImageClass.Width
    If srcHeight = 0& Then srcHeight = ImageClass.Height
    plgPts(0).X = SkewedTopLeftX: plgPts(0).Y = SkewedTopLeftY
    plgPts(1).X = SkewedTopRightX: plgPts(1).Y = SkewedTopRightY
    plgPts(2).X = SkewedBottomLeftX: plgPts(2).Y = SkewedBottomLeftY
    
    If EraseBkgBrushHandle Then
        hBrush = EraseBkgBrushHandle
    ElseIf EraseBkgARGBColor Then
        Call GdipCreateSolidFill(EraseBkgARGBColor, hBrush)
    End If
    
    If hBrush Then
        ReDim plgFillPts(0 To 3)
        plgFillPts(0) = plgPts(2)
        plgFillPts(1) = plgPts(0)
        plgFillPts(2) = plgPts(1)
        ' Since the last corner is never passed, we calculate it here
        plgFillPts(3).X = SkewedTopRightX + (SkewedBottomLeftX - SkewedTopLeftX)
        plgFillPts(3).Y = SkewedTopRightY + (SkewedBottomLeftY - SkewedTopLeftY)
        Call GdipFillPolygonI(hGraphics, hBrush, ByVal VarPtr(plgFillPts(0)), 4&, FillModeWinding)
    End If
    
    Call GdipSetSmoothingMode(hGraphics, SmoothingMode)
    Call GdipSetInterpolationMode(hGraphics, InterpolationMode)
    Call GdipSetCompositingQuality(hGraphics, CompositeMode)
    Call GdipSetPixelOffsetMode(hGraphics, PixelOffsetMode)

    PlgBltImageClassToDC = (GdipDrawImagePointsRectI(hGraphics, ImageClass.Handle, ByVal VarPtr(plgPts(0)), 3&, srcX, srcY, srcWidth, srcHeight, UnitPixel, ImageClass.ImageAttributesHandle) = 0&)
    
    If hBrush Then
        If EraseBkgBrushHandle <> hBrush Then GdipDeleteBrush hBrush
    End If
    GdipDeleteGraphics hGraphics

End Function

Public Function RenderImageClassToDC(ImageClass As cGDIpImage, ByVal hDC As Long, _
                        Optional ByVal destX As Long, Optional ByVal destY As Long, _
                        Optional ByVal destWidth As Long = 0&, Optional ByVal destHeight As Long = 0&, _
                        Optional ByVal srcX As Long = 0&, Optional ByVal srcY As Long = 0&, _
                        Optional ByVal srcWidth As Long = 0&, Optional ByVal srcHeight As Long = 0&, _
                        Optional ByVal InterpolationMode As InterpolationConstants = InterpolationModeDefault, _
                        Optional ByVal SmoothingMode As SmoothingConstants = SmoothingModeHighSpeed, _
                        Optional ByVal CompositeMode As CompositingModeConstants = CompositingQualityHighSpeed, _
                        Optional ByVal PixelOffsetMode As PixelOffsetModeConstants = PixelOffsetModeHighSpeed, _
                        Optional ByVal EraseBkgARGBColor As Long = 0&, _
                        Optional ByVal EraseBkgBrushHandle As Long = 0&) As Boolean

    ' For the various Modes, recommend googling a bit for the differences. Here is one: http://www.geekpedia.com/tutorial52_Antialiasing-or-Smoothing.html
    ' Parameters
    '   :: ImageClass - instantiated cGDIpImage class with a loaded image
    '   :: hDC - the target DC to render to
    '   :: destX,destY - the target top/left coords to render at
    '   :: destWidth,destHeight - the size of the target to be rendered
    '   :: srcx,srcy - the image top/left coords to render from
    '   :: srcWidth,srcHeight - the size of the image to be rendered
    '   :: InterpolationMode - end quality when scaling/rotating during rendering
    '   :: SmoothingMode - end quality of the rendering itself
    '   :: CompositeMode - how colors are rendered with background colors
    '   :: PixelOffsetMode - how pixels will be shifted to produce higher quality images
    '   :: EraseBkgARGBColor - ignored if EraseBkgBrushHandle is not null; fill color of erasing background
    '   :: EraseBkgBrushHandle - the brush used to erase the background
   
    Dim hGraphics As Long, hBrush As Long, Angle As Single
    
    If hDC = 0& Or (ImageClass Is Nothing) Then Exit Function
    If ImageClass.Handle = 0& Then Exit Function

    GdipCreateFromHDC hDC, hGraphics
    If hGraphics = 0& Then Exit Function

    If destWidth = 0& Then destWidth = ImageClass.Width
    If destHeight = 0& Then destHeight = ImageClass.Height
    If srcWidth = 0& Then srcWidth = ImageClass.Width
    If srcHeight = 0& Then srcHeight = ImageClass.Height
    
    If EraseBkgBrushHandle Then
        hBrush = EraseBkgBrushHandle
    ElseIf EraseBkgARGBColor Then
        Call GdipCreateSolidFill(EraseBkgARGBColor, hBrush)
    End If
    
    Call GdipSetSmoothingMode(hGraphics, SmoothingMode)
    Call GdipSetCompositingQuality(hGraphics, CompositeMode)
    Call GdipSetPixelOffsetMode(hGraphics, PixelOffsetMode)
    Call GdipSetInterpolationMode(hGraphics, InterpolationMode)
    
    ' keep angle btwn -359.99 to 359.99
    Angle = ImageClass.Rotation
    Angle = (Int(Angle) Mod 360) + (Angle - Int(Angle))
    
    If Angle = 0! Then  ' no rotation being used
        If hBrush Then Call GdipFillRectangleI(hGraphics, hBrush, destX, destY, destWidth, destHeight)
        RenderImageClassToDC = (GdipDrawImageRectRectI(hGraphics, ImageClass.Handle, destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, ImageClass.ImageAttributesHandle, 0&, 0&) = 0&)
    
    Else ' we are rotating
        
        If GdipRotateWorldTransform(hGraphics, Angle, 0&) = 0& Then
            If GdipTranslateWorldTransform(hGraphics, destX + destWidth / 2, destY + destHeight / 2, 1&) = 0& Then
               If hBrush Then Call GdipFillRectangleI(hGraphics, hBrush, -destWidth \ 2, -destHeight \ 2, destWidth, destHeight)
               RenderImageClassToDC = (GdipDrawImageRectRectI(hGraphics, ImageClass.Handle, -destWidth \ 2, -destHeight \ 2, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, ImageClass.ImageAttributesHandle, 0&, 0&) = 0&)
            End If
        End If
    
    End If
    
    If hBrush Then
        If EraseBkgBrushHandle <> hBrush Then GdipDeleteBrush hBrush
    End If
    GdipDeleteGraphics hGraphics

End Function

Public Function RenderImageClassToHGraphics(ImageClass As cGDIpImage, ByVal hGraphics As Long, _
                        Optional ByVal destX As Long, Optional ByVal destY As Long, _
                        Optional ByVal destWidth As Long = 0&, Optional ByVal destHeight As Long = 0&, _
                        Optional ByVal srcX As Long = 0&, Optional ByVal srcY As Long = 0&, _
                        Optional ByVal srcWidth As Long = 0&, Optional ByVal srcHeight As Long = 0&) As Boolean

    ' Parameters
    '   :: ImageClass - instantiated cGDIpImage class with a loaded image
    '   :: hGraphics - a GDI+ graphics handle, modified as needed (i.e., rotation, transformed, interpolation etc)
    '   :: destX,destY - the target top/left coords to render at
    '   :: destWidth,destHeight - the size of the target to be rendered
    '   :: srcx,srcy - the image top/left coords to render from
    '   :: srcWidth,srcHeight - the size of the image to be rendered
    
    If hGraphics = 0& Then Exit Function
    If ImageClass.Handle = 0& Then Exit Function

    If destWidth = 0& Then destWidth = ImageClass.Width
    If destHeight = 0& Then destHeight = ImageClass.Height
    If srcWidth = 0& Then srcWidth = ImageClass.Width
    If srcHeight = 0& Then srcHeight = ImageClass.Height
    
    RenderImageClassToHGraphics = (GdipDrawImageRectRectI(hGraphics, ImageClass.Handle, destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, ImageClass.ImageAttributesHandle, 0&, 0&) = 0&)

End Function

Public Function RenderToHGraphics(ByVal hBitmap As Long, ByVal hGraphics As Long, _
                        ByVal destX As Long, destY As Long, _
                        ByVal destWidth As Long, ByVal destHeight As Long, _
                        ByVal srcX As Long, ByVal srcY As Long, _
                        ByVal srcWidth As Long, ByVal srcHeight As Long, _
                        Optional ByVal hAttributes As Long) As Boolean

    ' Parameters
    '   :: hBitmap - any GDI+ image handle
    '   :: hGraphics - a GDI+ graphics handle, modified as needed (i.e., rotation, transformed, interpolation etc)
    '   :: destX,destY - the target top/left coords to render at
    '   :: destWidth,destHeight - the size of the target to be rendered
    '   :: srcx,srcy - the image top/left coords to render from
    '   :: srcWidth,srcHeight - the size of the image to be rendered
    '   :: hAttributes - a GDI+ attributes handle, optional
    
    If hGraphics = 0& Or hBitmap = 0& Then Exit Function
    If m_Token Is Nothing Then Exit Function
    If m_Token.Token = 0& Then Exit Function

    RenderToHGraphics = (GdipDrawImageRectRectI(hGraphics, hBitmap, destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, hAttributes, 0&, 0&) = 0&)

End Function

Public Function ResetWorldTransform(ByVal hGraphics As Long) As Boolean
    ' undoes RotateHGraphics
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If m_Token.Token Then ResetWorldTransform = (GdipResetWorldTransform(hGraphics) = 0&)
    End If
End Function

Public Function RotateHGraphics(ByVal hGraphics As Long, ByVal Angle As Single, ByVal AnchorX As Long, ByVal AnchorY As Long) As Boolean
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If m_Token.Token Then
            If GdipRotateWorldTransform(hGraphics, Angle, 0&) = 0& Then
                ' Anchor points are where the rotation will be translated from
                RotateHGraphics = (GdipTranslateWorldTransform(hGraphics, AnchorX, AnchorY, 1&) = 0&)
            End If
        End If
    End If

End Function

Public Function SetHGraphicsQuality(ByVal hGraphics As Long, _
                                    Optional ByVal InterpolationMode As InterpolationConstants = InterpolationModeDefault, _
                                    Optional ByVal SmoothingMode As SmoothingConstants = SmoothingModeDefault, _
                                    Optional ByVal CompositeMode As CompositingModeConstants = CompositingQualityDefault, _
                                    Optional ByVal PixelOffsetMode As PixelOffsetModeConstants = PixelOffsetModeDefault) As Boolean
    If Not (hGraphics = 0& Or (m_Token Is Nothing)) Then
        If m_Token.Token Then
            If GdipSetInterpolationMode(hGraphics, InterpolationMode) = 0& Then
                If GdipSetCompositingQuality(hGraphics, CompositeMode) = 0& Then
                    If GdipSetPixelOffsetMode(hGraphics, PixelOffsetMode) = 0& Then
                        SetHGraphicsQuality = (GdipSetSmoothingMode(hGraphics, SmoothingMode) = 0&)
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub Class_Initialize()
    Set m_Token = Nothing
End Sub
