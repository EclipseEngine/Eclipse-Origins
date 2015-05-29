VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIpPenBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Credit for these GDI+ classes go to LaVolpe
'http://www.vbforums.com/showthread.php?t=598771
' HOW TO USE THIS CLASS
' 1. Simply, instantiate a single copy for the life of your project
' 2. The cGDIpToken class must also be instantiated for any of these GDI+ classes
' 3. Before any function can be called, you must first call the AttachTokenClass method
' 4. When the class is destroyed it will also destroy all pens/brushes that have not yet been destroyed by you
'   :: Do not rely on this self-destruction. Creating many pens/brushes without destroying them can
'       eventually lead to low system resources. Ideally destroy each pen/brush as early as possible.
' 5. Pens and brushes are somehwat different than the old GDI. To change a color and many other attributes,
'       you simply call the appropriate Set____ function; no need to destroy & recreate.

' This class wraps many of the common GDI+ brush/pen calls.

' There are many, many more brush and pen methods. Add them as needed
' Gradient Brushes: http://www.com.it-berater.org/gdiplus/noframes/lineargradientbrush_functions.htm
' Hatch Brushes: http://www.com.it-berater.org/gdiplus/noframes/hatchbrush_functions.htm
' Solid Brushes: http://www.com.it-berater.org/gdiplus/noframes/solidbrush_functions.htm
' Pattern Brushes: http://www.com.it-berater.org/gdiplus/noframes/texturebrush_functions.htm
' Pens: http://www.com.it-berater.org/gdiplus/noframes/pen_functions.htm

Option Explicit

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type



Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As IUnknown, Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long


Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipGetPenColor Lib "gdiplus" (ByVal hPen As Long, ByRef pargb As Long) As Long
Private Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByRef pColor As Long) As Long

Private Declare Function GdipCreateHatchBrush Lib "GdiPlus.dll" (ByVal mHatchStyle As Long, ByVal mForecol As Long, ByVal mBackcol As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipGetHatchBackgroundColor Lib "GdiPlus.dll" (ByVal brush As Long, ByRef backcol As Long) As Long
Private Declare Function GdipGetHatchForegroundColor Lib "GdiPlus.dll" (ByVal brush As Long, ByRef forecol As Long) As Long
Private Declare Function GdipGetHatchStyle Lib "GdiPlus.dll" (ByVal brush As Long, ByRef hatchStyle As Long) As Long
Private Declare Function GdipCreateLineBrushI Lib "GdiPlus.dll" (ByRef point1 As POINTL, ByRef point2 As POINTL, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMode As Long, ByRef hBrush As Long) As Long
Private Declare Function GdipGetLineColors Lib "GdiPlus.dll" (ByVal brush As Long, ByRef Colors As Any) As Long
Private Declare Function GdipGetLineWrapMode Lib "GdiPlus.dll" (ByVal brush As Long, ByRef WrapMode As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectI Lib "gdiplus" (ByRef pRect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As Long, ByVal WrapMode As Long, ByRef lineGradient As Long) As Long
Private Declare Function GdipSetLineColors Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColor1 As Long, ByVal mColor2 As Long) As Long
Private Declare Function GdipSetLineWrapMode Lib "GdiPlus.dll" (ByVal brush As Long, ByVal WrapMode As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef pRect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal Angle As Single, ByVal isAngleScalable As Long, ByVal WrapMode As Long, ByRef hBrush As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mDashStyle As Long) As Long
Private Declare Function GdipGetPenDashStyle Lib "GdiPlus.dll" (ByVal mPen As Long, ByRef mDashStyle As Long) As Long
Private Declare Function GdipSetPenDashCap197819 Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mDashCap As Long) As Long
Private Declare Function GdipGetPenDashCap197819 Lib "GdiPlus.dll" (ByVal mPen As Long, ByRef mDashCap As Long) As Long
Private Declare Function GdipGetPenWidth Lib "GdiPlus.dll" (ByVal mPen As Long, ByRef nWidth As Single) As Long
Private Declare Function GdipSetPenWidth Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal nWidth As Single) As Long
Private Declare Function GdipSetPenColor Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal nColor As Long) As Long
Private Declare Function GdipCreateTexture Lib "GdiPlus.dll" (ByVal mImage As Long, ByVal mWrapMode As Long, ByRef mTexture As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal mImage As Long) As Long
Private Declare Function GdipGetBrushType Lib "GdiPlus.dll" (ByVal brush As Long, ByRef ptype As Long) As Long
Private Declare Function GdipSetPenEndCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mCap As Long) As Long
Private Declare Function GdipSetPenStartCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mCap As Long) As Long
Private Declare Function GdipGetPenEndCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByRef mCap As Long) As Long
Private Declare Function GdipGetPenStartCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByRef mCap As Long) As Long

' ToDo: GDI+ Paths
Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipSetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColors As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColors As Long, ByRef mCount As Long) As Long
Private Declare Function GdipGetPathWorldBounds Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mBounds As RECTL, ByVal mMatrix As Long, ByVal mPen As Long) As Long


Public Enum WrapModeConstants
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipY = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum

Public Enum HatchStylesConstant
    HatchStyleHorizontal = &H0
    HatchStyleVertical = &H1
    HatchStyleForwardDiagonal = &H2
    HatchStyleBackwardDiagonal = &H3
    HatchStyleCross = &H4
    HatchStyleDiagonalCross = &H5
    HatchStyle05Percent = &H6
    HatchStyle10Percent = &H7
    HatchStyle20Percent = &H8
    HatchStyle25Percent = &H9
    HatchStyle30Percent = &HA
    HatchStyle40Percent = &HB
    HatchStyle50Percent = &HC
    HatchStyle60Percent = &HD
    HatchStyle70Percent = &HE
    HatchStyle75Percent = &HF
    HatchStyle80Percent = &H10
    HatchStyle90Percent = &H11
    HatchStyleLightDownwardDiagonal = &H12
    HatchStyleLightUpwardDiagonal = &H13
    HatchStyleDarkDownwardDiagonal = &H14
    HatchStyleDarkUpwardDiagonal = &H15
    HatchStyleWideDownwardDiagonal = &H16
    HatchStyleWideUpwardDiagonal = &H17
    HatchStyleLightVertical = &H18
    HatchStyleLightHorizontal = &H19
    HatchStyleNarrowVertical = &H1A
    HatchStyleNarrowHorizontal = &H1B
    HatchStyleDarkVertical = &H1C
    HatchStyleDarkHorizontal = &H1D
    HatchStyleDashedDownwardDiagonal = &H1E
    HatchStyleDashedUpwardDiagonal = &H1F
    HatchStyleDashedHorizontal = &H20
    HatchStyleDashedVertical = &H21
    HatchStyleSmallConfetti = &H22
    HatchStyleLargeConfetti = &H23
    HatchStyleZigZag = &H24
    HatchStyleWave = &H25
    HatchStyleDiagonalBrick = &H26
    HatchStyleHorizontalBrick = &H27
    HatchStyleWeave = &H28
    HatchStylePlaid = &H29
    HatchStyleDivot = &H2A
    HatchStyleDottedGrid = &H2B
    HatchStyleDottedDiamond = &H2C
    HatchStyleShingle = &H2D
    HatchStyleTrellis = &H2E
    HatchStyleSphere = &H2F
    HatchStyleSmallGrid = &H30
    HatchStyleSmallCheckerBoard = &H31
    HatchStyleLargeCheckerBoard = &H32
    HatchStyleOutlinedDiamond = &H33
    HatchStyleSolidDiamond = &H34
End Enum
Public Enum EndCapEndConstants
    StartCap = 0
    EndCap = 1
End Enum
Public Enum PenEndCapConstants
    LineCapFlat = 0
    LineCapSquare = 1
    LineCapRound = 2
    LineCapTriangle = 3
    LineCapNoAnchor = &H10
    LineCapSquareAnchor = &H11
    LineCapRoundAnchor = &H12
    LineCapDiamondAnchor = &H13
    LineCapArrowAnchor = &H14
    LineCapCustom = &HFF
    LineCapAnchorMask = &HF0
End Enum
Public Enum DashCapConstants
    DashCapFlat = &H0
    DashCapRound = &H2
    DashCapTriangle = &H3
End Enum
Public Enum PenStyleConstants
    DashStyleSolid = &H0
    DashStyleDash = &H1
    DashStyleDot = &H2
    DashStyleDashDot = &H3
    DashStyleDashDotDot = &H4
End Enum
Public Enum GradientBrushStyleConstants ' gradient brush options
    LinearHorizontal = 0
    LinearVertical = 1
    LinearForwardDiagonal = 2
    LinearBackwardDiagonal = 3
End Enum
Public Enum BrushTypeConstants
    BrushTypeSolidColor = 0
    BrushTypeHatchFill = 1
    BrushTypeTextureFill = 2
    BrushTypePathGradient = 3
    BrushTypeLinearGradient = 4
End Enum


Private Const UnitPixel = &H2
Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type
Private Type POINTL
    X As Long
    Y As Long
End Type


Private m_Objects As Collection
Private m_Token As cGDIpToken

' MUST BE CALLED BEFORE ANY PEN/BRUSH METHODS ARE CALLED
Public Sub AttachTokenClass(TokenClass As cGDIpToken)
    If m_Token Is Nothing Then
        If TokenClass.Token Then
            Set m_Token = TokenClass
            m_Token.AddUser Me
            If m_Objects Is Nothing Then Set m_Objects = New Collection
        End If
    End If
End Sub

Public Function ARGBtoRGB(ByVal ARGBColor As Long, Optional ByRef Opacity As Byte) As Long

    ARGBtoRGB = (ARGBColor And &HFF) * &H10000 Or ((ARGBColor And &HFFFF&) \ &H100&) * &H100& Or ((ARGBColor And &HFFFFFF) \ &H10000)
    If ARGBColor < 0& Then
        Opacity = (ARGBColor And Not &H80000000) \ &H1000000 Or &H80
    Else
        Opacity = ARGBColor \ &H1000000
    End If
    
End Function

Public Sub Clear()
    Dim i As Long, sKey As String, hBrush As Long
    If Not m_Objects Is Nothing Then
        For i = m_Objects.count To 1 Step -1&
            If Not TypeOf m_Objects(i) Is IUnknown Then
                sKey = m_Objects(i)
                Select Case Left$(sKey, 1)
                Case "b" ' brush
                    hBrush = CLng(Mid$(sKey, 2))
                    If GetBrushType(hBrush) = BrushTypeTextureFill Then
                        sKey = "o" & CStr(hBrush)
                        m_Objects.Remove sKey
                    End If
                    GdipDeleteBrush hBrush
                Case "p" ' pen
                    GdipDeletePen CLng(Mid$(sKey, 2))
                End Select
                m_Objects.Remove i
            End If
        Next
    End If
End Sub

Public Function CreateBrushFromImage(ImageClass As cGDIpImage, Optional WithImageAttributes As Boolean = True, Optional ByVal WrapMode As WrapModeConstants = WrapModeTile) As Long

    Dim cNewImage As cGDIpImage, lHandle As Long
    
    ' to prevent the image from being deleted before the brush is deleted, we will clone the passed class
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If Not ImageClass Is Nothing Then
                If ImageClass.Handle Then
                    Set cNewImage = New cGDIpImage
                    If ImageClass.CloneImage(cNewImage, , , WithImageAttributes) Then
                        Call GdipCreateTexture(cNewImage.Handle, WrapMode, lHandle) ' create the image brush
                        If lHandle Then
                            m_Objects.Add "b" & CStr(lHandle), CStr(lHandle)        ' add brush handle to collection
                            m_Objects.Add cNewImage, "o" & CStr(lHandle)            ' add image class to collection
                            CreateBrushFromImage = lHandle                          ' cross-ref to the brush handle
                        End If
                    End If
                End If
            End If
        End If
    End If
                
End Function

Public Function CreateGradientBrushRect(ByVal RectLeft As Long, ByVal RectTop As Long, ByVal RectWidth As Long, ByVal RectHeight As Long, _
                                        ARGBStartColor As Long, ByVal ARGBEndColor As Long, _
                                        ByVal Gradient As GradientBrushStyleConstants, ByVal WrapMode As WrapModeConstants) As Long
    
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            Dim lHandle As Long, r As RECTL
            With r
                .Height = RectHeight
                .Left = RectLeft
                .Top = RectTop
                .Width = RectWidth
            End With
            Call GdipCreateLineBrushFromRectI(r, ARGBStartColor, ARGBEndColor, Gradient, WrapMode, lHandle)
            If lHandle Then
                m_Objects.Add "b" & CStr(lHandle), CStr(lHandle)
                CreateGradientBrushRect = lHandle
            End If
        End If
    End If
    
End Function
Public Function CreateGradientBrushRectAngled(ByVal RectLeft As Long, ByVal RectTop As Long, ByVal RectWidth As Long, ByVal RectHeight As Long, _
                                        ARGBStartColor As Long, ByVal ARGBEndColor As Long, _
                                        ByVal Angle As Single, ByVal AngleIsScalable As Long, _
                                        ByVal WrapMode As WrapModeConstants) As Long
    
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            Dim lHandle As Long, r As RECTL
            With r
                .Height = RectHeight
                .Left = RectLeft
                .Top = RectTop
                .Width = RectWidth
            End With
            Call GdipCreateLineBrushFromRectWithAngleI(r, ARGBStartColor, ARGBEndColor, Angle, AngleIsScalable, WrapMode, lHandle)
            If lHandle Then
                m_Objects.Add "b" & CStr(lHandle), CStr(lHandle)
                CreateGradientBrushRectAngled = lHandle
            End If
        End If
    End If
    
End Function

Public Function CreateGradientLineBrush(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
                                        ARGBStartColor As Long, ByVal ARGBEndColor As Long, ByVal WrapMode As WrapModeConstants) As Long
    
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            Dim lHandle As Long, pt1 As POINTL, pt2 As POINTL
            pt1.X = X1: pt1.Y = Y1
            pt2.X = X2: pt2.Y = Y2
            Call GdipCreateLineBrushI(pt1, pt2, ARGBStartColor, ARGBEndColor, WrapMode, lHandle)
            If lHandle Then
                m_Objects.Add "b" & CStr(lHandle), CStr(lHandle)
                CreateGradientLineBrush = lHandle
            End If
        End If
    End If
    
End Function

Public Function CreateHatchBrush(ByVal Style As HatchStylesConstant, ByVal ARGBForeColor As Long, ByVal ARGBBackColor As Long) As Long
    
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            Dim lHandle As Long
            Call GdipCreateHatchBrush(Style, ARGBForeColor, ARGBBackColor, lHandle)
            If lHandle Then
                m_Objects.Add "b" & CStr(lHandle), CStr(lHandle)
                CreateHatchBrush = lHandle
            End If
        End If
    End If
    
End Function

Public Function CreatePen(ByVal ARGBColor As Long, Optional ByVal PenWidth As Single = 1!, _
                            Optional PenStyle As PenStyleConstants = DashStyleSolid, _
                            Optional DashCapStyle As DashCapConstants = DashCapFlat) As Long
    
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If PenWidth > 0! Then
                Dim lHandle As Long
                Call GdipCreatePen1(ARGBColor, PenWidth, UnitPixel, lHandle)
                If lHandle Then
                    m_Objects.Add "p" & CStr(lHandle), CStr(lHandle)
                    GdipSetPenDashStyle lHandle, PenStyle
                    GdipSetPenDashCap197819 lHandle, DashCapStyle
                    CreatePen = lHandle
                End If
            End If
        End If
    End If
    
End Function

Public Function DestroyBrush(ByVal BrushHandle As Long) As Boolean

    Dim sObjKey As String, lType As BrushTypeConstants
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If BrushHandle Then
                On Error Resume Next
                lType = GetBrushType(BrushHandle)
                DestroyBrush = (GdipDeleteBrush(BrushHandle) = 0&)
                If Not Err Then
                    m_Objects.Remove CStr(BrushHandle)      ' delete brush from collection
                    If lType = BrushTypeTextureFill Then
                        sObjKey = "o" & CStr(BrushHandle)   ' get its image class key
                        m_Objects.Remove sObjKey
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function DestroyPen(ByVal PenHandle As Long) As Boolean

    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If PenHandle Then
                On Error Resume Next
                DestroyPen = (GdipDeletePen(PenHandle) = 0&)
                m_Objects.Remove CStr(PenHandle)
            End If
        End If
    End If

End Function

Public Function GetBrushType(ByVal BrushHandle As Long) As BrushTypeConstants
    
    Dim lRet As Long
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If BrushHandle Then
                If GdipGetBrushType(BrushHandle, lRet) = 0& Then GetBrushType = lRet
            End If
        End If
    End If
    
End Function

Public Function GetGradientLineBrushColors(ByVal BrushHandle As Long, _
                                            Optional ARGBStartColor As Long, Optional ARGBEndColor As Long) As Boolean
    
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            Dim lHandle As Long, Values(0 To 1) As Long
            If GdipGetLineColors(BrushHandle, Values(0)) = 0& Then
                ARGBStartColor = Values(0)
                ARGBEndColor = Values(1)
                GetGradientLineBrushColors = True
            End If
        End If
    End If
    
End Function

Public Function GetGradientLineBrushWrapMode(ByVal BrushHandle As Long, Optional WrapMode As WrapModeConstants) As Boolean
    
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            Dim lHandle As Long, lMode As Long
            If GdipGetLineWrapMode(BrushHandle, lMode) = 0& Then
                WrapMode = lMode
                GetGradientLineBrushWrapMode = True
            End If
        End If
    End If
    
End Function

Public Function GetHatchBrushColors(ByVal BrushHandle As Long, ByRef ARGBForeColor As Long, ByRef ARGBBackColor As Long) As Boolean

    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If BrushHandle Then
                If GdipGetHatchBackgroundColor(BrushHandle, ARGBBackColor) = 0& Then
                    Call GdipGetHatchForegroundColor(BrushHandle, ARGBForeColor)
                    GetHatchBrushColors = True
                End If
            End If
        End If
    End If

End Function

Public Function GetHatchStyle(ByVal BrushHandle As Long) As Long

    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If BrushHandle Then Call GdipGetHatchStyle(BrushHandle, GetHatchStyle)
        End If
    End If

End Function

Public Function GetPenProps(ByVal PenHandle As Long, Optional ByRef ARGBColor As Long, Optional Width As Single, _
                                Optional PenStyle As PenStyleConstants, Optional DashCapStyle As DashCapConstants, _
                                Optional EndCap As PenEndCapConstants, Optional StartCap As PenEndCapConstants) As Boolean
    Dim lValue As Long
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If PenHandle Then
                If GdipGetPenColor(PenHandle, ARGBColor) = 0& Then
                    Call GdipGetPenDashCap197819(PenHandle, lValue): DashCapStyle = lValue
                    Call GdipGetPenDashStyle(PenHandle, lValue): PenStyle = lValue
                    Call GdipGetPenWidth(PenHandle, Width)
                    Call GdipGetPenEndCap(PenHandle, lValue): EndCap = lValue
                    Call GdipGetPenStartCap(PenHandle, lValue): StartCap = lValue
                End If
            End If
        End If
    End If

End Function

Public Function GetSolidBrushColorARGB(ByVal BrushHandle As Long) As Long

    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If BrushHandle Then Call GdipGetSolidFillColor(BrushHandle, GetSolidBrushColorARGB)
        End If
    End If

End Function

Public Function RGBtoARGB(ByVal RGBcolor As Long, Optional ByVal Opacity As Byte = 255) As Long

    RGBtoARGB = (RGBcolor And &HFF) * &H10000 Or ((RGBcolor And &HFFFF&) \ &H100&) * &H100& Or ((RGBcolor And &HFFFFFF) \ &H10000)
    If Opacity > 127 Then
        RGBtoARGB = RGBtoARGB Or ((Opacity - 128) * &H1000000) Or &H80000000
    Else
        RGBtoARGB = RGBtoARGB Or (Opacity * &H1000000)
    End If
    
End Function

Public Function SetGradientColors(ByVal BrushHandle As Long, ByVal ARGBStartColor As Long, ByVal ARGBEndColor As Long) As Boolean

    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If BrushHandle Then SetGradientColors = (GdipSetLineColors(BrushHandle, ARGBStartColor, ARGBEndColor) = 0&)
        End If
    End If

End Function

Public Function SetGradientWrapMode(ByVal BrushHandle As Long, ByVal WrapMode As WrapModeConstants) As Boolean

    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If BrushHandle Then SetGradientWrapMode = (GdipSetLineWrapMode(BrushHandle, WrapMode) = 0&)
        End If
    End If

End Function

Public Function SetPenColor(ByVal PenHandle As Long, ByVal ARGBColor As Long) As Boolean

    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If PenHandle Then SetPenColor = (GdipSetPenColor(PenHandle, ARGBColor) = 0&)
        End If
    End If

End Function

Public Function SetPenWidth(ByVal PenHandle As Long, ByVal Width As Single) As Boolean

    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If PenHandle Then SetPenWidth = (GdipSetPenWidth(PenHandle, Width) = 0&)
        End If
    End If

End Function

Public Function CreateSolidBrush(ByVal ARGBColor As Long) As Long
    
    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            Dim lHandle As Long
            Call GdipCreateSolidFill(ARGBColor, lHandle)
            If lHandle Then
                m_Objects.Add "b" & CStr(lHandle), CStr(lHandle)
                CreateSolidBrush = lHandle
            End If
        End If
    End If
    
End Function
Public Function SetPenEndCaps(ByVal PenHandle As Long, ByVal WhichEnd As EndCapEndConstants, ByVal CapStyle As PenEndCapConstants) As Boolean

    If Not m_Token Is Nothing Then
        If m_Token.Token Then
            If PenHandle Then
                If WhichEnd = EndCap Then
                    SetPenEndCaps = (GdipSetPenEndCap(PenHandle, CapStyle) = 0&)
                Else
                    SetPenEndCaps = (GdipSetPenStartCap(PenHandle, CapStyle) = 0&)
                End If
            End If
        End If
    End If

End Function

Public Function SystemColorToARGB(ByVal SysColorValue As Long, Optional ByVal Opacity As Byte = 255) As Long
    
    SystemColorToARGB = RGBtoARGB(GetSysColor(SysColorValue And &HFF&), Opacity)

End Function

Public Function SystemColorToRGB(ByVal SysColorValue As Long) As Long

    SystemColorToRGB = GetSysColor(SysColorValue And &HFF&)

End Function

Private Sub Class_Terminate()
    If Not m_Token Is Nothing Then
        Call Clear
        m_Token.RemoveUser Me
        Set m_Token = Nothing
    End If
End Sub
