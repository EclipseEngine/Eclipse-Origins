VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIpImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Credit for these GDI+ classes go to LaVolpe
'http://www.vbforums.com/showthread.php?t=598771
' HOW TO USE THESE CLASSES
' 1. The first class that must be created and instantiated is cGDIpToken
'   -- do this in form load or sub main. Only create one instance of this class; make it Public to other forms
'   -- that class creates the GDI+ token if GDI+ can be started
'   -- you should check the .Token property of that class to ensure it is not zero
'   -- do not call any other methods in that class
'   -- that class must be kept alive until all images are no longer needed
' 2. When you want to load an image, use the cGDIpImage class
'   -- that class must be kept alive until you no longer need your image
'   -- that class can accept images as filenames (unicode filenames supported), arrays, stdPictures or DIB handles
'   -- that class' .Handle property will be non-zero if image is loaded successfully
'   -- that class offers image attributes like semi-transparency, lightness/darkness, grayscaling, rotation and more
'   -- the following image formats are supported: bmp,ico,cur,gif,jpg,wmf,emf,tiff,png
'       :: bmp includes those with alpha channels, including premultiplied alpha channels
'       :: gif includes transparent gifs and animated gifs (multi-frame gifs)
'       :: tiff includes multi-page tiffs
'       :: ico/cur includes both XP-alphablended icons and Vista PNG-encoded icons
'       :: each frame/page of tiff/gif and each ico/cur in a multi-icon resource can be rendered individually
'   -- the following images will be converted internally to 24/32bpp bitmaps. See known GDI+ issues for reasons why
'       :: all icons/cursors modified to 32bpp bitmaps
'       :: all metafiles modified to 24/32bpp bitmaps
'       :: multi-frame gifs and multi-page tiffs if image attributes/mirroring is used; conversion undone when not used
'   -- that class offers saving images to one of the following formats.
'       :: Images are always returned as byte arrays which can be saved to file as desired
'       :: exported formats are: bmp, jpg, png, tiff. TIFF files can be saved as multipage TIFFs
'   -- to convert an image to one of those formats and then use the image, follow these easy steps
'       :: save the image in the desired format and have byte array returned
'       :: re-load or load into a new cGDIpImage class the array
' 3. When you want to draw an image, use the cGDIpRenderer class
'   -- the class accepts any cGDIpImage class/image
'   -- the class can be created once and kept until your project closes; or it can be created/destroyed on demand
' 4. The cGDIpMultiImage class IS NOT to be used by you directly. It is written as a subordinate class for cGDIpImage
' 5. The cGDIpPenBrush class contains many brush and pen methods
'   -- Like the cGDIpRenderer class, this class can be created once and kept until your project closes;
'       or it can be created/destroyed on demand
' 6. Future additions expected:
'   -- a cGDIpPaths class that contains many Path functions/methods
'   -- a cGDIpText class that contains many Text functions/methods
'   -- a cGDIpRegions class that contains common region functions/methods.

' IMAGE SELECTION LOGIC AND ATTRIBUTES
'   Attributes are grayscaling, lightness, rotation, transparency, mirroring
'   These may be reset or not reset depending on the type of image selection
'   -- When a new image is selected into this class, all attributes are reset unless optional parameter is false
'   -- When a multi-image Index is changed, no attributes are reset

' CLASS HIERARCHY & CLASS DEPENDENCIES
'  *~cGDIpToken :: no class dependencies
'   +cGDIpImage :: requires cGDIpToken, cGDIpMultiImage, cGDIpRenderer
'  +~cGDIpMultiImage :: no class dependencies; used only by cGDIpImage
'   *cGDIpRenderer :: requires cGDIpToken, cGDIpImage
'   *cGDIpPenBrush :: requires cGDIpToken, cGDIpImage
' (*) are classes that can be created just once and kept around until your project closes
' (+) are classes that must be created for each instance of their purpose
' (~) are classes that should not be called by your project. Exception: the cGDIpToken.Token property can be queried


' GDI+ KNOWN ISSUES, version 1
' Though GDI+ makes working with PNGs, TIFFs, and transparency pretty easy, it isn't perfect.
' There are a number of workarounds provided in this class for some of the more severe shortfalls.
':: GDI+ can't display WMF/EMF with image attributes. Fix: Convert WMF/EMF to bitmap
':: GDI can't load cursors & many icon types. Fix: Convert icon/cursor to bitmap
'       A complete icon to bitmap conversion routine is provied
':: GDI can't load PNG-embedded icons. Fix: Parse PNGs out and load separately
':: GDI+ can select multipage/frame tiff/gif but this function breaks in many scenarios
'       Fix: Copy current frame/page to PNG and render PNG with attributes
':: GDI can't load bitmaps that use the alpha channel. Fix: Manually load those using Scan0 GDI+ API
':: GDI+ can crash if loaded image source is removed (i.e. file deleted, dib deleted, stream destroyed, etc)
'       Fix: This class keeps the source in a IStream or stdPicture until the GDI+ image is destroyed
':: GDI+ can crash if you END your project while GDI+ is loaded. Fix: None, close your app normally, don't use END
'       Note. The cGDIpToken class was designed to ensure all image classes are destroyed before token class is destroyed
'       Note. Destroying the cGDIpToken class automatically destroys all images attached to it via cGDIpImage classes
':: GDI+ and non-AutoRedraw VB hDCs can react oddly. I have seen images rendered to the DC then immediately erased.
'       Fix: Simply set AutoRedraw to True. This "bug" is not consistent enough to code around
':: GDI+ requires the souce bytes to remain until the GDI+ image object is destroyed
'       Fix: When possible the source is converted internally a special way to a GDI+ bitmap object that is sourceless
'       This fix reduces memory usage since the original data is not required to be kept around

Option Explicit

Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal length As Long, ByVal Fill As Byte)
Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As IUnknown, Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
Private Declare Function GdipGetImageBounds Lib "GdiPlus.dll" (ByVal nImage As Long, srcRect As RECTF, srcUnit As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imgAttr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imgAttr As Long, ByVal clrAdjust As Long, ByVal clrAdjustEnabled As Long, ByRef clrMatrix As Any, ByRef grayMatrix As Any, ByVal clrMatrixFlags As Long) As Long
Private Declare Function GdipSetImageAttributesColorKeys Lib "GdiPlus.dll" (ByVal mImageattr As Long, ByVal mType As Long, ByVal mEnableFlag As Long, ByVal mColorLow As Long, ByVal mColorHigh As Long) As Long
Private Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal hImage As Long, ByVal GUID As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal hImage As Long, PixelFormat As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal size As Long, Encoders As Any) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipSaveAddImage Lib "gdiplus" (ByVal pImage As Long, ByVal newImage As Long, ByRef encoderParams As Any) As Long
Private Declare Function GdipSaveAdd Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mEncoderParams As Any) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As RECTI, ByVal mFlags As Long, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapGetPixel Lib "GdiPlus.dll" (ByVal pbitmap As Long, ByVal X As Long, ByVal Y As Long, ByRef pColor As Long) As Long
Private Declare Function GdipBitmapSetPixel Lib "GdiPlus.dll" (ByVal pbitmap As Long, ByVal X As Long, ByVal Y As Long, ByVal pColor As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal hbm As Long, ByVal hpal As Long, ByRef pbitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromGraphics Lib "GdiPlus.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal ptarget As Long, ByRef pbitmap As Long) As Long
Private Declare Function GdipGetImageHorizontalResolution Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef resolution As Single) As Long
Private Declare Function GdipGetImageVerticalResolution Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef resolution As Single) As Long
Private Declare Function GdipGetImagePaletteSize Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef psize As Long) As Long
Private Declare Function GdipGetImagePalette Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef palette As Any, ByVal psize As Long) As Long

Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As Any, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Type BITMAPINFOHEADER   ' structure used within icon image data
    biSize As Long
    biWidth As Long
    biHeight As Long            ' always doubled for icons/cursors
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As Long
End Type
Private Type RECTF
    nLeft As Single
    nTop As Single
    nWidth As Single
    nHeight As Single
End Type
Private Type RECTI
    nLeft As Long
    nTop As Long
    nWidth As Long
    nHeight As Long
End Type
Private Type EncoderParameter
    GUID(0 To 3)   As Long
    NumberOfValues As Long
    type           As Long
    Value          As Long
End Type
'-- Encoder Parameters structure
Private Type EncoderParameters
    count     As Long
    Parameter(0 To 5) As EncoderParameter
End Type
Private Type ImageCodecInfo
    ClassID(0 To 3)   As Long
    FormatID(0 To 3)  As Long
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type
Private Type BitmapData
    Width As Long
    Height As Long
    Stride As Long
    PixelFormat As Long
    Scan0Ptr As Long
    ReservedPtr As Long
End Type

Public Enum ImageTypeConstants
    imageUnknown = 0
    imageBitmap = 1
    imageGIF = 2
    imageJPEG = 3
    imageWMF = 4
    imageEMF = 5
    imageIcon = 6
    imagePNG = 7
    imageTIFF = 8
    imageExIF = 9
    imageCursor = 10
    imageIconPNG = 38       ' imageIcon Or 32
    imageCursorPNG = 42     ' imageCurosr or 32
End Enum
Public Enum GrayScaleConstants
    attrGrayNone = 0
    attrGraySimpleAverage = 1
    attrGrayNTSCPAL = 2
    attrGrayCCIR709 = 3
    attrGrayRedMask = 4
    attrGrayGreenMask = 5
    attrGrayBlueMask = 6
    attrGrayBlueGreenMask = 7
    attrGrayRedGreenMask = 8
End Enum
Public Enum MirrorConstants
    attrMirrorNone = 0
    attrMirrorBoth = 1
    attrMirrorHorizontal = 2
    attrMirrorVertical = 3
End Enum
Public Enum LockModeConstants
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum
Public Enum TIFFColorDepthConstants
    TIFF_DepthAuto = 0
    TIFF_Depth24Bit = 1
    TIFF_Depth32Bit = 2
End Enum
Public Enum ImageColorFormatConstants
    PixelFormat32bppARGB = &H26200A 'Specifies that the format is 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components.
    PixelFormat32bppPARGB = &HE200B 'Specifies that the format is 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components. The red, green, and blue components are premultiplied according to the alpha component.
    PixelFormat32bppRGB = &H22009 'Specifies that the format is 32 bits per pixel; 8 bits each are used for the red, green, and blue components. The remaining 8 bits are not used.
    PixelFormat1bppIndexed = &H30101 'Specifies that the format is 1 bit per pixel, indexed.
    PixelFormat4bppIndexed = &H30402 'Specifies that the format is 4 bits per pixel, indexed.
    PixelFormat8bppIndexed = &H30803 'Specifies that the format is 8 bits per pixel, indexed.
    PixelFormat16bppARGB1555 = &H61007 'Specifies that the format is 16 bits per pixel; 1 bit is used for the alpha component, and 5 bits each are used for the red, green, and blue components.
    PixelFormat16bppGrayScale = &H101004 'Specifies that the format is 16 bits per pixel, grayscale.
    PixelFormat16bppRGB555 = &H21005 'Specifies that the format is 16 bits per pixel; 5 bits each are used for the red, green, and blue components. The remaining bit is not used.
    PixelFormat16bppRGB565 = &H21006 'Specifies that the format is 16 bits per pixel; 5 bits are used for the red component, 6 bits are used for the green component, and 5 bits are used for the blue component.
    PixelFormat48bppRGB = &H10300C 'Specifies that the format is 48 bits per pixel; 16 bits each are used for the red, green, and blue components.
    PixelFormat64bppARGB = &H34400D 'Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components.
    PixelFormat64bppPARGB = &H1C400E 'Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components. The red, green, and blue components are premultiplied according to the alpha component.
    PixelFormat24bppRGB = &H21808 'Specifies that the format is 24 bits per pixel; 8 bits each are used for the red, green, and blue components.
End Enum

Private Enum pvCleanUpEnum  ' used locally
    cuDestroyAll = -1&          ' destroy all objects related to current image
    cuDestroyClone = 1&         ' destroy clone only, if it exists (GIF/TIFF images only)
    cuResetAttrs = 2&           ' reset attributes only
    cuDestroyImage = 4&         ' destroy source image
    cuDestroyMultiImage = 8&    ' destroy mutli-image class
End Enum

Private Const png_Signature1 As Long = 1196314761   ' PNG signature is 8 bytes
Private Const png_Signature2 As Long = 169478669
Private Const bmp_Signature As Integer = &H4D42&    ' BMP signature is 2 bytes
Private Const UnitPixel As Long = 2&                ' GDI+ constant


Private m_Image(0 To 1) As Long             ' GDI+ image handle
Private m_Source(0 To 1) As IUnknown        ' IStream/stdPicture attached to m_Image
Private m_Token As cGDIpToken               ' reference to GDI Token class
Private m_Size As RECTF                     ' Image dimensions
Private m_Attr As Long                      ' Image Attributes (see pvModifyAttributes routine)
Private m_Alpha As Single                   ' Global transaprency level applied
Private m_TransColor As Long                ' optional extra color to be made transparent throughout (like TransparentBlt)
Private m_Angle As Single                   ' Image rotation applied
Private m_Lightness As Single               ' Image lightness/darkness level applied
Private m_GrayScale As GrayScaleConstants   ' Grayscale formula applied
Private m_Mirror As MirrorConstants         ' Mirrored state applied
Private m_ImageType As ImageTypeConstants   ' Image format type
Private m_OrigColorType As ImageColorFormatConstants ' what color depth image was loaded from
' There are some bugs with GDI+ and here is one. If applying attributes, flipping, cloning an image, and more,
' one can no longer select an image within a multi-image format; the result is always the 1st frame/page.
' So to workaround this bug, we create a "clone" that is actually a PNG from the current frame/page.
' Then that "clone" is used to render with mirroring/attributes. The routines herein remmove the
' clone when no longer needed.  A similar issue is with MetaFiles, they can be flipped, but cannot be
' rendered with attributes, so WMF/EMF will be converted to Bitmap for rendering.
Private m_SourceIndex As Long               ' ref to original image or PNG clone of TIFF,GIF,WMF,EMF as needed
' m_Image & m_Source usage for clarification
' Image Type        m_Image(0)  m_Source(0)     m_Image(1)  m_Source(1)
'   PNG             yes         never           never       never
'   JPG             yes         never           never       never
'   BMP             yes         never           never       never
'   GIF             yes         yes**1          Cloned**2   Cloned**2
'   TIFF            yes         yes**1          Cloned**2   Cloned**2
'   ICO/CUR         yes         yes**1          never       never
'   WMF/EMF         yes         yes**1          never       never
' Notes
' PNG, JPG, BMP, 1-page TIFFs are always converted to 1,4,8,24,32pARGB bpp bitmaps
' ICO, CUR, WMF, EMF are converted to 32bpp pARGB internally only
' (**1) Original source data/bytes are maintained so that multiple images from the resource can be extracted
'       and so that original format can be returned on demand
'       Note that the CUR/ICO/WMF/EMF source may be either stdPicture or IIStream depending on how the source was loaded
'       Exception: Single page TIFFs do not cache original source data
'       Note: ICO/CUR are cached for ease. One can easily add a BitmapToIcon routine to create the appropriate icon/cursor format
' (**2) When multiple images exist in resource, a clone (PNG format) is created if to be rendered with attributes.
'       Any created clone is destroyed when the attributes are reset
Private WithEvents m_MultiImage As cGDIpMultiImage  ' enables moving between images in multi-image source. Provides GIF properties
Attribute m_MultiImage.VB_VarHelpID = -1

Public Sub Clear()
    ' desroys GDI+ image & attributes and releases system resources
    Call pvCleanUp(cuDestroyAll)
End Sub

Public Function CloneImage(DestImageClass As cGDIpImage, Optional ByVal Width As Long, Optional ByVal Height As Long, _
                            Optional ByVal WithAttributes As Boolean = False) As Boolean

    ' function can apply rotation, grayscaling, scaling etc to a new image
    ' To apply the new image to the same class, simply pass this class as the DestImageClass parameter
    '       else ensure DestImageClass is initialized, i.e., it was set to a New cGDIpImage
    ' If desired, pass new width/height.
    ' Rotated sources will be trimmed to remove excess space around the rotated image
    
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be cloned if resizing/applying attributes
    ' NOTE: The end result will be a 24 or 32 bpp bitmap if resizing/applying attributes
    
    If m_Image(m_SourceIndex) = 0& Then Exit Function
    If DestImageClass Is Nothing Then Exit Function
    
    Dim BHI As BITMAPINFO, cRenderer As cGDIpRenderer
    Dim cX As Long, cY As Long, X As Long, Y As Long
    Dim clipX As Long, clipY As Long, clipCx As Long, clipCy As Long
    Dim hDib As Long, dibPtr As Long, hObj As Long
    Dim dDC As Long, tDC As Long, scanWidth As Long, Angle As Single
    Dim bData() As Byte, bTrim() As Byte, gifPal() As Long
    
    If Width < 1& Then Width = m_Size.nWidth        ' don't allow negative height/width
    If Height < 1& Then Height = m_Size.nHeight     ' set to actual size if 0 or less
    
    ' keep angle btwn -359.99 to 359.99; user may have supplied something like 728 for example
    If WithAttributes Then
        Angle = m_Angle
        Angle = (Int(Angle) Mod 360) + (Angle - Int(Angle))
        ' if not rotating, not mirroring and no attributes handle created, then cancel WithAttributes
        If Angle = 0! And m_Attr = 0& And m_Mirror = attrMirrorNone Then WithAttributes = False
    End If
    
    If Angle = 0! Then
        cX = Width: cY = Height
        If WithAttributes = False Then              ' applying attributes?
            If cX = m_Size.nWidth And cY = m_Size.nHeight Then
                If DestImageClass Is Me Then        ' if no attributes and no-resize, then this is easy
                    pvCleanUp cuResetAttrs          ' same class? Simply remove any attributes & reset to frame/page 0
                    If Not m_MultiImage Is Nothing Then m_MultiImage.Index = 0&
                    CloneImage = True
                    Exit Function
                Else                                ' else get the data in a stream & send it to the target class
                    If SaveAsOriginalFormat(bData()) = True Then
                        If DestImageClass.LoadPicture_Stream(bData, m_Token) Then
                            CloneImage = True
                            Exit Function
                        End If
                        Erase bData()
                    End If
                End If
            End If
        End If
    Else
        cX = Sqr(Width * Width + Height * Height)   ' max size required for 0-360 degree rotation
        cY = cX                                     ' any excess will be trimmed later
    End If
    
    With BHI.bmiHeader                              ' determine end result color depth
        Select Case m_ImageType
        Case imageIcon, imageIconPNG, imageCursor, imageCursorPNG
            .biBitCount = 32                        ' assume transparency exists else need too parse icon/cursor mask to verify
        Case imageGIF                               ' if GIF uses transparency index, assume transparency
            If m_MultiImage.GetGifFramePalette(gifPal(), X) Then
                If X = -1& Then .biBitCount = 24 Else .biBitCount = 32
                Erase gifPal()
            Else
                .biBitCount = 32
            End If
        Case Else
            Select Case m_OrigColorType             ' for all other formats, use its reported color depth
            Case PixelFormat32bppARGB, PixelFormat32bppPARGB, PixelFormat64bppARGB, PixelFormat64bppPARGB, PixelFormat16bppARGB1555
                .biBitCount = 32
            Case Else
                .biBitCount = 24
            End Select
        End Select
        ' tweak to 32bit if 24bit uses rotation, global transparency, or extra simple transparency
        If WithAttributes = True And .biBitCount = 24 Then
            If Not (Angle = 0! And m_Alpha = 0! And m_TransColor = 0&) Then .biBitCount = 32
        End If
        .biHeight = cY
        .biWidth = cX
        .biSize = 40&
        .biPlanes = 1
        .biSizeImage = pvByteAlignOnWord(.biBitCount, cX) * cY
    End With
    dDC = GetDC(GetDesktopWindow)                   ' create DIB and get DC so we can paint image into it
    hDib = CreateDIBSection(dDC, BHI, 0&, dibPtr, 0&, 0&)
    If hDib Then
        tDC = CreateCompatibleDC(dDC)
        ReleaseDC GetDesktopWindow(), dDC
        hObj = SelectObject(tDC, hDib)
        Set cRenderer = New cGDIpRenderer
        cRenderer.AttachTokenClass m_Token
        If WithAttributes Then
            cRenderer.RenderImageClassToDC Me, tDC, (cX - Width) \ 2, (cY - Height) \ 2, Width, Height, , , , , InterpolationModeHighQualityBicubic, SmoothingModeAntiAlias, CompositingQualityHighQuality, PixelOffsetModeHighQuality
        Else
            If m_Mirror Then GdipImageRotateFlip m_Image(m_SourceIndex), m_Mirror
            dDC = cRenderer.CreateHGraphicsFromHDC(tDC)
            cRenderer.SetHGraphicsQuality dDC, InterpolationModeHighQualityBicubic, SmoothingModeAntiAlias, CompositingQualityHighQuality, PixelOffsetModeHighQuality
            cRenderer.RenderToHGraphics m_Image(m_SourceIndex), dDC, 0&, 0&, Width, Height, m_Size.nLeft, m_Size.nTop, m_Size.nWidth, m_Size.nHeight
            cRenderer.DestroyHGraphics dDC
            If m_Mirror Then GdipImageRotateFlip m_Image(m_SourceIndex), m_Mirror
        End If
        Set cRenderer = Nothing
        SelectObject tDC, hObj
        
        If Angle <> 0! Then                 ' trim if rotated, else pass image to destination ImageClass
            Width = cX: Height = cY
            ReDim bData(0 To BHI.bmiHeader.biSizeImage - 1&)
            CopyMemory bData(0), ByVal dibPtr, BHI.bmiHeader.biSizeImage
            ' trimming rotated image; find real left edge; looking for non-zero Alpha byte
            scanWidth = Width * 4&
            For X = 3& To scanWidth - 1& Step 4&
                For Y = 0& To Height - 1&
                    If bData(Y * scanWidth + X) Then
                        clipX = X - 3&         ' trim left starting at this position
                        X = scanWidth
                        Exit For
                    End If
                Next
            Next
            ' find real bottom edge; looking for non-zero Alpha byte
            For Y = 0& To Height - 1&
                X = Y * scanWidth + 3&
                For X = X + clipX To X + scanWidth - 1& Step 4&
                    If bData(X) Then
                        clipY = Y           ' trim top starting at this posiiton
                        Y = Height
                        Exit For
                    End If
                Next
            Next
            ' find real right edge; looking for non-zero Alpha byte
            For X = scanWidth - 1& To clipX + 3& Step -4&
                For Y = clipY To Height - 1&
                    If bData(Y * scanWidth + X) Then
                        clipCx = X - clipX + 1& ' calc non-trimmed width in total bytes
                        X = 0&
                        Exit For
                    End If
                Next
            Next
            ' find real top edge; looking for non-zero Alpha byte
            For Y = Height - 1& To clipY Step -1&
                X = Y * scanWidth + 3& + clipX
                For X = X To X + clipCx - 1& Step 4&
                    If bData(X) Then
                        clipCy = Y - clipY + 1& ' calc number of non-trimmed rows
                        Y = clipY
                        Exit For
                    End If
                Next
            Next
            If (clipCy = Height And clipCx = scanWidth) Then  ' else we have some trimming to do
                Erase bData()
            Else
                DeleteObject hDib: hDib = 0&        ' get rid of rotated DIB
                Width = clipCx \ 4                  ' set adjusted width/height
                Height = clipCy
                cY = 0&
                ReDim bTrim(0 To clipCx * Height - 1&)
                For Y = clipY To clipY + clipCy - 1&
                    CopyMemory bTrim(cY * clipCx), bData(Y * scanWidth + clipX), clipCx
                    cY = cY + 1&
                Next
                Erase bData()                       ' create new trimmed DIB
                With BHI.bmiHeader
                    .biHeight = Height
                    .biWidth = Width
                    .biSizeImage = clipCx * Height
                End With                            ' transfer the data to new DIB
                hDib = CreateDIBSection(tDC, BHI, 0&, dibPtr, 0&, 0&)
                If hDib Then CopyMemory ByVal dibPtr, bTrim(0), BHI.bmiHeader.biSizeImage
                Erase bTrim()
            End If
        End If
        DeleteDC tDC
        If hDib Then
            CloneImage = DestImageClass.LoadPicture_DIBhandle(hDib, m_Token)
            DeleteObject hDib
        End If
    Else
        ReleaseDC GetDesktopWindow(), dDC
    End If
    
End Function

' Return the source image format
Public Property Get ColorFormat() As ImageColorFormatConstants
    ColorFormat = m_OrigColorType
End Property

' Sets/Gets color to be rendered transparently throughout image
Public Property Get ExtraTransparentColor() As Long
    ExtraTransparentColor = m_TransColor
End Property
Public Property Let ExtraTransparentColor(newVal As Long)
    ' Note: this color must be ARGB format.
    If newVal <> m_TransColor Then
        m_TransColor = newVal
        pvModifyAttributes
    End If
End Property

Public Function GetPixel(ByVal X As Long, ByVal Y As Long) As Long
    ' function returns an ARGB color. X,Y must be within 0 to Width-1 & Height-1
    If m_Image(m_SourceIndex) Then GdipBitmapGetPixel m_Image(m_SourceIndex), X, Y, GetPixel
End Function

Public Function GetScaledImageSizes(ByVal destWidth As Long, ByVal destHeight As Long, ScaledWidth As Long, ScaledHeight As Long, _
                                        Optional ByVal CanScaleUp As Boolean = True) As Boolean

    ' Function returns scaled (maintaining scale ratio) for passed destination width/height
    ' The CanScaleUp when set to false will never return scaled sizes > than 1:1
    If m_Image(m_SourceIndex) = 0& Then Exit Function
    
    Dim xRatio As Single, yRatio As Single
    xRatio = destWidth / m_Size.nWidth
    yRatio = destHeight / m_Size.nHeight
    If xRatio > yRatio Then xRatio = yRatio
    
    If xRatio > 1! And CanScaleUp = False Then
        ScaledWidth = m_Size.nWidth
        ScaledHeight = m_Size.nHeight
    Else
        ScaledWidth = m_Size.nWidth * xRatio
        ScaledHeight = m_Size.nHeight * xRatio
    End If
    GetScaledImageSizes = True

End Function

' Return the GDI+ image global lightness (values btwn -100 and 100)
Public Property Get GlobalLightnessPct() As Single
    GlobalLightnessPct = m_Lightness
End Property
Public Property Let GlobalLightnessPct(newVal As Single)
    If newVal >= -100! And newVal <= 100! Then
        If m_Lightness <> newVal Then
            m_Lightness = newVal
            pvModifyAttributes
        End If
    End If
End Property

' Return the GDI+ image global transparency (values btwn 0 and 100)
Public Property Get GlobalTransparencyPct() As Single
    GlobalTransparencyPct = m_Alpha
End Property
Public Property Let GlobalTransparencyPct(newVal As Single)
    If newVal >= 0! And newVal <= 100! Then
        If m_Alpha <> newVal Then
            m_Alpha = newVal
            pvModifyAttributes
        End If
    End If
End Property

' Return the GDI+ image grayscale formula
Public Property Get GrayScale() As GrayScaleConstants
    GrayScale = m_GrayScale
End Property
Public Property Let GrayScale(newVal As GrayScaleConstants)
    If newVal >= attrGrayNone And newVal <= attrGrayRedGreenMask Then
        If m_GrayScale <> newVal Then
            m_GrayScale = newVal
            pvModifyAttributes
        End If
    End If
End Property

' Return the GDI+ image handle
Public Property Get Handle() As Long
    Handle = m_Image(m_SourceIndex)
End Property

' Return the GDI+ image height
Public Property Get Height() As Long
    Height = m_Size.nHeight
End Property


' Return/Set the GDI+ image attributes handle
' (GDI+ object that includes grayscale, lightness, transparency & more)
Public Property Get ImageAttributesHandle() As Long
    ImageAttributesHandle = m_Attr
End Property

' Return the source image type
Public Property Get ImageType() As ImageTypeConstants
    ImageType = m_ImageType
End Property

Public Function LoadPicture_DIBhandle(DibHandle As Long, TokenClass As cGDIpToken, Optional ByVal ClearAttributes As Boolean = True) As Boolean

    If TokenClass Is Nothing Then Exit Function
    If TokenClass.Token = 0& Then Exit Function
    
    Dim hImage As Long, bData() As Byte
    Dim clrUsed As Long, clrImportant As Long
    Dim BHI As BITMAPINFO, dDC As Long
    Dim lClearFlag As pvCleanUpEnum
    
    On Error GoTo ExitRoutine
    
    If m_Token Is Nothing Then
        Set m_Token = TokenClass
        TokenClass.AddUser Me
    End If
    
    dDC = GetDC(GetDesktopWindow())
    ' GDI+ does not support alpha channel bitmaps so we support it manually
    ' Use GetDIBits to fill in the bitmap info header structure
    BHI.bmiHeader.biSize = 40&
    If GetDIBits(dDC, DibHandle, 0&, 0&, ByVal 0&, BHI, 0&) Then
        If BHI.bmiHeader.biBitCount = 32& Then
            With BHI.bmiHeader
                .biSizeImage = .biWidth * .biHeight * 4&
                ReDim bData(0 To .biSizeImage - 1&)
            End With
            ' call API again to transfer the pixel data to our array
            If GetDIBits(dDC, DibHandle, 0&, BHI.bmiHeader.biHeight, bData(0), BHI, 0&) Then
                ReleaseDC GetDesktopWindow(), dDC: dDC = 0&
                If ClearAttributes Then lClearFlag = cuDestroyAll Else lClearFlag = cuDestroyAll Xor cuResetAttrs
                If pvProcessAlphaBitmap(bData(), BHI.bmiHeader.biWidth, BHI.bmiHeader.biHeight, lClearFlag) = True Then
                    LoadPicture_DIBhandle = True
                    Exit Function
                End If
            End If
            Erase bData()
        End If
    End If
    If dDC Then ReleaseDC GetDesktopWindow(), dDC
    
    ' create the stream to cache our bytes & load the image from the stream
    If GdipCreateBitmapFromHBITMAP(DibHandle, 0&, hImage) = 0& Then
        If ClearAttributes = False Then lClearFlag = cuResetAttrs Else lClearFlag = 0&
        Call pvCleanUp(cuDestroyAll Xor lClearFlag)  ' destroy previous
        m_Image(0) = hImage         ' cache latest image
        m_ImageType = imageBitmap
        Call GdipGetImagePixelFormat(hImage, m_OrigColorType)  ' get size,type and apply attributes
        Call GdipGetImageBounds(hImage, m_Size, UnitPixel)
        LoadPicture_DIBhandle = True
    End If
    
ExitRoutine:
End Function

Public Function LoadPicture_FileName(filename As String, TokenClass As cGDIpToken, Optional ClearAttributes As Boolean = True) As Boolean

    ' Unicode compatible if the FileName passes contains unicode characters
    ' Routine simply opens the file, reads the data, and passes to the LoadPicture_Stream function
    If TokenClass Is Nothing Then Exit Function
    If TokenClass.Token = 0& Then Exit Function
    
    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    Const FILE_ATTRIBUTE_NORMAL = &H80&
    Const INVALID_HANDLE_VALUE As Long = -1&
    
    Dim hImage As Long, IStream As IUnknown
    Dim flags As Long, hFile As Long, bData() As Byte
    
    On Error GoTo ExitRoutine
    flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
            Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    If IsWindowUnicode(GetDesktopWindow()) = 0& Then
        hFile = CreateFileA(filename, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, flags, 0&)
    Else
        hFile = CreateFileW(StrPtr(filename), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, flags, 0&)
    End If
    
    If Not (hFile = INVALID_HANDLE_VALUE Or hFile = 0&) Then
        flags = GetFileSize(hFile, 0&)
        ReDim bData(0 To flags - 1&)
        ReadFile hFile, bData(0), flags, flags, ByVal 0&
        CloseHandle hFile
        If flags > UBound(bData) Then LoadPicture_FileName = LoadPicture_Stream(bData(), TokenClass, ClearAttributes)
    End If
    
ExitRoutine:
End Function

Public Function LoadPicture_FromNothing(Width As Long, Height As Long, graphics As Long, TokenClass As cGDIpToken) As Boolean

    ' Unicode compatible if the FileName passes contains unicode characters
    ' Routine simply opens the file, reads the data, and passes to the LoadPicture_Stream function
    If TokenClass Is Nothing Then Exit Function
    If TokenClass.Token = 0& Then Exit Function
    
    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    Const FILE_ATTRIBUTE_NORMAL = &H80&
    Const INVALID_HANDLE_VALUE As Long = -1&
    
    Dim hImage As Long, IStream As IUnknown
    Dim flags As Long, hFile As Long, bData() As Byte
    
    hImage = pvCreateSourcelessImage2(hImage, Width, Height, 1&, TokenClass, graphics)
                    If hImage Then
                        m_Image(0) = hImage
                        m_Size.nWidth = Width
                        m_Size.nHeight = Height
                        LoadPicture_FromNothing = True
                    End If
                    
                    
                    Exit Function
    
ExitRoutine:
End Function

Public Function LoadPicture_stdPicture(Picture As StdPicture, TokenClass As cGDIpToken, Optional ClearAttributes As Boolean = True) As Boolean

    If TokenClass Is Nothing Then Exit Function
    If TokenClass.Token = 0& Then Exit Function
    If Picture Is Nothing Then Exit Function
    
    Dim lClearFlag As pvCleanUpEnum
    Dim oldType As Long, oldSource As IUnknown

    If m_Token Is Nothing Then
        Set m_Token = TokenClass
        TokenClass.AddUser Me
    End If

    If ClearAttributes Then lClearFlag = cuDestroyAll Else lClearFlag = cuDestroyAll Xor cuResetAttrs
    Select Case Picture.type
    Case vbPicTypeBitmap    ' pass off to another function. VB bitmaps are DIBs
        LoadPicture_stdPicture = LoadPicture_DIBhandle(Picture.Handle, TokenClass, ClearAttributes)
        
    Case vbPicTypeEMetafile, vbPicTypeMetafile
        If pvMetaFileTo32bpp(Picture, 0&, lClearFlag) = True Then ' convert metafile to bitmap
            If Picture.type = vbPicTypeEMetafile Then m_ImageType = imageEMF Else m_ImageType = imageWMF
            LoadPicture_stdPicture = True
        End If
        
    Case vbPicTypeIcon
        Set oldSource = m_Source(0)     ' back up current image type
        Set m_Source(0) = Picture
        oldType = m_ImageType
        m_ImageType = imageIcon
        If pvCreateClone(&H80000000, lClearFlag) Then ' create the icon
            LoadPicture_stdPicture = True
        Else
            m_ImageType = oldType       ' rollback
            Set m_Source(0) = oldSource
        End If
    Case Else
        
    End Select

End Function

Public Function LoadPicture_Stream(stream() As Byte, TokenClass As cGDIpToken, Optional ByVal ClearAttributes As Boolean = True) As Boolean
    
    ' must pass a 1D, zero-bound array
    
    If TokenClass Is Nothing Then Exit Function
    If TokenClass.Token = 0& Then Exit Function
    
    Dim hImage As Long, IStream As IUnknown, IBackup As IUnknown, tPIC As StdPicture
    Dim hDib As Long, lClearFlag As pvCleanUpEnum, newMultiImage As cGDIpMultiImage
    Dim lSize As Long, lValue As Long, lCount As Long, lLength As Long
    Dim arrPtr As Long, Index As Long, bIsIcon As Boolean, lType As Long
    
    On Error GoTo ExitRoutine
    lValue = Not stream()
    Debug.Assert App.hInstance
    If lValue = -1& Then Exit Function ' empty array
    If LBound(stream) <> 0 Then Exit Function ' must be zero bound; just lazy on my part

    If m_Token Is Nothing Then
        Set m_Token = TokenClass
        TokenClass.AddUser Me
    End If
    If ClearAttributes Then lClearFlag = cuDestroyAll Else lClearFlag = cuDestroyAll Xor cuResetAttrs

    ' /// Testing for icons/cursors. GDI+ can't seem to load cursors and has many issues with various types of icons
    ' See the pvIconTo32bpp routine for more info
    lLength = UBound(stream) + 1&
    If lLength > 5& Then
        On Error Resume Next
        CopyMemory lType, stream(2), 2&
        If lType = 1 Or lType = 2 Then ' validate: must be 1 or 2 to be an icon/cursor
            CopyMemory lCount, stream(4), 2& ' count of icon/cursor in resource
            If lCount > 0& Then
                If lCount * 16& + 6& <= lLength Then  ' minimial bytes required for now
                    For Index = 0& To lCount - 1&
                        arrPtr = 14& + Index * 16&
                        CopyMemory lSize, stream(arrPtr), 4&        ' bytes in resource
                        CopyMemory lValue, stream(arrPtr + 4&), 4&  ' offset in resource
                        If lValue < 1& Or lSize < 1& Then Exit For  ' valid entries?
                        If lValue + lSize > lLength Then Exit For   ' enough bytes in resource?
                        ' let's validate the header size, bitcount, width, height
                        CopyMemory lSize, stream(lValue), 4&
                        If lSize = png_Signature1 Then
                            CopyMemory lSize, stream(lValue + 4&), 4&
                            If lSize <> png_Signature2 Then Exit For    ' invalid header size
                        Else
                            If lSize < 40& Then Exit For                ' invalid header size
                            CopyMemory lSize, stream(lValue + 4&), 4&
                            If lSize < 1& Then Exit For                 ' invalid width
                            CopyMemory lSize, stream(lValue + 8&), 4&
                            If lSize < 1& Then Exit For                 ' invalid height
                            CopyMemory lSize, stream(lValue + 14&), 2&
                            Select Case (lSize And &HFFFF&)             ' validate bit count
                                Case 1, 2, 4, 8 ' supported
                                Case 16, 24, 32 ' supported
                                Case Else       ' not supported or not an icon format
                                    Exit For
                            End Select
                        End If
                        If Err Then
                            Err.Clear                               ' overflow?
                            Exit For
                        End If
                    Next
                    If (Index = lCount) Then                        ' an icon/cursor file
                        Set IBackup = m_Source(0)
                        Set IStream = pvStreamFromArray(VarPtr(stream(0)), UBound(stream) + 1&)
                        If Not IStream Is Nothing Then
                            Set m_Source(0) = IStream
                            lValue = m_ImageType                    ' back up if icon fails to load
                            m_ImageType = imageIcon
                            If pvCreateClone(&H80000000, lClearFlag) Then   ' create image from icon
                                LoadPicture_Stream = True
                                Exit Function
                            End If
                            m_ImageType = lValue                    ' roll back
                        End If
                        Set m_Source(0) = IBackup
                    End If
                End If
            End If
        End If
        On Error GoTo ExitRoutine
    End If
    
    ' see if this is a 32bit bitmap & process if it is
    If pvProcessAlphaBitmap(stream(), 0&, 0&, lClearFlag) = True Then
        LoadPicture_Stream = True
    Else
        Set IStream = pvStreamFromArray(VarPtr(stream(0)), UBound(stream) + 1&)
        If Not IStream Is Nothing Then
            If GdipLoadImageFromStream(IStream, hImage) = 0& Then
                lType = pvGetImageType(hImage)                      ' what type image do we have?
                If lType = imageWMF Or lType = imageEMF Then
                    If pvMetaFileTo32bpp(IStream, hImage, lClearFlag) = False Then Exit Function
                    m_ImageType = lType
                Else
                    Call GdipGetImagePixelFormat(hImage, lValue)    ' get original pixel format
                    If lType = imageGIF Or lType = imageTIFF Then
                        Set newMultiImage = New cGDIpMultiImage
                        newMultiImage.frSetImage hImage, lType
                        If newMultiImage.count = 1& Then
                            If lType = imageTIFF Then               ' single frame TIFF; make sourceless
                                hImage = pvCreateSourcelessImage(hImage, 0&, 0&, m_Mirror * Abs(ClearAttributes = False))
                                If hImage = 0& Then Exit Function
                            End If
                        End If
                        lClearFlag = lClearFlag Xor cuDestroyMultiImage
                        Set m_MultiImage = newMultiImage
                    Else
                        hImage = pvCreateSourcelessImage(hImage, 0&, 0&, m_Mirror * Abs(ClearAttributes = False)) ' convert to sourceless
                        If hImage = 0& Then Exit Function
                    End If
                    Call pvCleanUp(lClearFlag)                      ' clear existing image
                    m_Image(m_SourceIndex) = hImage
                    m_OrigColorType = lValue
                    Call GdipGetImageBounds(hImage, m_Size, UnitPixel)
                    m_ImageType = lType
                    If m_ImageType = imageGIF Or m_ImageType = imageTIFF Then   ' cache original data for gifs & multipage tiffs
                        If m_ImageType = imageGIF Or m_MultiImage.count > 1& Then Set m_Source(m_SourceIndex) = IStream
                        If (m_Attr Or m_Mirror) Then Call pvModifyAttributes  ' create clone if necessary
                    End If
                End If
                LoadPicture_Stream = True
            End If
        End If
    End If
    
ExitRoutine:
End Function

Public Function LockImageBits(ByVal LockMode As LockModeConstants, _
                            Optional ByRef PixelFormat As ImageColorFormatConstants = PixelFormat32bppARGB, _
                            Optional ByVal X As Long, Optional ByVal Y As Long, _
                            Optional ByRef Width As Long, Optional ByRef Height As Long, _
                            Optional ByRef StrideOut As Long) As Long
            
    ' Used to edit pixel data. Image cannot be modified and any updates don't take effect until unlocked
    ' Parameters
    '   :: [in] LockMode - must be one of the LockModeConstants
    '   :: [in/out] PixelFormat - getting lower level formats may fail; example trying to get a 8bpp format from a 32bpp image won't work
    '                   on return, the format is what was retrieved
    '   :: [in] X,Y - pixel coords to start the locking
    '   :: [in/out] Width,Height - size of image to be locked. On return, actual width/height locked
    '   :: [out] StrideOut - the scanwidth for the returned pixel format
    ' Return value: A pointer to the pixel data
    
    ' To unlock the bits, YOU MUST call UnLockImageBits and pass the exact values that were returned to you
    ' and supply a pointer to your updated array in the ScanPtr parameter if updating pixels: i.e., VarPtr(myPixels(0,0))
    
    If m_Image(m_SourceIndex) Then
        Dim BD As BitmapData, r As RECTI
        With r
            .nHeight = Height
            .nLeft = X
            .nTop = Y
            .nWidth = Width
        End With
        If GdipBitmapLockBits(m_Image(m_SourceIndex), r, LockMode, PixelFormat, BD) = 0& Then
            With BD
                PixelFormat = .PixelFormat
                Width = .Width
                Height = .Height
                StrideOut = .Stride
                LockImageBits = .Scan0Ptr
            End With
        End If
    End If
        
End Function

' Return the GDI+ image mirrored state
Public Property Get Mirrored() As MirrorConstants
    Mirrored = (m_Mirror \ 2&)
End Property
Public Property Let Mirrored(newVal As MirrorConstants)
    ' FYI: GDI+ actually modifies the bits of the GDI+ image when applying mirroring
    ' This is actually a faster way to render because the bits are already in the
    ' proper format and color-lossless. For our purposes, we may need to un-mirror
    ' on occassion to save/render without attributes applied in several of the routines.
    If newVal > -1& And newVal < 4& Then
        If (newVal <> m_Mirror \ 2) Then
            If newVal = attrMirrorNone Then
                If (m_Attr = 0& And m_SourceIndex = 1&) Then ' clone exists; but no longer needed
                    Call pvCleanUp(cuDestroyClone)
                    m_Mirror = attrMirrorNone
                End If
            Else
                ' do we need to create a clone; we do if m_multimage and not an icon/cursor
                If (m_ImageType = imageGIF Or m_ImageType = imageTIFF) Then
                    If m_MultiImage.count > 1& And m_SourceIndex = 0& Then ' not yet cloned
                        Debug.Assert (m_Image(1) = 0&)          ' testing purposes. Not executed when compiled
                        m_Mirror = (newVal * 2&)                ' set mirror state
                        pvCreateClone m_MultiImage.Index, 0&    ' create clone, got a multi-image GIF/TIF
                        Exit Property                           ' clone routine flips/applies attributes
                    End If
                End If
            End If
            GdipImageRotateFlip m_Image(m_SourceIndex), (m_Mirror Xor (newVal * 2))
            m_Mirror = newVal * 2&
        End If
    End If
End Property

Public Property Get MultiImageInfo() As cGDIpMultiImage
    ' allows user to navigate btwn multi-frame/page images
    ' also allows retrieval of animated GIF frame delays and loop count
    If m_MultiImage Is Nothing Then
        Set MultiImageInfo = New cGDIpMultiImage
    Else
        Set MultiImageInfo = m_MultiImage
    End If
End Property

Public Sub ResetImageAttributes()
    ' Removes mirroring, lightness, transparency, grayscaling, & rotation
    Call pvCleanUp(cuResetAttrs)
End Sub

' Return the GDI+ image angle
Public Property Get Rotation() As Single
    Rotation = m_Angle
End Property
Public Property Let Rotation(newVal As Single)
    m_Angle = newVal
End Property

Public Function SaveAsBMP(outArray() As Byte, Optional ApplyImageAttributes As Boolean = False) As Boolean

    ' saves image as a bitmap
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved
    
    Dim uEncCLSID(0 To 3) As Long, IIStream As IUnknown
    Dim cNewImage As cGDIpImage, hImage As Long
    Const MimeType As String = "image/bmp"
    
    If m_Image(m_SourceIndex) Then
        If pvGetEncoderClsID(MimeType, uEncCLSID) <> -1& Then
            Set IIStream = pvStreamFromArray(0&, 0&)
            If Not IIStream Is Nothing Then
                If ApplyImageAttributes Then
                    If Not (m_Attr = 0& And m_Angle = 0!) Then ' any attributes/rotation to apply?
                        Set cNewImage = New cGDIpImage
                        CloneImage cNewImage, , , True
                        hImage = cNewImage.Handle
                    Else
                        hImage = m_Image(m_SourceIndex)
                    End If
                Else
                    hImage = m_Image(m_SourceIndex)
                    ' mirroring will be transfered when image is saved; no attributes to apply so un-mirror as needed
                    If m_Mirror Then GdipImageRotateFlip hImage, m_Mirror
                End If
                If hImage Then
                    If GdipSaveImageToStream(hImage, IIStream, uEncCLSID(0), ByVal 0&) = 0& Then
                        SaveAsBMP = pvStreamToArray(ObjPtr(IIStream), outArray())
                    End If
                    If ApplyImageAttributes = False Then ' if we un-mirrored above, re-mirror now
                        If m_Mirror Then GdipImageRotateFlip hImage, m_Mirror
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function SaveAsJPG(outArray() As Byte, Optional ByVal QualityPct As Long = 70&, Optional ApplyImageAttributes As Boolean = False) As Boolean

    ' saves image as a JPG
    ' QualityPct must be between 10 and 100
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved
    
    Dim uEncCLSID(0 To 3) As Long, IIStream As IUnknown, hImage As Long
    Dim uEncParams As EncoderParameters, cNewImage As cGDIpImage
    Const MimeType As String = "image/jpeg"
    Const JPGEncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
    Const EncoderParameterValueTypeLong As Long = &H4&
    
    If m_Image(m_SourceIndex) Then
        If pvGetEncoderClsID(MimeType, uEncCLSID) <> -1& Then
            Set IIStream = pvStreamFromArray(0&, 0&)
            If Not IIStream Is Nothing Then
                If ApplyImageAttributes Then
                    If Not (m_Attr = 0& And m_Angle = 0!) Then ' any attributes/rotation to apply?
                        Set cNewImage = New cGDIpImage
                        CloneImage cNewImage, , , True
                        hImage = cNewImage.Handle
                    Else
                        hImage = m_Image(m_SourceIndex)
                    End If
                Else
                    hImage = m_Image(m_SourceIndex)
                    ' mirroring will be transfered when image is saved; no attributes to apply so un-mirror as needed
                    If m_Mirror Then GdipImageRotateFlip hImage, m_Mirror
                End If
                If hImage Then
                    uEncParams.count = 1
                    With uEncParams.Parameter(0)
                        .NumberOfValues = 1
                        .type = EncoderParameterValueTypeLong
                        Call CLSIDFromString(StrPtr(JPGEncoderQuality), .GUID(0))
                        If QualityPct < 10& Then
                            QualityPct = 10&
                        ElseIf QualityPct > 100& Then
                            QualityPct = 100&
                        End If
                        .Value = VarPtr(QualityPct)
                    End With
                
                    If GdipSaveImageToStream(hImage, IIStream, uEncCLSID(0), uEncParams) = 0& Then
                        SaveAsJPG = pvStreamToArray(ObjPtr(IIStream), outArray())
                    End If
                    If ApplyImageAttributes = False Then ' if we un-mirrored above, re-mirror now
                        If m_Mirror Then GdipImageRotateFlip hImage, m_Mirror
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function SaveAsOriginalFormat(outArray() As Byte) As Boolean
    
    ' Saves image from the original image format, no attributes applied
    ' This does not guarantee the returned byte array will be the same as the original source
    ' For example. PNGs may have many chunks in the original, but when we get the data from
    '   the converted internal bitmap, those chunks will no longer be there, but the returned
    '   bytes will be a PNG format.
    
    Dim IIStream As IUnknown, IPIC As IPicture
    Dim lSize As Long, bFalse As Boolean
    
    If m_Source(0) Is Nothing Then  ' a sourceless bitmap
        Select Case m_ImageType     ' convert to original format type
            Case imagePNG: SaveAsOriginalFormat = SaveAsPNG(outArray())
            Case imageJPEG: SaveAsOriginalFormat = SaveAsJPG(outArray(), 100&)
            Case imageBitmap: SaveAsOriginalFormat = SaveAsBMP(outArray())
            Case imageTIFF: SaveAsOriginalFormat = SaveAsTIFF(outArray, , True, TIFF_DepthAuto) ' single-page TIFFs
            Case Else: Debug.Assert (m_ImageType = imageUnknown) ' shouldn't be anything else; Assert for testing purposes only
        End Select
    Else
        If TypeOf m_Source(0) Is StdPicture Then
            Set IPIC = m_Source(0)
            Set IIStream = pvStreamFromArray(0&, 0&)
            If (IPIC Is Nothing Or IIStream Is Nothing) Then Exit Function
            ' have VB's stdPicture save to our stream
            IPIC.SaveAsFile ByVal ObjPtr(IIStream), bFalse, lSize
            If lSize = 0& Then Exit Function
        Else
            Set IIStream = m_Source(0)
        End If
        SaveAsOriginalFormat = pvStreamToArray(ObjPtr(IIStream), outArray)
    End If

End Function

Public Function SaveAsPNG(outArray() As Byte, Optional ApplyImageAttributes As Boolean = False) As Boolean

    ' saves image as a PNG
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved
    
    Dim uEncCLSID(0 To 3) As Long, IIStream As IUnknown, hImage As Long
    Dim cNewImage As cGDIpImage
    Const MimeType As String = "image/png"
    
    If m_Image(m_SourceIndex) Then
        If pvGetEncoderClsID(MimeType, uEncCLSID) <> -1& Then
            Set IIStream = pvStreamFromArray(0&, 0&)
            If Not IIStream Is Nothing Then
                If ApplyImageAttributes Then
                    If Not (m_Attr = 0& And m_Angle = 0!) Then ' any attributes/rotation to apply?
                        Set cNewImage = New cGDIpImage
                        CloneImage cNewImage, , , True
                        hImage = cNewImage.Handle
                    Else
                        hImage = m_Image(m_SourceIndex)
                    End If
                Else
                    hImage = m_Image(m_SourceIndex)
                    ' mirroring will be transfered when image is saved; no attributes to apply so un-mirror as needed
                    If m_Mirror Then GdipImageRotateFlip hImage, m_Mirror
                End If
                If hImage Then
                    If GdipSaveImageToStream(hImage, IIStream, uEncCLSID(0), ByVal 0&) = 0& Then
                        SaveAsPNG = pvStreamToArray(ObjPtr(IIStream), outArray())
                    End If
                    If ApplyImageAttributes = False Then ' if we un-mirrored above, re-mirror now
                        If m_Mirror Then GdipImageRotateFlip hImage, m_Mirror
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function SaveAsTIFF(outArray() As Byte, Optional ArrayOfImageClasses As Variant, _
                            Optional ByVal UseCompression As Boolean = True, _
                            Optional ByVal ColorDepth As TIFFColorDepthConstants = TIFF_DepthAuto, _
                            Optional ApplyImageAttributes As Boolean = False) As Boolean

    ' saves image a single/multi-page TIFF
    
    ' Parameters:
    '   outArray() :: byte array will be resized if function returns true
    '   ArrayOfImageClasses :: a 1D zero-bound array of cGDIpImage classes to create
    '                          a multipage-tiff else this class' image is used
    '   UseCompression :: if true tiff is compressed else it is not
    '   ColorDepth :: Color depth to save tiff page(s) in
    '       Notes regarding TIFF_DepthAuto:
    '           Icon, cursor always saved in 32bpp
    '           GIF will be saved as 24bpp if no attributes applied and GIF frame has no transparency index else 32bpp
    '   ApplyImageAttributes :: whether any/all images will be saved with attributes or not (i.e., rotation, grayscale, etc)
    
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved

    Dim uEncCLSID(0 To 3) As Long, gifPal() As Long, IIStream As IUnknown
    Dim uEncParams As EncoderParameters
    Dim cNewImages() As cGDIpImage, bUnMirror() As Boolean
    Dim Index As Long, lParamVal As Long, hImage As Long
    Dim lDepth As Long, lCompress As Long, nrImages As Long
    
    Const MimeType As String = "image/tiff"
    Const TIFFCompress As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
    Const TIFFBitDepth As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
    Const SaveFlag As String = "{292266fc-ac40-47bf-8cfc-a85b89a655de}"
    Const EncoderValueCompressionLZW As Long = &H2
    Const EncoderValueMultiFrame As Long = 18&
    Const EncoderValuePageDims As Long = 23&
    Const EncoderValueFlush As Long = 20&
    Const EncoderParameterValueTypeLong As Long = &H4&

    ' /// This section is used for sanity checks....
    If IsMissing(ArrayOfImageClasses) Then      ' is an array of images provided?
        If m_Image(m_SourceIndex) Then          ' if not, we use this class
            ReDim cNewImages(0 To 0)
            ReDim bUnMirror(0 To 0)
            If ApplyImageAttributes = True Then ' include rotation, grayscaling, mirroring, etc
                If (m_Attr = 0& And m_Angle = 0!) = True Then
                    Set cNewImages(0) = Me
                Else
                    Set cNewImages(0) = New cGDIpImage      ' create new class with attributes
                    If CloneImage(cNewImages(0), , , True) Then nrImages = 1&
                End If
            Else
                Set cNewImages(0) = Me          ' no attributes used at all
                nrImages = 1&: bUnMirror(0) = Not (m_Mirror = attrMirrorNone)
            End If
        End If
    Else
        ' validate passed classes have something to save & its an array of what we expect
        On Error Resume Next
        If Not VarType(ArrayOfImageClasses) = (vbObject Or vbArray) Then Exit Function
        nrImages = UBound(ArrayOfImageClasses) + 1&
        If Err Then
            Err.Clear
            Exit Function
        Else
            If LBound(ArrayOfImageClasses) <> 0 Then Exit Function
        End If
        ' check for invalid objects, empty objects, classes without images, etc
        ReDim cNewImages(0 To nrImages - 1&)
        ReDim bUnMirror(0 To nrImages - 1&)
        For Index = 0& To nrImages - 1&
            If Not ArrayOfImageClasses(Index) Is Nothing Then
                If TypeOf ArrayOfImageClasses(Index) Is cGDIpImage Then
                    ' ok, got a image class with valid handle
                    If ApplyImageAttributes = True Then
                        If (ArrayOfImageClasses(Index).ImageAttributesHandle = 0& And ArrayOfImageClasses(Index).Rotation = 0!) = True Then
                            ' only mirroring will be applicable, don't create clone
                            Set cNewImages(lParamVal) = ArrayOfImageClasses(Index)
                            lParamVal = lParamVal + 1&
                        Else ' create clone with applied attributes
                            Set cNewImages(lParamVal) = New cGDIpImage
                            If ArrayOfImageClasses(Index).CloneImage(cNewImages(lParamVal), , , True) Then lParamVal = lParamVal + 1&
                        End If
                    ElseIf ArrayOfImageClasses(Index).Handle Then
                        ' no attributes applied
                        Set cNewImages(lParamVal) = ArrayOfImageClasses(Index)
                        bUnMirror(lParamVal) = Not (cNewImages(lParamVal).Mirrored = attrMirrorNone): lParamVal = lParamVal + 1&
                    End If
                End If
            End If
            If Err Then
                Err.Clear
                Exit Function               ' unexpected error
            End If
        Next
        If lParamVal < nrImages Then
            If lParamVal = 0& Then Exit Function
            nrImages = lParamVal
        End If
        On Error GoTo 0
    End If
    
    ' /// This section creates the TIFF if no errors are encountered
    If nrImages Then    ' failure
    
        If pvGetEncoderClsID(MimeType, uEncCLSID) <> -1& Then   ' get TIFF encoder class
            Set IIStream = pvStreamFromArray(0&, 0&)            ' create IStream
            If Not IIStream Is Nothing Then
                
                If UseCompression Then                          ' use compression to keep bytes size low
                    With uEncParams.Parameter(0)
                        lCompress = EncoderValueCompressionLZW
                        .NumberOfValues = 1
                        .type = EncoderParameterValueTypeLong
                         CLSIDFromString StrPtr(TIFFCompress), .GUID(0)
                        .Value = VarPtr(lCompress)
                    End With
                    uEncParams.count = 1                        ' keep count
                End If
                
                With uEncParams.Parameter(uEncParams.count)     ' set up bit depth parameter
                    .NumberOfValues = 1
                    .type = EncoderParameterValueTypeLong
                    CLSIDFromString StrPtr(TIFFBitDepth), .GUID(0)
                   .Value = VarPtr(lDepth)                      ' will be applied per page
                End With
                uEncParams.count = uEncParams.count + 1         ' keep count
                
                If nrImages > 1& Then                           ' multiple images
                    lParamVal = EncoderValueMultiFrame          ' set multipage flag
                    With uEncParams.Parameter(uEncParams.count)
                        .NumberOfValues = 1
                        .type = EncoderParameterValueTypeLong
                        CLSIDFromString StrPtr(SaveFlag), .GUID(0)
                       .Value = VarPtr(lParamVal)
                    End With
                    uEncParams.count = uEncParams.count + 1     ' keep count
                    
                    For Index = 0 To nrImages - 1&              ' begin writing multiple pages
                        ' determine destination color depth
                        Select Case cNewImages(Index).ImageType
                        Case imageIcon, imageIconPNG, imageCursor, imageCursorPNG
                            ' without processing icon mask to verify, must assume icons/cursors/gifs use alpha channel
                            lDepth = 32&
                        Case imageGIF   ' if gif has transparency index, assume uses transparency
                            If cNewImages(Index).MultiImageInfo.GetGifFramePalette(gifPal(), lDepth) Then
                                If lDepth = -1& Then lDepth = 24& Else lDepth = 32&
                                Erase gifPal()
                            Else
                                lDepth = 32&
                            End If
                        Case Else
                            Select Case ColorDepth
                            Case TIFF_DepthAuto
                                Select Case cNewImages(Index).ColorFormat
                                Case PixelFormat24bppRGB, PixelFormat32bppRGB: lDepth = 24&
                                Case PixelFormat16bppRGB555, PixelFormat16bppRGB565, PixelFormat16bppGrayScale, PixelFormat48bppRGB: lDepth = 24&
                                Case PixelFormat1bppIndexed: lDepth = 1&
                                Case PixelFormat4bppIndexed: lDepth = 4&
                                Case PixelFormat8bppIndexed: lDepth = 8&
                                Case Else: lDepth = 32&
                                End Select
                            Case TIFF_Depth24Bit: lDepth = 24&
                            Case Else: lDepth = 32&
                            End Select
                        End Select
                        
                        ' un-mirror if mirrored and not wanting to save with attributes
                        If bUnMirror(Index) Then GdipImageRotateFlip cNewImages(Index).Handle, cNewImages(Index).Mirrored
                        If hImage = 0& Then ' already wrote first image?
                            If GdipSaveImageToStream(cNewImages(Index).Handle, IIStream, uEncCLSID(0&), uEncParams) = 0& Then
                                hImage = cNewImages(Index).Handle   ' cache first image handle
                                lParamVal = EncoderValuePageDims    ' set flag for non-1st page, additional pages
                            ElseIf lDepth < 32& Then ' failure, try to save as 32bit
                                lDepth = 32&
                                If GdipSaveImageToStream(cNewImages(Index).Handle, IIStream, uEncCLSID(0&), uEncParams) = 0& Then
                                    hImage = cNewImages(Index).Handle
                                    lParamVal = EncoderValuePageDims
                                End If
                            End If
                        Else        ' saving additional pages now
                            If GdipSaveAddImage(hImage, cNewImages(Index).Handle, uEncParams) Then
                                If lDepth < 32& Then ' failure try again with 32bpp
                                    lDepth = 32&
                                    Call GdipSaveAddImage(hImage, cNewImages(Index).Handle, uEncParams)
                                End If
                            End If
                        End If
                        If bUnMirror(Index) Then GdipImageRotateFlip cNewImages(Index).Handle, cNewImages(Index).Mirrored
                    Next
                    If hImage Then
                        lParamVal = EncoderValueFlush       ' tell GDI+ we are done adding pages
                        If GdipSaveAdd(hImage, uEncParams) = 0& Then SaveAsTIFF = pvStreamToArray(ObjPtr(IIStream), outArray())
                    End If
                    
                Else                                            ' single image; return array
                
                    Select Case cNewImages(0).ImageType
                    Case imageIcon, imageIconPNG, imageCursor, imageCursorPNG
                        lDepth = 32&
                    Case imageGIF
                        If cNewImages(Index).MultiImageInfo.GetGifFramePalette(gifPal(), lDepth) Then
                            If lDepth = -1& Then lDepth = 24& Else lDepth = 32&
                        Else
                            lDepth = 32&
                        End If
                    Case Else
                        Select Case ColorDepth
                        Case TIFF_DepthAuto
                            Select Case cNewImages(0).ColorFormat
                            Case PixelFormat24bppRGB, PixelFormat32bppRGB: lDepth = 24&
                            Case PixelFormat16bppRGB555, PixelFormat16bppRGB565, PixelFormat16bppGrayScale, PixelFormat48bppRGB: lDepth = 24&
                            Case PixelFormat1bppIndexed: lDepth = 1&
                            Case PixelFormat4bppIndexed: lDepth = 4&
                            Case PixelFormat8bppIndexed: lDepth = 8&
                            Case Else: lDepth = 32&
                            End Select
                        Case TIFF_Depth24Bit: lDepth = 24&
                        Case Else: lDepth = 32&
                        End Select
                    End Select
                    ' attempt to save
                    If bUnMirror(0) Then GdipImageRotateFlip cNewImages(0).Handle, cNewImages(0).Mirrored
                    SaveAsTIFF = (GdipSaveImageToStream(cNewImages(0).Handle, IIStream, uEncCLSID(0&), uEncParams) = 0&)
                    If SaveAsTIFF = False Then
                        If lDepth < 32& Then ' failure, try again with 32bpp
                            lDepth = 32&
                            SaveAsTIFF = (GdipSaveImageToStream(cNewImages(0).Handle, IIStream, uEncCLSID(0&), uEncParams) = 0&)
                        End If
                    End If
                    If bUnMirror(0) Then GdipImageRotateFlip cNewImages(0).Handle, cNewImages(0).Mirrored
                    If SaveAsTIFF Then SaveAsTIFF = pvStreamToArray(ObjPtr(IIStream), outArray())
                End If
            End If
        End If
    End If
    
End Function

Public Function SetPixel(ByVal X As Long, ByVal Y As Long, ByVal ARGBColor As Long) As Boolean
    ' function sets an ARGB color. X,Y must be within Width-1 & Height-1
    If m_Image(m_SourceIndex) Then SetPixel = (GdipBitmapSetPixel(m_Image(m_SourceIndex), X, Y, ARGBColor) = 0&)
End Function

Public Function UnLockImageBits(ByVal PixelFormat As ImageColorFormatConstants, _
                                ByVal Width As Long, ByVal Height As Long, _
                                ByVal Stride As Long, ByVal ScanPtr As Long) As Boolean
    
    ' Used to edit pixel data. Unlocks a previous call to LockImageBits and updates the image if necessary
    ' Parameters
    ' ... if you locked the bits for reading, pass back the same parameters that were received during the LockImageBits call
    '   :: PixelFormat - the PixelFormat of the pixels supplied for updating
    '   :: Width,Height - the size of image to be unlocked/updated
    '   :: Stride - the scanwidth of the pixel format supplied for updating
    '   :: ScanPtr - the original function return value of LockImageBits or a valid pointer to a new set of pixels
    ' CAUTION: Passing invalid parameters can crash. Pixels data must be Word aligned
    
    If m_Image(m_SourceIndex) Then
        Dim BD As BitmapData
        With BD
            .Height = Height
            .Width = Width
            .Scan0Ptr = ScanPtr
            .Stride = Stride
            .PixelFormat = PixelFormat
        End With
        UnLockImageBits = (GdipBitmapUnlockBits(m_Image(m_SourceIndex), BD) = 0&)
    End If
    
End Function

' Returns the image width
Public Property Get Width() As Long
    Width = m_Size.nWidth
End Property

Private Function pvByteAlignOnWord(ByVal bitDepth As Byte, ByVal Width As Long) As Long
    ' function to align any bit depth on dWord boundaries
    pvByteAlignOnWord = (((Width * bitDepth) + &H1F&) And Not &H1F&) \ &H8&
End Function

Private Sub pvCleanUp(mode As pvCleanUpEnum)
    ' m_Image & m_Source usage for clarification
    ' Image Type        m_Image(0)  m_Source(0)     m_Image(1)  m_Source(1)
    '   PNG             yes         never           never       never
    '   JPG             yes         never           never       never
    '   BMP             yes         never           never       never
    '   GIF             yes         yes**1          Cloned**2   Cloned**2
    '   TIFF            yes         yes**1          Cloned**2   Cloned**2
    '   ICO/CUR         yes         yes**1          never       never
    '   WMF/EMF         yes         yes**1          never       never
    ' Notes
    ' PNG, JPG, BMP, 1-page TIFFs are always converted to 1,4,8,24,32pARGB bpp bitmaps
    ' ICO, CUR, WMF, EMF are converted to 32bpp pARGB internally only
    ' (**1) Original source data/bytes are maintained so that multiple images from the resource can be extracted
    '       and so that original format can be returned on demand
    '       Note that the CUR/ICO/WMF/EMF source may be either stdPicture or IIStream depending on how the source was loaded
    '       Exception: Single page TIFFs do not cache original source data
    '       Note: ICO/CUR are cached for ease. One can easily add a BitmapToIcon routine to create the appropriate icon/cursor format
    ' (**2) When multiple images exist in resource, a clone (PNG format) is created if to be rendered with attributes.
    '       Any created clone is destroyed when the attributes are reset
    
    If (mode And cuResetAttrs) Then
        If m_Attr Then                  ' destroy GDI+ attributes object
            GdipDisposeImageAttributes m_Attr
            m_Attr = 0&
        End If
        If (mode And cuDestroyImage) = 0& Then ' not a full pvCleanUp?
            If m_Mirror Then            ' if mirrored & not cloned, then unmirror
                If m_Image(1) = 0& Then GdipImageRotateFlip m_Image(0), m_Mirror
            End If
        End If
        mode = mode Or cuDestroyClone   ' get rid of clone if needed
        m_Lightness = 0!                ' reset all attributes
        m_Alpha = 0!
        m_Angle = 0!
        m_TransColor = 0&
        m_Mirror = attrMirrorNone
        m_GrayScale = attrGrayNone
    End If
    
    If (mode And cuDestroyClone) Then   ' destroy clone if it exists
        If m_Image(1) Then
            GdipDisposeImage m_Image(1)
            m_Image(1) = 0&
            m_SourceIndex = 0&
            Set m_Source(1) = Nothing
        End If
    End If
    
    If (mode And cuDestroyImage) Then   ' destroy core image
        If m_Image(0) Then
            GdipDisposeImage m_Image(0)
            m_Image(0) = 0&
            Set m_Source(0) = Nothing
        End If
        m_Size.nHeight = 0!             ' reset its properties
        m_Size.nWidth = 0!
        If (mode And cuDestroyMultiImage) Then
            Set m_MultiImage = Nothing
            m_ImageType = imageUnknown
        End If
    End If
    
End Sub

Private Function pvCreateClone(Index As Long, resetFlags As pvCleanUpEnum) As Boolean

    ' Cloning rules
    ' GIF/TIFF
    '   :: more than 1 page/frame, flipping and/or attributes applied
    ' ICO/CUR
    '   :: never cloned, however, new image created on demand via pvIconTo32bpp called from this routine
    
    
    Dim uEncCLSID(0 To 3) As Long
    Dim hImage As Long, lColorType As Long, IIStream As IUnknown
    
    If (m_ImageType = imageIcon Or m_ImageType = imageCursor Or ((m_ImageType And 32&) = 32&)) Then ' icons/cursors
        pvCreateClone = pvIconTo32bpp(Index, (resetFlags Or cuDestroyClone Or cuDestroyImage) And Not cuDestroyMultiImage)
        ' when icon is processed any existing clone is cleaned up
    
    Else ' multi-frame/page GIF/TIFF
        ' create clone only if absolutely needed
        Call GdipGetImagePixelFormat(m_Image(0), lColorType)
        If ((m_Attr = 0& And m_Mirror = attrMirrorNone)) Then
            hImage = m_Image(m_SourceIndex)
            
        Else
            ' to workaround the bug with multi-page/frame formats, we create a PNG image to be rendered in place of the actual image
            ' See comments near top of module regarding this bug
            Set IIStream = pvStreamFromArray(0&, 0&)
            If pvGetEncoderClsID("image/png", uEncCLSID) <> -1& Then
                If GdipSaveImageToStream(m_Image(0), IIStream, uEncCLSID(0), ByVal 0&) = 0& Then
                    Call GdipLoadImageFromStream(IIStream, hImage)
                End If
                If hImage Then
                    Call pvCleanUp(resetFlags Or cuDestroyClone)
                    m_SourceIndex = 1&
                    m_Image(m_SourceIndex) = hImage
                    Set m_Source(m_SourceIndex) = IIStream
                    If m_Mirror Then GdipImageRotateFlip m_Image(m_SourceIndex), m_Mirror
                    pvCreateClone = True
                End If
            End If
            
        End If
        m_OrigColorType = lColorType
        Call GdipGetImageBounds(hImage, m_Size, UnitPixel) ' retrieve dimensions/offsets
    End If
    
End Function

Private Function pvCreateSourcelessImage2(hImage As Long, Width As Long, Height As Long, flipFlag As Long, TokenClass As cGDIpToken, hGraphics1 As Long) As Long

    ' function creates a stand-alone GDI+ image from a linked image
    ' The routine here follows the instructions given at this MS KB article link
    '       http://support.microsoft.com/kb/814675
    Dim tSize As RECTF, tSizeI As RECTI, BHI As BITMAPINFO
    Dim tBMPsrc As BitmapData, tBMPdst As BitmapData
    Dim cRender As cGDIpRenderer, pal() As Byte
    Dim oldImage As Long, newImage As Long, hObj As Long
    Dim dDC As Long, tDC As Long, Depth As Long, hGraphics As Long, dibPtr As Long
    
    If Width = 0& Or Height = 0& Then               ' get size of image
        Call GdipGetImageBounds(hImage, tSize, UnitPixel)
        Width = tSize.nWidth: Height = tSize.nHeight
    End If
    Call GdipGetImagePixelFormat(hImage, Depth)
    Select Case Depth
        Case PixelFormat1bppIndexed, PixelFormat4bppIndexed, PixelFormat8bppIndexed
            With BHI.bmiHeader
                If Depth = PixelFormat1bppIndexed Then
                    .biBitCount = 1
                ElseIf Depth = PixelFormat4bppIndexed Then
                    .biBitCount = 4
                Else
                    .biBitCount = 8
                End If
                Call GdipGetImagePaletteSize(hImage, .biClrImportant)
                ReDim pal(0 To .biClrImportant + 7&)
                GdipGetImagePalette hImage, pal(0), .biClrImportant
                CopyMemory BHI.bmiColors(0), pal(8), .biClrImportant
                .biClrImportant = .biClrImportant \ 4&
                .biClrUsed = 2 ^ .biBitCount
                Erase pal()
            End With
        Case PixelFormat24bppRGB, PixelFormat16bppGrayScale, PixelFormat16bppRGB555, PixelFormat16bppRGB565, PixelFormat48bppRGB
            BHI.bmiHeader.biBitCount = 24
                
        Case Else
            Set cRender = New cGDIpRenderer
            cRender.AttachTokenClass TokenClass
            oldImage = m_Image(m_SourceIndex)                   ' using the cGDIpRenderer class
            m_Image(m_SourceIndex) = hImage                     ' which requires this class to reflect correct image handle
            hGraphics = hGraphics1 ' create a generic graphics object
            If hGraphics Then                                   ' then create new GDI+ bitmap
                If GdipCreateBitmapFromGraphics(Width, Height, hGraphics, newImage) = 0& Then
                    cRender.DestroyHGraphics hGraphics
                    Call GdipGetImageBounds(hImage, tSize, UnitPixel)
                    m_Image(m_SourceIndex) = newImage           ' create graphics object around new bitmap
                    hGraphics = cRender.CreateGraphicsFromImageClass(Me)
                    If hGraphics Then                           ' draw the passed bitmap onto the new bitmap & clean up
                        Call cRender.RenderToHGraphics(hImage, hGraphics, 0&, 0&, Width, Height, tSize.nLeft, tSize.nTop, tSize.nWidth, tSize.nHeight)
                        cRender.DestroyHGraphics hGraphics
                        If flipFlag Then GdipImageRotateFlip newImage, flipFlag
                        pvCreateSourcelessImage2 = newImage
                    End If
                Else
                    cRender.DestroyHGraphics hGraphics
                End If
            End If
            m_Image(m_SourceIndex) = oldImage
    End Select
    
    If BHI.bmiHeader.biBitCount Then                            ' handle paletted & 24bit bitmaps here
        tSizeI.nHeight = Height: tSizeI.nWidth = Width
        With BHI.bmiHeader
            .biHeight = Height
            .biPlanes = 1
            .biSize = 40
            .biWidth = Width
            If .biBitCount = 24 Then
                If GdipBitmapLockBits(hImage, tSizeI, ImageLockModeRead, PixelFormat24bppRGB, tBMPsrc) Then .biBitCount = 0& 'flag meaning failure
            Else
                If GdipBitmapLockBits(hImage, tSizeI, ImageLockModeRead, Depth, tBMPsrc) Then .biBitCount = 0&
            End If
        End With
        If BHI.bmiHeader.biBitCount Then
            dDC = GetDC(GetDesktopWindow())
            hObj = CreateDIBSection(dDC, BHI, 0&, dibPtr, 0&, 0&)
            ReleaseDC GetDesktopWindow(), dDC
            If hObj Then
                CopyMemory ByVal dibPtr, ByVal tBMPsrc.Scan0Ptr, tBMPsrc.Stride * Height
                GdipBitmapUnlockBits newImage, tBMPsrc
                GdipCreateBitmapFromHBITMAP hObj, 0&, newImage
                DeleteObject hObj
                GdipImageRotateFlip newImage, (flipFlag Xor 6&)
                pvCreateSourcelessImage2 = newImage
            End If
        End If
    End If
    GdipDisposeImage hImage
    
End Function
Private Function pvCreateSourcelessImage(hImage As Long, Width As Long, Height As Long, flipFlag As Long) As Long

    ' function creates a stand-alone GDI+ image from a linked image
    ' The routine here follows the instructions given at this MS KB article link
    '       http://support.microsoft.com/kb/814675
    Dim tSize As RECTF, tSizeI As RECTI, BHI As BITMAPINFO
    Dim tBMPsrc As BitmapData, tBMPdst As BitmapData
    Dim cRender As cGDIpRenderer, pal() As Byte
    Dim oldImage As Long, newImage As Long, hObj As Long
    Dim dDC As Long, tDC As Long, Depth As Long, hGraphics As Long, dibPtr As Long
    
    If Width = 0& Or Height = 0& Then               ' get size of image
        Call GdipGetImageBounds(hImage, tSize, UnitPixel)
        Width = tSize.nWidth: Height = tSize.nHeight
    End If
    Call GdipGetImagePixelFormat(hImage, Depth)
    Select Case Depth
        Case PixelFormat1bppIndexed, PixelFormat4bppIndexed, PixelFormat8bppIndexed
            With BHI.bmiHeader
                If Depth = PixelFormat1bppIndexed Then
                    .biBitCount = 1
                ElseIf Depth = PixelFormat4bppIndexed Then
                    .biBitCount = 4
                Else
                    .biBitCount = 8
                End If
                Call GdipGetImagePaletteSize(hImage, .biClrImportant)
                ReDim pal(0 To .biClrImportant + 7&)
                GdipGetImagePalette hImage, pal(0), .biClrImportant
                CopyMemory BHI.bmiColors(0), pal(8), .biClrImportant
                .biClrImportant = .biClrImportant \ 4&
                .biClrUsed = 2 ^ .biBitCount
                Erase pal()
            End With
        Case PixelFormat24bppRGB, PixelFormat16bppGrayScale, PixelFormat16bppRGB555, PixelFormat16bppRGB565, PixelFormat48bppRGB
            BHI.bmiHeader.biBitCount = 24
                
        Case Else
            Set cRender = New cGDIpRenderer
            cRender.AttachTokenClass m_Token
            oldImage = m_Image(m_SourceIndex)                   ' using the cGDIpRenderer class
            m_Image(m_SourceIndex) = hImage                     ' which requires this class to reflect correct image handle
            hGraphics = cRender.CreateGraphicsFromImageClass(Me) ' create a generic graphics object
            If hGraphics Then                                   ' then create new GDI+ bitmap
                If GdipCreateBitmapFromGraphics(Width, Height, hGraphics, newImage) = 0& Then
                    cRender.DestroyHGraphics hGraphics
                    Call GdipGetImageBounds(hImage, tSize, UnitPixel)
                    m_Image(m_SourceIndex) = newImage           ' create graphics object around new bitmap
                    hGraphics = cRender.CreateGraphicsFromImageClass(Me)
                    If hGraphics Then                           ' draw the passed bitmap onto the new bitmap & clean up
                        Call cRender.RenderToHGraphics(hImage, hGraphics, 0&, 0&, Width, Height, tSize.nLeft, tSize.nTop, tSize.nWidth, tSize.nHeight)
                        cRender.DestroyHGraphics hGraphics
                        If flipFlag Then GdipImageRotateFlip newImage, flipFlag
                        pvCreateSourcelessImage = newImage
                    End If
                Else
                    cRender.DestroyHGraphics hGraphics
                End If
            End If
            m_Image(m_SourceIndex) = oldImage
    End Select
    
    If BHI.bmiHeader.biBitCount Then                            ' handle paletted & 24bit bitmaps here
        tSizeI.nHeight = Height: tSizeI.nWidth = Width
        With BHI.bmiHeader
            .biHeight = Height
            .biPlanes = 1
            .biSize = 40
            .biWidth = Width
            If .biBitCount = 24 Then
                If GdipBitmapLockBits(hImage, tSizeI, ImageLockModeRead, PixelFormat24bppRGB, tBMPsrc) Then .biBitCount = 0& 'flag meaning failure
            Else
                If GdipBitmapLockBits(hImage, tSizeI, ImageLockModeRead, Depth, tBMPsrc) Then .biBitCount = 0&
            End If
        End With
        If BHI.bmiHeader.biBitCount Then
            dDC = GetDC(GetDesktopWindow())
            hObj = CreateDIBSection(dDC, BHI, 0&, dibPtr, 0&, 0&)
            ReleaseDC GetDesktopWindow(), dDC
            If hObj Then
                CopyMemory ByVal dibPtr, ByVal tBMPsrc.Scan0Ptr, tBMPsrc.Stride * Height
                GdipBitmapUnlockBits newImage, tBMPsrc
                GdipCreateBitmapFromHBITMAP hObj, 0&, newImage
                DeleteObject hObj
                GdipImageRotateFlip newImage, (flipFlag Xor 6&)
                pvCreateSourcelessImage = newImage
            End If
        End If
    End If
    GdipDisposeImage hImage
    
End Function

'   -.-.-.-.-.-.-.-.-.-.-.-.-.-.-
'   pvGetEncdoerClsID Routine
'   -.-.-.-.-.-.-.-.-.-.-.-.-.-.-
' Routine is a helper function for the SavePicture routine
Private Function pvGetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
  
  Dim Num      As Long
  Dim SizeA    As Long
  Dim SizeB    As Long
  Dim LIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim buffer() As Byte, sMime As String
    
    '-- Get the encoder array size
    Call GdipGetImageEncodersSize(Num, SizeA)
    If (SizeA = 0&) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim buffer(1 To SizeA) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageEncoders(Num, SizeA, buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), buffer(1), (Len(ICI(1)) * Num))
    
    SizeA = Len(strMimeType)
    sMime = String$(SizeA, vbNullChar)
    '-- Loop through all the codecs
    For LIdx = Num To 1& Step -1&
        '-- Must convert the pointer into a usable string
        With ICI(LIdx)
            SizeB = lstrlenW(ByVal .MimeType)
            If SizeA = SizeB Then
                Call CopyMemory(ByVal StrPtr(sMime), ByVal .MimeType, SizeA * 2&)
                If sMime = strMimeType Then
                    CopyMemory ClassID(0), .ClassID(0), 16&
                    Exit For
                End If
            End If
        End With
    Next LIdx
    pvGetEncoderClsID = LIdx
End Function

Private Function pvGetImageType(hImage As Long) As ImageTypeConstants
    
    ' http://com.it-berater.org/gdiplus/noframes/GdiPlus_constants.htm
    Const ImageFormatBMP As String = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatEMF As String = "{B96B3CAC-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatGIF As String = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatIcon As String = "{B96B3CB5-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatJPEG As String = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatMemoryBMP As String = "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatPNG As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatTIFF As String = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatUndefined As String = "{B96B3CA9-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatWMF As String = "{B96B3CAD-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatEXIF As String = "{B96B3CB2-0728-11D3-9D7B-0000F81EF32E}"
    ' ^^ note to self: haven't seen any images return this yet. Curious
    Dim GUID(0 To 3) As Long, sGUID As String, lRet As Long
    
    If GdipGetImageRawFormat(hImage, VarPtr(GUID(0))) = 0& Then
        sGUID = String$(40, vbNullChar)
        lRet = StringFromGUID2(VarPtr(GUID(0)), StrPtr(sGUID), 40&)
        Select Case Left$(sGUID, lRet - 1&)
            Case ImageFormatPNG: pvGetImageType = imagePNG
            Case ImageFormatJPEG: pvGetImageType = imageJPEG
            Case ImageFormatWMF, ImageFormatEMF: pvGetImageType = imageWMF
            Case ImageFormatGIF: pvGetImageType = imageGIF
            Case ImageFormatTIFF: pvGetImageType = imageTIFF
            Case ImageFormatEXIF: pvGetImageType = imageExIF
            Case ImageFormatBMP, ImageFormatMemoryBMP: pvGetImageType = imageBitmap
            Case ImageFormatIcon: pvGetImageType = imageIcon
            Case Else: pvGetImageType = imageUnknown
        End Select
    Else
        pvGetImageType = imageUnknown
    End If

End Function

Private Function pvIconTo32bpp(Index As Long, resetFlags As pvCleanUpEnum) As Boolean

    ' GDI+ has some serious issues with icons/cursors. So a nearly complete icon parser/converter is provided below
    ' - doesn't appear to even load cursors
    ' - doesn't like 24/32bpp icons
    ' - can't load PNG-embedded icons (Vista+)
    ' - even has some issues with 1-4 bpp icons
    
    ' So to support icon/cursor formats, all formats are converted to 32bpp images
    ' Supported bit depths are 1, 2, 4, 8, 16, 24, 32 and PNGs
    
    Dim X As Long, Y As Long, bBits() As Byte, bMask() As Byte, pow2(0 To 10) As Long
    Dim scanWidth As Long, tgtScanWidth As Long, arrOffset As Long
    Dim hImage As Long, hDib As Long, dibPtr As Long, dDC As Long
    Dim imageOffset As Long, bitShift As Long, bitAnd As Long, bitDepth As Long
    Dim srcPtr As Long, dstPtr As Long, palIndex As Long, lColor As Long, bFalse As Boolean
    Dim BHI As BITMAPINFO, IIStream As IUnknown, IIBackup As IUnknown, IPIC As IPicture
    

    ' Test for PNG encoded icon. Offset=location in icon source where icon starts
    If TypeOf m_Source(0) Is StdPicture Then
        Set IIStream = pvStreamFromArray(0&, 0&)
        If Not IIStream Is Nothing Then
            Set IPIC = m_Source(0)
            IPIC.SaveAsFile ByVal ObjPtr(IIStream), bFalse, X ' have VB's stdPicture save icon to stream for us
            If X = 0& Then Exit Function ' failed
            If pvStreamToArray(ObjPtr(IIStream), bBits()) = False Then Exit Function
            Set IIStream = Nothing
            Set IPIC = Nothing
        End If
    Else
        If pvStreamToArray(ObjPtr(m_Source(0)), bBits()) = False Then Exit Function
    End If
    
    arrOffset = (Index And &HFF&) * 16& + 6&    ' get icon directory
    CopyMemory Y, bBits(arrOffset + 8&), 4&     ' get size of resource
    CopyMemory arrOffset, bBits(arrOffset + 12&), 4& ' get resource offset from the icon directory
    CopyMemory X, bBits(arrOffset), 4&          ' test for PNG header1
    If X = png_Signature1 Then
        CopyMemory X, bBits(arrOffset + 4&), 4& ' test for PNG header2
        If X = png_Signature2 Then
            ' we have a PNG, try to create GDI+ image
            Set IIStream = pvStreamFromArray(VarPtr(bBits(arrOffset)), Y)
            If Not IIStream Is Nothing Then
                If GdipLoadImageFromStream(IIStream, hImage) = 0& Then
                    hImage = pvCreateSourcelessImage(hImage, 0&, 0&, 6&)
                    If hImage Then
                        Set IIBackup = m_Source(0)  ' ensure cleanup doesn't destroy source
                        Call pvCleanUp(resetFlags)
                        Set m_Source(0) = IIBackup
                        m_Image(m_SourceIndex) = hImage
                        m_OrigColorType = PixelFormat32bppARGB
                        m_ImageType = m_ImageType Or 32 ' add PNG flag
                        pvIconTo32bpp = True
                    End If
                End If
            End If
        End If
    Else    ' palettized or 16/24/32 bit icon
            
        CopyMemory BHI, bBits(arrOffset), 40&   ' get 1st 40 bytes of embedded header
        BHI.bmiHeader.biCompression = 0&        ' ensure non-compressed bits returned when needed
        pow2(0) = 1&                        ' build a power of two lookup table
        For X = 1& To 10&                   ' Note. 2^10 is 16bit max shift value for 2-byte packed value
            pow2(X) = pow2(X - 1) * 2&
        Next
        With BHI.bmiHeader                  ' update header and set offsets/pointers needed to parse icon
            .biHeight = .biHeight \ 2       ' icons/cursors have double heights
            imageOffset = .biSize + arrOffset  ' where pixel data starts
            bitDepth = .biBitCount
            If .biBitCount < 9 Then         ' paletted icon/cursor
                Select Case bitDepth
                    Case 1, 4, 8 ' no issues
                    Case 2: .biBitCount = 4 ' should never see non-png 2 bit icon/cursor
                    Case Else: Exit Function ' unsupported; flat out wrong
                End Select
                ' fix up .biClrUsed if needed; adjust imageOffset, copy colors into our local header
                If .biClrUsed = 0& Then .biClrUsed = pow2(bitDepth)
                CopyMemory BHI.bmiColors(0), bBits(imageOffset), .biClrUsed * 4&
                imageOffset = imageOffset + .biClrUsed * 4&
                .biClrUsed = pow2(.biBitCount)
            End If
            scanWidth = pvByteAlignOnWord(bitDepth, .biWidth)   ' source scan width
            tgtScanWidth = .biWidth * 4&                        ' scan width of our 32bit DIB
            .biSizeImage = tgtScanWidth * .biHeight             ' calc target pixel image size
            ReDim bMask(0 To .biSizeImage - 1&)                 ' target pixel data
            
            Select Case .biBitCount
            Case 1, 4, 8                                        ' paletted bitmaps
                bitAnd = pow2(bitDepth) - 1&                    ' calc AND mask for packed bytes
                For Y = 0& To .biHeight - 1&                    ' unpack bytes to 32bpp format
                    srcPtr = Y * scanWidth + imageOffset        ' where current source row starts
                    dstPtr = Y * tgtScanWidth
                    bitShift = 8& - bitDepth                    ' calc bit shifting start
                    For X = 0& To .biWidth - 1&                 ' get palette index from packed byte
                        palIndex = ((bBits(srcPtr) \ pow2(bitShift)) And bitAnd)
                        CopyMemory bMask(dstPtr), BHI.bmiColors(palIndex), 4& ' copy color to target
                        dstPtr = dstPtr + 4&                    ' move target pointer along
                        If bitShift = 0& Then                   ' done with this byte?
                            bitShift = 8& - bitDepth            ' reset
                            srcPtr = srcPtr + 1&                ' move source pointer along
                        Else
                            bitShift = bitShift - bitDepth      ' adjust bit shifting value
                        End If
                    Next
                Next
                Select Case .biBitCount
                    Case 1: lColor = PixelFormat1bppIndexed
                    Case Is < 5: lColor = PixelFormat4bppIndexed
                    Case Else: lColor = PixelFormat8bppIndexed
                End Select
            Case 32                                             ' true color with alpha channel
                For Y = 0& To .biHeight - 1&                    ' move source pixels to target
                    srcPtr = Y * scanWidth + imageOffset        ' where current source row starts
                    dstPtr = Y * tgtScanWidth
                    For X = dstPtr To dstPtr + tgtScanWidth - 1& Step 4& ' transfer pixel data
                        CopyMemory bMask(X), bBits(srcPtr), 4&  ' ARGB format in source; no mask parsing required
                        srcPtr = srcPtr + 4&                    ' move source pointer along
                    Next
                Next
                lColor = PixelFormat32bppARGB
                
            Case 24                                             ' true color without alpha channel
                For Y = 0& To .biHeight - 1&                    ' move source pixels to target
                    srcPtr = Y * scanWidth + imageOffset        ' where current source row starts
                    dstPtr = Y * tgtScanWidth
                    For X = dstPtr To dstPtr + tgtScanWidth - 1& Step 4& ' transfer pixel data
                        CopyMemory bMask(X), bBits(srcPtr), 3&
                        srcPtr = srcPtr + 3&                    ' move source pointer along
                    Next
                Next
                lColor = PixelFormat24bppRGB
                
            Case 16                                             ' 16bpp images, assumes 5-bits for each R,G,B (i.e, 555 format)
                For Y = 0& To .biHeight - 1&
                    srcPtr = Y * scanWidth + imageOffset        ' where current source row starts
                    dstPtr = Y * tgtScanWidth
                    bitShift = 0&                               ' reset
                    For X = 0& To .biWidth * 3& - 1&            ' get next 16bits (Integer)
                        If bitShift = 0& Then CopyMemory lColor, bBits(srcPtr), 2&
                        palIndex = ((lColor \ pow2(bitShift)) And &H1F&) ' extract RGB component value
                        If palIndex = &H1F& Then                ' set value to target component
                            bMask(dstPtr) = 255
                        Else
                            bMask(dstPtr) = (palIndex * 256&) \ &H1F&
                        End If
                        dstPtr = dstPtr + 1&                    ' move to next target byte
                        If bitShift = 10& Then                  ' done with this integer?
                            bitShift = 0&                       ' reset
                            srcPtr = srcPtr + 2&                ' move source pointer along
                            dstPtr = dstPtr + 1&                ' skip over the target alpha byte
                        Else
                            bitShift = bitShift + 5&            ' adjust bit shifting value
                        End If
                    Next
                Next
                lColor = PixelFormat16bppRGB555
                
            Case Else
                Exit Function
            End Select
            
            ' need to apply mask now
            If .biBitCount < 32& Then
                imageOffset = imageOffset + scanWidth * .biHeight ' where mask starts in source data
                scanWidth = pvByteAlignOnWord(1, .biWidth)        ' mask scan width
                For Y = 0& To .biHeight - 1&
                    srcPtr = Y * scanWidth + imageOffset        ' where current mask row starts
                    bitShift = 7&                               ' reset
                    dstPtr = Y * tgtScanWidth + 3&              ' where current row's alpha starts
                    For X = 0 To .biWidth - 1&                  ' make pixel is transparent if needed
                        If ((bBits(srcPtr) \ pow2(bitShift)) And 1&) = 0 Then bMask(dstPtr) = 255
                        dstPtr = dstPtr + 4&                    ' move target pointer along
                        If bitShift = 0& Then                   ' done shifting
                            bitShift = 7&                       ' reset
                            srcPtr = srcPtr + 1&                ' move source pointer along
                        Else
                            bitShift = bitShift - 1&            ' adjust bit shifting value
                        End If
                    Next
                Next
            End If
            
            .biBitCount = 32                                    ' prepare to create DIB to wrap GDI+ image around
            .biClrUsed = 0&
            .biClrImportant = 0&
            dDC = GetDC(GetDesktopWindow)                       ' create the DIB
            hDib = CreateDIBSection(dDC, BHI, 0&, dibPtr, 0&, 0&)
            ReleaseDC GetDesktopWindow(), dDC
        End With
        If hDib Then
            CopyMemory ByVal dibPtr, bMask(0), BHI.bmiHeader.biSizeImage ' copy the unpacked pixel data to DIB
            Erase bMask()                                   ' create the GDI+ image
            If GdipCreateBitmapFromScan0(BHI.bmiHeader.biWidth, BHI.bmiHeader.biHeight, tgtScanWidth, PixelFormat32bppARGB, ByVal dibPtr, hImage) = 0& Then
                hImage = pvCreateSourcelessImage(hImage, BHI.bmiHeader.biWidth, BHI.bmiHeader.biHeight, 0&)
                If hImage Then
                    Set IIStream = m_Source(0)
                    Call pvCleanUp(resetFlags)
                    Set m_Source(0) = IIStream
                    m_Image(m_SourceIndex) = hImage
                    m_OrigColorType = lColor
                    m_ImageType = (m_ImageType And Not 32) ' remove any PNG flag
                    pvIconTo32bpp = True
                End If
            End If
            DeleteObject hDib
        End If
    End If
    If pvIconTo32bpp = True Then
        If (Index And &H80000000) Then ' called from the Load function not m_MultiImage
            CopyMemory X, bBits(2), 2&  ' get type of resource
            If (X And &HFF&) = 1& Then
                m_ImageType = imageIcon Or (m_ImageType And 32&)
            Else
                m_ImageType = imageCursor Or (m_ImageType And 32&)
            End If
            CopyMemory X, bBits(4), 2&  ' get count in resource
            Set m_MultiImage = New cGDIpMultiImage
            m_MultiImage.frSetImage (X And &HFF&), m_ImageType
        End If
        GdipImageRotateFlip hImage, (m_Mirror Xor 6)
        Call GdipGetImageBounds(hImage, m_Size, UnitPixel) ' retrieve dimensions/offsets
    End If
    
End Function

Private Function pvMetaFileTo32bpp(metaSource As IUnknown, theImage As Long, resetFlags As pvCleanUpEnum) As Boolean

    ' function converts a metafile Stream/stdPicture to bitmap so we can render it with image attributes

    Dim hImage As Long, xyRatio As Single
    Dim cX As Long, cY As Long, lColorType As Long
    Dim bFalse As Boolean, lSize As Long, dibPtr As Long
    Dim IIStream As IUnknown, IPIC As IPicture, Bounds As RECTF
    Dim BHI As BITMAPINFO, cRenderer As cGDIpRenderer
    Dim dDC As Long, tDC As Long, scanWidth As Long
    Dim hObj As Long, hDib As Long, hGraphics As Long
    Dim bBits() As Byte
    Const PICTURE_TRANSPARENT = &H2
    
    If TypeOf metaSource Is StdPicture Then        ' called from LoadPicture_stdPicture
        Set IPIC = metaSource
        Set IIStream = IPIC                                         ' new source if load is successful
        If (IPIC.Attributes And PICTURE_TRANSPARENT) Then lColorType = PixelFormat32bppARGB Else lColorType = PixelFormat24bppRGB
        hImage = IPIC.Handle                                        ' make non-zero
        cX = IPIC.Width * 1440& / 2540& / Screen.TwipsPerPixelX     ' get size in pixels
        cY = IPIC.Height * 1440& / 2540& / Screen.TwipsPerPixelY
    Else                                        ' called from LoadPicture_Stream
        Call GdipGetImageHorizontalResolution(theImage, xyRatio)     ' calc size
        If xyRatio > 0! Then
            Call GdipGetImageBounds(theImage, Bounds, UnitPixel)        ' get reported size & convert to himetric to pixels
            cX = Bounds.nWidth * (2540! / xyRatio) * 1440! / 2540! / Screen.TwipsPerPixelX
            If GdipGetImageVerticalResolution(theImage, xyRatio) = 0& Then
                If xyRatio > 0! Then
                    cY = Bounds.nHeight * (2540! / xyRatio) * 1440! / 2540! / Screen.TwipsPerPixelY
                    lColorType = PixelFormat32bppARGB
                    hImage = theImage
                    Set IIStream = metaSource
                End If
            End If
        End If
    End If
    If hImage = 0& Then                                         ' bug out
        If theImage Then GdipDisposeImage theImage
    Else
        With BHI.bmiHeader                                      ' create DIB to render metafile to
            If lColorType = PixelFormat24bppRGB Then .biBitCount = 24 Else .biBitCount = 32
            .biHeight = cY
            .biWidth = cX
            .biPlanes = 1
            .biSize = 40
            .biSizeImage = pvByteAlignOnWord(.biBitCount, cX) * cY
        End With
        dDC = GetDC(GetDesktopWindow())
        tDC = CreateCompatibleDC(dDC)
        ReleaseDC GetDesktopWindow(), dDC
        hDib = CreateDIBSection(tDC, BHI, 0&, dibPtr, 0&, 0&)
        If hDib = 0& Then                                       ' bug out
            If theImage Then GdipDisposeImage theImage
            DeleteDC tDC
        Else
            ' Ok, metafile conversion logic. When a metafile is drawn, only the pixels in the DIB that are touched
            ' by the metafile convert any alpha byte to zero.  So to keep transparency while converting to bitmap,
            ' we will prefill the DIB with opaque white including alpha. Now when drawn, the metafile changes touched
            ' alpha values to zero and we know where the transparent pixels are: they have an alpha of 255.
            If BHI.bmiHeader.biBitCount = 32 Then FillMemory ByVal dibPtr, BHI.bmiHeader.biSizeImage, 255
            hObj = SelectObject(tDC, hDib)
            If IPIC Is Nothing Then                             ' rendering from stream object not stdPicture object
                Set cRenderer = New cGDIpRenderer
                cRenderer.AttachTokenClass m_Token
                hGraphics = cRenderer.CreateHGraphicsFromHDC(tDC)
                If hGraphics Then
                    cRenderer.RenderToHGraphics hImage, hGraphics, 0&, 0&, cX, cY, Bounds.nLeft, Bounds.nTop, Bounds.nWidth, Bounds.nHeight
                    cRenderer.DestroyHGraphics hGraphics
                Else
                    GdipDisposeImage hImage: hImage = 0&    ' but out; theImage=hImage
                End If
                Set cRenderer = Nothing
            Else                                                ' rendering from stdPicture object
                IPIC.Render tDC + 0&, 0&, 0&, cX + 0&, cY + 0&, 0&, IPIC.Height, IPIC.Width, -IPIC.Height, ByVal 0&
            End If
            SelectObject tDC, hObj
            DeleteDC tDC
            
            If hImage = 0& Then                                 ' continue?
                DeleteObject hDib
            Else
                If IPIC Is Nothing Then GdipDisposeImage theImage ' destroy passed theImage
                hImage = 0&
                If BHI.bmiHeader.biBitCount = 24 Then           ' create new bitmap from DIB; no alpha checks necessary. Done.
                    GdipCreateBitmapFromHBITMAP hDib, 0&, hImage
                Else                                            ' GDI+ can't do alpha bitmaps, so we do it manually
                    scanWidth = cX * 4&
                    ReDim bBits(0 To BHI.bmiHeader.biSizeImage - 1&)
                    CopyMemory bBits(0), ByVal dibPtr, BHI.bmiHeader.biSizeImage
                    For lSize = 3& To BHI.bmiHeader.biSizeImage - 1& Step 4& ' toggle alpha to opaque and vice versa
                        bBits(lSize) = bBits(lSize) Xor 255
                    Next
                    CopyMemory ByVal dibPtr, bBits(0), BHI.bmiHeader.biSizeImage
                    Erase bBits()
                                                                ' create new GDI+ image & convert to stand alone image
                    GdipCreateBitmapFromScan0 cX, cY, scanWidth, PixelFormat32bppPARGB, ByVal dibPtr, hImage
                    If hImage Then hImage = pvCreateSourcelessImage(hImage, cX, cY, 6&)
                End If
                DeleteObject hDib                               ' no longer needed
                If hImage Then                                  ' success?
                    Call pvCleanUp(resetFlags)                  ' replace existing image/properties with new
                    Set m_Source(0) = IIStream
                    m_Image(m_SourceIndex) = hImage
                    m_OrigColorType = lColorType
                    If m_Mirror Then GdipImageRotateFlip m_Image(m_SourceIndex), m_Mirror
                    Call GdipGetImageBounds(hImage, m_Size, UnitPixel) ' retrieve dimensions/offsets
                    pvMetaFileTo32bpp = True
                End If
            End If
        End If
        
    End If
    
End Function

Private Sub pvModifyAttributes()
    
    Dim clrMatrix(0 To 4, 0 To 4) As Single
    Dim r As Single, G As Single, b As Single
    Const ColorAdjustTypeBitmap As Long = &H1&
    
    ' about attributes.
    ' The following are added to a GDI+ attributes handle:
    '       GrayScale, GlobalTranparency, GlobalLightness, ExtraTransparency
    ' Mirroring is done inside the GDI+ image itself
    ' Rotation is done via DC WorldTransformation on-the-fly when rendering
    ' So to determine if any attributes are in play, 3 variables must be checked:
    '   Inplay = ((Me.ImageAttributesHandle=0& And Me.Rotation=0! And Me.Mirrored=attrMirrorNone) = False)
    
    If m_Image(0) Then
        If m_Attr Then
            GdipDisposeImageAttributes m_Attr
            m_Attr = 0&
        End If
        
        If m_GrayScale > attrGrayNone Then ' apply grayscale ratios
            Select Case m_GrayScale
            Case attrGrayNTSCPAL
                r = 0.299: G = 0.587: b = 0.114 ' standard weighted average
            Case attrGraySimpleAverage
                r = 0.333: G = 0.334: b = r     ' pure average
            Case attrGrayCCIR709
                r = 0.213: G = 0.715: b = 0.072 ' CCIR709
            Case attrGrayRedMask
                r = 0.8: G = 0.1: b = G     ' personal preferences: could be r=1:g=0:b=0 or other weights
            Case attrGrayGreenMask
                r = 0.1: G = 0.8: b = r     ' personal preferences: could be r=0:g=1:b=0 or other weights
            Case attrGrayBlueMask
                r = 0.1: G = r: b = 0.8     ' personal preferences: could be r=0:g=0:b=1 or other weights
            Case attrGrayBlueGreenMask
                r = 0.1: G = 0.45: b = G    ' personal preferences: could be r=0:g=.5:b=.5 or other weights
            Case attrGrayRedGreenMask
                r = 0.45: G = r: b = 0.1    ' personal preferences: could be r=.5:g=.5:b=0 or other weights
            End Select
            ' grayscale the image
            clrMatrix(0, 0) = r: clrMatrix(1, 0) = r: clrMatrix(2, 0) = r
            clrMatrix(0, 1) = G: clrMatrix(1, 1) = G: clrMatrix(2, 1) = G
            clrMatrix(0, 2) = b: clrMatrix(1, 2) = b: clrMatrix(2, 2) = b
            clrMatrix(3, 3) = CSng((100! - m_Alpha) / 100!)  ' global blending; value between 0 & 1
            clrMatrix(4, 4) = 1! ' required; cannot be anything else
        End If
        If m_Lightness <> 0! Then ' add/subtract light intensity
            If clrMatrix(4, 4) = 0! Then
                clrMatrix(0, 0) = 1!: clrMatrix(1, 1) = 1!: clrMatrix(2, 2) = 1!
                clrMatrix(3, 3) = CSng((100! - m_Alpha) / 100!)  ' global blending; value between 0 & 1
                clrMatrix(4, 4) = 1! ' required; cannot be anything else
            End If
            clrMatrix(0, 4) = m_Lightness / 100! ' red added/subtracted brightness
            clrMatrix(1, 4) = clrMatrix(0, 4)    ' same for blue
            clrMatrix(2, 4) = clrMatrix(0, 4)    ' same for green
        End If
        If m_Alpha <> 0! Then ' add global transparency
            If clrMatrix(4, 4) = 0! Then
                clrMatrix(0, 0) = 1!: clrMatrix(1, 1) = 1!: clrMatrix(2, 2) = 1!
                clrMatrix(3, 3) = CSng((100! - m_Alpha) / 100!) ' global blending; value between 0 & 1
                clrMatrix(4, 4) = 1! ' required; cannot be anything else
            End If
        End If
        If clrMatrix(4, 4) = 1! Then ' create attributes?
            If GdipCreateImageAttributes(m_Attr) = 0& Then
                If GdipSetImageAttributesColorMatrix(m_Attr, ColorAdjustTypeBitmap, 1&, clrMatrix(0, 0), clrMatrix(0, 0), 0&) Then
                    GdipDisposeImageAttributes m_Attr
                    m_Attr = 0&
                End If
            End If
        End If
        If m_TransColor Then
            If m_Attr = 0& Then Call GdipCreateImageAttributes(m_Attr)
            If m_Attr Then
                If GdipSetImageAttributesColorKeys(m_Attr, 1&, 1&, m_TransColor, m_TransColor) Then ' failure
                    If clrMatrix(4, 4) = 0! Then
                        GdipDisposeImageAttributes m_Attr
                        m_Attr = 0&
                    End If
                End If
            End If
        End If
        
        If (m_Attr Or m_Mirror) Then ' need a clone?
            ' If so, no clone must exist and attributes and/or mirroring must be applied
            ' also only applies to GIFs/TIFFs, and only if more than one frame/page exists
            If (m_Image(1) = 0&) And (m_ImageType = imageGIF Or m_ImageType = imageTIFF) Then
                If m_MultiImage.count > 1& Then Call pvCreateClone(0, 0&)
            End If
        ElseIf m_Image(1) Then
            Call pvCleanUp(cuDestroyClone)
        End If
        
    End If
    
End Sub

Private Function pvProcessAlphaBitmap(inArray() As Byte, Width As Long, Height As Long, resetFlag As pvCleanUpEnum) As Boolean

    ' This is a non-foolproof method of determining whether image uses ARGB or pARGB alpha channel

    Dim X As Long, Y As Long, scanWidth As Long, hDib As Long, dibPtr As Long, dDC As Long
    Dim b As Long, lColorFormat As Long, hImage As Long, hdrSize As Long
    Dim tPIC As StdPicture, nullAlphas As Long, fullAlphas As Long
    
    If Width = 0& And Height = 0& Then      ' called from LoadPicture_Stream
        CopyMemory b, inArray(0), 2&        ' magic number
        If b <> bmp_Signature Then Exit Function  ' not a bitmap
        CopyMemory X, inArray(28), 2&       ' bit depth
        If X <> 32 Then Exit Function       '  not processed here
        CopyMemory b, inArray(26), 2&       ' validate planes (must be 1)
        If b <> 1& Then Exit Function
        CopyMemory Width, inArray(18), 4&
        CopyMemory Height, inArray(22), 4&  ' validate sizes
        If Width < 1 Or Height < 1 Then Exit Function
        CopyMemory hdrSize, inArray(14), 4&
        ' validate image size compared file size
        b = Width * Height * 4 + hdrSize + 13
        If b > UBound(inArray) Then Exit Function
        hdrSize = hdrSize + 14&
    Else
        dibPtr = VarPtr(inArray(0))         ' called from LoadPicture_DIBhandle
    End If
    
    scanWidth = Width * 4&                  ' bytes per row
    For Y = 0& To Height - 1                ' loop thru image, abort early if possible
        X = hdrSize + Y * scanWidth
        For X = X + 3& To X + scanWidth - 1& Step 4&
            If inArray(X) = 0 Then
                nullAlphas = nullAlphas + 1& ' track zero alpha bytes
            ElseIf inArray(X) = 255 Then
                fullAlphas = fullAlphas + 1& ' track fully opaque bytes
            Else
                For b = X - 3& To X - 1&
                    If inArray(b) > inArray(X) Then
                        ' pARGB can never have a R,G,B value > alpha value so this must be ARGB
                        lColorFormat = PixelFormat32bppARGB ' ARGB format; done checking
                        X = X + scanWidth   ' exit X loop
                        Y = Height          ' exit Y loop
                        Exit For
                    End If
                Next
            End If
        Next
    Next
    ' decision time
    b = Width * Height ' total number of bytes/pixels
    ' if the image is not 100% zero alpha or 100% full alpha then it uses the alpha channel
    If ((nullAlphas = b) Or (fullAlphas = b)) Then  ' 32bpp with no alpha in play
        lColorFormat = PixelFormat32bppRGB
    Else
        ' if we didn't detect ARGB formatting then assume pARGB formatting
        ' This logic isn't 100% foolproof but exceptions to this logic would be extremely rare
        If lColorFormat = 0& Then lColorFormat = PixelFormat32bppPARGB
    End If
    ' because GDI+ won't load alphachannel bitmaps we will create a DIB and have GDI+ load from it
    If dibPtr = 0& Then
        dDC = GetDC(GetDesktopWindow())
        hDib = CreateDIBSection(dDC, inArray(14), 0&, dibPtr, 0&, 0&)
        ReleaseDC GetDesktopWindow(), dDC
        If hDib Then
            CopyMemory ByVal dibPtr, inArray(hdrSize), scanWidth * Height
        Else
            Exit Function
        End If
    End If
    If lColorFormat = PixelFormat32bppRGB Then
        GdipCreateBitmapFromHBITMAP hDib, 0&, hImage
    Else
        Call GdipCreateBitmapFromScan0(Width, Height, scanWidth, lColorFormat, ByVal dibPtr, hImage)
        If hImage Then hImage = pvCreateSourcelessImage(hImage, Width, Height, 0&)
    End If
    If hDib Then DeleteObject hDib
    If hImage Then
        Call pvCleanUp(resetFlag)
        m_Image(m_SourceIndex) = hImage
        m_ImageType = imageBitmap
        If lColorFormat = PixelFormat32bppRGB Then
            If m_Mirror Then GdipImageRotateFlip hImage, m_Mirror
        Else
            GdipImageRotateFlip hImage, (m_Mirror Xor 6&)
        End If
        m_OrigColorType = lColorFormat
        Call GdipGetImageBounds(hImage, m_Size, UnitPixel)
        pvProcessAlphaBitmap = True
    End If

End Function

Private Function pvStreamFromArray(ArrayPtr As Long, length As Long) As stdole.IUnknown
    
    ' Purpose: Create an IStream-compatible IUnknown interface containing the
    ' passed byte aray. This IUnknown interface can be passed to GDI+ functions
    ' that expect an IStream interface -- neat hack
    
    On Error GoTo HandleError
    Dim o_hMem As Long
    Dim o_lpMem  As Long
     
    If ArrayPtr = 0& Then
        CreateStreamOnHGlobal 0&, 1&, pvStreamFromArray
    ElseIf length <> 0& Then
        o_hMem = GlobalAlloc(&H2&, length)
        If o_hMem <> 0 Then
            o_lpMem = GlobalLock(o_hMem)
            If o_lpMem <> 0 Then
                CopyMemory ByVal o_lpMem, ByVal ArrayPtr, length
                Call GlobalUnlock(o_hMem)
                Call CreateStreamOnHGlobal(o_hMem, 1&, pvStreamFromArray)
            End If
        End If
    End If
    
HandleError:
End Function

Private Function pvStreamToArray(hStream As Long, arrayBytes() As Byte) As Boolean

    ' Return the array contained in an IUnknown interface (stream)
    
    Dim o_hMem As Long, o_lpMem As Long
    Dim o_lngByteCount As Long
    
    If hStream Then
        If GetHGlobalFromStream(ByVal hStream, o_hMem) = 0 Then
            o_lngByteCount = GlobalSize(o_hMem)
            If o_lngByteCount > 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    ReDim arrayBytes(0 To o_lngByteCount - 1)
                    CopyMemory arrayBytes(0), ByVal o_lpMem, o_lngByteCount
                    GlobalUnlock o_hMem
                    pvStreamToArray = True
                End If
            End If
        End If
    End If
    
End Function

Private Sub Class_Terminate()
    Call pvCleanUp(cuDestroyAll)
    If Not m_Token Is Nothing Then
        m_Token.RemoveUser Me
        Set m_Token = Nothing
    End If
End Sub

Private Sub m_MultiImage_FrameChanged(Index As Long)
    ' this callback event is triggered when user changes the MutiImage index
    ' we call routine that creates clone if needed or parses desired icon from resource
    pvCreateClone Index, 0&
End Sub
