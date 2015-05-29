VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIpMultiImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Credit for these GDI+ classes go to LaVolpe
'http://www.vbforums.com/showthread.php?t=598771

' DO NOT CREATE THIS CLASS IN YOUR PROJECT
' IT IS ONLY CREATED/USED BY THE cGDIpImage class

Option Explicit

Public Event FrameChanged(Index As Long)
' Index is zero-bound, 1st frame is zero

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long

Private Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal Image As Long, ByRef dimensionID As Any, ByVal frameIndex As Long) As Long
Private Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, ByVal propSize As Long, ByRef buffer As Any) As Long
Private Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, ByRef size As Long) As Long
Private Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal Image As Long, ByRef dimensionID As Any, ByRef count As Long) As Long
Private Declare Function GdipGetImagePaletteSize Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef psize As Long) As Long
Private Declare Function GdipGetImagePalette Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef palette As Any, ByVal psize As Long) As Long

Private Type PropertyItem
    id As Long
    length As Long
    type As Long
    Value As Long
End Type
Private Const PropertyTagTypeLong = 4&
Private Const PropertyTagTypeShort = 3&

Private m_Index As Long             ' current multi-image index
Private m_Count As Long             ' number of frames/pages/icons
Private m_GUID(0 To 3) As Long      ' either TIFF or GIF GUID for pages/frames
Private m_Image As Long             ' reference to source image
Private m_IsGIF As Boolean          ' whether sourdce is GIF or not
Private m_FrameDelays() As Long     ' GIF only: Each frame's delay value.

Public Property Get count() As Long
    count = m_Count
End Property

Public Function GetGifFramePalette(thePalette() As Long, ByRef TransparencyIndex As Long) As Boolean
    
    ' function returns the palette, sized appropriately, and the transparency index
    ' If no transparency index exists then the TransparencyIndex parameter will be -1&
    
    Dim palSize As Long, tPal() As Long, X As Long
    Const PaletteFlagsHasAlpha As Long = 1&
    
    ' GDI+ palette structure looks like:
    ' Long: palette flags enumeration (http://www.com.it-berater.org/gdiplus/noframes/GdiPlus_enumerations.htm)
    ' Long: size of palette in bytes
    ' Array of Longs: one each for each palette entry
    
    If m_Count Then
       If m_IsGIF Then
            GdipGetImagePaletteSize m_Image, palSize
            If palSize Then
                 ReDim tPal(0 To palSize \ 4& + 1&)
                 If GdipGetImagePalette(m_Image, tPal(0), palSize) = 0& Then
                     ReDim thePalette(0 To palSize \ 4& - 1&)
                     CopyMemory thePalette(0), tPal(2), palSize
                     If (tPal(0) And PaletteFlagsHasAlpha) Then   ' palette contains alpha channels
                         For X = 0& To UBound(thePalette)
                            ' positive values should be negative since alpha channel is supplied with color
                            ' so a non-negative value should be the index. But we will validate correctly vs
                            ' just checking to see if the palette value is => -1
                             If (thePalette(X) And &HFFFFFF) = thePalette(X) Then
                                 TransparencyIndex = X
                                 Exit For
                             End If
                         Next
                     Else
                         TransparencyIndex = -1&    ' no transparency index
                     End If
                     GetGifFramePalette = True
                 End If
             End If
        End If
    End If
End Function

Public Function GetGifLoopCount() As Long

    Dim bPropData() As Byte
    Dim lPropSize As Long
    Dim tPropItem As PropertyItem
    Const PropertyTagLoopCount As Long = &H5101
    
    If m_Image Then
        If m_IsGIF Then
            On Error GoTo errorhandler
            If GdipGetPropertyItemSize(m_Image, PropertyTagLoopCount, lPropSize) = 0& Then
                ' get the loop count if it exists
                ReDim bPropData(0 To lPropSize - 1&)
                If GdipGetPropertyItem(m_Image, PropertyTagLoopCount, lPropSize, bPropData(0)) = 0& Then
                    CopyMemory tPropItem, bPropData(0), 16&
                    If tPropItem.type = PropertyTagTypeLong Then
                        CopyMemory GetGifLoopCount, ByVal tPropItem.Value, 4&
                    ElseIf tPropItem.type = PropertyTagTypeShort Then
                        CopyMemory GetGifLoopCount, ByVal tPropItem.Value, 2&
                    End If
                End If
            End If
        End If
    End If
errorhandler:
End Function

Public Property Get GifFrameDelay(ByVal Index As Long) As Long
    ' Index is zero-bound, 1st frame is zero
    If m_Count Then
        If m_IsGIF Then
            If Index >= 0 And Index < m_Count Then
                GifFrameDelay = m_FrameDelays(Index)
                If GifFrameDelay < m_FrameDelays(m_Count) Then GifFrameDelay = m_FrameDelays(m_Count)
            End If
        End If
    End If

End Property
Public Property Let GifFrameDelay(ByVal Index As Long, ByVal DelayTime As Long)
    ' Index is zero-bound, 1st frame is zero
    If m_Count Then
        If m_IsGIF Then
            If Index >= 0 And Index < m_Count Then
                If DelayTime > -1 Then m_FrameDelays(Index) = DelayTime
            End If
        End If
    End If
End Property

' for your use.
' Many animated GIFs are encoded with a zero frame delay. Being zero, you
' should set the frame delay yourself. By default, this value is 10ms within
' this class. Setting it higher or lower is your option. Must be set each time
' after an animated GIF is loaded.
' Any frame with a delay less than the GifMimimumFrameDelay will be set to the minimum
Public Property Get GifMinimumFrameDelay() As Long
    If m_IsGIF Then
        If m_Count Then GifMinimumFrameDelay = m_FrameDelays(m_Count)
    End If
End Property
Public Property Let GifMinimumFrameDelay(newVal As Long)
    If m_IsGIF Then
        If m_Count Then
            If (newVal And &HFFFF&) > 0& Then m_FrameDelays(m_Count) = (newVal And &HFFFF&)
        End If
    End If
End Property

Public Property Get Index() As Long
    Index = m_Index
End Property
Public Property Let Index(newVal As Long)
    ' Index is zero-bound, 1st frame is zero
    If newVal >= 0& And newVal < m_Count Then pvChangeFrame newVal
End Property

Public Sub MoveFirst()
    pvChangeFrame 0&
End Sub

Public Sub MoveLast()
    pvChangeFrame m_Count - 1&
End Sub

Public Sub MoveNext()
    If m_Count > 1& Then
        If m_Index = m_Count - 1& Then
            pvChangeFrame 0&
        Else
            pvChangeFrame m_Index + 1&
        End If
    End If
End Sub

Public Sub MovePrevious()
    If m_Count > 1& Then
        If m_Index = 0& Then
            pvChangeFrame m_Count - 1&
        Else
            pvChangeFrame m_Index - 1&
        End If
    End If
End Sub

Friend Sub frSetImage(ByVal lParam As Long, ByVal ImageType As Long)
    
    Const FrameDimensionTime As String = "{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"
    Const FrameDimensionPage As String = "{7462DC86-6180-4C7E-8E3F-EE7333A7A483}"
    
    Erase m_FrameDelays
    m_Index = 0&
    m_IsGIF = False
    If ImageType = 2& Then ' GIF
        m_Image = lParam
        CLSIDFromString StrPtr(FrameDimensionTime), m_GUID(0)
        Call GdipImageGetFrameCount(lParam, m_GUID(0), m_Count)
        Call pvExtractFrameDelays
        m_IsGIF = True
    ElseIf ImageType = 8& Then ' TIFF
        m_Image = lParam
        CLSIDFromString StrPtr(FrameDimensionPage), m_GUID(0)
        Call GdipImageGetFrameCount(lParam, m_GUID(0), m_Count)
    Else    ' ICO or CUR
        ' lParam is the count
        m_Count = lParam
    End If
    
End Sub

Private Sub pvChangeFrame(newIndex As Long)

    ' Index is zero-bound, 1st frame is zero
    If m_Count Then
        If newIndex <> m_Index Then
            If m_Image Then
                If GdipImageSelectActiveFrame(m_Image, m_GUID(0), newIndex) = 0& Then
                    m_Index = newIndex
                    RaiseEvent FrameChanged(newIndex)
                End If
            Else ' icon/cursor
                m_Index = newIndex
                RaiseEvent FrameChanged(newIndex)
            End If
        End If
    End If
    
End Sub

Private Sub pvExtractFrameDelays()

    Dim bPropData() As Byte
    Dim i As Long
    Dim lPropSize As Long
    Dim tPropItem As PropertyItem
    Const PropertyTagFrameDelay As Long = &H5100&
    
    If m_Image Then
        On Error GoTo errorhandler
        ReDim m_FrameDelays(0 To m_Count)
        If GdipGetPropertyItemSize(m_Image, PropertyTagFrameDelay, lPropSize) = 0& Then
            ' get the frame delays
            ReDim bPropData(0 To lPropSize - 1&)
            If GdipGetPropertyItem(m_Image, PropertyTagFrameDelay, lPropSize, bPropData(0)) = 0& Then
                CopyMemory tPropItem, bPropData(0), 16&
                If tPropItem.type = PropertyTagTypeLong Then
                    If tPropItem.length = m_Count * PropertyTagTypeLong Then
                        CopyMemory m_FrameDelays(0), ByVal tPropItem.Value, tPropItem.length
                        For i = 0& To m_Count - 1&
                            ' gif frame delays are in 1/10 seconds, convert to 1/1000
                            m_FrameDelays(i) = (m_FrameDelays(i) And &HFFFF&) * 10&
                        Next
                    End If
                End If
            End If
        End If
        m_FrameDelays(m_Count) = 10& ' minimal frame delay. See GifMinimumFrameDelay
    End If
errorhandler:
End Sub

