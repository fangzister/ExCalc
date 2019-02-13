VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'**模 块 名：clsGDI
'**说    明：E_mail:snyga.jik@163.com 版权所有2017 - 2018(C)1
'**创 建 人：格式化 QQ:65464145
'**日    期：2017-07-19 09:40:44
'**描    述：
'**版    本：V1.0.0
'*************************************************************************
Option Explicit

Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal StartAngle As Single, ByVal sweepAngle As Single) As GpStatus
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll " (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As GpUnit, pen As Long) As GpStatus
Private Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Private Declare Function GdipResetClip Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GdipCreateLineBrush Lib "gdiplus" (Point1 As POINTF, Point2 As POINTF, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Private Declare Function CreateSolidBrush Lib "gdi32 " (ByVal crColor As Long) As Long
Private Declare Function FloodFill Lib "gdi32 " (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32 " (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As GpStatus
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As GpStatus
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As GpStatus
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As GpStatus
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As GpStatus
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hDCSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GdipMeasureString Lib "gdiplus" (ByVal Graphics As Long, ByVal str As Long, ByVal Length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, boundingBox As RECTF, codepointsFitted As Long, linesFilled As Long) As GpStatus
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As GpStatus
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As GpStatus
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, brush As Long) As GpStatus
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As GpStatus
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As Long, ByVal unit As GpUnit, createdfont As Long) As GpStatus
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal Graphics As Long, ByVal Mode As TextRenderingHint) As GpStatus
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal Graphics As Long, ByVal str As Long, ByVal Length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal brush As Long) As GpStatus
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMd As SmoothingMode) As GpStatus
Private Declare Function GdipCreatePathGradient Lib "gdiplus" (Points As POINTF, ByVal Count As Long, ByVal WrapMd As WrapMode, polyGradient As Long) As GpStatus
Private Declare Function GdipAddPathStringI Lib "gdiplus" (ByVal Path As Long, ByVal str As Long, ByVal Length As Long, ByVal family As Long, ByVal style As Long, ByVal emSize As Single, layoutRect As RECTL, ByVal StringFormat As Long) As GpStatus
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal brushmode As FillMode, Path As Long) As GpStatus
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal Graphics As Long, ByVal brush As Long, ByVal Path As Long) As GpStatus
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, Graphics As Long) As GpStatus
Private Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hwnd As Long, Graphics As Long) As GpStatus
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Private Declare Function GdipGetDC Lib "gdiplus" (ByVal Graphics As Long, hDC As Long) As GpStatus
Private Declare Function GdipReleaseDC Lib "gdiplus" (ByVal Graphics As Long, ByVal hDC As Long) As GpStatus
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal Stream As Any, Bitmap As Long) As GpStatus
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, Image As Long) As GpStatus
Private Declare Function GdipLoadImageFromFileICM Lib "gdiplus" (ByVal FileName As Long, Image As Long) As GpStatus
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, Image As Long) As GpStatus
Private Declare Function GdipLoadImageFromStreamICM Lib "gdiplus" (ByVal Stream As Any, Image As Long) As GpStatus
Private Declare Function GdipCloneImage Lib "gdiplus" (ByVal Image As Long, cloneImage As Long) As GpStatus
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, Bitmap As Long) As GpStatus
Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal Bitmap As Long, ByVal X As Long, ByVal Y As Long, Color As Long) As GpStatus
Private Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal Bitmap As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As GpStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, Bitmap As Long) As GpStatus
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As Long, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipSetClipRect Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal CombineMd As CombineMode) As GpStatus
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long

Public Enum ImageFileFormat
    Bmp = 1
    Jpg = 2
    Png = 3
    Gif = 4
End Enum

Private Const HALFTONE = 4
Private Const SRCCOPY As Long = &HCC0020
Private Const ULW_OPAQUE = &H4
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_BOTTOM As Long = 1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const DT_SINGLELINE = &H20
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112
Private Const WS_EX_TRANSPARENT = &H20&
Private Const GW_CHILD = 5
Private Const EncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

'////////////////////////////// PNG
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Type PALETTEENTRY
    peRed   As Byte
    peGreen   As Byte
    peBlue   As Byte
    peFlags   As Byte
End Type

Private Type LOGPALETTE
    palVersion   As Integer
    palNumEntries   As Integer
    palPalEntry(255)   As PALETTEENTRY     '   Enough   for   256   colors
End Type

'//////////////////////////////

Private Declare Function GetDeviceCaps Lib "gdi32 " (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32 " (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32 " (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32 " (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32 " (ByVal hDC As Long) As Long

Private Const RC_PALETTE      As Long = &H100
Private Const SIZEPALETTE      As Long = 104
Private Const RASTERCAPS      As Long = 38
Private Const GMEM_MOVEABLE As Long = &H2
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PicBmp
    Size   As Long
    Type   As Long
    hBmp   As Long
    hPal   As Long
    Reserved   As Long
End Type

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
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
    bmiColors As RGBQUAD
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Enum CombineMode
    CombineModeReplace      ' 0
    CombineModeIntersect    ' 1
    CombineModeUnion        ' 2
    CombineModeXor          ' 3
    CombineModeExclude      ' 4
    CombineModeComplement   ' 5 (Exclude From)
End Enum

Private Enum GpStatus
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
End Enum

Private Enum StringAlignment
    StringAlignmentNear = 0
    StringAlignmentCenter = 1
    StringAlignmentFar = 2
End Enum

Private Enum FontStyle
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

'Unit
Private Enum GpUnit
    UnitWorld
    UnitDisplay
    UnitPixel
    UnitPoint
    UnitInch
    UnitDocument
    UnitMillimeter
End Enum

Private Enum TextRenderingHint
    TextRenderingHintSystemDefault = 0            ' Glyph with system default rendering hint
    TextRenderingHintSingleBitPerPixelGridFit     ' Glyph bitmap with hinting
    TextRenderingHintSingleBitPerPixel            ' Glyph bitmap without hinting
    TextRenderingHintAntiAliasGridFit             ' Glyph anti-alias bitmap with hinting
    TextRenderingHintAntiAlias                    ' Glyph anti-alias bitmap without hinting
    TextRenderingHintClearTypeGridFit              ' Glyph CT bitmap with hinting
End Enum

Private Type RECTF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

'Quality
Private Enum QualityMode
    QualityModeInvalid = -1
    QualityModeDefault = 0
    QualityModeLow = 1
    QualityModeHigh = 2
End Enum

Private Enum SmoothingMode
    SmoothingModeInvalid = QualityModeInvalid
    SmoothingModeDefault = QualityModeDefault
    SmoothingModeHighSpeed = QualityModeLow
    SmoothingModeHighQuality = QualityModeHigh
    SmoothingModeNone
    SmoothingModeAntiAlias
End Enum

'=================================
'Fill / Wrap
Private Enum FillMode
    FillModeAlternate
    FillModeWinding
End Enum

Private Enum WrapMode
    WrapModeTile
    WrapModeTileFlipX
    WrapModeTileFlipY
    WrapModeTileFlipXY
    WrapModeClamp
End Enum

'Point Structure
Private Type POINTL
    X As Long
    Y As Long
End Type

Private Type POINTF
    X As Single
    Y As Single
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1
    EncoderParameterValueTypeASCII = 2
    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
    EncoderParameterValueTypeRational = 5
    EncoderParameterValueTypeLongRange = 6
    EncoderParameterValueTypeUndefined = 7
    EncoderParameterValueTypeRationalRange = 8
End Enum

Private Type EncoderParameter
    GUID(0 To 3)        As Long
    NumberOfValues      As Long
    Type                As EncoderParameterValueType
    Value               As Long
End Type

Private Type EncoderParameters
    Count               As Long
    Parameter           As EncoderParameter
End Type

Private Type ImageCodecInfo
    ClassID(0 To 3)     As Long
    FormatID(0 To 3)    As Long
    CodecName           As Long
    DllName             As Long
    FormatDescription   As Long
    FilenameExtension   As Long
    MimeType            As Long
    Flags               As Long
    Version             As Long
    SigCount            As Long
    SigSize             As Long
    SigPattern          As Long
    SigMask             As Long
End Type

'****************** 获取文本尺寸大小 长度/高度 ******************
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

'Private Type Size
'    cx   As Long
'    cy   As Long
'End Type

'************************************
Private Type tImageListLocal '本地图片列表 有的话就不重复画了
    Path As String
    Image As Long
End Type

Private m_Graphics As Long
Private m_GraphicsWidth As Long '画板宽度
Private m_GraphicsHeight As Long '画板高度

Private m_pImageListLocal() As tImageListLocal

Private m_hDC As Long
Private m_DrawBinary() As Byte
Private m_DrawPicture As StdPicture
Private m_Font As StdFont

Public Property Get Font() As StdFont
    Font = m_Font
End Property

Public Property Let Font(ByVal StdValue As StdFont)
    m_Font = StdValue
End Property

'图像
Public Property Get Picture() As StdPicture
    If m_hDC = 0 Then Exit Property
    
    Set Picture = hDCToPicture(m_hDC, 0, 0, m_GraphicsWidth, m_GraphicsHeight)
End Property

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

'画板
Public Property Get Graphics() As Long
    Graphics = m_Graphics
End Property

'画板宽度
Public Property Get Width() As Long
    Width = m_GraphicsWidth
End Property

'画板高度
Public Property Get Height() As Long
    Height = m_GraphicsHeight
End Property

Public Property Let Height(ByVal lonValue As Long)
    m_GraphicsHeight = lonValue
End Property

'创建画板   返回 Graphics
Public Function CreateGraphics(ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long) As Long
    Dim mBITMAP, oldBitmap As Long
    Dim blendFunc32bpp As BLENDFUNCTION

    Dim tempBI As BITMAPINFO
    Dim tempBlend As BLENDFUNCTION
    Dim curWinLong As Long
    Dim winSize As Size
    Dim srcPoint As POINTAPI
    
    If hDC = 0 Then Exit Function
    If Width <= 0 Or Height <= 0 Then Exit Function
    
    DeleteGraphics
    
    m_hDC = hDC 'GetDC(vbNull)   ' hdc
    
    m_GraphicsWidth = Width
    m_GraphicsHeight = Height

    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32
        .biWidth = m_GraphicsWidth * Screen.TwipsPerPixelX
        .biHeight = m_GraphicsHeight * Screen.TwipsPerPixelY
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
    End With

    'm_hDC = CreateCompatibleDC(hdc)
    'm_hDC = hdc

    mBITMAP = CreateDIBSection(m_hDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    oldBitmap = SelectObject(m_hDC, mBITMAP)

    Call GdipCreateFromHDC(m_hDC, m_Graphics)
    
    DeleteObject mBITMAP
    DeleteObject oldBitmap
End Function

'删除画板
Public Function DeleteGraphics() As Boolean
    Dim i As Long
    
    If m_Graphics = 0 Then Exit Function
    
    Call GdipDeleteGraphics(m_Graphics)
    
    m_hDC = 0
    m_Graphics = 0
    
    For i = 1 To UBound(m_pImageListLocal)
        GdipDisposeImage m_pImageListLocal(i).Image
    Next
End Function

'画文本
Public Sub DrawText(ByVal sText As String, _
                    Optional ByVal Rect As String = "0,0,100,100", _
                    Optional ByVal BrushXY As String = "0,0,100,100", _
                    Optional ByVal FontColor As String = "&HFFFFFFFF,&HFF70DBFF", _
                    Optional ByVal FontSize As String = 15, _
                    Optional ByVal FontName As String = "黑体", _
                    Optional ByVal FontBold As Boolean, _
                    Optional ByVal FontItalic As Boolean, _
                    Optional ByVal FontUnderline As Boolean, _
                    Optional ByVal FontStyleStrikeout As Boolean, _
                    Optional ByRef recTextRect As String)
    Dim str1 As String
    Dim brush As Long
    Dim fontfam As Long, strformat As Long, curFont As Long, rclayout As RECTF, rclayout2 As RECTL, rf2 As RECTF
    Dim strPath As Long
    Dim FontStyleEx As Long

    Dim color1 As Long, color2 As Long
    Dim p() As String, p3() As String, p1 As POINTF, p2 As POINTF
    Dim i As Long, k As Long
    
    recTextRect = ""
    If m_Graphics = 0 Then Exit Sub

    '======================= 判断 RECT
    p = Split(Rect, "|")
    For i = 0 To UBound(p)
        p3 = Split(p(i), "=")
        If UBound(p3) > 0 Then
            If InStr("," & p3(1) & ",", "," & Len(sText) & ",") > 0 Then
                Rect = p3(0)
                Exit For
            End If
        End If
    Next

    '======================= 判断 字体大小
    p = Split(FontSize, "|")
    FontSize = ""
    For i = 0 To UBound(p)
        p3 = Split(p(i), "=")
        If UBound(p3) > 0 Then
            If InStr("," & p3(1) & ",", "," & Len(sText) & ",") > 0 Then
                FontSize = Val(p3(0))
                Exit For
            End If
        End If
    Next

    If UBound(p) = -1 Then
        ReDim p(0)
        p(0) = 25
    End If

    If Len(FontSize) = 0 Then
        FontSize = Val(p(0))
    End If

    '================================ 判断是否渐变
    FontColor = Replace$(UCase(FontColor), "&H", "")
    p = Split(FontColor, ",")
    If UBound(p) = 0 Then '没渐变
        BrushXY = ""

        color1 = Val("&H" & Trim$(p(0)))

    ElseIf UBound(p) = 1 Then '渐变
        color1 = Val("&H" & Trim$(p(0)))
        color2 = Val("&H" & Trim$(p(1)))

    Else
        BrushXY = ""
        color1 = &HFF000000
    End If

    '================================= 判断字体外观
    If FontBold = True Then FontStyleEx = FontStyleEx Or FontStyle.FontStyleBold
    If FontItalic = True Then FontStyleEx = FontStyleEx Or FontStyle.FontStyleItalic
    If FontUnderline = True Then FontStyleEx = FontStyleEx Or FontStyle.FontStyleUnderline
    If FontStyleStrikeout = True Then FontStyleEx = FontStyleEx Or FontStyle.FontStyleStrikeout
    
    GdipSetSmoothingMode m_Graphics, SmoothingModeAntiAlias

    GdipCreateFontFamilyFromName StrPtr(FontName), 0, fontfam
    GdipCreateStringFormat 0, 0, strformat

    '=================================渐变刷子
    If Len(BrushXY) = 0 Then
        GdipCreateSolidFill color1, brush
    Else
        p = Split(BrushXY, ",")

        p1.X = Val(p(0))
        p1.Y = Val(p(1))
        p2.X = Val(p(2))
        p2.Y = Val(p(3))

        GdipCreateLineBrush p1, p2, color1, color2, WrapModeTileFlipXY, brush
    End If

    GdipSetStringFormatAlign strformat, StringAlignmentNear
    GdipCreateFont fontfam, FontSize, FontStyleEx, UnitPixel, curFont

    GdipSetTextRenderingHint m_Graphics, TextRenderingHintClearTypeGridFit
    
    'GdipAddPathStringI strPath, StrPtr(sText), -1, fontfam, FontStyle.FontStyleBold, 55, rclayout2, strformat
    'GdipCreatePath FillModeAlternate, strPath

    str1 = Rect
    rclayout = GetRectCfg(str1)

    GdipDrawString m_Graphics, StrPtr(sText), -1, curFont, rclayout, strformat, brush
    
    GdipMeasureString m_Graphics, StrPtr(sText), -1, curFont, rclayout, strformat, rf2, 0, 0
    
    With rf2
        recTextRect = .Left & "," & .Top & "," & .Right & "," & .Bottom
    End With

    Call GdipDeleteFontFamily(fontfam)
    Call GdipDeleteStringFormat(strformat)
    Call GdipDeleteFont(curFont)
    Call GdipDeleteBrush(brush)
End Sub

'画矩形
Public Sub DrawRect(ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal h As Long, _
                    Optional ByVal BackColor As OLE_COLOR, _
                    Optional ByVal BorderColor As OLE_COLOR = -1)

    Dim pen As Long, brush As Long

    If m_Graphics = 0 Then Exit Sub

    GdipCreateSolidFill BackColor, brush
    GdipFillRectangleI Graphics, brush, X, Y, w, h

    If BorderColor <> -1 Then
        GdipCreatePen1 BorderColor, 1, UnitPixel, pen
        w = w - 1
        h = h - 1
    End If
    GdipDrawRectangleI m_Graphics, pen, X, Y, w, h

    GdipDeletePen pen
    GdipDeleteBrush brush
End Sub

Public Sub DrawRoundRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional RoundSize As Long = 30, Optional ByVal lColor As OLE_COLOR)
    Dim pen As Long
    
    GdipCreatePen1 lColor, 1, UnitPixel, pen  '创建画笔
    
    GdipDrawArcI m_Graphics, pen, X, Y, RoundSize * 2, RoundSize * 2, 180, 90
    GdipDrawLineI m_Graphics, pen, X + RoundSize, Y, X + Width - RoundSize, Y
    GdipDrawArcI m_Graphics, pen, X + Width - RoundSize * 2, Y, RoundSize * 2, RoundSize * 2, 270, 90
    GdipDrawLineI m_Graphics, pen, X + Width, Y + RoundSize, X + Width, Y + Height - RoundSize
    GdipDrawArcI m_Graphics, pen, X + Width - RoundSize * 2, Y + Height - RoundSize * 2, RoundSize * 2, RoundSize * 2, 0, 90
    GdipDrawLineI m_Graphics, pen, X + RoundSize, Y + Height, X + Width - RoundSize, Y + Height
    GdipDrawArcI m_Graphics, pen, X, Y + Height - RoundSize * 2, RoundSize * 2, RoundSize * 2, 90, 90
    GdipDrawLineI m_Graphics, pen, X, Y + RoundSize, X, Y + Height - RoundSize
    
    GdipDeletePen pen
End Sub

'从路径获取图片
Public Function getImageFormPath(ByVal Path As String) As StdPicture
    
End Function

'从二进制获取图片
Public Function getImageFormByte(ByVal Path As String) As StdPicture
    
End Function

'计算图片长宽
'obj  可输入 stdPicture / 路径 / 二进制
'w , h  不指定  则返回原图片大小   指定 x 则返回 y 对应大小 反之一样
Public Function getImageWH(ByVal obj As Variant, Optional ByVal w As Long, Optional ByVal h As Long) As String
    Dim sType As String
    Dim Image As Long, imgW As Long, imgH As Long
    Dim lW As Long, lH As Long, Rate As Single
    Dim sResult As String
    Dim IStream As Object, Bye() As Byte
    
    sType = TypeName(obj)
    Select Case sType
        Case "Picture"
        If obj = 0 Or obj Is Nothing Then Exit Function
        GdipCreateBitmapFromHBITMAP obj.handle, obj.hPal, Image
        
        Case "String"
        If Mid$(obj & Space$(5), 2, 2) = ":\" Then
            GdipLoadImageFromFile StrPtr(obj), Image
        End If
        
        Case "Byte()"
        Bye = obj
        
        CreateStreamOnHGlobal Bye(0), False, IStream
        GdipLoadImageFromStream IStream, Image
        
        'GlobalFree IStream
        Case Else
        
    End Select
    
    If Image = 0 Then Exit Function
    
    '/////////////////// 画图像
    GdipGetImageWidth Image, imgW
    GdipGetImageHeight Image, imgH

    If w <= 0 And h <= 0 Then '按图像默认大小输出
        lW = imgW
        lH = imgH

    ElseIf w > 0 And h > 0 Then '按指定大小输出
        lW = w
        lH = h

    ElseIf w > 0 And h <= 0 Then '按 x 比例缩放
        Rate = imgW / imgH

        lW = w
        lH = lW / Rate

    ElseIf h > 0 And w <= 0 Then '按 y 比例缩放
        Rate = imgH / imgW

        lH = h
        lW = lH / Rate

    Else
        Debug.Assert False

    End If
    GdipDisposeImage Image
    
    sResult = lW & "," & lH
    getImageWH = sResult
End Function

'画图像
Private Function DrawImage(Optional ByVal Path As String, _
                           Optional ByVal X As Long, Optional ByVal Y As Long, _
                           Optional ByVal w As Long, Optional ByVal h As Long) As Long

    Dim Image As Long, imgW As Long, imgH As Long
    Dim k As Long, i As Long
    Dim isOK As Boolean
    Dim lW As Long, lH As Long, Rate As Single
    Dim IStream As Object

    If m_Graphics = 0 Then Exit Function

    If Path = "{byte}" Then '二进制
        On Error Resume Next
        k = UBound(m_DrawBinary)
        If Err Then Exit Function
        On Error GoTo 0

        CreateStreamOnHGlobal m_DrawBinary(0), False, IStream
        GdipLoadImageFromStream IStream, Image
        
    ElseIf Path = "{picture}" Then 'stdPicture
        If m_DrawPicture = 0 Or m_DrawPicture Is Nothing Then Exit Function
        GdipCreateBitmapFromHBITMAP m_DrawPicture.handle, m_DrawPicture.hPal, Image

    ElseIf Len(Path) > 0 And IsFileExist(Path) = True Then '从路径画

        '/////////////////// 遍历图片
        For i = 1 To UBound(m_pImageListLocal)
            If m_pImageListLocal(i).Path = Path Then
                isOK = True
                Image = m_pImageListLocal(i).Image
                Exit For
            End If
        Next

        If isOK = False Then
            GdipLoadImageFromFile StrPtr(Path), Image

            ReDim Preserve m_pImageListLocal(UBound(m_pImageListLocal) + 1)
            With m_pImageListLocal(UBound(m_pImageListLocal))
                .Path = Path
                .Image = Image
            End With
        End If

    Else
        Exit Function
    End If

    '/////////////////// 画图像
    GdipGetImageWidth Image, imgW
    GdipGetImageHeight Image, imgH

    If w <= 0 And h <= 0 Then '按图像默认大小输出
        lW = imgW
        lH = imgH

    ElseIf w > 0 And h > 0 Then '按指定大小输出
        lW = w
        lH = h

    ElseIf w > 0 And h <= 0 Then '按 x 比例缩放
        Rate = imgW / imgH

        lW = w
        lH = lW / Rate

    ElseIf h > 0 And w <= 0 Then '按 y 比例缩放
        Rate = imgH / imgW

        lH = h
        lW = lH / Rate

    Else
        Debug.Assert False

    End If

    GdipDrawImageRect m_Graphics, Image, X, Y, lW, lH
    'GdipDrawImageRectRectI m_Graphics, Image, x, y, imgW, imgH, 0, 0, imgW, imgH, UnitPixel
    
    'If Not (IStream Is Nothing) Then GlobalFree IStream
    GdipDisposeImage Image
    DrawImage = Image
End Function

'从路径画图像
Public Function DrawImageFormPath(ByVal Path As String, _
                                  Optional ByVal X As Long, Optional ByVal Y As Long, _
                                  Optional ByVal w As Long, Optional ByVal h As Long) As Long
                                  
    DrawImageFormPath = DrawImage(Path, X, Y, w, h)
End Function

'从图像画图像
Public Function DrawImageFormPicture(ByVal Picture As StdPicture, _
                                     Optional ByVal X As Long, Optional ByVal Y As Long, _
                                     Optional ByVal w As Long, Optional ByVal h As Long) As Long
    
    
    Set m_DrawPicture = Picture
    DrawImageFormPicture = DrawImage("{picture}", X, Y, w, h)
    Set m_DrawPicture = Nothing
End Function

'从二进制画图像
Public Function DrawImageFormByte(ByRef Binary() As Byte, _
                                  Optional ByVal X As Long, Optional ByVal Y As Long, _
                                  Optional ByVal w As Long, Optional ByVal h As Long) As Long
    m_DrawBinary = Binary
    DrawImageFormByte = DrawImage("{byte}", X, Y, w, h)
    m_DrawBinary = ""
End Function

Private Function GetRectCfg(ByVal sText As String) As RECTF
    Dim p() As String
    Dim Rct As RECTF

    p = Split(sText, ",")
    ReDim Preserve p(3)

    Rct.Left = Val(p(0))
    Rct.Top = Val(p(1))
    Rct.Right = Val(p(2))
    Rct.Bottom = Val(p(3))

    GetRectCfg = Rct
End Function

Private Function GetRectCfg2(ByVal sText As String) As RECTF
    Dim p() As String
    Dim Rct As RECTF

    p = Split(sText, ",")
    ReDim Preserve p(3)

    Rct.Left = Val(p(0))
    Rct.Top = Val(p(1))
    Rct.Right = Val(p(2))
    Rct.Bottom = Val(p(3))

    GetRectCfg2 = Rct
End Function

'创建 位图图片
Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long, Pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With

    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    Set CreateBitmapPicture = IPic
End Function

'hDC 转 Picture
Public Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, R As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE

    hDCMemory = CreateCompatibleDC(hDCSrc)
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        R = RealizePalette(hDCMemory)
    End If

    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    hBmp = SelectObject(hDCMemory, hBmpPrev) '恢复旧的位图

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0) '选择调色板
    End If

    R = DeleteDC(hDCMemory) '删除内存 DC
    Set hDCToPicture = CreateBitmapPicture(hBmp, hPal)
End Function


Public Function SaveStdPicToFile(StdPic As StdPicture, ByVal FileName As String, _
                              Optional ByVal FileFormat As ImageFileFormat = Jpg, _
                              Optional ByVal JpgQuality As Long = 80) As Boolean
                              
    Dim CLSID(3)        As Long
    Dim Bitmap          As Long
    Dim Token           As Long
    Dim Gsp             As GdiplusStartupInput

    Gsp.GdiplusVersion = 1                      'GDI+ 1.0版本
    GdiplusStartup Token, Gsp                   '初始化GDI+
    GdipCreateBitmapFromHBITMAP StdPic.handle, StdPic.hPal, Bitmap
    If Bitmap <> 0 Then                          '说明我们成功的将StdPic对象转换为GDI+的Bitmap对象了
        Select Case FileFormat
        Case ImageFileFormat.Bmp
            If Not GetEncoderClsID("Image/bmp", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), ByVal 0) = 0)
            End If
        Case ImageFileFormat.Jpg                    'JPG格式可以设置保存的质量
            Dim aEncParams()        As Byte
            Dim uEncParams          As EncoderParameters
            If GetEncoderClsID("Image/jpeg", CLSID) <> -1 Then
                uEncParams.Count = 1                                        ' 设置自定义的编码参数，这里为1个参数
                If JpgQuality < 0 Then
                    JpgQuality = 0
                ElseIf JpgQuality > 100 Then
                    JpgQuality = 100
                End If
                ReDim aEncParams(1 To Len(uEncParams))
                With uEncParams.Parameter
                    .NumberOfValues = 1
                    .Type = EncoderParameterValueTypeLong                   ' 设置参数值的数据类型为长整型
                    Call CLSIDFromString(StrPtr(EncoderQuality), .GUID(0))  ' 设置参数唯一标志的GUID，这里为编码品质
                    .Value = VarPtr(JpgQuality)                                ' 设置参数的值：品质等级，最高为100，图像文件大小与品质成正比
                End With
                CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), aEncParams(1)) = 0)
            End If
        Case ImageFileFormat.Png
            If Not GetEncoderClsID("Image/png", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), ByVal 0) = 0)
            End If
        Case ImageFileFormat.Gif
            If Not GetEncoderClsID("Image/gif", CLSID) = -1 Then                '如果原始的图像是24位，则这个函数会调用系统的调色板来将图像转换为8位，转换的效果会不尽人意,但也有可能系统不自动转换，保存失败
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), ByVal 0) = 0)
            End If
        End Select
    End If
    GdipDisposeImage Bitmap      '注意释放资源
    GdiplusShutdown Token       '关闭GDI+。
End Function


Private Function GetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
    Dim Num         As Long
    Dim Size        As Long
    Dim i           As Long
    Dim Info()      As ImageCodecInfo
    Dim Buffer()    As Byte
    GetEncoderClsID = -1
    GdipGetImageEncodersSize Num, Size               '得到解码器数组的大小
    If Size <> 0 Then
       ReDim Info(1 To Num) As ImageCodecInfo       '给数组动态分配内存
       ReDim Buffer(1 To Size) As Byte
       GdipGetImageEncoders Num, Size, Buffer(1)            '得到数组和字符数据
       CopyMemory Info(1), Buffer(1), (Len(Info(1)) * Num)     '复制类头
       For i = 1 To Num             '循环检测所有解码
           If (StrComp(PtrToStrW(Info(i).MimeType), strMimeType, vbTextCompare) = 0) Then         '必须把指针转换成可用的字符
               CopyMemory ClassID(0), Info(i).ClassID(0), 16  '保存类的ID
               GetEncoderClsID = i      '返回成功的索引值
               Exit For
           End If
       Next
    End If
End Function

Private Function PtrToStrW(ByVal lpsz As Long) As String
    Dim Out         As String
    Dim Length      As Long
    Length = lstrlenW(lpsz)
    If Length > 0 Then
        Out = StrConv(String$(Length, vbNullChar), vbUnicode)
        CopyMemory ByVal Out, ByVal lpsz, Length * 2
        PtrToStrW = StrConv(Out, vbFromUnicode)
    End If
End Function



'RGB颜色转换成ARGB
Public Function RGB2ARGB(ByVal Color As Long, Optional ByVal Alpha As Long = &HFF000000) As Long
    RGB2ARGB = Alpha Or ((Color And &HFF0000) \ &H10000) Or (Color And &HFF00&) Or ((Color And &HFF&) * &H10000)
End Function

'判断文件是否存在
Private Function IsFileExist(ByVal szFileName As String) As Boolean
    On Error Resume Next
    Dim ret As Long
    IsFileExist = False
    ret = GetFileAttributes(szFileName)
    If ret = -1 Then Exit Function
    IsFileExist = IIf((ret And FILE_ATTRIBUTE_DIRECTORY), False, True)
End Function

'类初始化
Private Sub Class_Initialize()
    Dim uInput As GdiplusStartupInput
    Dim Token As Long

    uInput.GdiplusVersion = 1
    If GdiplusStartup(Token, uInput) <> 0 Then
        MsgBox "GDI+ 初始化错误。程序即将关闭。", vbCritical, "Init GDI+ Error"
        End
    End If
    
    Set m_Font = New StdFont
    ReDim m_pImageListLocal(0)
End Sub

Private Sub Class_Terminate()
    DeleteGraphics
End Sub



'--------------------------------------------------------------------------------
' Procedure  :       ScaleImageAsFile
' Description:       缩放图片后另存为文件
'
' Parameters :       SourceImage (String)   源图片路径
'                    ZoomRate (Single)      缩放比例
'                    OutputFile (String)    输出文件路径
'                    ImageType (ImageFileFormat = Jpg)  输出图片格式
'--------------------------------------------------------------------------------
Public Function ScaleImageAsFile(ByRef oPicBox As Object, ByVal SourceImage As String, ByVal zoomRate As Single, ByVal OutputFile As String, Optional ByVal ImageType As ImageFileFormat = Jpg, Optional ByVal JpgQuality As Long = 80) As Boolean
    Dim w As Long
    Dim h As Long
    Dim cw As Single
    Dim ch As Single
    Dim oPic As StdPicture
           
    With oPicBox
        .AutoRedraw = True
        .AutoSize = True
        .ScaleMode = ScaleModeConstants.vbPixels
        
        Set .Picture = LoadPicture(SourceImage)
        
        w = .ScaleWidth
        h = .ScaleHeight
        
        cw = w * zoomRate
        ch = h * zoomRate
        
        SetStretchBltMode .hDC, HALFTONE
        StretchBlt .hDC, 0, 0, cw, ch, .hDC, 0, 0, w, h, SRCCOPY
    End With
    
    Set oPic = Me.hDCToPicture(oPicBox.hDC, 0, 0, cw, ch)
    ScaleImageAsFile = Me.SaveStdPicToFile(oPic, OutputFile, ImageType, JpgQuality)
    
    Set oPic = Nothing
End Function


