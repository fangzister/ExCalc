VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.UserControl SaveWebImage 
   BackColor       =   &H00FFC0FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   120
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1920
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      ExtentX         =   2778
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "SaveWebImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lparam As Any) As Long

Private Declare Function FindWindowEx _
                Lib "user32" _
                Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                       ByVal hWnd2 As Long, _
                                       ByVal lpsz1 As String, _
                                       ByVal lpsz2 As String) As Long

Private Declare Function DeleteFile _
                Lib "kernel32" _
                Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Const UnitPixel      As Long = 2

Private Const EncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

Private Type GdiplusStartupInput

    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long

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

Private Declare Function GdiplusStartup _
                Lib "gdiplus" (Token As Long, _
                               inputbuf As GdiplusStartupInput, _
                               Optional ByVal outputbuf As Long = 0) As Long

Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)

Private Declare Function GdipSaveImageToFile _
                Lib "gdiplus" (ByVal hImage As Long, _
                               ByVal sFilename As Long, _
                               clsidEncoder As Any, _
                               encoderParams As Any) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long

Private Declare Function GdipCreateBitmapFromHBITMAP _
                Lib "gdiplus" (ByVal hbm As Long, _
                               ByVal hPal As Long, _
                               Bitmap As Long) As Long

Private Declare Function GdipGetImageEncodersSize _
                Lib "gdiplus" (numEncoders As Long, _
                               Size As Long) As Long

Private Declare Function GdipGetImageEncoders _
                Lib "gdiplus" (ByVal numEncoders As Long, _
                               ByVal Size As Long, _
                               Encoders As Any) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long

Private Declare Function CLSIDFromString _
                Lib "ole32" (ByVal lpszProgID As Long, _
                             pCLSID As Any) As Long

Public Enum WebImageFileFormat

    Bmp = 1
    Jpg = 2
    Png = 3
    Gif = 4

End Enum

Private Const WM_PRINT = &H317

Private m_bIsDebug        As Boolean '是否调试模式

Private m_pAccountInfo(3) As tAccountInfo

Private m_CurAccountType  As eAccountType

Private m_CurOpenWebState As eOpenWebState '打开网页状态

Private Enum eOpenWebState

    ewsNoop = 0
    ewsOpenSnapURL
    
    ewsInputAccountInfo
    ewsCheckLgoin
    ewsReOpenSnapURL
    ewsOpenComplete
    ewsError

End Enum

Public Enum eAccountType

    eatSinaWeibo = 0
    eatShouWeibo = 1
    eat163Weibo = 2
    eatQQWeibo = 3

End Enum

Private Type tAccountInfo '账号信息

    AccountType As eAccountType
    AccountValid As Boolean '账号是否有效
    UserName As String
    PassWord As String

End Type

Private m_VerifyCode As String

Public Property Get VerifyCode() As String
    VerifyCode = m_VerifyCode
End Property

Public Property Let VerifyCode(ByVal StrValue As String)
    m_VerifyCode = StrValue
End Property

Public Property Get IsDebug() As Boolean
    IsDebug = m_bIsDebug
End Property

Public Property Let IsDebug(ByVal bValue As Boolean)
    m_bIsDebug = bValue
End Property

Private Sub UserControl_Initialize()

    With Web1
        .Navigate "about:blank"
    End With
    
    pic1.AutoRedraw = True
    pic1.Move -10000, -10000
    
    Timer1.Enabled = False
    Timer1.Interval = 1000
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next

    If m_bIsDebug = True Then
        Web1.Move 0, 0, UserControl.Width, UserControl.Height
        
    Else
        UserControl.Width = 420
        UserControl.Height = 420
    End If

End Sub

Private Sub Web1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    Dim sHTML As String

    If URL = "about:blank" Then Exit Sub
    If Not pDisp Is Web1.object Then Exit Sub

    If m_CurOpenWebState = ewsInputAccountInfo Then Exit Sub

    sHTML = Web1.Document.documentElement.outerHTML

    If InStr(LCase(URL), "http://weibo.com/") > 0 And InStr(sHTML, "<!--注册登录header-->") > 0 Then
        m_CurAccountType = eatSinaWeibo
        
        If m_pAccountInfo(m_CurAccountType).AccountValid = False Then
            m_CurOpenWebState = ewsError
            MsgBox "新浪微博账号登录信息不完整！", vbInformation, "提示："
            
        Else
            m_CurOpenWebState = ewsInputAccountInfo
            Web1.Navigate "http://weibo.com/login.php"
            
            Timer1.Enabled = True
            
        End If

    Else
        m_CurOpenWebState = ewsOpenComplete '页面加载完成
    End If

End Sub

Private Sub Web1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Cancel = True
End Sub

'保存网页图片
Public Function SaveWebImageToPath(ByVal URL As String, _
                                   ByVal Path As String, _
                                   Optional ByVal FileFormat As WebImageFileFormat = Jpg, _
                                   Optional ByVal JpgQuality As Long = 80) As Boolean

    Dim hwnd As Long

    Dim p()  As String
    
    Web1.Silent = True '静默模式 啥脚本错误就不要提示了
    Timer1.Enabled = False
    
OpenSnapURL:
    Web1.Navigate URL
    m_CurOpenWebState = ewsOpenSnapURL
    
    Do

        If m_CurOpenWebState = ewsOpenComplete Then '页面打开完成

            Exit Do
            
        ElseIf m_CurOpenWebState = ewsReOpenSnapURL Then '重新打开示功图页面
            GoTo OpenSnapURL
        
        ElseIf m_CurOpenWebState = ewsError Then '出错

            Exit Function
            
        End If
        
        DoEvents
        Sleep 1
    Loop

    Timer1.Enabled = False
    
    p = GetBrowHwnd(UserControl.hwnd)

    If UBound(p) = -1 Then
        MsgBox "不可能找不到web句柄的啊啊啊！", vbInformation, "提示："

        Exit Function

    End If

    BrowserFullSize
    
    Const WM_SETFOCUS = &H7
    
    SendMessage pic1.hwnd, WM_SETFOCUS, 0, 0
    
    hwnd = Val(p(0))
    SendMessage hwnd, WM_PRINT, pic1.hDC, 0
    
    Set pic1.Picture = pic1.Image
    pic1.Refresh
    
    SaveStdPicToFile pic1.Picture, Path, FileFormat, JpgQuality
    
    'SavePicture pic1.Picture, Path
    SaveWebImageToPath = True
End Function

Private Sub BrowserFullSize()

    Dim mW As Long, mH As Long

    With Web1.Document
        mW = .documentElement.ScrollWidth
        mH = .documentElement.scrollHeight + 20 * Screen.TwipsPerPixelY
    End With

    With Web1
        .Move 0, 0, mW * Screen.TwipsPerPixelX, mH * Screen.TwipsPerPixelY
        'pic1.Width = .Width
        'pic1.Height = .Height
        
        pic1.Move UserControl.Width + 10000, UserControl.Height + 10000, .Width, .Height
    End With

End Sub

'获取 Webbrowser 句柄
Private Function GetBrowHwnd(ByVal hwnd As Long) As String()

    Dim p() As String

    Dim j   As Long, k As Long

    p = Split("")

    Do '获得所有子句柄
        DoEvents
        j = FindWindowEx(hwnd, j, "Shell Embedding", vbNullString)

        If j = 0 Then Exit Do

        k = FindWindowEx(j, 0, "Shell DocObject View", vbNullString)

        If k > 0 Then
            '            k = FindWindowEx(k, 0, "Internet Explorer_Server", vbNullString)
            '            If k > 0 Then
            ReDim Preserve p(UBound(p) + 1)
            p(UBound(p)) = k
            '            End If
        End If

    Loop

    GetBrowHwnd = p
End Function

Private Function SaveStdPicToFile(StdPic As StdPicture, _
                                  ByVal FileName As String, _
                                  Optional ByVal FileFormat As WebImageFileFormat = Jpg, _
                                  Optional ByVal JpgQuality As Long = 80) As Boolean

    Dim CLSID(3) As Long

    Dim Bitmap   As Long

    Dim Token    As Long

    Dim Gsp      As GdiplusStartupInput

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

            Dim aEncParams() As Byte

            Dim uEncParams   As EncoderParameters

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

    Dim Num      As Long

    Dim Size     As Long

    Dim i        As Long

    Dim Info()   As ImageCodecInfo

    Dim Buffer() As Byte

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

    Dim Out    As String

    Dim Length As Long

    Length = lstrlenW(lpsz)

    If Length > 0 Then
        Out = StrConv(String$(Length, vbNullChar), vbUnicode)
        CopyMemory ByVal Out, ByVal lpsz, Length * 2
        PtrToStrW = StrConv(Out, vbFromUnicode)
    End If

End Function
