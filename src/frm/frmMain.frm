VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "表达式计算器"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   1050
   ClientWidth     =   11280
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   11280
   StartUpPosition =   2  '屏幕中心
   Begin ExCalc.SaveWebImage ucSaveWebImage 
      Height          =   420
      Left            =   10560
      TabIndex        =   9
      Top             =   5640
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.PictureBox statusbar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EDEDF1&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   11280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6810
      Width           =   11280
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   8220
         TabIndex        =   12
         Top             =   60
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   2
         Left            =   2460
         TabIndex        =   13
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   1
         Left            =   1380
         TabIndex        =   11
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   540
      End
   End
   Begin VB.PictureBox pnlContainer 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   600
      ScaleHeight     =   4215
      ScaleWidth      =   8835
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   8835
      Begin VB.PictureBox pnl2 
         BackColor       =   &H00000000&
         Height          =   2835
         Left            =   3180
         ScaleHeight     =   185
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   2775
         Begin RichTextLib.RichTextBox txtResult 
            Height          =   855
            Left            =   540
            TabIndex        =   7
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   1508
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            TextRTF         =   $"frmMain.frx":12EC2
         End
         Begin MSComctlLib.ListView lvwFile2 
            Height          =   1095
            Left            =   600
            TabIndex        =   5
            Top             =   1260
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1931
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.PictureBox pnl1 
         BackColor       =   &H00000000&
         Height          =   2835
         Left            =   60
         ScaleHeight     =   185
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   60
         Width           =   2775
         Begin RichTextLib.RichTextBox txtExpression 
            Height          =   855
            Left            =   240
            TabIndex        =   6
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   1508
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            OLEDragMode     =   0
            OLEDropMode     =   1
            TextRTF         =   $"frmMain.frx":12F5F
         End
      End
      Begin VB.PictureBox splitter 
         Height          =   135
         Left            =   4080
         ScaleHeight     =   75
         ScaleWidth      =   435
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4020
         Width           =   495
      End
      Begin VB.PictureBox splitShadow 
         Height          =   135
         Left            =   4920
         ScaleHeight     =   75
         ScaleWidth      =   435
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4020
         Width           =   495
      End
   End
   Begin VB.Menu mnuTmp 
      Caption         =   "TempTestMenu"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSelectFolder 
         Caption         =   "选择目录(&O)..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileList 
         Caption         =   "列出(&T)"
         Begin VB.Menu mnuFileListFiles 
            Caption         =   "全部(&A)"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuFileListSameType 
            Caption         =   "同类(&S)"
         End
         Begin VB.Menu mnuFileListSubFiles 
            Caption         =   "子目录文件(&L)"
         End
         Begin VB.Menu mnuFileListSubFolderAndFileCount 
            Caption         =   "下级目录及文件数量(&C)"
         End
         Begin VB.Menu mnuFileListEmptyFolders 
            Caption         =   "空目录(&T)"
         End
         Begin VB.Menu mnuFileListFileExists 
            Caption         =   "测试文件是否存在(&E)"
         End
         Begin VB.Menu mnuFileListSubFolder 
            Caption         =   "下级目录(&F)"
         End
      End
      Begin VB.Menu mnuFileCopyListedFiles 
         Caption         =   "将列表中的文件复制到(&C)..."
      End
      Begin VB.Menu mnuFileListFileProperty 
         Caption         =   "显示属性(&E)"
         Begin VB.Menu mnuFileListFilePropertySimple 
            Caption         =   "仅属性(&1)"
         End
         Begin VB.Menu mnuFileListFilePropertyEvidence 
            Caption         =   "固定电子证据清单格式(&D)"
         End
         Begin VB.Menu mnuFileListFilePropertyEvidenceList 
            Caption         =   "生成固定电子证据清单(&E)"
         End
      End
      Begin VB.Menu mnuFileDeleteListedFiles 
         Caption         =   "删除列表中的文件(&D)..."
      End
      Begin VB.Menu mnuFileGetHeader 
         Caption         =   "读取前10个字节(&H)"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "重命名(&R)..."
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileReload 
         Caption         =   "重新加载最后文件(&L)"
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "撤销(&U)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "还原(&R)"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewExchange 
         Caption         =   "交换输入输出(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "字体(&F)"
         Begin VB.Menu mnuViewFontAdd 
            Caption         =   "放大(&A)"
         End
         Begin VB.Menu mnuViewFontMinus 
            Caption         =   "缩小(&M)"
         End
         Begin VB.Menu mnuViewFontNormal 
            Caption         =   "正常(&N)"
         End
         Begin VB.Menu mnuViewFontSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewFontCustom 
            Caption         =   "自定义(&C)..."
         End
      End
      Begin VB.Menu mnuViewPos 
         Caption         =   "视图(&V)"
         Begin VB.Menu mnuViewPosUpDown 
            Caption         =   "上下(&1)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewPosLeftRight 
            Caption         =   "左右(&2)"
         End
      End
   End
   Begin VB.Menu mnuString 
      Caption         =   "字符串(&S)"
      Begin VB.Menu mnuStringLen 
         Caption         =   "计算字符串长度(&L)"
      End
      Begin VB.Menu mnuStringGetLineCount 
         Caption         =   "计算行数(&N)"
      End
      Begin VB.Menu mnuStringTransform 
         Caption         =   "转置(&T)"
         Begin VB.Menu mnuStringTransColumn2Row 
            Caption         =   "列转行(&1)"
         End
         Begin VB.Menu mnuStringTransRow2Column 
            Caption         =   "行转列(&2)"
         End
      End
      Begin VB.Menu mnuStringSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStringFind 
         Caption         =   "查找替换(&H)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuStringCaps 
         Caption         =   "大小写转换(&C)"
         Begin VB.Menu mnuStringCapsLineUpper 
            Caption         =   "行首字母大写(&1)"
         End
         Begin VB.Menu mnuStringCapsLineLower 
            Caption         =   "行首字母小写(&2)"
         End
         Begin VB.Menu mnuStringCapsSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStringCapsWordUpper 
            Caption         =   "单词首字母大写(&3)"
         End
         Begin VB.Menu mnuStringCapsWordLower 
            Caption         =   "单词首字母小写(&4)"
         End
         Begin VB.Menu mnuStringCapsSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStringCapsAllUpper 
            Caption         =   "全部转换为大写(&5)"
         End
         Begin VB.Menu mnuStringCapsAllLower 
            Caption         =   "全部转换为小写(&6)"
         End
      End
      Begin VB.Menu mnuStringSort 
         Caption         =   "排序(&S)"
         Begin VB.Menu mnuStringSortAsc 
            Caption         =   "字符升序(&A)"
         End
         Begin VB.Menu mnuStringSortDesc 
            Caption         =   "字符降序(&D)"
         End
         Begin VB.Menu mnuStringSortNumberAsc 
            Caption         =   "数字升序(&N)"
         End
         Begin VB.Menu mnuStringSortNumberDesc 
            Caption         =   "数字降序(&M)"
         End
         Begin VB.Menu mnuStringSortFileAsc 
            Caption         =   "文件名升序(&F)"
         End
         Begin VB.Menu mnuStringSortFileDesc 
            Caption         =   "文件名降序(&E)"
         End
         Begin VB.Menu mnuStringSortExtAsc 
            Caption         =   "扩展名升序(&1)"
         End
         Begin VB.Menu mnuStringSortExtDesc 
            Caption         =   "扩展名降序(&2)"
         End
         Begin VB.Menu mnuStringSortSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStringSortRandom 
            Caption         =   "随机(&R)"
         End
      End
      Begin VB.Menu mnuStringDelete 
         Caption         =   "删除(&D)"
         Begin VB.Menu mnuStringDeleteEmptyLines 
            Caption         =   "删除空行(&E)"
            Index           =   0
         End
         Begin VB.Menu mnuStringDeleteEmptyLines 
            Caption         =   "删除空行（含空格与制表符）(&T)"
            Index           =   1
         End
         Begin VB.Menu mnuStringDeleteSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStringDeleteDupe 
            Caption         =   "删除重复行(&D)"
         End
         Begin VB.Menu mnuStringDeleteSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStringDeleteInvisibleChars 
            Caption         =   "删除首尾不可见字符(&H)"
         End
      End
      Begin VB.Menu mnuStringFormat 
         Caption         =   "格式化字符串(&F)..."
      End
      Begin VB.Menu mnuStringGenerate 
         Caption         =   "生成序列(&G)..."
      End
      Begin VB.Menu mnuStringFetchEvidenceReportMsg 
         Caption         =   "提取电子数据取证报告特定信息(&R)"
      End
   End
   Begin VB.Menu mnuEncoding 
      Caption         =   "编码(&C)"
      Begin VB.Menu mnuEncodingAscII 
         Caption         =   "AscII(&A)"
         Begin VB.Menu mnuEncodingAscIIEncode 
            Caption         =   "获取字符Ascii值(&1)"
         End
         Begin VB.Menu mnuEncodingAscIIDecode 
            Caption         =   "将Ascii值转换为字符(&2)"
         End
      End
      Begin VB.Menu mnuEncodingHTMLEntity 
         Caption         =   "HTML字符实体(&T)"
         Begin VB.Menu mnuEncodingHTMLEntityReplace 
            Caption         =   "翻译HTML字符实体(&1)"
         End
         Begin VB.Menu mnuEncodingHTMLEntityEncode 
            Caption         =   "转换为HTML字符实体(&2)"
         End
      End
      Begin VB.Menu mnuEncodingURL 
         Caption         =   "URL编码(&R)"
         Begin VB.Menu mnuEncodingURLEncode 
            Caption         =   "URLEncode(&1)"
         End
         Begin VB.Menu mnuEncodingURLDecode 
            Caption         =   "URLDecode(&2)"
         End
         Begin VB.Menu mnuEncodingURLSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEncodingURLEncodeUTF8 
            Caption         =   "URLEncodeUTF-8(&3)"
         End
         Begin VB.Menu mnuEncodingURLDecodeUTF8 
            Caption         =   "URLDecodeUTF-8(&4)"
         End
      End
      Begin VB.Menu mnuEncodingUnicode 
         Caption         =   "Unicode编码(&U)"
         Begin VB.Menu mnuEncodingUnicodeEncode 
            Caption         =   "UnicodeEncode(&1)"
         End
         Begin VB.Menu mnuEncodingUnicodeDecode 
            Caption         =   "UnicodeDecode(&2)"
         End
      End
      Begin VB.Menu mnuEncodingBase64 
         Caption         =   "Base64编码(&B)"
         Begin VB.Menu mnuEncodingBase64Encode 
            Caption         =   "Base64加密(&1)"
         End
         Begin VB.Menu mnuEncodingBase64Decode 
            Caption         =   "Base64解密(&2)"
         End
      End
      Begin VB.Menu mnuEncodingHex 
         Caption         =   "十六进制(&X)"
         Begin VB.Menu mnuEncodingHexEncodeX 
            Caption         =   "编码0x(&1)"
         End
         Begin VB.Menu mnuEncodingHexDecodeX 
            Caption         =   "解码0x(&2)"
         End
      End
      Begin VB.Menu mnuEncodingHash 
         Caption         =   "哈希(&H)"
         Begin VB.Menu mnuEncodingHashText 
            Caption         =   "文本(&T)"
            Begin VB.Menu mnuEncodingHashTextMD5 
               Caption         =   "MD5(&M)"
            End
            Begin VB.Menu mnuEncodingHashTextSHA1 
               Caption         =   "SHA-1(&1)"
            End
            Begin VB.Menu mnuEncodingHashTextSHA256 
               Caption         =   "SHA-256(&2)"
            End
            Begin VB.Menu mnuEncodingHashTextSHA384 
               Caption         =   "SHA-384(&3)"
            End
            Begin VB.Menu mnuEncodingHashTextSHA512 
               Caption         =   "SHA-512(&5)"
            End
            Begin VB.Menu mnuEncodingHashTextALL 
               Caption         =   "全部(&A)"
            End
         End
         Begin VB.Menu mnuEncodingHashFile 
            Caption         =   "文件(&F)"
            Begin VB.Menu mnuEncodingHashFileMD5 
               Caption         =   "MD5(&M)"
            End
            Begin VB.Menu mnuEncodingHashFileSHA1 
               Caption         =   "SHA-1(&1)"
            End
            Begin VB.Menu mnuEncodingHashFileSHA256 
               Caption         =   "SHA-256(&2)"
            End
            Begin VB.Menu mnuEncodingHashFileSHA384 
               Caption         =   "SHA-384(&3)"
            End
            Begin VB.Menu mnuEncodingHashFileSHA512 
               Caption         =   "SHA-512(&5)"
            End
            Begin VB.Menu mnuEncodingHashFileALL 
               Caption         =   "全部(&A)"
            End
         End
      End
   End
   Begin VB.Menu mnuMath 
      Caption         =   "数学(&M)"
      Begin VB.Menu mnuMathEval 
         Caption         =   "计算表达式(&E)"
      End
      Begin VB.Menu mnuMathUnaryEquation 
         Caption         =   "一元一次方程求解(&U)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuMathSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMathLCM 
         Caption         =   "最小公倍数(&L)"
      End
      Begin VB.Menu mnuMathGCD 
         Caption         =   "最大公约数(&G)"
      End
      Begin VB.Menu mnuMathArrangement 
         Caption         =   "排列(&A)"
      End
      Begin VB.Menu mnuMathCombination 
         Caption         =   "组合(&C)"
      End
      Begin VB.Menu mnuMathPower 
         Caption         =   "乘方(&P)"
      End
      Begin VB.Menu mnuMathFactorial 
         Caption         =   "阶乘(&F)"
      End
      Begin VB.Menu mnuMathReciprocal 
         Caption         =   "倒数(&R)"
      End
      Begin VB.Menu mnuMathSum 
         Caption         =   "求和(&S)"
      End
      Begin VB.Menu mnuMathAvg 
         Caption         =   "平均数(&V)"
      End
      Begin VB.Menu mnuMathSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMathRedix 
         Caption         =   "进制转换(&R)..."
      End
      Begin VB.Menu mnuMathMesurement 
         Caption         =   "单位换算(&M)..."
      End
   End
   Begin VB.Menu mnuDate 
      Caption         =   "日期(&D)"
      Begin VB.Menu mnuDateInsert 
         Caption         =   "插入(&I)"
         Begin VB.Menu mnuDateInsertToday 
            Caption         =   "当前日期(&D)"
         End
         Begin VB.Menu mnuDateInsertNow 
            Caption         =   "当前时间(&T)"
         End
      End
      Begin VB.Menu mnuDateFormat 
         Caption         =   "格式化(&F)..."
      End
      Begin VB.Menu mnuDateAdd 
         Caption         =   "增加时间(&A)"
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "图片(&I)"
      Begin VB.Menu mnuImageBatch 
         Caption         =   "批处理(&B)..."
      End
      Begin VB.Menu mnuImageEvidenceProperty 
         Caption         =   "电子证据清单格式属性(&D)"
      End
   End
   Begin VB.Menu mnuURL 
      Caption         =   "URL(&U)"
      Begin VB.Menu mnuDateStampToDate 
         Caption         =   "时间戳转日期(&S)..."
      End
      Begin VB.Menu mnuURLOpen 
         Caption         =   "打开URL(&O)"
      End
      Begin VB.Menu mnuURLSaveWebImage 
         Caption         =   "保存网页为图片(&S)..."
      End
      Begin VB.Menu mnuURLSaveWebMHT 
         Caption         =   "保存网页为mht文件(&M)..."
      End
      Begin VB.Menu mnuURLQueryLocate 
         Caption         =   "查询归属地(&L)"
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "搜索(&Q)"
      Begin VB.Menu mnuQueryKeysBy 
         Caption         =   "搜索多个关键词"
         Begin VB.Menu mnuQueryBy 
            Caption         =   "mnuQueryBy"
            Index           =   0
         End
      End
      Begin VB.Menu mnuQueryKeyAt 
         Caption         =   "在多个网站搜索"
         Begin VB.Menu mnuQueryAt 
            Caption         =   "mnuQueryAt"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuExcel 
      Caption         =   "Excel(&X)"
      Begin VB.Menu mnuExcelMergeWorkbook 
         Caption         =   "合并工作簿(&M)"
      End
      Begin VB.Menu mnuExcelCalcRow 
         Caption         =   "统计行数(&C)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpOpenProject 
         Caption         =   "打开工程(&V)"
      End
      Begin VB.Menu mnuHelpOpenProjectFolder 
         Caption         =   "打开工程目录(&F)"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpReleaseHistory 
         Caption         =   "更新历史(&R)"
      End
      Begin VB.Menu mnuHelpReleaseHistoryEdit 
         Caption         =   "编辑更新日志(&E)"
      End
      Begin VB.Menu mnuHelpTodo 
         Caption         =   "TODO(&D)"
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpCheckUpdate 
         Caption         =   "检查更新(&C)"
      End
      Begin VB.Menu mnuHelpSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
      Begin VB.Menu mnuHelpTest 
         Caption         =   "测试(&T)"
      End
      Begin VB.Menu mnuHelpShowIni 
         Caption         =   "查看ini文件(&I)"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Begin VB.Menu mnuTrayShow 
         Caption         =   "显示(&O)"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "退出(&X)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'TODO
'CTRL+D 删除当前行
'拖拽文件时，按住CTRL键Append（自动确定是否要加Crlf），不按则替换全部
'时间戳转日期毫秒数应用.分隔
'查找替换时，选定文本时不应选中换行符

Enum AppendModeEnum
    APPEND_NONE = 0&
    APPEND_CRLF_START = 1&
    APPEND_CRLF_END = 2&
    APPEND_SPACE_START = 3&
    APPEND_SPACE_END = 4&
End Enum

Dim m_ProjectPath          As String
Dim m_UpdateURL            As String
Dim m_UpdateFiles          As String
Dim m_b_expression_changed As Boolean
Dim spMain                 As SplitPane
Public g_StringKeyWithIn   As String
Public g_StringKeyStart    As String

Dim m_Undo As String
Dim m_Redo As String

Public Sub BlockAt()
    Dim lft      As String
    Dim rgt      As String
    Dim s      As String
    Dim nIndex As Long
    Dim nLen   As Long
    Dim t      As String
    
    With txtExpression
        s = .Text
        t = .SelText
        nIndex = .SelStart
        nLen = .SelLength
        
        If nLen = 0 Then
            lft = Left$(s, nIndex)
            rgt = Right$(s, Len(s) - nIndex)
            s = lft & "()" & rgt
            nIndex = nIndex + 1
        Else
            lft = Left$(s, nIndex)
            rgt = Right$(s, Len(s) - nIndex - nLen)
            s = lft & "(" & t & ")" & rgt
            nIndex = nIndex + Len(t) + 2
        End If
        
        .Text = s
        
        .SelStart = nIndex
        .SetFocus
    End With

End Sub

Function GetFunction(ByVal FunctionDef As String, ByVal VarDef As String) As Variant
    Dim lft             As String
    Dim rgt             As String
    Dim equ             As Long
    Dim p()             As String
    Dim X               As Double
    Dim a               As String
    Dim u               As Long
    Dim i               As Long
    Dim xPos            As Long
    Dim tmp             As String
    Dim Y               As Double
    Dim lastFunctionDef As String

    '2x = 10
    '2/4 = 10/x
    '
    p = Split(FunctionDef, "=")
    lft = p(0)
    rgt = p(1)
    
    u = Len(lft)
    
    '只有一边有未知数的情况
    '让未知数在左边
    xPos = InStr(lft, VarDef)
    
    '左边没有未知数
    If xPos = 0 Then
        tmp = lft
        lft = rgt
        rgt = tmp
    End If
    
    '暂时不考虑括号
    '看左边有没有+-号
    
    '左边只有未知数
    If lft = VarDef Then
        GetFunction = EvalExpression(rgt)
        Exit Function
    End If
    
    '左边无加减号
    If InStr(lft, "+") = 0 And InStr(lft, "-") = 0 Then

        '看左边有没有除号
        If InStr(lft, "/") > 0 Then
            '左边有除号 结果=(右边)*被除数
            '去掉除号
            lft = Replace$(lft, "/", "")
            
            '            Math.Exp
            Y = Val(Replace$(lft, VarDef, ""))
            lastFunctionDef = Y & "/(" & rgt & ")"
        Else
            '左边只有数字和未知数
            '结果 = 把右边用括号括起来，除以左边的数字
            
            '去掉乘号
            lft = Replace$(lft, "*", "")
            
            '获取左边的数字
            Y = Val(Replace$(lft, VarDef, ""))
            lastFunctionDef = "(" & rgt & ")/" & Y
        End If
        
        GetFunction = EvalExpression(lastFunctionDef)
    Else
        SetResult "暂不支持"
    End If
End Function

Public Function ReplaceOperator(Source As String) As String
    Dim s As String
    Dim i As Long
    Dim l As Long
   
    Const CNOP As String = "×＝÷＋"
    Const ENOP As String = "*=/+"

    l = Len(CNOP)
    
    s = Source
    
    For i = 1 To l
        s = Replace(s, Mid$(CNOP, i, 1), Mid$(ENOP, i, 1))
    Next

    ReplaceOperator = s
End Function

Public Sub SetResult(ByVal Result As String, _
                     Optional ByVal SelectResult As Boolean = False, _
                     Optional ByVal AppendResult As Boolean = False, _
                     Optional ByVal AppendMode = AppendModeEnum.APPEND_NONE)
    Dim bSetFont As Boolean
    
    txtResult.Visible = True
    lvwFile2.Visible = False
    
    bSetFont = (txtResult.Text = "")
    
    If AppendResult Then
        Select Case AppendMode
        Case AppendModeEnum.APPEND_NONE
            txtResult.Text = txtResult.Text & Result
        Case AppendModeEnum.APPEND_CRLF_START
            If Len(txtResult.Text) > 0 Then
                txtResult.Text = txtResult.Text & vbCrLf & Result
            Else
                txtResult.Text = Result
            End If
        Case AppendModeEnum.APPEND_CRLF_END
            txtResult.Text = txtResult.Text & Result & vbCrLf
        Case AppendModeEnum.APPEND_SPACE_START
            txtResult.Text = txtResult.Text & " " & Result
        Case AppendModeEnum.APPEND_SPACE_END
            txtResult.Text = txtResult.Text & Result & " "
        End Select
    Else
        txtResult.Text = Result
    End If
    
    If bSetFont Then
        txtResult.Font.Name = txtExpression.Font.Name
        txtResult.Font.Size = txtExpression.Font.Size
    End If
    
    If SelectResult Then
        txtResult.SelStart = 0
        txtResult.SelLength = Len(txtResult.Text)
        txtResult.SetFocus
    End If
End Sub

Private Function EvalExpression(ByVal ExpressionDef As String) As Double
    Dim script As ScriptControl
    
    Set script = New ScriptControl
    script.language = "VBScript"
        
    EvalExpression = script.Eval(ExpressionDef)
    
    Set script = Nothing
End Function

Private Function GetCommand() As Boolean
    Dim p() As String
    Dim ret() As String
    Dim u   As Long
    Dim i   As Long
   
    GetCommandLine u, p

    If u > 0 Then
        ReDim ret(u - 1) As String

        For i = 0 To u - 1
            ret(i) = p(i + 1)
        Next

        txtExpression.Text = Join(ret, vbCrLf)
        GetCommand = True
    End If

End Function

Private Sub GoResult(ResultText As String)
    On Error Resume Next

    With txtResult
        .Text = ResultText
        .SelStart = 0
        .SelLength = Len(ResultText)
        .SetFocus
    End With
End Sub

Private Sub InsertAt(ByVal NewChar As String)
    Dim sLft   As String
    Dim sRgt   As String
    Dim s      As String
    Dim nIndex As Long
    Dim nLen   As Long
    
    With txtExpression
        s = .Text
    
        nIndex = .SelStart
        nLen = .SelLength
        
        If nLen = 0 Then
            sLft = Left$(s, nIndex)
            sRgt = Right$(s, Len(s) - nIndex)
        Else
            sLft = Left$(s, nIndex)
            sRgt = Right$(s, Len(s) - nIndex - nLen)
        End If
        
        s = sLft & NewChar & sRgt
        .Text = s
        
        .SelStart = nIndex + Len(NewChar)
        .SetFocus
    End With
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       LettersOf
' Description:       返回包含Source中所有不重复字母的数组
' Created by :       fangzi
' Date-Time  :       6/14/2017-13:30:28
'
' Parameters :       Source (String)    要查找字母的字符串
'                    ReturnArray() (String) '返回结果
' Returns    :       返回结果数组的长度
'--------------------------------------------------------------------------------
Private Function LettersOf(ByVal Source As String, ByRef ReturnArray() As String) As Long
    Dim i  As Long
    Dim u  As Long
    Dim c  As String
    Dim al As ArrayList
   
    Set al = New ArrayList
    
    u = Len(Source)
    
    For i = 1 To u
        c = Mid$(Source, i, 1)

        If IsLetter(c) Then
            If Not al.Contains(c) Then
                al.Append c
            End If
        End If

    Next
    
    ReturnArray = al.ToStringArray
    LettersOf = UBound(ReturnArray) + 1
    Set al = Nothing
End Function

Private Function LoadLastFiles() As String
    Dim f   As String
    Dim fso As FileSystemObject
    Dim s   As String
    
    f = App.Path & "\lastFiles.txt"
    Set fso = New FileSystemObject
    
    If fso.FileExists(f) Then
        s = modStrings.LoadText(f)
        
        LoadLastFiles = s
    End If

    Set fso = Nothing
End Function

Private Sub LoadQueryMenu(ini As INIProfile)
    Dim keys() As String
    Dim i As Long
    Dim s As String
    Dim p() As String
    
    keys = ini.GetAllKeys("QueryBy")
    
    
    For i = 0 To UBound(keys)
        s = ini.GetString("QueryBy", keys(i))
        p = Split(s, "|")
        
        If i > 0 Then
            Load mnuQueryBy(i)
        End If
        
        mnuQueryBy(i).Caption = p(0)
        mnuQueryBy(i).Tag = p(1)
    Next
End Sub

Private Sub LoadProfile()
    Dim ini As INIProfile
    Dim fs  As Long
    Dim fn  As String

    Set ini = New INIProfile

    With ini
        .ExeFolderPath = App.Path
        .Name = App.title
        
        m_ProjectPath = .GetString("App", "Project")
        m_UpdateURL = .GetString("Update", "URL")
        m_UpdateFiles = .GetString("Update", "Files")
        
        g_StringKeyWithIn = .GetString("String", "KeyWithin")
        g_StringKeyStart = .GetString("String", "KeyStart")
                
        Set spMain = New SplitPane
        
        If .GetString("FormMain", "Layout", 0) = 0 Then
            spMain.VerticalLayout = True
            mnuViewPosUpDown.Checked = True
            mnuViewPosLeftRight.Checked = False
        Else
            spMain.VerticalLayout = False
            mnuViewPosLeftRight.Checked = True
            mnuViewPosUpDown.Checked = False
        End If
        
        spMain.Init pnlContainer, pnl1, pnl2, splitter, splitShadow, 0.5
        
        Me.WindowState = .GetLong("FormMain", "WindowState", 0)
        fs = .GetLong("FormMain", "FontSize")

        If fs > 0 Then
            txtExpression.Font.Size = fs
            txtResult.Font.Size = fs
        End If
        
        fn = .GetString("FormMain", "FontName")

        If Len(fn) > 0 Then
            txtExpression.Font.Name = fn
            txtResult.Font.Name = fn
            
            txtExpression.Font.Charset = 134
            txtResult.Font.Charset = 134
        End If

        LoadQueryMenu ini
        
    End With
    
    Set ini = Nothing
End Sub

Private Sub SaveProfile()
    Dim ini As INIProfile
    
    Set ini = New INIProfile
    
    With ini
        .ExeFolderPath = App.Path
        .Name = App.title
        
        If Me.WindowState <> vbMinimized Then
            .SetLong "FormMain", "WindowState", Me.WindowState
        End If
        
        .SetLong "FormMain", "FontSize", txtExpression.Font.Size
        .SetString "FormMain", "FontName", txtExpression.Font.Name
        
        .SetString "FormMain", "Layout", IIf(spMain.VerticalLayout, "0", "1")
        
        
    End With
    
    Set ini = Nothing
End Sub

Private Sub SetStatus(Index As Integer, ByVal Text As String)
    lblStatus(Index).Caption = Text
End Sub


Private Sub Form_Load()
    Dim s As String
    
    Me.Caption = App.title & " - V" & GetAppVersion
    
    modListView.AddLvwHeads lvwFile2, "序号=60=NO|路径=0=Path|名称=200=Name|文件数=60=FileCount|创建时间=128=DateCreate|修改时间=128=DateModify"
    
    'If Len(s) > 0 Then
    '运行模式
    If App.LogMode = LogModeConstants.vbLogOff Then
        mnuTmp.Visible = False
        mnuHelpTest.Visible = False
        mnuHelpReleaseHistoryEdit.Visible = False
        mnuHelpShowIni.Visible = False
        If Not GetCommand Then
            mnuFileReload_Click
        End If
    Else    '调试模式
        mnuFileReload_Click
    End If
    
    txtExpression.Refresh
    txtResult.Refresh
    
    txtExpression_Change
    txtResult_Change
    
    m_Undo = txtExpression.Text
    
    LoadProfile
    
    lvwFile2.Visible = False
    mnuTray.Visible = False
    
    modTrayIcon.AddTrayIcon Me, Me.Caption, , mnuTray
    SetTrayIcon SetIconFromFile(Me.hwnd)
    
    SetStatus 2, "就绪"
    '    GlobalMouseHook Me.hWnd
    m_b_expression_changed = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ini As INIProfile
    Dim q As Long
    
    Set ini = New INIProfile
    q = ini.GetLong("App", "SaveOnExit")
    If q = 1 Then
        If txtExpression.Text <> "" And m_b_expression_changed Then
            Select Case MsgBox("是否保存内容", vbYesNoCancel Or vbQuestion, "退出前保存")
            Case vbYes
                modStrings.SaveAs txtExpression.Text, App.Path & "\lastfiles.txt"
            Case vbCancel
                Cancel = 1
            End Select
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim w As Single
    Dim h As Single
    
    On Error Resume Next

    w = Me.ScaleWidth
    h = Me.ScaleHeight - statusbar.Height
    
    If Me.WindowState = vbNormal Then
        pnlContainer.Move 0, 0, w, h
    ElseIf Me.WindowState = vbMaximized Then
        pnlContainer.Move 60, 60, w - 120, h - 120
    ElseIf Me.WindowState = vbMinimized Then
        ShowWindow Me.hwnd, SW_HIDE
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveProfile
    Set spMain = Nothing
End Sub

Private Sub lvwFile2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'modListView.SortByHead lvwFile2, ColumnHeader, 1
    'Exit Sub
    Dim tp As Long
    Dim md As Long
    
    Select Case ColumnHeader.Key

    Case "NO", "FileCount"
        tp = LVItemTypes.lvNumber
    Case Else
        tp = LVItemTypes.lvAlphabetic
    End Select
    
    If lvwFile2.SortOrder = lvwAscending Then
        md = lvDescending
    Else
        md = lvAscending
    End If
    
    modListView.LVSortI lvwFile2, ColumnHeader.Index - 2, tp, md
End Sub

Private Sub lvwFile2_DblClick()
    Dim itm As ListItem
    
    Set itm = lvwFile2.SelectedItem
    
    Shell itm.ListSubItems("Path"), vbMaximizedFocus
End Sub

Private Sub mnuDateAdd_Click()
    Dim s        As String
    Dim p()      As String
    Dim i        As Long
    Dim bInvalid As Boolean
    
    s = Trim$(txtExpression.Text)

    If Len(s) = 0 Then
        MsgBox "请输入日期", vbExclamation
        Exit Sub
    End If
    
    bInvalid = True
    p = Split(s, vbCrLf)

    For i = 0 To UBound(p)
        If IsDate(Trim$(p(i))) Then
            bInvalid = False
            Exit For
        End If
    Next
    
    If bInvalid Then
        MsgBox "请输入至少一个合法日期", vbExclamation
        Exit Sub
    End If
    
    dlgDateAdd.Show vbModeless, Me
End Sub

Private Sub mnuDateFormat_Click()
    dlgDateFormat.Show vbModeless, Me
End Sub

Private Sub mnuDateInsertNow_Click()
    txtExpression.SelText = Format(Now(), "yyyy-MM-dd HH:mm:ss")
End Sub

Private Sub mnuDateInsertToday_Click()
    txtExpression.SelText = Format(Now(), "yyyy-MM-dd")
End Sub

Private Sub mnuDateStampToDate_Click()
    dlgTimeStamp.Show vbModeless, Me
End Sub

Private Sub mnuEditRedo_Click()
    m_Undo = txtExpression.Text
    txtExpression.Text = m_Redo
End Sub

Private Sub mnuEditUndo_Click()
    m_Redo = txtExpression.Text
    txtExpression.Text = m_Undo
End Sub

Private Sub mnuEncodingAscIIDecode_Click()
    Dim s    As String
    Dim i    As Long
    Dim j    As Long
    Dim u    As Long
    Dim k    As Long
    Dim p()  As String
    Dim ns() As String
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    u = UBound(p)

    For i = 0 To u
        ns = Split(p(i), " ")
        k = UBound(ns)

        For j = 0 To k
            ns(j) = Chr(ns(j))
        Next

        p(i) = Join(ns, " ")
    Next
    
    SetResult Join(p, vbCrLf)
End Sub

Private Sub mnuEncodingAscIIEncode_Click()
    Dim s  As String
    Dim u  As Long
    Dim sb As StringBuffer
    Dim i  As Long
    Dim c  As String
    
    s = txtExpression.Text
    u = Len(s)

    For i = 1 To u
        c = Mid$(s, i, 1)
        sb.AddText Asc(c)
    Next
    
    SetResult sb.ToString
    Set sb = Nothing
End Sub

Private Sub mnuEncodingBase64Decode_Click()
    SetResult modBase64.Decode64(txtExpression.Text)
End Sub

Private Sub mnuEncodingBase64Encode_Click()
    SetResult modBase64.Encode64(txtExpression.Text)
End Sub

Private Sub mnuEncodingHashALL_Click()
    Dim s As String
    
    s = txtExpression.Text
    SetResult "MD5:"
    SetResult mMD5.MD5FormString(s, True), False, True, AppendModeEnum.APPEND_CRLF_START
    
    SetResult "SHA-1:", False, True, AppendModeEnum.APPEND_CRLF_START
    SetResult modHash.CreateSHA1HashString(s), False, True, AppendModeEnum.APPEND_CRLF_START
    
    SetResult "SHA-256:", False, True, AppendModeEnum.APPEND_CRLF_START
    SetResult modHash.CreateSHA256HashString(s), False, True, AppendModeEnum.APPEND_CRLF_START
    
    SetResult "SHA-384:", False, True, AppendModeEnum.APPEND_CRLF_START
    SetResult modHash.CreateSHA384HashString(s), False, True, AppendModeEnum.APPEND_CRLF_START
    
    SetResult "SHA-512:", False, True, AppendModeEnum.APPEND_CRLF_START
    SetResult modHash.CreateSHA512HashString(s), False, True, AppendModeEnum.APPEND_CRLF_START
End Sub

Private Sub mnuEncodingHashFileALL_Click()
    MsgBox "todo"
End Sub

Private Sub mnuEncodingHashFileMD5_Click()
    Dim s   As String
    Dim p() As String
    Dim fso As FileSystemObject
    Dim i As Long
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    
    Set fso = New FileSystemObject
    
    For i = 0 To UBound(p)
        s = p(i)
        If fso.FileExists(s) Then
            p(i) = UCase$(mMD5.MD5FormFile(s)) & vbTab & s
        Else
            p(i) = "【文件不存在】" & vbTab & s
        End If
    Next
    
    SetResult Join(p, vbCrLf)

    Set fso = Nothing
End Sub

Private Sub mnuEncodingHashFileSHA1_Click()
    Dim s   As String
    Dim p() As String
    Dim fso As FileSystemObject
    Dim i As Long
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    
    Set fso = New FileSystemObject
    
    For i = 0 To UBound(p)
        s = p(i)
        If fso.FileExists(s) Then
            p(i) = UCase$(modHash.CreateSHA1HashFile(s)) & vbTab & s
        Else
            p(i) = "【文件不存在】" & vbTab & s
        End If
    Next
    
    SetResult Join(p, vbCrLf)

    Set fso = Nothing
End Sub

Private Sub mnuEncodingHashFileSHA256_Click()
    Dim s   As String
    Dim p() As String
    Dim fso As FileSystemObject
    Dim i As Long
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    
    Set fso = New FileSystemObject
    
    For i = 0 To UBound(p)
        s = p(i)
        If fso.FileExists(s) Then
            p(i) = UCase$(modHash.CreateSHA256HashFile(s)) & vbTab & s
        Else
            p(i) = "【文件不存在】" & vbTab & s
        End If
    Next
    
    SetResult Join(p, vbCrLf)

    Set fso = Nothing
End Sub

Private Sub mnuEncodingHashFileSHA384_Click()
    Dim s   As String
    Dim p() As String
    Dim fso As FileSystemObject
    Dim i As Long
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    
    Set fso = New FileSystemObject
    
    For i = 0 To UBound(p)
        s = p(i)
        If fso.FileExists(s) Then
            p(i) = UCase$(modHash.CreateSHA384HashFile(s)) & vbTab & s
        Else
            p(i) = "【文件不存在】" & vbTab & s
        End If
    Next
    
    SetResult Join(p, vbCrLf)

    Set fso = Nothing
End Sub

Private Sub mnuEncodingHashFileSHA512_Click()
    Dim s   As String
    Dim p() As String
    Dim fso As FileSystemObject
    Dim i As Long
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    
    Set fso = New FileSystemObject
    
    For i = 0 To UBound(p)
        s = p(i)
        If fso.FileExists(s) Then
            p(i) = UCase$(modHash.CreateSHA512HashFile(s)) & vbTab & s
        Else
            p(i) = "【文件不存在】" & vbTab & s
        End If
    Next
    
    SetResult Join(p, vbCrLf)

    Set fso = Nothing
End Sub

Private Sub mnuEncodingHashTextMD5_Click()
    Dim s   As String
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    s = txtExpression.Text

    If fso.FileExists(s) Then
        SetResult UCase$(mMD5.MD5FormFile(s))
    Else
        SetResult UCase$(mMD5.MD5FormString(s))
    End If

    Set fso = Nothing
End Sub

Private Sub mnuEncodingHashTextSHA1_Click()
    SetResult modHash.CreateSHA1HashString(txtExpression.Text)
End Sub

Private Sub mnuEncodingHashTextSHA256_Click()
    SetResult modHash.CreateSHA256HashString(txtExpression.Text)
End Sub

Private Sub mnuEncodingHashTextSHA384_Click()
    SetResult modHash.CreateSHA384HashString(txtExpression.Text)
End Sub

Private Sub mnuEncodingHashTextSHA512_Click()
    SetResult modHash.CreateSHA512HashString(txtExpression.Text)
End Sub

Private Sub mnuEncodingHTMLEntityEncode_Click()
    MsgBox "todo"
End Sub

Private Sub mnuEncodingHTMLEntityReplace_Click()
    Dim s As String
    Dim ret As String
    
    s = txtExpression.Text
    ret = modStrings.ReplaceHTMLEntity(s)
    txtResult.Text = ret
End Sub

Private Sub mnuEncodingUnicodeDecode_Click()
    txtResult.Text = modStrings.UnicodeDecode(txtExpression.Text)
End Sub

Private Sub mnuEncodingUnicodeEncode_Click()
    txtResult.Text = modStrings.UnicodeEncode(txtExpression.Text)
End Sub

Private Sub mnuEncodingURLDecode_Click()
    txtResult.Text = modStrings.URLDecode(txtExpression.Text)
End Sub

Private Sub mnuEncodingURLDecodeUTF8_Click()
    txtResult.Text = modStrings.URLDecode(txtExpression.Text, True)
End Sub

Private Sub mnuEncodingURLEncode_Click()
    txtResult.Text = modStrings.URLEncode(txtExpression.Text)
End Sub

Private Sub mnuEncodingURLEncodeUTF8_Click()
    txtResult.Text = modStrings.URLEncode(txtExpression.Text, True)
End Sub

Private Sub mnuExcelCalcRow_Click()
    Dim p()        As String
    Dim s          As String
    Dim i          As Long
    Dim j          As Long
    Dim u          As Long
    Dim c          As Long
    Dim fso        As FileSystemObject
    Dim xls        As Excel.Application
    Dim bOK        As Excel.Workbook
    Dim sht        As Excel.Worksheet
    Dim total      As Long
    Dim sColDesc   As String
    Dim nRowCount  As Long
    Dim sCountDesc As String
    Dim nTitle     As Long
    
    p = Split(txtExpression.Text, vbCrLf)
    u = UBound(p)
        
    Set fso = New FileSystemObject
    
    Me.MousePointer = vbHourglass
    SetStatus 2, "正在计算..."
    pb.Max = u
    
    For i = 0 To u
        pb.Value = i
        
        s = Trim$(p(i))
        sCountDesc = ""

        If fso.FileExists(s) Then
            If xls Is Nothing Then
                Set xls = New Excel.Application
                xls.Visible = False
                xls.DisplayAlerts = False
            End If
            
            c = c + 1   '成功的文件数
            
            Set bOK = xls.Workbooks.Open(s)

            For Each sht In bOK.Worksheets

                nRowCount = sht.Range("A65536").End(xlUp).row
                
                For j = 2 To 255

                    If j > 26 Then
                        sColDesc = Chr$(64 + (j \ 26)) & Chr$(65 + (j Mod 26))
                    Else
                        sColDesc = Chr(64 + j)
                    End If
                    
                    nRowCount = Max(nRowCount, sht.Range(sColDesc & "65536").End(xlUp).row)
                    
                    Debug.Print sColDesc, nRowCount
                    
                    If nRowCount = 1 Then
                        If sht.Range(sColDesc & "1").Value = "" Then
                            nRowCount = 0
                        End If
                    End If

                Next
                
                sCountDesc = sCountDesc & nRowCount & ","
                total = total + nRowCount
            Next
            
            sCountDesc = vbTab & "[" & Left$(sCountDesc, Len(sCountDesc) - 1) & "]"
            bOK.Close
            
            Set sht = Nothing
            Set bOK = Nothing
        End If
        
        p(i) = s & sCountDesc
    Next
    
    nTitle = Val(InputBox("输入标题行数量"))
    
    total = total - nTitle * c
    
    SetResult Join(p, vbCrLf) & vbCrLf & "总行数（不含标题行）=" & total
    
    Me.MousePointer = vbNormal
    SetStatus 2, "就绪"
    
    If Not xls Is Nothing Then
        xls.Quit
        Set xls = Nothing
    End If

    Set fso = Nothing
End Sub

Private Sub mnuExcelMergeWorkbook_Click()
    Dim xls               As Excel.Application
    Dim rootBok           As Excel.Workbook
    Dim bOK               As Excel.Workbook
    Dim rootSht           As Excel.Worksheet
    Dim sht               As Excel.Worksheet
    Dim fso               As FileSystemObject
    Dim i                 As Long
    Dim j                 As Long
    Dim lastRow           As Long
    Dim sParam            As String
    Dim firstRow          As Long
    Dim nRowCount         As Long
    Dim sColDesc          As String
    Dim sPath             As String
    Dim sFolder           As String
    Dim sExcel            As String
    Dim oFile             As File
    Dim sFileNameTemplate As String
    Dim sFilename         As String
    Dim p()               As String
    Dim bDelfile          As Boolean
    
    If Len(Trim$(txtExpression.Text)) = 0 Then
        MsgBox "请先将要合并的文件完整路径添加到表达式区域", vbExclamation
        Exit Sub
    End If
    
    sParam = InputBox("请输入标题行序号", "合并Excel工作簿", "1")

    If Len(sParam) = 0 Or (IsNumeric(sParam) = False) Then
        Exit Sub
    End If
    
    If MsgBox("合并后是否删除原文件？", vbQuestion Or vbYesNo) = vbYes Then
        bDelfile = True
    End If
    
    p = Split(txtExpression.Text, vbCrLf)
    
    firstRow = CInt(sParam)
    firstRow = firstRow + 1
    
    Set fso = New FileSystemObject
    Set xls = New Excel.Application
    
    xls.Visible = False
    xls.DisplayAlerts = False
    
    sPath = Trim$(p(0))
    Set oFile = fso.GetFile(sPath)
    sFileNameTemplate = fso.GetBaseName(sPath)
    sFolder = fso.GetParentFolderName(sPath)
    
    Set rootBok = xls.Workbooks.Open(sPath)
    
    xls.CopyObjectsWithCells = True
    
    For i = 1 To UBound(p)
        sExcel = Trim$(p(i))

        If fso.FileExists(sExcel) Then
            Set bOK = xls.Workbooks.Open(sExcel)
            Set sht = bOK.Worksheets(1)
            
            nRowCount = sht.Range("A65536").End(xlUp).row
            
            For j = 2 To 255

                If j > 26 Then
                    sColDesc = Chr$(64 + (j \ 26)) & Chr$(65 + (j Mod 26))
                Else
                    sColDesc = Chr(64 + j)
                End If
                
                nRowCount = Max(nRowCount, sht.Range(sColDesc & "65536").End(xlUp).row)
            Next

            sht.Rows(firstRow & ":" & nRowCount).Copy
            
            Set rootSht = rootBok.Worksheets(1)
            lastRow = rootSht.Range("A65536").End(xlUp).row
            lastRow = lastRow + 1
            
            rootSht.Activate
            rootSht.Range("A" & lastRow).Select
            'rootSht.Paste
            rootSht.PasteSpecial xlPasteValuesAndNumberFormats
                        
            bOK.Saved = True
            bOK.Close
            
            If bDelfile Then
                fso.DeleteFile sExcel
            End If
        End If

    Next
    
    rootBok.Worksheets(1).Range("A1").Select
    
    sFilename = sFileNameTemplate & "-合并.xls"
    rootBok.Save
    xls.Visible = True
    
    On Error GoTo hRename

    oFile.Name = sFilename
    GoTo hClear
hRename:
    sFilename = InputBox("文件已存在，请修改文件名", "合并完成", sFilename)
    oFile.Name = sFilename
hClear:
    
    SetResult oFile.Path
    
    MsgBox "合并完成", vbInformation
    Set sht = Nothing
    Set bOK = Nothing
    Set rootSht = Nothing
    Set rootBok = Nothing
    Set xls = Nothing
    Set fso = Nothing
    Set oFile = Nothing
End Sub

Private Sub mnuFileCopyListedFiles_Click()
    Dim s       As String
    Dim p()     As String
    Dim i       As Long
    Dim u       As Long
    Dim fso     As FileSystemObject
    Dim sFolder As String
    
    s = Trim$(txtExpression.Text)

    If Len(s) = 0 Then
        MsgBox "没有文件", vbExclamation
        Exit Sub
    End If
    
    p = Split(s, vbCrLf)
    
    Set fso = New FileSystemObject
    
    sFolder = modBrowseDirectory.BrowseDirectory(Me.hwnd, "选择保存目录", modBrowseDirectory.GetSpecialFolder(Desktop), True, True)

    If Len(sFolder) = 0 Then Exit Sub
    sFolder = sFolder & "\"
    
    For i = 0 To UBound(p)
        p(i) = Replace$(p(i), vbCr, "")
        p(i) = Replace$(p(i), vbLf, "")
        
        fso.CopyFile p(i), sFolder, True
    Next

    Set fso = Nothing
End Sub

Private Sub mnuFileDeleteListedFiles_Click()
    Dim s     As String
    Dim p()   As String
    Dim c     As Long
    Dim fso   As FileSystemObject
    Dim f     As String
    Dim i     As Long
    Dim suc   As Long
    Dim fal   As Long
    Dim ret() As String
    
    s = DeleteBlankLines(txtExpression.Text, False)
        
    If Len(s) = 0 Then
        MsgBox "列表中没有文件", vbInformation
        Exit Sub
    End If
    
    p = Split(s, vbCrLf)
    
    c = UBound(p) + 1

    If c > 0 Then
        If MsgBox("确实要删除列表中的" & c & "个文件吗？", vbYesNo Or vbQuestion, "确认删除文件") = vbYes Then
            Set fso = New FileSystemObject
            ReDim ret(0 To c - 1) As String

            For i = 0 To UBound(p)
                f = Replace(p(i), vbLf, "")
                f = Replace(f, vbCr, "")

                On Error Resume Next

                If fso.FileExists(f) Then
                    fso.DeleteFile f
                    suc = suc + 1
                Else
                    ret(fal) = f
                    fal = fal + 1
                End If
            Next
            
            If fal = 0 Then
                MsgBox "成功删除" & suc & "个文件", vbInformation, "完成"
            Else
                If MsgBox("成功删除" & suc & "个文件，" & fal & "个文件未能删除，是否在结果列表显示无法删除的文件？", vbInformation Or vbYesNo, "部分文件未能删除") = vbYes Then
                    ReDim Preserve ret(0 To fal - 1) As String
                    SetResult Join(ret, vbCrLf)
                End If
            End If
            
            Set fso = Nothing
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileGetHeader_Click()
    Dim fso As FileSystemObject
    Dim p() As String
    Dim i As Long
    Dim u As Long
    Dim s As String
    Dim ts As TextStream
    Dim v As String
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    u = UBound(p)
    
    MsgBox "TODO"
    Exit Sub
    If u = -1 Then Exit Sub
    
    Set fso = New FileSystemObject
    
    For i = 0 To u
        If Len(p(i)) > 0 Then
            If fso.FileExists(p(i)) Then
                Open p(i) For Binary Access Read As #1
                Debug.Print v
                p(i) = v & vbTab & p(i)
                ts.Close
                txtResult.Text = p(i)
                Exit Sub
            Else
                p(i) = "【文件不存在】" & vbTab & p(i)
            End If
        End If
    Next
    
    SetResult Join(p, vbCrLf)
    Set fso = Nothing
    Set ts = Nothing
End Sub

Private Sub mnuFileListEmptyFolders_Click()
    MsgBox "TODO"
    
End Sub

Private Sub mnuFileListFileExists_Click()
    Dim fso As FileSystemObject
    Dim p() As String
    Dim i As Long
    Dim u As Long
    Dim s As String
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    u = UBound(p)
    
    If u = -1 Then Exit Sub
    
    Set fso = New FileSystemObject
    For i = 0 To u
        If Len(p(i)) > 0 Then
            If fso.FileExists(p(i)) Then
                p(i) = "【文件存在】" & vbTab & p(i)
            Else
                p(i) = "【文件不存在】" & vbTab & p(i)
            End If
        End If
    Next
    
    SetResult Join(p, vbCrLf)
    Set fso = Nothing
End Sub

Private Sub mnuFileListFilePropertyEvidence_Click()
    Dim p()   As String
    Dim s     As String
    Dim i     As Long
    Dim j     As Long
    Dim u     As Long
    Dim fso   As FileSystemObject
    Dim oFile As File
    Dim sType As String
    Dim sDate As String
    Dim sSize As String
    Dim sMD5  As String
    
    Me.MousePointer = vbHourglass
    
    SetStatus 2, "正在计算..."
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    
    Set fso = New FileSystemObject
    
    u = UBound(p)
    
    If u > 0 Then
        pb.Max = u
    End If
    
    For i = 0 To u
        pb.Value = i
        s = p(i)

        If Len(s) > 0 Then
            If Left$(s, 1) = """" Then
                s = Mid$(s, 2)
            End If
    
            If Right$(s, 1) = """" Then
                s = Left$(s, Len(s) - 1)
            End If
            
            If fso.FileExists(s) Then
                Set oFile = fso.GetFile(s)
                
                Select Case LCase$(fso.GetExtensionName(s))
                Case "jpg"
                    sType = "jpg图片"
                Case "rar"
                    sType = "rar压缩包"
                Case "mp4"
                    sType = "mp4视频"
                Case "mp3"
                    sType = "mp3音频"
                Case "doc"
                    sType = "word文档"
                Case "docx"
                    sType = "word文档"
                Case "xls"
                    sType = "excel表格"
                Case "xlsx"
                    sType = "excel表格"
                Case Else
                    sType = oFile.Type
                End Select
                
                sDate = Format(oFile.DateLastModified, "yyyy-MM-dd HH:mm:ss")
                sSize = Format(oFile.Size, "0,0")
                
                If oFile.Size > 2147483648# Then
                    sMD5 = "不支持2G以上的文件MD5"
                Else
                    sMD5 = mMD5.MD5FormFile(oFile.Path, True, md5, False)
                End If
                
                j = j + 1
                
                p(i) = j & vbTab & oFile.Name & vbTab & sType & vbTab & sDate & vbTab & sSize & vbTab & sMD5 & vbTab & "导出文件"
            Else
                p(i) = "文件不存在"
            End If
        End If

    Next
    
    If j > 0 Then
        s = "序号" & vbTab & "名称" & vbTab & "文件类型" & vbTab & "创建时间" & vbTab & "物理大小(字节)" & vbTab & "MD5值" & vbTab & "原始路径"
        SetResult s & vbCrLf & Join(p, vbCrLf), True
    Else
        SetResult Join(p, vbCrLf)
    End If
    
    Me.MousePointer = vbDefault
    SetStatus 2, "就绪"
    pb.Value = 0
    
    Set fso = Nothing
    Set oFile = Nothing
End Sub

Private Sub mnuFileListFilePropertyEvidenceList_Click()
    MsgBox "todo"
End Sub

Private Sub mnuFileListFilePropertySimple_Click()
    Dim p()   As String
    Dim s     As String
    Dim i     As Long
    Dim u     As Long
    Dim fso   As FileSystemObject
    Dim f     As File
    Dim ret() As String
    Dim j     As Long
    Dim bHash As Boolean
    Dim tm    As String
    
    Me.MousePointer = vbHourglass
    
    SetStatus 2, "正在计算..."
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    
    Set fso = New FileSystemObject
    
    u = UBound(p)
    ReDim ret(0 To u) As String
    
    If MsgBox("是否计算MD5值？", vbYesNo Or vbQuestion) = vbYes Then
        bHash = True
    End If

    If u > 0 Then
        pb.Max = u
    End If
    
    For i = 0 To u
        pb.Value = i
        s = p(i)

        If Left$(s, 1) = """" Then
            s = Mid$(s, 2)
        End If

        If Right$(s, 1) = """" Then
            s = Left$(s, Len(s) - 1)
        End If
        
        If fso.FileExists(s) Then
            Set f = fso.GetFile(s)
            
            tm = Format(f.DateCreated, "yyyy-MM-dd HH:mm:ss")

            ret(j) = (j + 1) & vbTab & f.Name & vbTab & vbTab & tm & vbTab & Format(f.Size, "0,0")

            If bHash Then
                ret(j) = ret(j) & vbTab & mMD5.MD5FormFile(f.Path, True, md5, False)
            End If
            
            j = j + 1
        End If

    Next
    
    If j > 0 Then
        ReDim Preserve ret(0 To j - 1) As String
        SetResult Join(ret, vbCrLf), True
    End If
    
    Me.MousePointer = vbDefault
    SetStatus 2, "就绪"
    
    Set fso = Nothing
    Set f = Nothing
End Sub

Private Sub mnuFileListFiles_Click()
    Dim p()   As String
    Dim s     As String
    Dim i     As Long
    Dim u     As Long
    Dim fso   As FileSystemObject
    Dim fod   As Folder
    Dim f     As File
    Dim ret() As String
    Dim j     As Long
    
    s = DeleteBlankLines(txtExpression.Text, False)
    
    If Len(s) = 0 Then
        MsgBox "请在表达式框填入要列出文件的目录路径", vbExclamation
        Exit Sub
    End If
    
    p = Split(s, vbCrLf)
    u = UBound(p)
    
    Set fso = New FileSystemObject
    
    '如果表达式中只有一个文件
    If UBound(p) = 0 Then
        p(0) = Replace$(p(0), vbCr, "")

        If Left$(p(0), 1) = """" And Right$(p(0), 1) = """" Then
            p(0) = Mid$(p(0), 2, Len(p(0)) - 2)
        End If

        If fso.FileExists(p(0)) Then
            Set fod = fso.GetFolder(fso.GetParentFolderName(p(0)))

            If fod.Files.Count > 0 Then
                ReDim ret(0 To fod.Files.Count - 1)

                For Each f In fod.Files

                    ret(j) = f.Path
                    j = j + 1
                Next

                txtExpression.Text = Join(ret, vbCrLf)
                GoTo done:
            End If
        End If
    End If
    
    For i = 0 To u
        s = Trim$(p(i))

        If fso.FolderExists(s) Then
            Set fod = fso.GetFolder(s)
            
            If fod.Files.Count > 0 Then
                ReDim ret(0 To fod.Files.Count - 1)

                For Each f In fod.Files

                    ret(j) = f.Path
                    j = j + 1
                Next

                p(i) = Join(ret, vbCrLf)
            Else
                p(i) = ""
            End If
        End If
    Next
    
    modSort.ShellSortAsc p
    
    txtExpression.Text = Join(p, vbCrLf)

done:
    Set fso = Nothing
    Set fod = Nothing
    Set f = Nothing
End Sub

Private Sub mnuFileListSameType_Click()
    Dim p()   As String
    Dim s     As String
    Dim i     As Long
    Dim u     As Long
    Dim fso   As FileSystemObject
    Dim fod   As Folder
    Dim f     As File
    Dim ret() As String
    Dim j     As Long
    Dim ext   As String
    
    s = DeleteBlankLines(txtExpression.Text, False)
    
    If Len(s) = 0 Then
        MsgBox "请在表达式框填入要列出文件的目录路径", vbExclamation
        Exit Sub
    End If
    
    p = Split(s, vbCrLf)
    u = UBound(p)
    
    Set fso = New FileSystemObject
    
    '如果表达式中只有一个文件
    If UBound(p) = 0 Then
        p(0) = Replace$(p(0), vbCr, "")

        If fso.FileExists(p(0)) Then
            ext = LCase$(fso.GetExtensionName(p(0)))
            Set fod = fso.GetFolder(fso.GetParentFolderName(p(0)))

            If fod.Files.Count > 0 Then
                ReDim ret(0 To fod.Files.Count - 1)

                For Each f In fod.Files

                    If LCase$(fso.GetExtensionName(f.Name)) = ext Then
                        ret(j) = f.Path
                        j = j + 1
                    End If

                Next

                If j > 0 Then
                    ReDim Preserve ret(0 To j - 1) As String
                    txtExpression.Text = Join(ret, vbCrLf)
                End If

                GoTo done:
            End If
        End If
    End If
    
    For i = 0 To u
        s = Trim$(p(i))

        If fso.FolderExists(s) Then
            Set fod = fso.GetFolder(s)
            
            If fod.Files.Count > 0 Then
                ReDim ret(0 To fod.Files.Count - 1)

                For Each f In fod.Files

                    ret(j) = f.Path
                    j = j + 1
                Next

                p(i) = Join(ret, vbCrLf)
            Else
                p(i) = ""
            End If
        End If
    Next
    
    txtExpression.Text = Join(p, vbCrLf)

done:
    Set fso = Nothing
    Set fod = Nothing
    Set f = Nothing
End Sub

Private Sub ListSubFiles(Folder As Folder, ByRef pRet() As String)
    Dim fd    As Folder
    Dim fs As Files
    Dim f As File
    Dim nFileCount As Long
    Dim nFolderCount As Long
    Dim u As Long
    
    '文件
    nFileCount = Folder.Files.Count
    
    If nFileCount > 0 Then
        u = UBound(pRet)
        If u = -1 Then
            ReDim pRet(0 To nFileCount - 1) As String
        Else
            ReDim Preserve pRet(0 To u + nFileCount) As String
        End If
        
        For Each f In Folder.Files
            u = u + 1
            pRet(u) = f.Path
        Next
    End If
    
    '子目录
    nFolderCount = Folder.SubFolders.Count
    If nFolderCount > 0 Then
        For Each fd In Folder.SubFolders
            ListSubFiles fd, pRet
        Next
    End If
End Sub

Private Sub mnuFileListSubFiles_Click()
    Dim p()   As String
    Dim s     As String
    Dim i     As Long
    Dim j     As Long
    Dim u     As Long
    Dim c     As Long
    Dim fso   As FileSystemObject
    Dim fod   As Folder
    Dim tmp() As String
    Dim sFiles As String
    
    s = DeleteBlankLines(txtExpression.Text, False)
    
    If Len(s) = 0 Then
        MsgBox "请在表达式框填入一个目录路径", vbExclamation
        Exit Sub
    End If
    
    p = Split(s, vbCrLf)
    u = UBound(p)
    
    m_Undo = s
    
    Set fso = New FileSystemObject
    
    For i = 0 To u
        If fso.FolderExists(p(i)) Then
            ReDim tmp(-1 To -1) As String
            Set fod = fso.GetFolder(p(i))
            ListSubFiles fod, tmp
            If UBound(tmp) > -1 Then
                p(i) = Join(tmp, vbCrLf)
            Else
                p(i) = ""
            End If
        End If
    Next
    
    txtExpression.Text = Join(p, vbCrLf)

    Set fso = Nothing
    Set fod = Nothing
End Sub

Private Sub mnuFileListSubFolder_Click()
    Dim p()   As String
    Dim s     As String
    Dim i     As Long
    Dim j     As Long
    Dim u     As Long
    Dim c     As Long
    Dim fso   As FileSystemObject
    Dim fod   As Folder
    Dim fd    As Folder
    Dim tmp() As String
    
    s = DeleteBlankLines(txtExpression.Text, False)
    
    If Len(s) = 0 Then
        MsgBox "请在表达式框填入一个目录路径", vbExclamation
        Exit Sub
    End If
    
    p = Split(s, vbCrLf)
    u = UBound(p)
    
    Set fso = New FileSystemObject
    
    For i = 0 To u
        If fso.FolderExists(p(i)) Then
            Set fod = fso.GetFolder(p(i))
            c = fod.SubFolders.Count
            If c > 0 Then
                j = 0
                ReDim tmp(1 To c) As String
                For Each fd In fod.SubFolders
                    j = j + 1
                    tmp(j) = fd.Path
                Next
                p(i) = Join(tmp, vbCrLf)
            End If
        End If
    Next
    
    txtExpression.Text = Join(p, vbCrLf)

    Set fso = Nothing
    Set fod = Nothing
    Set fd = Nothing
End Sub

Private Sub mnuFileListSubFolderAndFileCount_Click()
    Dim p()     As String
    Dim s       As String
    Dim i       As Long
    Dim u       As Long
    Dim fso     As FileSystemObject
    Dim fod     As Folder
    Dim sfd     As Folder
    Dim j       As Long
    Dim pItem() As String
    Dim sItem   As String
    Dim c       As String
    Dim iColor  As Long
    Dim sfc     As Long
    Dim sf      As Folders
   
    s = Trim$(txtExpression.Text)

    If Len(s) = 0 Then
        MsgBox "请在表达式框填入要列出文件的目录路径", vbExclamation
        Exit Sub
    End If
    
    p = Split(s, vbCrLf)
    u = UBound(p)
    
    Set fso = New FileSystemObject
    
    c = Chr$(0)
    
    '遍历表达式框中的目录
    For i = 0 To u
        s = Trim$(p(i))

        '如果目录存在
        If fso.FolderExists(s) Then
            Set fod = fso.GetFolder(s)
            '如果当前目录有子目录
            Set sf = fod.SubFolders
            sfc = sf.Count

            If sfc > 0 Then
                ReDim ret(0 To sfc - 1) As String
                
                j = 0

                '遍历子目录
                For Each sfd In sf

                    ReDim pItem(0 To 5) As String
                    'ret(j) = (j + 1) & ":" & sfd.Files.Count & "个文件" & vbTab & sfd.Name & vbTab & sfd.DateCreated & vbTab & sfd.DateLastModified
                    j = j + 1
                    
                    pItem(0) = "NO=" & j
                    pItem(1) = "Path=" & sfd.Path
                    pItem(2) = "Name=" & sfd.Name
                    pItem(3) = "FileCount=" & sfd.Files.Count
                    pItem(4) = "DateCreate=" & Format(sfd.DateCreated, "yyyy-MM-dd HH:mm:ss")
                    pItem(5) = "DateModify=" & Format(sfd.DateLastModified, "yyyy-MM-dd HH:mm:ss")
                            
                    If sfd.Files.Count = 0 Then
                        iColor = vbRed
                    Else
                        iColor = vbBlack
                    End If

                    modListView.AddLvwItem lvwFile2, Join(pItem, c), , , iColor
                Next

                'p(i) = Join(ret, vbCrLf)
            Else
                'p(i) = "【没有子目录】"
            End If

        Else
            '输出目录不存在
            'p(i) = "【目录不存在】" & s
        End If
    Next
    
    txtResult.Visible = False
    lvwFile2.Visible = True
    'txtExpression.Text = Join(p, vbCrLf)
    
    Set fso = Nothing
    Set sfd = Nothing
    Set fod = Nothing
    Set sf = Nothing
End Sub

Private Sub mnuFileReload_Click()
    Dim s As String
    
    s = LoadLastFiles

    If Len(s) > 0 Then
        txtExpression.Text = s
    End If
End Sub

Public Sub mnuFileRename_Click()
    Dim sSource       As String
    Dim sDest         As String
    Dim fso           As FileSystemObject
    Dim ps()          As String
    Dim pd()          As String
    Dim ret()         As String
    Dim s             As String
    Dim d             As String
    Dim i             As Long
    Dim u             As Long
    Dim cfOperation   As ConflictEnum
    Dim bDealAsSame   As Boolean
    Dim ConflictCount As Long
    
    sSource = Trim$(txtExpression.Text)

    If Len(sSource) = 0 Then
        MsgBox "请在表达式框填入待重命名的文件列表，在结果框填入对应的新文件名列表", vbExclamation
        Exit Sub
    End If
    
    sDest = Trim$(txtResult.Text)

    If Len(sDest) = 0 Then
        MsgBox "请将新文件名列表填入结果框", vbExclamation
        Exit Sub
    End If
    
    ps = Split(sSource, vbCrLf)
    pd = Split(sDest, vbCrLf)
    
    u = UBound(ps)

    If u <> UBound(pd) Then
        MsgBox "原文件数量与新文件数量不等", vbExclamation
        Exit Sub
    End If
    
    ReDim ret(0 To u) As String
    
    Set fso = New FileSystemObject
    
    '先遍历寻找冲突
    For i = 0 To u
        If fso.FileExists(pd(i)) Then
            ConflictCount = ConflictCount + 1
        End If
    Next
    
    '开始重命名
    For i = 0 To u
        s = ps(i)
        d = pd(i)
    
        '源文件不存在
        If fso.FileExists(s) = False And fso.FolderExists(s) = False Then
            ret(i) = "【原文件不存在】" & vbTab & pd(i)
        Else

            '新文件已存在
            If StrComp(s, d, vbBinaryCompare) = 0 Then
                '第一次冲突，或者之前处理时选定了相同处理
                ConflictCount = ConflictCount - 1
                
                If bDealAsSame = False Then
                    dlgRename.Conflict s, d, ConflictCount, cfOperation, bDealAsSame
                End If
                
                Select Case cfOperation

                Case ConflictEnum.OverwriteSource
                    fso.CopyFile s, d, True
                    fso.DeleteFile s, True
                    ret(i) = "【重名，覆盖原文件】" & vbTab & pd(i)

                Case ConflictEnum.Ignore
                    ret(i) = "【重名，忽略】" & vbTab & pd(i)

                Case ConflictEnum.DelSource
                    fso.DeleteFile s, True
                    ret(i) = "【重名，删除原文件】" & vbTab & pd(i)
                End Select

            Else    '新文件不存在，正常重命名
                On Error Resume Next
                Name s As d
                If Err Then
                    ret(i) = "【" & Err.Description & "】" & vbTab & pd(i)
                    Err.Clear
                    On Error GoTo 0
                Else
                    ret(i) = "【成功】" & vbTab & pd(i)
                End If
            End If
        End If

    Next
    
    SetResult Join(ret, vbCrLf)
    
    Set fso = Nothing
End Sub

Private Sub mnuFileSave_Click()
    Dim p As String
    
    p = modBrowseDirectory.OpenFile(Me.hwnd, "保存", App.Path, "LastFiles", "文本文件|*.txt", True)
    
    modStrings.SaveAs txtExpression.Text, p
End Sub

Private Sub mnuFileSelectFolder_Click()
    txtExpression.Text = modBrowseDirectory.BrowseDirectory(Me.hwnd, "选择目录")
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpCheckUpdate_Click()
    Dim v    As String
    Dim u    As String
    Dim fs() As String
    Dim i    As Long
    Dim sv   As String
    Dim fso  As FileSystemObject
    
    MsgBox "TODO"
    Exit Sub
    
    v = GetAppVersion
    
    u = DownloadFile(m_UpdateURL & "server/currentVersion.txt")
    sv = modStrings.LoadText(u)
    Set fso = New FileSystemObject
    
    fso.DeleteFile u
    
    Set fso = Nothing
    
    If v = sv Then
        MsgBox "当前版本[" & v & "]已是最新", vbInformation
        Exit Sub
    End If
    
    MsgBox "有新版本，点击确定开始更新", vbInformation
    
    Unload Me
    ShellOpen App.Path & "\update.exe"
End Sub

Private Sub mnuHelpOpenProject_Click()
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    
    If fso.FileExists(m_ProjectPath) Then
        ShellOpen m_ProjectPath
        Unload Me
    End If

    Set fso = Nothing
End Sub

Private Sub mnuHelpOpenProjectFolder_Click()
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    
    If fso.FileExists(m_ProjectPath) Then
        ShellOpen fso.GetParentFolderName(m_ProjectPath)
    End If

    Set fso = Nothing
End Sub

Private Sub mnuHelpReleaseHistory_Click()
    frmReleaseHistory.Show vbModal, Me
End Sub

Private Sub mnuHelpReleaseHistoryEdit_Click()
    ShellOpen App.Path & "\releasehistory.txt"
End Sub

Private Sub mnuHelpShowIni_Click()
    Dim ini As INIProfile
    Set ini = New INIProfile
    ini.ExeFolderPath = App.Path
    ini.Name = App.title
    ShellOpen ini.ProfilePath
    Set ini = Nothing
End Sub

Private Sub mnuHelpTest_Click()
    Dim k As Long
    Dim s As String
    
    s = "a"
    k = 1
    TestA k
    TestA s
End Sub

Private Sub TestA(ByVal aaa As Variant)
    Debug.Print TypeName(aaa)
End Sub

Private Sub mnuHelpTodo_Click()
    Dim s As String
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    s = App.Path & "\todolist.txt"
    If Not fso.FileExists(s) Then
        fso.CreateTextFile s
    End If
    ShellOpen s
    Set fso = Nothing
End Sub

Private Sub mnuImageBatch_Click()
    dlgBatchImage.Show vbModeless, Me
End Sub

Private Sub mnuImageEvidenceProperty_Click()
    Dim fso           As FileSystemObject
    Dim fImageList    As String
    Dim fEvidenceFile As String
    Dim cmd           As String
    
    fImageList = App.Path & "\plugins\filelist.txt"
    fEvidenceFile = App.Path & "\plugins\list.txt"
    Set fso = New FileSystemObject

    If fso.FileExists(fEvidenceFile) Then
        fso.DeleteFile fEvidenceFile
    End If
    
    SaveAs txtExpression.Text, fImageList
    cmd = "cmd.exe /c python.exe " & App.Path & "\plugins\ImageEvidence.py " & App.Path & "\plugins\"
    OpenPlugin cmd
    
    SetResult LoadText(fEvidenceFile)
    fso.DeleteFile fEvidenceFile
    fso.DeleteFile fImageList
    Set fso = Nothing
End Sub

Private Sub mnuMathArrangement_Click()
    Dim m As Long
    Dim n As Long
    Dim s As String
    Dim a As Variant
    
    On Error GoTo ErrHandle

    s = txtExpression.Text
    a = Split(s, " ")
    m = a(0)
    n = a(1)
    
    SetResult modMath.Arrangement(m, n), True, True, AppendModeEnum.APPEND_CRLF_END
    
    Exit Sub

ErrHandle:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuMathAvg_Click()
    Dim i As Long
    Dim s As String
    Dim a As Variant
    Dim v As Variant
    Dim c As Long
   
    s = Trim$(txtExpression.Text)
    a = Split(s, " ")
    v = 0
    
    For i = 0 To UBound(a)
        If Len(a(i)) > 0 Then
            If IsNumeric(a(i)) Then
                v = v + a(i)
                c = c + 1
            End If
        End If
    Next
    
    If c = 0 Then
        txtResult.Text = 0
    Else
        txtResult.Text = v / c
    End If
End Sub

Private Sub mnuMathCombination_Click()
    Dim m As Long
    Dim n As Long
    Dim s As String
    Dim a As Variant
   
    On Error GoTo ErrHandle

    s = txtExpression.Text
    a = Split(s, " ")
    m = a(0)
    n = a(1)
    
    SetResult modMath.Combination(m, n), True, True, AppendModeEnum.APPEND_CRLF_END
    
    Exit Sub

ErrHandle:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuMathEval_Click()
    On Error GoTo ErrHandle

    Dim script As ScriptControl
    Dim s      As String
    Dim p()    As String
    Dim i      As Long
    
    s = ReplaceOperator(txtExpression.Text)
    p = Split(s, vbCrLf)
    
    Set script = New ScriptControl
    
    script.language = "VBScript"
    
    For i = 0 To UBound(p)
        s = p(i)
        SetResult script.Eval(s), True, True, AppendModeEnum.APPEND_CRLF_START
    Next
    
    Set script = Nothing

    Exit Sub

ErrHandle:
    MsgBox Err.Description, vbExclamation, "表达式错误"
    Set script = Nothing
End Sub

Private Sub mnuMathFactorial_Click()
    Dim v As Long

    On Error GoTo ErrHandle

    v = CLng(txtExpression.Text)
    
    SetResult modMath.Factorial(v), True, True, AppendModeEnum.APPEND_CRLF_END

    Exit Sub

ErrHandle:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuMathGCD_Click()
    Dim s As String
    Dim a As Variant
    Dim m As Long
    Dim n As Long
    Dim v As Long

    On Error GoTo ErrHandle

    s = txtExpression.Text
    a = Split(s, " ")

    If UBound(a) <> 1 Then
        MsgBox "表达式错误，求最小公倍数的表达式为：x y", vbExclamation

        Exit Sub

    End If

    m = a(0)
    n = a(1)
    
    v = modMath.GCD(m, n)
    
    SetResult v, True, True, AppendModeEnum.APPEND_CRLF_END

    Exit Sub

ErrHandle:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuMathLCM_Click()
    Dim s As String
    Dim a As Variant
    Dim m As Long
    Dim n As Long
    Dim v As Long

    On Error GoTo ErrHandle

    s = txtExpression.Text
    a = Split(s, " ")

    If UBound(a) <> 1 Then
        MsgBox "表达式错误，求最小公倍数的表达式为：x y", vbExclamation
        Exit Sub
    End If

    m = a(0)
    n = a(1)
    
    v = modMath.LCM(m, n)
    
    SetResult v, True, True, AppendModeEnum.APPEND_CRLF_END

    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub mnuMathMesurement_Click()
    dlgMesurement.Show vbModeless, Me
End Sub

Private Sub mnuMathPower_Click()
    Dim p() As String
    Dim a   As Long
    Dim b   As Long
    
    On Error Resume Next

    p = Split(txtExpression.Text, " ")
    a = p(0)
    b = p(1)
    
    txtResult.Text = a ^ b
    
    If Err Then
        MsgBox Err.Description, vbCritical
    End If
End Sub

Private Sub mnuMathReciprocal_Click()
    SetResult 1 / EvalExpression(txtExpression.Text)
End Sub

Private Sub mnuMathRedix_Click()
    dlgRedix.Show vbModeless, Me
End Sub

Private Sub mnuMathSum_Click()
    Dim i As Long
    Dim s As String
    Dim a As Variant
    Dim v As Variant
    
    s = Trim$(txtExpression.Text)
    s = Replace$(s, ",", " ")
    s = Replace$(s, vbCrLf, " ")
    s = Replace$(s, vbTab, " ")
    
    a = Split(s, " ")
    v = 0
    
    For i = 0 To UBound(a)
        If Len(a(i)) > 0 Then
            If IsNumeric(a(i)) Then
                v = v + a(i)
            End If
        End If
    Next

    txtResult.Text = v
End Sub

Private Sub mnuMathUnaryEquation_Click()
    Dim s                As String
    Dim vLft             As Double
    Dim vRgt             As Double
    Dim script           As ScriptControl
    Dim i                As Long
    Dim xMult            As Double
    Dim vars()           As String
    Dim varsCount        As Long
    
    Const ALL_OPERATIONS As String = "+-*/\^%()"
    Const ALL_VARIABLES  As String = "abcdefghijklmnopqrstuvwxyz"
   
    
    s = Trim$(txtExpression.Text)
    's = "2x = 10"
    s = Replace$(s, " ", "")
    
    If modStrings.SubCount(s, "=") <> 1 Then
        SetResult "表达式不是方程"

        Exit Sub

    End If
    
    varsCount = LettersOf(s, vars)

    If varsCount = 0 Then
        SetResult "表达式不是方程"
        Exit Sub
    End If
    
    If varsCount > 1 Then
        SetResult "目前仅支持一元一次方程求解"
        Exit Sub
    End If
    
    SetResult vars(0) & " = " & GetFunction(s, vars(0))
End Sub

Private Sub mnuQueryBy_Click(Index As Integer)
    Dim u As String
    Dim s As String
    Dim p() As String
    Dim i As Long
    
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    If UBound(p) > 10 Then
        If MsgBox("将要同时打开" & (UBound(p) + 1) & "个搜索页面，是否继续？", vbYesNo, "批量搜索") <> vbYes Then
            Exit Sub
        End If
    End If
    
    For i = 0 To UBound(p)
        If Len(p(i)) > 0 Then
            u = Replace(mnuQueryBy(Index).Tag, "{0}", p(i))
            ShellOpen u
        End If
    Next
End Sub

Private Sub mnuStringCapsAllLower_Click()
    GoResult LCase$(txtExpression.Text)
End Sub

Private Sub mnuStringCapsAllUpper_Click()
    GoResult UCase$(txtExpression.Text)
End Sub

Private Sub mnuStringCapsLineLower_Click()
    Dim i   As Long
    Dim p() As String
    Dim s   As String
    
    p = Split(txtExpression.Text, vbCrLf)

    For i = 0 To UBound(p)
        s = p(i)
        p(i) = LCase$(Left$(s, 1)) & Mid$(s, 2)
    Next
    
    GoResult Join(p, vbCrLf)
End Sub

Private Sub mnuStringCapsLineUpper_Click()
    Dim i   As Long
    Dim p() As String
    Dim s   As String
    
    p = Split(txtExpression.Text, vbCrLf)

    For i = 0 To UBound(p)
        s = p(i)
        p(i) = UCase$(Left$(s, 1)) & Mid$(s, 2)
    Next
    
    GoResult Join(p, vbCrLf)
End Sub

Private Sub mnuStringCapsWordLower_Click()
    Dim i      As Long
    Dim s      As String
    Dim a      As Integer
    Dim u      As Long
    Dim sRow   As String
    Dim j      As Long
    Dim c      As String
    Dim bStart As Boolean
    
    s = txtExpression.Text
    u = Len(s)

    For i = 1 To u
        c = Mid$(s, i, 1)
        a = Asc(c)

        If modStrings.IsLetter(c) Or (a > 47 And a < 58) Then
            If bStart = False Then
                bStart = True
                Mid(s, i, 1) = LCase$(c)
            End If
        Else
            bStart = False
        End If
    Next
    
    GoResult s
End Sub

Private Sub mnuStringCapsWordUpper_Click()
    Dim i      As Long
    Dim s      As String
    Dim u      As Long
    Dim a      As Integer
    Dim sRow   As String
    Dim j      As Long
    Dim c      As String
    Dim bStart As Boolean
    
    s = txtExpression.Text
    u = Len(s)

    For i = 1 To u
        c = Mid$(s, i, 1)
        a = Asc(c)

        If modStrings.IsLetter(c) Or (a > 47 And a < 58) Then
            If bStart = False Then
                bStart = True
                Mid(s, i, 1) = UCase$(c)
            End If
        Else
            bStart = False
        End If
    Next

    GoResult s
End Sub

Private Sub mnuStringDeleteDupe_Click()
    Dim s   As String
    Dim p() As String
    Dim ret As Variant
    
    '    s = modStrings.DeleteDuplicateLines(txtExpression.Text)
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    ret = Array_unique(p)

    txtResult.Text = Join(ret, vbCrLf)
End Sub

Private Sub mnuStringDeleteEmptyLines_Click(Index As Integer)
    Dim s As String

    s = modStrings.DeleteBlankLines(txtExpression.Text, (Index = 0))
    txtResult.Text = s
End Sub

Private Sub mnuStringDeleteInvisibleChars_Click()
    Dim s   As String
    Dim p() As String
    Dim i   As Long
    Dim u   As Long
   
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    u = UBound(p)
    
    For i = 0 To u
        p(i) = TrimEx(p(i))
    Next
    
    SetResult Join(p, vbCrLf)
End Sub

Private Sub mnuStringFetchEvidenceReportMsg_Click()
    Dim s As String
    Dim R As String
    Dim i As Long
    Dim j As Long
    Dim p() As String
    Dim pRow() As String
    Dim sItem As String
    Dim sProperty As String
    Dim pRet() As String
    Dim oBase As EvidenceIMBase
    
    s = txtExpression.Text
    R = DeleteBlankLines(s)
    
    p = Split(R, "序号: 1")
    ReDim pRet(0 To UBound(p)) As String
    
    For i = 0 To UBound(p)
        sItem = p(i)
        If Len(sItem) > 0 Then
            Set oBase = New EvidenceIMBase
            oBase.Init sItem
            sItem = oBase.ToString
            If Len(sItem) > 0 Then
                pRet(j) = "（" & j + 1 & "）" & sItem
                j = j + 1
            End If
            Set oBase = Nothing
        End If
    Next
    ReDim Preserve pRet(0 To j - 1) As String
    SetResult Join(pRet, vbCrLf)
End Sub

Private Sub mnuStringFind_Click()
    dlgFindReplace.Show vbModeless, Me
End Sub

Private Sub mnuStringFormat_Click()
    Dim i   As Long
    Dim p() As String
    Dim u   As Long
    Dim f   As String
    
    f = InputBox("输入格式化字符串", "格式化字符串")

    If Len(f) = 0 Then Exit Sub
    
    p = Split(txtExpression.Text, vbCrLf)
    u = UBound(p)

    For i = 0 To u
        p(i) = Format$(p(i), f)
    Next

    SetResult Join(p, vbCrLf)
End Sub

Private Sub mnuStringGenerate_Click()
    dlgStringGenerator.Show vbModal, Me
End Sub

Private Sub mnuStringGetLineCount_Click()
    Dim p() As String

    p = Split(txtExpression.Text, vbCrLf)
    SetResult UBound(p) + 1
End Sub

Private Sub mnuStringLen_Click()
    If txtExpression.SelLength = 0 Then
        txtResult.Text = Len(txtExpression.Text)
    Else
        txtResult.Text = "选定文本长度为：" & txtExpression.SelLength
    End If

End Sub

Private Sub mnuStringSortAsc_Click()
    Dim p() As String
    
    p = Split(txtExpression.Text, vbCrLf)
    
    modSort.ShellSortAsc p
    SetResult Join(p, vbCrLf)
End Sub

Private Sub mnuStringSortDesc_Click()
    Dim p() As String
    
    p = Split(txtExpression.Text, vbCrLf)
    
    modSort.ShellSortDesc p
    SetResult Join(p, vbCrLf)
End Sub

Private Sub mnuStringSortExtAsc_Click()
    Dim p() As String
    Dim e() As String
    Dim i   As Long
    Dim u   As Long
    Dim f   As String
    Dim v   As Long
    
    p = Split(txtExpression.Text, vbCrLf)
    u = UBound(p)
    
    ReDim e(0 To u) As String
    
    For i = 0 To u
        f = p(i)
        v = InStrRev(f, ".")

        If v > 0 Then
            e(i) = Right(f, Len(f) - v) & "." & Left$(f, v)
        Else
            e(i) = f
        End If
    Next
    
    modSort.ShellSortAsc e
    
    For i = 0 To u
        f = e(i)
        v = InStr(f, ".")

        If v > 0 Then
            e(i) = Mid$(f, v + 1) & Left$(f, v - 1)
        End If
    Next

    SetResult Join(e, vbCrLf)
End Sub

Private Sub mnuStringSortExtDesc_Click()
    Dim p() As String
    Dim e() As String
    Dim i   As Long
    Dim u   As Long
    Dim f   As String
    Dim v   As Long
    
    p = Split(txtExpression.Text, vbCrLf)
    u = UBound(p)
    
    ReDim e(0 To u) As String
    
    For i = 0 To u
        f = p(i)
        v = InStrRev(f, ".")

        If v > 0 Then
            e(i) = Right(f, Len(f) - v) & "." & Left$(f, v)
        Else
            e(i) = f
        End If
    Next
    
    modSort.ShellSortDesc e
    
    For i = 0 To u
        f = e(i)
        v = InStr(f, ".")

        If v > 0 Then
            e(i) = Mid$(f, v + 1) & Left$(f, v - 1)
        End If
    Next

    SetResult Join(e, vbCrLf)
End Sub

Private Sub mnuStringSortFileAsc_Click()
    Dim s   As String
    Dim p() As String
    Dim i   As Long
    Dim u   As Long
    
    MsgBox "todo"
    Exit Sub
    s = txtExpression.Text
    p = Split(s, vbCrLf)
    u = UBound(p)
    
    For i = 0 To u
        s = p(i)

        If Len(s) > 0 Then
            
        End If
    Next

End Sub

Private Sub mnuStringSortFileDesc_Click()
    MsgBox "todo"
End Sub

Private Sub mnuStringSortNumberAsc_Click()
    Dim p() As String
    Dim i   As Long
    
    p = Split(txtExpression.Text, vbCrLf)
    modSort.NumSortAZ p, 0, UBound(p)
    SetResult Join(p, vbCrLf)
End Sub

Private Sub mnuStringSortNumberDesc_Click()
    Dim p() As String
    
    p = Split(txtExpression.Text, vbCrLf)
    
    modSort.NumSortZA p, 0, UBound(p)
    SetResult Join(p, vbCrLf)
End Sub

Private Sub mnuStringSortRandom_Click()
    Dim p() As String
    
    p = Split(txtExpression.Text, vbCrLf)
    
    modSort.ShellSortRandom p
    SetResult Join(p, vbCrLf)
End Sub

Private Sub mnuStringTransColumn2Row_Click()
    Dim p() As String
    Dim sep  As String
    
    sep = " "
    '    Select Case cboTransString.ListIndex
    '    Case 0  'tab
    '        r = vbTab
    '    Case 1  '[,]
    '        r = ","
    '    Case 2  '[空格]
    '        r = " "
    '    Case 3  '[自定义]
    '        r = txtTransSeperator.Text
    '    Case Else
    '        Exit Sub
    '    End Select
    
    p = Split(txtExpression.Text, vbCrLf)
    GoResult Join(p, sep)
End Sub

Private Sub mnuStringTransRow2Column_Click()
    Dim p() As String
    Dim sep As String
    
    sep = " "
    '    Select Case cboTransString.ListIndex
    '    Case 0  'tab
    '        r = vbTab
    '    Case 1  '[,]
    '        r = ","
    '    Case 2  '[空格]
    '        r = " "
    '    Case 3  '[自定义]
    '        r = txtTransSeperator.Text
    '    Case Else
    '        Exit Sub
    '    End Select
    
    p = Split(txtExpression.Text, sep)
    GoResult Join(p, vbCrLf)
End Sub

Private Sub DelCurrentLine()
    Dim s          As String
    Dim sLft       As String
    Dim sRgt       As String
    Dim nIndex     As Long
    Dim p()        As String
    Dim nStartLine As Long
    Dim nEndLine   As Long
    Dim sSel       As String
    Dim i          As Long
    
    s = txtExpression.Text
    
    nIndex = txtExpression.SelStart
    
    sLft = Left$(s, nIndex)
    sRgt = Mid$(s, nIndex + txtExpression.SelLength)
    sSel = txtExpression.SelText
    
    nStartLine = modStrings.SubCount(sLft, vbCrLf)
    nEndLine = modStrings.SubCount(sSel, vbCrLf) + nStartLine
    
    nIndex = 0

    For i = 0 To nStartLine
        nIndex = InStr(nIndex + 1, s, vbCrLf)
    Next

    MsgBox nIndex
End Sub

Private Sub mnuTmp_Click()
    mnuFileListSubFiles_Click
End Sub

Private Sub mnuTrayExit_Click()
    Unload Me
End Sub

Private Sub mnuTrayShow_Click()
    modTrayIcon.ActivationWindow Me.hwnd
End Sub

Private Sub mnuURLQueryLocate_Click()
    Dim p() As String
    Dim i   As Long
    Dim u   As String
    Dim s   As String
   
    MsgBox "todo"

    Exit Sub
     
    s = "http://"
    p = Split(txtExpression.Text, vbCrLf)

    For i = 0 To UBound(p)
        u = Trim$(p(i))
        ShellOpen u
    Next

End Sub

Private Sub mnuURLOpen_Click()
    Dim p() As String
    Dim i   As Long
    Dim u   As String
    
    p = Split(txtExpression.Text, vbCrLf)

    For i = 0 To UBound(p)
        u = Trim$(p(i))

        If StartsWith(u, "http://") Or StartsWith(u, "https://") Then
            ShellOpen u
        End If
    Next
End Sub

Private Sub mnuURLSaveWebImage_Click()
    Dim f   As String
    Dim p() As String
    Dim u   As String
    Dim i   As Long
    Dim s   As String
    Dim c   As Long
    
    s = Trim$(txtExpression.Text)
    
    If Len(s) = 0 Then Exit Sub
    
    p = Split(s, vbCrLf)
    
    f = modBrowseDirectory.BrowseDirectory(Me.hwnd, "选择保存目录")

    If Len(f) = 0 Then Exit Sub
    
    For i = 0 To UBound(p)
        u = Trim$(p(i))

        If modStrings.StartsWith(u, "http://") Or modStrings.StartsWith(u, "https://") Then
            s = f & "\" & (i + 1) & ".jpg"

            If ucSaveWebImage.SaveWebImageToPath(u, s) Then
                c = c + 1
            End If
        End If
    Next
    
    MsgBox "成功保存" & c & "个文件", vbInformation
End Sub

Private Sub mnuURLSaveWebMHT_Click()
    Dim f   As String
    Dim p() As String
    Dim u   As String
    Dim i   As Long
    Dim j   As Long
    Dim s   As String
    Dim c   As Long
    
    s = Trim$(txtExpression.Text)
    
    If Len(s) = 0 Then Exit Sub
    
    p = Split(s, vbCrLf)
    
    f = modBrowseDirectory.BrowseDirectory(Me.hwnd, "选择保存目录")

    If Len(f) = 0 Then Exit Sub
    
    For i = 0 To UBound(p)
        u = Trim$(p(i))

        If modStrings.StartsWith(u, "http://") Or modStrings.StartsWith(u, "https://") Then
            j = j + 1
            s = f & "\" & j & ".mht"

            If modURL.SavePageToMHT(u, s) Then
                c = c + 1
            End If
        End If
    Next
    
    MsgBox "成功保存" & c & "个文件", vbInformation
End Sub

Private Sub mnuViewExchange_Click()
    Dim a As String
    Dim b As String
    Dim t As String
    
    a = txtExpression.Text
    b = txtResult.Text
    
    t = a
    a = b
    b = t
    txtExpression.Text = a
    txtResult.Text = b
End Sub

Private Sub mnuViewFontAdd_Click()
    txtExpression.Font.Size = txtExpression.Font.Size + 1
    txtResult.Font.Size = txtResult.Font.Size + 1
End Sub

Private Sub mnuViewFontCustom_Click()
    modCommonDialog.SetFont txtExpression
    Set txtResult.Font = txtExpression.Font
End Sub

Private Sub mnuViewFontMinus_Click()

    If txtExpression.Font.Size > 2 Then
        txtExpression.Font.Size = txtExpression.Font.Size - 1
    End If

    If txtResult.Font.Size > 2 Then
        txtResult.Font.Size = txtResult.Font.Size - 1
    End If
End Sub

Private Sub mnuViewFontNormal_Click()
    txtExpression.Font.Size = 12
    txtResult.Font.Size = 12
End Sub

Private Sub mnuViewPosLeftRight_Click()
    spMain.VerticalLayout = False
    mnuViewPosLeftRight.Checked = True
    mnuViewPosUpDown.Checked = False
End Sub

Private Sub mnuViewPosUpDown_Click()
    spMain.VerticalLayout = True
    mnuViewPosUpDown.Checked = True
    mnuViewPosLeftRight.Checked = False
End Sub

Private Sub pnl1_Resize()
    txtExpression.Move 0, 0, pnl1.ScaleWidth, pnl1.ScaleHeight
End Sub

Private Sub pnl2_Resize()
    txtResult.Move 0, 0, pnl2.ScaleWidth, pnl2.ScaleHeight
    lvwFile2.Move 0, 0, pnl2.ScaleWidth, pnl2.ScaleHeight
End Sub

Private Sub pnlContainer_Resize()
    spMain.DoLayout 0.5
End Sub

Private Sub splitter_DblClick()
    If spMain.VerticalLayout Then
        spMain.DoLayout (Me.ScaleHeight - splitter.Height) * 0.5
    Else
        spMain.DoLayout (Me.ScaleWidth - splitter.Width) * 0.5
    End If
End Sub

Private Sub statusbar_Resize()
    Dim w            As Long
    Dim h            As Long
    Dim mBorderColor As Long
    Dim sx           As Long
    Dim sy           As Long
    Dim sp1          As Long
    Dim sp2          As Long
    Dim sp3          As Long
   
    statusbar.Cls
    
    mBorderColor = &HFFFFFF
    
    w = statusbar.ScaleWidth
    h = statusbar.ScaleHeight
    
    sx = Screen.TwipsPerPixelX
    sy = Screen.TwipsPerPixelY
        
    '边框
    statusbar.Line (0, 0)-(w, 0), &H919191
    statusbar.Line (0, sy)-(0, h - sy), mBorderColor '左
    statusbar.Line (sx, sy)-(w, sy), mBorderColor  '上
    statusbar.Line (w - sx, sy)-(w - sx, h - sy), mBorderColor  '右
    statusbar.Line (0, h - sy)-(w, h - sy), mBorderColor    '下
    
    '分隔线
    sp1 = 1800
    sp2 = sp1 * 2
    sp3 = sp2 + 0.3 * w
    
    statusbar.Line (sp1, sy * 4)-(sp1, h - sy * 4), &HB4B4B4
    statusbar.Line (sp2, sy * 4)-(sp2, h - sy * 4), &HB4B4B4
    
    lblStatus(1).Move sp1 + sx * 6, lblStatus(0).Top
    lblStatus(2).Move sp2 + sx * 6, lblStatus(0).Top

    pb.Move w - pb.Width - 6 * sx
End Sub

Private Sub txtExpression_Change()
    Dim lines As Long
    Dim t     As String
    
    lines = UBound(Split(txtExpression.Text, vbCrLf)) + 1
    t = "表达式：共" & lines & "行"
    SetStatus 0, t
    m_b_expression_changed = True
End Sub

Private Sub txtExpression_GotFocus()
    If App.LogMode = LogModeConstants.vbLogOff Then
        modMouseHook.HookMouse txtExpression
    End If
End Sub

Private Sub txtExpression_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
    Case vbCtrlMask
        Select Case KeyCode
        Case vbKeyV
            txtExpression.SelText = Clipboard.GetText(vbCFText)
        Case vbKeyL
            mnuFileListFiles_Click
        Case vbKeyE
            mnuViewExchange_Click
        Case vbKeyR '
            '
            'Case vbKeyZ
            '
        Case vbKeyTab   'CTRL+TAB
            '切换焦点
            txtResult.SetFocus
        Case vbKeyReturn 'CTRL+Enter
            mnuMathEval_Click
        Case Else
            Exit Sub
        End Select
    Case Else
        Select Case KeyCode
        Case 9  '输入tab
            txtExpression.SelText = Chr$(9)
        Case Else
            Exit Sub
        End Select
    End Select

    KeyCode = 0
    Shift = 0
End Sub

Private Sub txtExpression_LostFocus()
    If App.LogMode = LogModeConstants.vbLogOff Then
        modMouseHook.UnHookMouse txtExpression
    End If
End Sub

Private Sub txtExpression_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim d   As Variant
    Dim i   As Long
    Dim p() As String
    
    If Data.GetFormat(vbCFFiles) Then
        ReDim p(1 To Data.Files.Count) As String
    
        For Each d In Data.Files
            i = i + 1
            p(i) = d
        Next
        
        modSort.ShellSortAsc p
        
        With txtExpression
            '空内容
            If Len(.Text) = 0 Then
                .Text = Join(p, vbCrLf)
            Else
    
                '未选定文本
                If .SelLength = 0 Then
    
                    '光标在开始位置
                    If .SelStart = 0 Then
                        .Text = Join(p, vbCrLf) & vbCrLf & .Text
                    Else    '光标在中间
                        .SelText = Join(p, vbCrLf)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub txtResult_Change()
    Dim lines As Long
    Dim t     As String

    lines = UBound(Split(txtResult.Text, vbCrLf)) + 1
    t = "结果：共" & lines & "行"
    SetStatus 1, t
End Sub

Private Sub txtResult_GotFocus()
    If App.LogMode = LogModeConstants.vbLogOff Then
        modMouseHook.HookMouse txtResult
    End If
End Sub

Private Sub txtResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
    Case vbCtrlMask
        Select Case KeyCode
        Case vbKeyV
            txtResult.SelText = Clipboard.GetText(vbCFText)
        Case vbKeyL
            mnuFileListFiles_Click
        Case vbKeyE
            mnuViewExchange_Click
        Case vbKeyR
            '
        Case vbKeyTab   'CTRL+TAB
            '切换焦点
            txtExpression.SetFocus
        Case Else
            Exit Sub
        End Select
    Case Else
        Select Case KeyCode
        Case 9  '输入tab
            txtResult.SelText = Chr$(9)
        Case Else
            Exit Sub
        End Select
    End Select

    KeyCode = 0
    Shift = 0
End Sub

Private Sub txtResult_LostFocus()
    If App.LogMode = LogModeConstants.vbLogOff Then
        modMouseHook.UnHookMouse txtResult
    End If
End Sub
