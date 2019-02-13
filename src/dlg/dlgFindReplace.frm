VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form dlgFindReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找替换"
   ClientHeight    =   7380
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   17745
   Icon            =   "dlgFindReplace.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   492
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1183
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdClearHistory 
      Caption         =   "清空历史"
      Height          =   375
      Left            =   5880
      TabIndex        =   51
      Top             =   900
      Width           =   1755
   End
   Begin VB.PictureBox pnl 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5655
      Index           =   2
      Left            =   12660
      ScaleHeight     =   5655
      ScaleWidth      =   4935
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   660
      Width           =   4935
      Begin VB.TextBox txtNumberIncCount 
         Height          =   330
         Left            =   1740
         TabIndex        =   47
         Top             =   120
         Width           =   420
      End
      Begin VB.TextBox txtNumberIncIndex 
         Height          =   330
         Left            =   360
         TabIndex        =   46
         Top             =   120
         Width           =   420
      End
      Begin VB.CommandButton cmdNumberInc 
         Caption         =   "执行"
         Height          =   330
         Left            =   3720
         TabIndex        =   45
         Top             =   120
         Width           =   1140
      End
      Begin VB.TextBox txtNumberIncZeroize 
         Height          =   330
         Left            =   3180
         TabIndex        =   44
         Top             =   120
         Width           =   420
      End
      Begin VB.CommandButton cmdNumberPreserve 
         Caption         =   "仅保留数值"
         Height          =   360
         Left            =   300
         TabIndex        =   43
         Top             =   660
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "组数值加"
         Height          =   180
         Left            =   900
         TabIndex        =   50
         Top             =   195
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "第"
         Height          =   180
         Left            =   120
         TabIndex        =   49
         Top             =   195
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补零位数"
         Height          =   180
         Left            =   2400
         TabIndex        =   48
         Top             =   195
         Width           =   720
      End
   End
   Begin VB.ComboBox txtReplace 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   5580
   End
   Begin VB.ComboBox txtFind 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   5580
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "执行（&E)"
      Default         =   -1  'True
      Height          =   990
      Left            =   5880
      TabIndex        =   20
      Top             =   5760
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      Caption         =   "查找选项"
      Height          =   1455
      Left            =   5880
      TabIndex        =   39
      Top             =   2640
      Width           =   1755
      Begin VB.CheckBox chkFindFilename 
         Caption         =   "仅查找文件名"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkIgnoreCase 
         Caption         =   "区分大小写(&S)"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1515
      End
      Begin VB.CheckBox chkUseRegexp 
         Caption         =   "正则表达式(&X)"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "执行结果输出"
      Height          =   1455
      Left            =   5880
      TabIndex        =   40
      Top             =   4200
      Width           =   1755
      Begin VB.OptionButton optReplaceAll 
         Caption         =   "同时替换"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1020
         Width           =   1335
      End
      Begin VB.OptionButton optReplaceExpression 
         Caption         =   "替换表达式"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optReplaceResult 
         Caption         =   "输出到结果"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "查找目标"
      Height          =   1155
      Left            =   5880
      TabIndex        =   38
      Top             =   1380
      Width           =   1755
      Begin VB.OptionButton optFindExpression 
         Caption         =   "表达式"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.OptionButton optFindResult 
         Caption         =   "结果"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1410
      End
   End
   Begin VB.PictureBox pnl 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4995
      Index           =   0
      Left            =   240
      ScaleHeight     =   4995
      ScaleWidth      =   5190
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1320
      Width           =   5190
      Begin VB.CommandButton cmdClearExtension 
         Caption         =   "清空扩展名"
         Height          =   330
         Left            =   1500
         TabIndex        =   11
         Top             =   660
         Width           =   1170
      End
      Begin VB.CommandButton cmdClearFilename 
         Caption         =   "清空文件名"
         Height          =   330
         Left            =   180
         TabIndex        =   10
         Top             =   660
         Width           =   1170
      End
      Begin VB.TextBox txtInsertChar 
         Height          =   330
         Left            =   2580
         TabIndex        =   8
         Top             =   120
         Width           =   1680
      End
      Begin VB.TextBox txtInsertCharAt 
         Height          =   330
         Left            =   900
         TabIndex        =   7
         Top             =   120
         Width           =   480
      End
      Begin VB.CommandButton cmdInsertChar 
         Caption         =   "执行"
         Height          =   330
         Left            =   4320
         TabIndex        =   9
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "在第(&I)"
         Height          =   180
         Left            =   180
         TabIndex        =   34
         Top             =   180
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "个字符后插入"
         Height          =   180
         Left            =   1440
         TabIndex        =   35
         Top             =   180
         Width           =   1080
      End
   End
   Begin VB.PictureBox pnl 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5655
      Index           =   1
      Left            =   7920
      ScaleHeight     =   5655
      ScaleWidth      =   4455
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   660
      Width           =   4455
      Begin VB.ComboBox cboKeyWithin 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelKeyIn 
         Caption         =   "删除包含"
         Height          =   360
         Left            =   2100
         TabIndex        =   14
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdDelKeyRight 
         Caption         =   "删除右侧"
         Height          =   360
         Left            =   2100
         TabIndex        =   17
         Top             =   600
         Width           =   990
      End
      Begin VB.ComboBox cboKeyStart 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   570
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelKeyLeft 
         Caption         =   "删除左侧"
         Height          =   360
         Left            =   3120
         TabIndex        =   18
         Top             =   600
         Width           =   990
      End
      Begin VB.CheckBox chkKeepKey 
         Caption         =   "保留符号"
         Height          =   255
         Left            =   3180
         TabIndex        =   15
         Top             =   240
         Width           =   1035
      End
      Begin VB.CheckBox chkFindKeyReverse 
         Caption         =   "从右侧查找"
         Height          =   255
         Left            =   2100
         TabIndex        =   19
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "特定符号"
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "特定符号"
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   690
         Width           =   720
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5835
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   10292
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "插入删除(&1)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "符号(&2)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "数值(&3)"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdPopupReplace 
      Caption         =   ">"
      Height          =   330
      Left            =   6600
      TabIndex        =   4
      Top             =   480
      Width           =   360
   End
   Begin VB.CommandButton cmdPopupFind 
      Caption         =   ">"
      Height          =   330
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   360
   End
   Begin VB.CommandButton cmdRenameFiles 
      Caption         =   "重命名"
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   6900
      Width           =   1635
   End
   Begin VB.CommandButton cmdExchange 
      Caption         =   "↑↓"
      Height          =   690
      Left            =   7020
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdFindAll 
      Caption         =   "查找全部(&A)"
      Height          =   375
      Left            =   960
      TabIndex        =   22
      Top             =   6900
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   6420
      TabIndex        =   24
      Top             =   6900
      Width           =   1215
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   180
      Left            =   180
      TabIndex        =   41
      Top             =   7020
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "替换(&R):"
      Height          =   180
      Left            =   180
      TabIndex        =   33
      Top             =   555
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查找(&F):"
      Height          =   180
      Left            =   180
      TabIndex        =   32
      Top             =   195
      Width           =   720
   End
   Begin VB.Menu popFind 
      Caption         =   "查找"
      Visible         =   0   'False
      Begin VB.Menu popFindInsertParentheses 
         Caption         =   "插入圆括号(&C)"
      End
      Begin VB.Menu popFindInsertBrackets 
         Caption         =   "插入方括号(&B)"
      End
      Begin VB.Menu popFindSep1 
         Caption         =   "-"
      End
      Begin VB.Menu popFindNumber 
         Caption         =   "数字(&N)          \d"
      End
      Begin VB.Menu popFindNumberGroup 
         Caption         =   "数字组(&D)        (\d)"
      End
      Begin VB.Menu popFindSep2 
         Caption         =   "-"
      End
      Begin VB.Menu popFindLastSubstr 
         Caption         =   "最后一个子串(&L)  {L}"
      End
   End
   Begin VB.Menu popReplace 
      Caption         =   "替换"
      Visible         =   0   'False
      Begin VB.Menu popReplaceCopyFind 
         Caption         =   "复制查找文本(&C)"
      End
      Begin VB.Menu popReplaceInsertSerial 
         Caption         =   "插入序号(&S)      \{0}"
      End
      Begin VB.Menu popReplaceCNSerial2Albert 
         Caption         =   "中文转数字(&A)    {n}"
      End
      Begin VB.Menu popReplaceSep1 
         Caption         =   "-"
      End
      Begin VB.Menu popReplaceExpSub 
         Caption         =   "整个表达式(&0)    $0"
         Index           =   0
      End
      Begin VB.Menu popReplaceExpSub 
         Caption         =   "子表达式1(&1)     $1"
         Index           =   1
      End
      Begin VB.Menu popReplaceExpSub 
         Caption         =   "子表达式2(&2)     $2"
         Index           =   2
      End
      Begin VB.Menu popReplaceExpSub 
         Caption         =   "子表达式3(&3)     $3"
         Index           =   3
      End
      Begin VB.Menu popReplaceExpSub 
         Caption         =   "子表达式4(&4)     $4"
         Index           =   4
      End
   End
End
Attribute VB_Name = "dlgFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents shdE As RichTextBox
Attribute shdE.VB_VarHelpID = -1
Dim WithEvents shdR As RichTextBox
Attribute shdR.VB_VarHelpID = -1

Dim objLastFocus    As ComboBox

Dim objFso As FileSystemObject

Private Const INI_SECTION_FIND As String = "FindText"
Private Const INI_SECTION_REPLACE As String = "ReplaceText"

Private Sub ClearHistory()
    Dim ini As INIProfile
    
    Set ini = New INIProfile

    With ini
        .ExeFolderPath = App.Path
        .Name = App.title
        
        .DeleteSection INI_SECTION_FIND
        .DeleteSection INI_SECTION_REPLACE
        
    End With
    
    txtFind.Clear
    txtReplace.Clear
    txtFind.Text = ""
    txtReplace.Text = ""
    
    Set ini = Nothing
End Sub

Private Sub DeleteFindHistory(ByVal Value As String)
    Dim ini As INIProfile
    
    Set ini = New INIProfile

    With ini
        .ExeFolderPath = App.Path
        .Name = App.title
    
        .DeleteKey INI_SECTION_FIND, .GetKeyByValue(INI_SECTION_FIND, Value)
    End With
    
    Set ini = Nothing
End Sub

Private Sub DeleteReplaceHistory(ByVal Value As String)
    Dim ini As INIProfile
    
    Set ini = New INIProfile

    With ini
        .ExeFolderPath = App.Path
        .Name = App.title
    
        .DeleteKey INI_SECTION_REPLACE, .GetKeyByValue(INI_SECTION_FIND, Value)
    End With
    LoadFindHistory ini
    
    Set ini = Nothing
End Sub

Private Sub SaveHistory(ByVal sFind As String, ByVal sReplace As String)
    Dim ini As INIProfile
    Dim sKeyFind As String
    Dim sKeyReplace As String
    Dim i As Long
    Dim j As Long
    Dim s As String
    
    Set ini = New INIProfile

    With ini
        .ExeFolderPath = App.Path
        .Name = App.title
        If Len(sFind) > 0 Then
            If .IsValueExists(INI_SECTION_FIND, sFind) = False Then
                sKeyFind = .GetNextKey(INI_SECTION_FIND)
                .SetString INI_SECTION_FIND, sKeyFind, sFind
            End If
        End If
        If Len(sReplace) > 0 Then
            If .IsValueExists(INI_SECTION_REPLACE, sReplace) = False Then
                sKeyReplace = .GetNextKey(INI_SECTION_REPLACE)
                .SetString INI_SECTION_REPLACE, sKeyReplace, sReplace
            End If
        End If
    End With
    
    Set ini = Nothing
End Sub

Private Sub LoadFindHistory(ByRef ini As INIProfile)
    Dim p() As String
    Dim i As Long
    Dim s As String
    
    p = ini.GetAllKeys(INI_SECTION_FIND)
    txtFind.Clear
    
    For i = 0 To UBound(p)
        s = ini.GetString("FindText", p(i))
        If Len(s) > 0 Then
            txtFind.AddItem s
        End If
    Next
End Sub

Private Sub LoadReplaceHistory(ByRef ini As INIProfile)
    Dim p() As String
    Dim i As Long
    Dim s As String
    
    p = ini.GetAllKeys(INI_SECTION_REPLACE)
    txtReplace.Clear
    
    For i = 0 To UBound(p)
        s = ini.GetString("ReplaceText", p(i))
        If Len(s) > 0 Then
            txtReplace.AddItem s
        End If
    Next
End Sub

Private Sub LoadHistory()
    Dim ini As INIProfile
    
    Set ini = New INIProfile

    With ini
        .ExeFolderPath = App.Path
        .Name = App.title
    End With
    
    LoadFindHistory ini
    LoadReplaceHistory ini
    
    Set ini = Nothing
End Sub

Private Function IsCNNumber(ByVal Char As String)
    Dim s As String
    s = "一二三四五六七八九十百零〇"
    IsCNNumber = InStr(s, Char) > 0
End Function

Private Sub SingletonFSO()
    If objFso Is Nothing Then
        Set objFso = New FileSystemObject
    End If
End Sub

Private Sub cmdClearExtension_Click()
    Dim s          As String
    Dim p()        As String
    Dim i          As Long
    Dim dotIndex   As Long
    Dim slashIndex As Long
    
    s = frmMain!txtExpression.Text
    p = Split(s, vbCrLf)
    
    For i = 0 To UBound(p)
        If Len(p(i)) > 0 Then
            slashIndex = InStrRev(p(i), "\")

            dotIndex = InStrRev(p(i), ".")

            If dotIndex > 0 Then
                p(i) = Left$(p(i), slashIndex) & Mid$(p(i), slashIndex + 1, dotIndex - slashIndex - 1)
            End If
        End If
    Next
    
    s = Join(p, vbCrLf)
    
    If optReplaceResult.Value Then
        frmMain.SetResult s
    ElseIf optReplaceExpression.Value Then
        frmMain!txtExpression.Text = s
    Else
        frmMain!txtExpression.Text = s
        frmMain.SetResult s
    End If
    
    optFindResult.Value = True
End Sub

Private Sub cmdClearFilename_Click()
    Dim s          As String
    Dim p()        As String
    Dim i          As Long
    Dim dotIndex   As Long
    Dim slashIndex As Long
    
    s = frmMain!txtExpression.Text
    p = Split(s, vbCrLf)
    
    For i = 0 To UBound(p)
        If Len(p(i)) > 0 Then
            slashIndex = InStrRev(p(i), "\")

            dotIndex = InStrRev(p(i), ".")
            p(i) = Left$(p(i), slashIndex) & Mid$(p(i), dotIndex)
        End If
    Next
    
    s = Join(p, vbCrLf)
    
    If optReplaceResult.Value Then
        frmMain.SetResult s
    ElseIf optReplaceExpression.Value Then
        frmMain!txtExpression.Text = s
    Else
        frmMain!txtExpression.Text = s
        frmMain.SetResult s
    End If
    
    optFindResult.Value = True
End Sub


Private Sub FindAllByReg()
    Dim s   As String
    Dim reg As RegExp
    Dim mc  As MatchCollection
    
    If optFindExpression.Value Then
        s = frmMain!txtExpression.Text
    Else
        s = frmMain!txtResult.Text
    End If
    
    Set reg = New RegExp
    
    reg.Pattern = txtFind.Text
    reg.Global = True
    reg.MultiLine = True
    reg.IgnoreCase = (chkIgnoreCase.Value <> 1)
    
    Set mc = reg.Execute(s)
    
    lblResult.Caption = mc.Count
    
    Set mc = Nothing
    Set reg = Nothing
End Sub


Private Function ReplaceSymbol(ByVal Source As String, ByVal Find As String, ByVal Replacement As String, ByVal Compare As VbCompareMethod)
    Dim p()        As String
    Dim u          As Long
    Dim R          As String
    Dim k          As String
    Dim zcount     As Long
    Dim znumberFmt As String
    Dim nStartNum  As Integer
    Dim nkp        As Long
    Dim i As Long
    Dim sSubstr As String
    Dim nIndex As Long
    Dim sLft As String
    Dim sRgt As String
    
    p = Split(Source, vbCrLf)
    u = UBound(p)
    
    k = "{L}"  '最后一个
    nkp = InStr(1, Find, k)
    sSubstr = Left(Find, nkp - 1)
        
    For i = 0 To u
        nIndex = InStrRev(p(i), sSubstr, , Compare)
        sLft = Left$(p(i), nIndex - 1)
        sRgt = Mid$(p(i), nIndex + Len(sSubstr))
        
        
        p(i) = sLft & Replacement & sRgt
    Next
    
    ReplaceSymbol = Join(p, vbCrLf)
End Function

'替换自定义标记
Private Function ReplaceMark(ByVal Source As String, ByVal Find As String, ByVal Replacement As String)
    Dim i          As Long
    Dim p()        As String
    Dim u          As Long
    Dim R          As String
    Dim k          As String
    Dim zcount     As Long
    Dim znumberFmt As String
    Dim nStartNum  As Integer
    Dim nkp        As Long
    
    p = Split(Source, vbCrLf)
    u = UBound(p)
    
    k = "\{s"  '数字序号
    nkp = InStr(1, Replacement, k)

    If nkp > 0 Then
        zcount = Len(Trim$(str(u + 1)))
        nStartNum = Mid$(Replacement, nkp + Len(k), InStr(nkp, Replacement, "}") - nkp - Len(k))
        k = k & nStartNum & "}"
        nStartNum = Int(nStartNum)

        If zcount > 1 Then
            znumberFmt = String(zcount, "0")
        Else
            znumberFmt = ""
        End If

        For i = 0 To u
            R = Replace$(Replacement, k, Format$(i + nStartNum + 1, znumberFmt))
            p(i) = Replace$(p(i), Find, R)
        Next
    End If
    
    ReplaceMark = Join(p, vbCrLf)
End Function

Private Function ReplaceAllFilename(ByVal Source As String, ByVal sFind As String, ByVal sRepl As String, Optional ByVal CheckFileExists As Boolean = True) As String
    Dim p() As String
    Dim i As Long
    Dim s As String
    
    p = Split(Source, vbCrLf)
    
    SingletonFSO
    
    For i = 0 To UBound(p)
        p(i) = ReplaceFilename(p(i), sFind, sRepl, "", CheckFileExists)
    Next
    ReplaceAllFilename = Join(p, vbCrLf)
End Function

Private Function ReplaceFilename(ByVal Path As String, ByVal sFind As String, ByVal sRepl As String, Optional ByVal OldName As String = "", Optional ByVal CheckFileExists As Boolean = True) As String
    Dim sExt        As String
    Dim sFolder     As String
    
    SingletonFSO
    
    If CheckFileExists = True And objFso.FileExists(Path) = False Then
        ReplaceFilename = Path
    Else
        sFolder = objFso.GetParentFolderName(Path)
        If Len(OldName) = 0 Then
            OldName = objFso.GetBaseName(Path)
        End If
        sExt = objFso.GetExtensionName(Path)
        If Len(sExt) > 0 Then
            sExt = "." & sExt
        End If
        
        ReplaceFilename = sFolder & "\" & Replace$(OldName, sFind, sRepl) & sExt
    End If
End Function


Private Function ReplaceFilenameAs(ByVal Path As String, ByVal NewName As String, Optional ByVal CheckFileExists As Boolean = True) As String
    Dim sExt        As String
    Dim sFolder     As String
    
    SingletonFSO
    
    If CheckFileExists = True And objFso.FileExists(Path) = False Then
        ReplaceFilenameAs = Path
    Else
        sFolder = objFso.GetParentFolderName(Path)
        sExt = objFso.GetExtensionName(Path)
        If Len(sExt) > 0 Then
            sExt = "." & sExt
        End If
        
        ReplaceFilenameAs = sFolder & "\" & NewName & sExt
    End If
End Function

'获取要替换的字符
Private Function GetReplaceArea(ByVal Source As String, ByVal sLft As String, ByVal sRgt As String) As String
    Dim nLft As Long
    Dim nRgt As Long
    Dim c As String
    Dim u As Long
    Dim i As Long
    Dim sToReplace As String
    
    u = Len(Source)
    
    If u = 0 Then Exit Function
    
    nLft = InStr(Source, sLft) + 1
    For i = nLft To u
        c = Mid$(Source, i, 1)
        If Not IsCNNumber(c) Then
            Exit For
        End If
    Next
    nRgt = i
    
    sToReplace = Mid$(Source, nLft, nRgt - nLft)
    
    GetReplaceArea = Replace$(Source, sLft & sToReplace, sLft & modStrings.CNSerial2Albert(sToReplace), 1, 1)
    
End Function

'替换文件名中的特殊标记
Private Function ReplaceFilenameMark(ByVal Source As String, ByVal Find As String, ByVal Replacement As String, Optional ByVal CheckFileExists As Boolean = True) As String
    Dim p() As String
    Dim i As Long
    Dim u As Long
    Dim nLft As Long
    Dim nRgt As Long
    Dim sLft As String
    Dim sRgt As String
    Dim sVal As String
    Dim s As String
    
    p = Split(Source, vbCrLf)
    u = UBound(p)
    
    nLft = InStr(Replacement, "{")
    nRgt = InStr(nLft + 1, Replacement, "}")
    
    If nRgt <= nLft Then
        MsgBox "替换内容中标记语法有误", vbExclamation
        txtReplace.SetFocus
        Exit Function
    End If
    
    sLft = Left$(Replacement, nLft - 1)
    sRgt = Mid$(Replacement, nRgt + 1)
    
    sVal = Mid$(Replacement, nLft + 1, 1)
    
    '汉字转数字
    If sVal = "n" Then
        For i = 0 To UBound(p)
            p(i) = ReplaceFilenameAs(p(i), GetReplaceArea(objFso.GetBaseName(p(i)), sLft, sRgt), CheckFileExists)
        Next
        
        ReplaceFilenameMark = Join(p, vbCrLf)
    End If
End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmdClearHistory_Click()
    ClearHistory
End Sub

Private Sub cmdDelKeyIn_Click()
    Dim s   As String
    Dim k   As String
    Dim a   As String
    Dim b   As String
    Dim i   As Long
    Dim p() As String
    Dim f   As String
    Dim m   As String
    Dim e   As String
   
    k = cboKeyWithin.Text
    
    If Len(k) = 0 Then
        Exit Sub
    End If
    
    If k = "空格" Then
        k = " "
    End If
    
    a = Left$(k, 1)
    b = Right$(k, 1)
    
    If optFindResult.Value Then
        s = frmMain!txtResult.Text
    Else
        s = frmMain!txtExpression.Text
    End If
    
    p = Split(s, vbCrLf)
    
    '找文件名，找表达式中的文件名
    If chkFindFilename.Value = 1 And optFindResult.Value = 0 Then
        SingletonFSO
        
        For i = 0 To UBound(p)
            s = p(i)
            If Len(s) > 0 Then
                m = objFso.GetBaseName(s)
                m = modStrings.ReplaceMidKey(m, a, , b, (chkKeepKey.Value = 1))
    
                p(i) = ReplaceFilenameAs(s, m)
            End If
        Next
    Else
        For i = 0 To UBound(p)
            If Len(p(i)) > 0 Then
                p(i) = modStrings.ReplaceMidKey(p(i), a, , b, (chkKeepKey.Value = 1))
            End If
        Next
    End If
        
    frmMain!txtResult.Text = Join(p, vbCrLf)
End Sub

Private Sub cmdDelKeyLeft_Click()
    Dim s   As String
    Dim k   As String
    Dim a   As String
    Dim b   As String
    Dim i   As Long
    Dim p() As String
    Dim fso As FileSystemObject
    Dim f   As String
    Dim m   As String
    Dim e   As String
    Dim v   As Long
    
    k = cboKeyStart.Text
    
    If Len(k) = 0 Then
        Exit Sub
    End If

    If k = "空格" Then
        k = " "
    End If
    
    a = Left$(k, 1)
    b = Right$(k, 1)
    
    If optFindResult.Value Then
        s = frmMain!txtResult.Text
    Else
        s = frmMain!txtExpression.Text
    End If
    
    p = Split(s, vbCrLf)
    
    Set fso = New FileSystemObject

    For i = 0 To UBound(p)
        If Len(p(i)) > 0 Then
            If optFindResult.Value Or fso.FileExists(p(i)) Then
                f = fso.GetParentFolderName(p(i))

                If Right$(f, 1) <> "\" Then
                    f = f & "\"
                End If

                m = fso.GetBaseName(p(i))
                e = fso.GetExtensionName(p(i))

                If Len(e) > 0 Then
                    e = "." & e
                End If
            Else
                f = ""
                m = p(i)
                e = ""
            End If
                        
            If chkFindKeyReverse.Value = 1 Then
                v = InStrRev(m, k, , vbBinaryCompare)

                If v > 0 Then
                    m = Mid$(m, v + 1)
                End If
            Else
                v = InStr(1, m, k, vbBinaryCompare)
                m = Mid$(m, v + 1)
            End If
            
            p(i) = f & m & e
        End If
    Next
    
    frmMain!txtResult.Text = Join(p, vbCrLf)
    
    Set fso = Nothing
End Sub

Private Sub cmdDelKeyRight_Click()
    Dim s   As String
    Dim k   As String
    Dim a   As String
    Dim b   As String
    Dim i   As Long
    Dim p() As String
    Dim fso As FileSystemObject
    Dim f   As String
    Dim m   As String
    Dim e   As String
    Dim v   As Long
    
    k = cboKeyStart.Text
    
    If Len(k) = 0 Then
        Exit Sub
    End If

    If k = "空格" Then
        k = " "
    End If
    
    a = Left$(k, 1)
    b = Right$(k, 1)
    
    If optFindResult.Value Then
        s = frmMain!txtResult.Text
    Else
        s = frmMain!txtExpression.Text
    End If
    
    p = Split(s, vbCrLf)
    
    Set fso = New FileSystemObject

    For i = 0 To UBound(p)
        If Len(p(i)) > 0 Then
            If optFindResult.Value Or fso.FileExists(p(i)) Then
                f = fso.GetParentFolderName(p(i))

                If Right$(f, 1) <> "\" Then
                    f = f & "\"
                End If

                m = fso.GetBaseName(p(i))
                e = fso.GetExtensionName(p(i))

                If Len(e) > 0 Then
                    e = "." & e
                End If
            Else
                f = ""
                m = p(i)
                e = ""
            End If
            
            If chkFindKeyReverse.Value = 1 Then
                v = InStrRev(m, k, , vbBinaryCompare)

                If v > 0 Then
                    m = Left$(m, v - 1)
                End If
            Else
                v = InStr(1, m, k, vbBinaryCompare)

                If v > 0 Then
                    m = Left$(m, v - 1)
                End If
            End If
            
            p(i) = f & m & e
        End If
    Next
    
    frmMain!txtResult.Text = Join(p, vbCrLf)
    
    Set fso = Nothing
End Sub

Private Sub cmdExchange_Click()
    Dim s As String

    s = txtReplace.Text
    txtReplace.Text = txtFind.Text
    txtFind.Text = s
End Sub

Private Sub cmdFindAll_Click()
    If chkUseRegexp.Value = 1 Then
        FindAllByReg
    Else
        lblResult.Caption = modStrings.SubCount(frmMain!txtExpression.Text, txtFind.Text, (chkIgnoreCase.Value = 1))
    End If
End Sub

Private Sub cmdInsertChar_Click()
    Dim s   As String
    Dim p() As String
    Dim i   As Long
    Dim c   As String
    Dim d   As Long
    Dim f   As String
   
    c = txtInsertChar.Text
    
    If IsNumeric(txtInsertCharAt.Text) Then
        d = CLng(Val(txtInsertCharAt.Text))
    Else
        d = 0
    End If
    
    If optFindExpression.Value Then
        s = frmMain!txtExpression.Text
    Else
        s = frmMain!txtResult.Text
    End If
    
    p = Split(s, vbCrLf)
    
    For i = 0 To UBound(p)
        If Len(p(i)) > 0 Then
            f = Left$(p(i), InStrRev(p(i), "\"))
            s = Mid$(p(i), InStrRev(p(i), "\") + 1)
            
            If d = 0 Then
                '开头
                p(i) = f & c & s
            ElseIf d > 0 Then
                '中间
                p(i) = f & Left$(s, d) & c & Mid$(s, d)
            ElseIf d = -1 Then
                '末尾
                p(i) = f & s & c
            Else
                '倒数
                If -d > Len(s) Then
                    p(i) = f & c & s
                Else
                    p(i) = f & Left$(s, Len(s) + d + 1) & c & Right$(s, -(d + 1))
                End If
            End If
        End If
    Next
    
    frmMain.SetResult Join(p, vbCrLf)
End Sub

Private Function NumberInc(ByVal s As String, ByVal nStartIndex As Long, ByVal nIncCount As Long, ByVal nZeroize As Long) As String
    Dim j As Long
    Dim k As Long
    Dim c As String
    Dim numberStart         As Long
    Dim bufferNumber        As String
    Dim isCurrentCharNumber As Boolean
    Dim numberIndex         As Long
    Dim numberCurrent       As Long
    
    
    For j = 1 To Len(s)
        c = Mid$(s, j, 1)

        If Asc(c) > 47 And Asc(c) < 58 Then    '当前是数字
            bufferNumber = bufferNumber & c '存入buffer
            isCurrentCharNumber = True
        Else

            '当前不是数字了
            If isCurrentCharNumber Then '前面的又是数字
                k = k + 1

                If k = nStartIndex Then
                    Exit For
                Else
                    bufferNumber = ""   '清空buffer
                End If
            End If

            isCurrentCharNumber = False '当前不是数字
        End If
    Next

    numberCurrent = CLng(bufferNumber)
    
    If nZeroize > 0 Then
        NumberInc = Left$(s, j - 1 - Len(bufferNumber)) & Format(numberCurrent + nIncCount, String(nZeroize, "0")) & Mid$(s, j)
    Else
        NumberInc = Left$(s, j - 1 - Len(bufferNumber)) & numberCurrent + nIncCount & Mid$(s, j)
    End If
End Function

'选定位置数值增加
Private Sub cmdNumberInc_Click()
    Dim nIncCount   As Long
    Dim nZeroize    As Long
    Dim s           As String
    Dim f           As String
    Dim p()         As String
    Dim i           As Long
    Dim nStartIndex As Long
    
    If IsNumeric(txtNumberIncIndex.Text) Then
        nStartIndex = Val(txtNumberIncIndex.Text)

        If nStartIndex < 1 Then
            MsgBox "要增加的数值组数从1开始", vbExclamation
            txtNumberIncIndex.SelStart = 0
            txtNumberIncIndex.SelLength = Len(txtNumberIncIndex.Text)
            txtNumberIncIndex.SetFocus

            Exit Sub
        End If

    Else
        txtNumberIncIndex.Text = "1"
        nStartIndex = 1
    End If
    
    If IsNumeric(txtNumberIncCount.Text) Then
        nIncCount = Val(txtNumberIncCount.Text)
    End If
    
    If IsNumeric(txtNumberIncZeroize.Text) Then
        nZeroize = Val(txtNumberIncZeroize.Text)

        If nZeroize < 0 Then
            MsgBox "补零位数不能小于0", vbExclamation
            txtNumberIncZeroize.SelStart = 0
            txtNumberIncZeroize.SelLength = Len(txtNumberIncZeroize.Text)
            txtNumberIncZeroize.SetFocus

            Exit Sub
        End If
    End If
    
    If optFindExpression.Value Then
        s = frmMain!txtExpression.Text
    Else
        s = frmMain!txtResult.Text
    End If
    
    p = Split(s, vbCrLf)

    '取第一行的指定组的数值
    If chkFindFilename.Value = 1 Then
        SingletonFSO
        For i = 0 To UBound(p)
            If Len(p(i)) > 0 Then
                p(i) = ReplaceFilenameAs(p(i), NumberInc(objFso.GetBaseName(p(i)), nStartIndex, nIncCount, nZeroize))
            End If
        Next
    Else
        For i = 0 To UBound(p)
            If Len(p(i)) > 0 Then
                p(i) = NumberInc(p(i), nStartIndex, nIncCount, nZeroize)
            End If
        Next
    End If
    
    If optReplaceExpression.Value Then
        frmMain!txtExpression.Text = Join(p, vbCrLf)
    Else
        frmMain.SetResult Join(p, vbCrLf)
    End If

End Sub

Private Sub cmdPopupFind_Click()
    PopupMenu popFind
End Sub

Private Sub cmdPopupReplace_Click()
    PopupMenu popReplace
End Sub

Private Sub cmdRenameFiles_Click()
    frmMain.mnuFileRename_Click
End Sub

Private Sub cmdExecute_Click()
    Dim v        As String
    Dim s        As String
    Dim sFind    As String
    Dim sReplace As String
    
    sFind = txtFind.Text
    sReplace = txtReplace.Text
    
    '在表达式中查找
    If optFindExpression.Value Then
        s = frmMain!txtExpression.Text
    Else
        '在结果中查找
        s = frmMain!txtResult.Text
    End If
    
    '使用正则
    If (chkUseRegexp.Value = 1) Then
        On Error Resume Next
        v = RegReplace(s, sFind, sReplace, (chkIgnoreCase.Value = 1))
        
        If Err Then
            Err.Clear
            On Error GoTo 0
            MsgBox "正则表达式有误", vbExclamation
            txtFind.SetFocus
            Exit Sub
        End If
    Else
        '替换特定标记
        If InStr(1, txtReplace.Text, "\{") Then
            v = ReplaceMark(s, sFind, sReplace)
        ElseIf InStr(1, txtFind.Text, "{L}") Then
            v = ReplaceSymbol(s, sFind, sReplace, 1 - chkIgnoreCase.Value)
        Else
            If chkFindFilename.Value = 1 Then
                SingletonFSO
                If InStr(1, txtReplace.Text, "{") Then
                    v = ReplaceFilenameMark(s, sFind, sReplace, optFindResult.Value = 0)
                Else
                    v = ReplaceAllFilename(s, sFind, sReplace, optFindResult.Value = 0)
                End If

            Else
                v = Replace$(s, sFind, sReplace)
            End If
            
        End If
    End If

    If optReplaceResult.Value Then
        frmMain.SetResult v
    ElseIf optReplaceExpression Then
        frmMain!txtExpression.Text = v
    Else
        frmMain!txtExpression.Text = v
        frmMain.SetResult v
    End If
    
    '执行一次以后自动调到搜索结果
    optFindResult.Value = True
    
    '保存历史
    SaveHistory txtFind.Text, txtReplace.Text
End Sub


Private Sub Form_Load()
    Dim ini As INIProfile
    Dim p() As String
    Dim i   As Long
    
    Set shdE = frmMain!txtExpression
    Set shdR = frmMain!txtResult
    
    lblResult.Caption = ""

    LoadHistory
    
    If Len(shdE.SelText) > 0 Then
        txtFind.Text = shdE.SelText
    ElseIf Len(shdR.SelText) > 0 Then
        txtFind.Text = shdR.SelText
    End If
    
    p = Split(frmMain.g_StringKeyWithIn, "|")

    For i = 0 To UBound(p)
        cboKeyWithin.AddItem p(i)
    Next

    cboKeyWithin.ListIndex = 0
    
    p = Split(frmMain.g_StringKeyStart, "|")

    For i = 0 To UBound(p)
        cboKeyStart.AddItem p(i)
    Next

    cboKeyStart.ListIndex = 0
    
    popFind.Visible = False
    popReplace.Visible = False
    
    txtNumberIncZeroize.Text = Len(CStr((UBound(Split(frmMain!txtExpression.Text, vbCrLf)) + 1)))
    
    For i = 0 To pnl.UBound
        With pnl(i)
            .Move 16, 88, 346
            .Visible = False
            .BackColor = &H8000000F
        End With
    Next
    
    pnl(0).Visible = True
    Me.Width = 7800
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set shdE = Nothing
    Set shdR = Nothing
    Set objFso = Nothing
End Sub

Private Sub popFindInsertBrackets_Click()
    txtFind.SelText = "[" & txtFind.SelText & "]"
    chkUseRegexp.Value = 1
End Sub

Private Sub popFindInsertParentheses_Click()
    txtFind.SelText = "(" & txtFind.SelText & ")"
    chkUseRegexp.Value = 1
End Sub

Private Sub popFindLastSubstr_Click()
    Dim s As String

    s = txtFind.SelText
    
    If Len(s) > 0 Then
        txtFind.SelText = txtFind.SelText & "{L}"
    Else
        txtFind.SelStart = Len(txtFind.Text)
        txtFind.Text = txtFind.Text & "{L}"
    End If

    chkUseRegexp.Value = 0
End Sub

Private Sub popFindNumber_Click()
    Dim s As String

    s = txtFind.SelText
    
    If Len(s) > 0 Then
        txtFind.SelText = "\d{" & Len(s) & "}"
    Else
        txtFind.SelText = "\d"
    End If

    chkUseRegexp.Value = 1
End Sub

Private Sub popFindNumberGroup_Click()
    Dim s As String

    s = txtFind.SelText
    
    If Len(s) > 0 Then
        txtFind.SelText = "(\d{" & Len(s) & "})"
    Else
        txtFind.SelText = "(\d)"
    End If

    chkUseRegexp.Value = 1
End Sub

Private Sub popReplaceCNSerial2Albert_Click()
    Dim i As Long
    Dim c As String
    Dim R As String
    Dim n As Long
    Dim bChanged  As Boolean
    Dim v As String
    
    R = txtReplace.Text

    n = Len(R)
    If n = 0 Then
        R = txtFind.Text
        n = Len(R)
        For i = 1 To n
            c = Mid$(R, i, 1)
            If IsCNNumber(c) Then
                If bChanged = False Then
                    v = v & "{n}"
                    bChanged = True
                End If
            Else
                v = v & c
            End If
            
        Next
        txtReplace.Text = v
    Else
        txtReplace.SelText = "{n}"
    End If
End Sub

Private Sub popReplaceCopyFind_Click()
    txtReplace.SelText = txtFind.Text
End Sub

Private Sub popReplaceExpSub_Click(Index As Integer)
    txtReplace.SelText = "$" & Index
    chkUseRegexp.Value = 1
End Sub

Private Sub popReplaceInsertSerial_Click()
    txtReplace.SelText = "\{s0}"
    chkUseRegexp.Value = 0
End Sub

Private Sub shdE_GotFocus()
    optFindExpression.Value = True
End Sub

Private Sub shdE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        objLastFocus.Text = Replace(shdE.SelText, vbCrLf, "")
    End If
End Sub

Private Sub shdR_GotFocus()
    optFindResult.Value = True
End Sub

Private Sub shdR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        objLastFocus.Text = Replace(shdR.SelText, vbCrLf, "")
    End If
End Sub

Private Sub shdR_SelChange()
    Dim nIndex     As Long
    Dim s          As String
    Dim sLeft      As String
    Dim nLeftCrlf  As Long
    Dim nRightCrlf As Long
    Dim nLastSlash As Long
    Dim sLine      As String
   
    Select Case TabStrip1.SelectedItem.Index
    Case 1  '插入删除
        Exit Sub

        '字符序号改为光标位置
        s = shdR.Text
        nIndex = shdR.SelStart
        
        If nIndex = 0 Then
            txtInsertCharAt.Text = "0"
            Exit Sub
        End If
        
        '看光标左边是否有vbCrlf，如果有就不是第一行，从当前行取，如果无，就是第一行，直接取
        sLeft = Left$(s, nIndex)
        nLeftCrlf = InStrRev(sLeft, vbCrLf)

        If nLeftCrlf = 0 Then
            nRightCrlf = InStr(s, vbCrLf)
        Else
            nRightCrlf = InStr(nLeftCrlf + 2, s, vbCrLf)
        End If

        '光标在第一行
        If nLeftCrlf = 0 Then
            If nRightCrlf = 0 Then
                sLine = s
            Else
                sLine = Left$(s, nRightCrlf - 2)
            End If
        Else    '光标不在第一行
            '光标在最后一行
            If nRightCrlf = 0 Then
                sLine = Mid$(s, nLeftCrlf)
            Else    '光标在中间
                sLine = Mid$(s, nLeftCrlf, nRightCrlf - nLeftCrlf - 2)
            End If
        End If

        Debug.Assert False
        Debug.Print "|" & sLine & "|"

        '仅在文件名中查找
        If chkFindFilename.Value = 1 Then
            nLastSlash = InStrRev(sLine, "\")
        Else
            
        End If

    End Select
End Sub

Private Sub TabStrip1_Click()
    Dim i As Long

    For i = 0 To pnl.UBound
        pnl(i).Visible = (i = TabStrip1.SelectedItem.Index - 1)
    Next
End Sub

Private Sub txtFind_GotFocus()
    Set objLastFocus = txtFind
End Sub


Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim nIndex As Long
    If KeyCode = vbKeyDelete Then
        nIndex = txtFind.ListIndex
        If nIndex > -1 Then
            modComboBoxHelper.CloseDropdownList txtFind
            DeleteFindHistory txtFind.List(nIndex)
            txtFind.RemoveItem nIndex
            txtFind.Refresh
        End If
    End If
End Sub

Private Sub txtNumberIncCount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyExecute Then
        cmdNumberInc_Click
    End If
End Sub

Private Sub txtReplace_GotFocus()
    Set objLastFocus = txtReplace
End Sub

Private Sub txtReplace_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim nIndex As Long
    If KeyCode = vbKeyDelete Then
        nIndex = txtReplace.ListIndex
        If nIndex > -1 Then
            modComboBoxHelper.CloseDropdownList txtReplace
            DeleteReplaceHistory txtReplace.List(nIndex)
            txtReplace.RemoveItem nIndex
            txtReplace.Refresh
        End If
    End If
End Sub
