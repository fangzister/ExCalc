VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form dlgBatchImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "批处理图片"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   5460
      TabIndex        =   17
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox oPic 
      Height          =   195
      Left            =   6000
      ScaleHeight     =   135
      ScaleWidth      =   195
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1080
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3855
      Begin VB.OptionButton optImageType 
         Caption         =   "BMP"
         Height          =   375
         Index           =   3
         Left            =   2880
         TabIndex        =   3
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton optImageType 
         Caption         =   "GIF"
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   2
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton optImageType 
         Caption         =   "PNG"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton optImageType 
         Caption         =   "JPG"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSComctlLib.Slider sldQuality 
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   780
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      TickStyle       =   2
   End
   Begin VB.TextBox txtSuffix 
      Height          =   330
      Left            =   1860
      TabIndex        =   15
      Top             =   2880
      Width           =   2100
   End
   Begin VB.TextBox txtPrefix 
      Height          =   330
      Left            =   1860
      TabIndex        =   12
      Top             =   2460
      Width           =   2100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CheckBox chkMakeCopy 
      Caption         =   "生成副本"
      Height          =   375
      Left            =   1020
      TabIndex        =   11
      Top             =   1980
      Width           =   1215
   End
   Begin MSComctlLib.Slider sldZoomRate 
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   180
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      TickStyle       =   2
      TextPosition    =   1
   End
   Begin VB.Label lblQuality 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   180
      Left            =   4200
      TabIndex        =   19
      Top             =   937
      Width           =   540
   End
   Begin VB.Label lblZoomRate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   180
      Left            =   4200
      TabIndex        =   18
      Top             =   337
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输出格式"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   1537
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图像质量"
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   930
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称后缀"
      Height          =   180
      Left            =   1020
      TabIndex        =   14
      Top             =   2955
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称前缀"
      Height          =   180
      Left            =   1020
      TabIndex        =   13
      Top             =   2535
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缩放比例"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   337
      Width           =   720
   End
End
Attribute VB_Name = "dlgBatchImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMakeCopy_Click()

    Dim b As Boolean
    
    b = (chkMakeCopy.Value = 1)
    txtPrefix.Enabled = b
    txtSuffix.Enabled = b
    
    If b Then
        txtPrefix.SetFocus
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim oGDI         As GDI
    Dim p()          As String
    Dim i            As Long
    Dim fso          As FileSystemObject
    Dim sSourceImg   As String
    Dim iZoomRate    As Single
    Dim it           As ImageFileFormat
    Dim iff(0 To 3)  As Long
    Dim exts(0 To 3) As String
    Dim jq           As Long
    Dim sOutFile     As String
    Dim pf           As String
    Dim sf           As String
    Dim ext          As String
    
    p = Split(frmMain!txtExpression.Text, vbCrLf)
    
    iZoomRate = sldZoomRate.Value / 10
    
    iff(0) = ImageFileFormat.Jpg
    iff(1) = ImageFileFormat.Png
    iff(2) = ImageFileFormat.Gif
    iff(3) = ImageFileFormat.Bmp
    exts(0) = "jpg"
    exts(1) = "png"
    exts(2) = "gif"
    exts(3) = "bmp"
    
    For i = 0 To optImageType.UBound
        If optImageType(i).Value Then
            it = iff(i)
            ext = exts(i)

            Exit For
        End If
    Next
    
    If it = ImageFileFormat.Jpg Then
        jq = sldQuality.Value * 10
    End If
    
    Set oGDI = New GDI
    Set fso = New FileSystemObject
    
    pf = txtPrefix.Text
    sf = txtSuffix.Text
    
    frmMain!pb.Max = UBound(p)
    
    For i = 0 To UBound(p)
        frmMain!pb.Value = i
        sSourceImg = p(i)
        
        If fso.FileExists(sSourceImg) Then
            If (chkMakeCopy.Value = 1) Then
                sOutFile = fso.GetParentFolderName(sSourceImg) & "\" & pf & fso.GetBaseName(sSourceImg) & sf & "." & ext
            Else
                sOutFile = sSourceImg
            End If
            
            If oGDI.ScaleImageAsFile(oPic, sSourceImg, iZoomRate, sOutFile, it, jq) Then
                p(i) = sOutFile
            Else
                p(i) = "【处理失败】" & sSourceImg
            End If
        Else
            p(i) = "【文件不存在】" & sSourceImg
        End If
    Next
    
    frmMain.SetResult Join(p, vbCrLf)
    
    Set oGDI = Nothing
    Set fso = Nothing
End Sub

Private Sub Form_Load()
    sldZoomRate.Value = 5
    sldQuality.Value = 8
    txtSuffix.Text = "_zoom"
    
    lblZoomRate.Caption = sldZoomRate.Value * 10 & "%"
    lblQuality.Caption = sldQuality.Value * 10
End Sub

Private Sub sldQuality_Change()
    lblQuality.Caption = sldQuality.Value * 10
End Sub

Private Sub sldZoomRate_Change()
    lblZoomRate.Caption = sldZoomRate.Value * 10 & "%"
End Sub

