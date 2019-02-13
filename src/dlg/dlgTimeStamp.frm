VERSION 5.00
Begin VB.Form dlgTimeStamp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "时间戳转换"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dlgTimeStamp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdToTimeStamp 
      Caption         =   "输出时间戳（毫秒）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   1740
      Width           =   1920
   End
   Begin VB.CommandButton cmdToDateTime 
      Caption         =   "输出日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   1920
   End
   Begin VB.TextBox txtStart 
      Height          =   375
      Left            =   1020
      TabIndex        =   0
      Top             =   255
      Width           =   3075
   End
   Begin VB.CommandButton cmdToTimeStamp 
      Caption         =   "输出时间戳（秒）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1920
   End
   Begin VB.ComboBox cboFormat 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   3075
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "初始日期"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Top             =   300
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输出格式"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   825
      Width           =   780
   End
End
Attribute VB_Name = "dlgTimeStamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdToTimeStamp_Click(Index As Integer)

    Dim p() As String

    Dim i   As Long

    Dim s   As String

    Dim R() As String

    Dim v   As String

    Dim d   As String

    Dim bMs As Boolean
    
    v = frmMain!txtExpression.Text
    
    If Len(v) = 0 Then
        MsgBox "请输入时间值，支持多行", vbExclamation

        Exit Sub

    End If
    
    d = txtStart.Text

    If IsDate(d) = False Then
        MsgBox "初始日期格式错误", vbExclamation

        Exit Sub

    End If
    
    bMs = (Index = 1)
    
    p = Split(v, vbCrLf)
    ReDim R(0 To UBound(p)) As String
    
    For i = 0 To UBound(p)
        s = Trim$(p(i))
        R(i) = modDateTime.DateToTimeStamp(d, cboFormat.Text, s, bMs)
    Next

    frmMain.SetResult Join(R, vbCrLf)
End Sub

Private Sub cmdToDateTime_Click()

    Dim s         As String

    Dim p()       As String

    Dim R()       As String

    Dim i         As Long

    Dim u         As Long

    Dim v         As Variant

    Dim dateStart As String

    Dim DateValue As String

    Dim reg       As RegExp

    Dim mc        As MatchCollection

    Dim m         As Match

    Dim fmt       As String
    
    s = frmMain!txtExpression.Text
    p = Split(s, vbCrLf)
    
    u = UBound(p)

    If u < 0 Then Exit Sub
    
    dateStart = txtStart.Text

    If IsDate(dateStart) = False Then
        MsgBox "初始日期格式错误", vbExclamation

        Exit Sub

    End If
    
    fmt = cboFormat.Text
    
    ReDim R(0 To u) As String
    
    Set reg = New RegExp
    
    For i = 0 To u
        s = p(i)
        reg.Pattern = "\d{10,13}"
        reg.MultiLine = False
        reg.Global = False
        
        If reg.Test(s) Then
            Set mc = reg.Execute(s)
            Set m = mc.Item(0)
            v = m.Value

            If Len(v) = 10 Then
                v = v & "000"
            ElseIf Len(v) <> 13 Then
                R(i) = s
            End If
            
            If Len(R(i)) = 0 Then
                DateValue = modDateTime.TimeStampToDate(v, fmt, dateStart)
                R(i) = reg.Replace(s, DateValue)
            End If

        Else
            R(i) = s
        End If
        
    Next
    
    frmMain.SetResult Join(R, vbCrLf)
End Sub

Private Sub Form_Load()

    With cboFormat
        .AddItem "yyyy-MM-dd HH:mm:ss"
        .AddItem "yyyy-MM-dd"
        .AddItem "yyyy-M-d"
        .AddItem "yyyyMMdd"
        .AddItem "yyyy年M月d日 H:m:s"
        .AddItem "yyyy年M月d日 H时m分s秒"
        .AddItem "yyyyMMdd_HHmmss"
        .AddItem "yyyyMMdd_HHmm"
        .AddItem "yyyyMMddHHmmss"
        .AddItem "yyyyMMddHHmm"
        .ListIndex = 0
    End With
    
    txtStart.Text = "1970-01-01"
End Sub
