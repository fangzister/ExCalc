VERSION 5.00
Begin VB.Form dlgDateAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "时间计算"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dlgDateAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   300
      Width           =   780
   End
   Begin VB.ComboBox cboFormat 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Text            =   "cboFormat"
      Top             =   780
      Width           =   2475
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "减"
      Height          =   360
      Index           =   1
      Left            =   2340
      TabIndex        =   4
      Top             =   1260
      Width           =   1200
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "加"
      Height          =   360
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   1260
      Width           =   1200
   End
   Begin VB.ComboBox cboInterval 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输出格式"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "时间单位"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "dlgDateAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click(Index As Integer)

    Dim s   As String

    Dim p() As String

    Dim R() As String

    Dim i   As Long

    Dim u   As Long

    Dim d   As String

    Dim v   As Long

    Dim n   As String

    Dim f   As String
    
    s = Trim$(frmMain!txtExpression.Text)
    p = Split(s, vbCrLf)
    n = GetSelectedDateTimeInterval
    f = cboFormat.Text
    v = Val(txtValue.Text)
    
    u = UBound(p)
    ReDim R(0 To u) As String
    
    If Index = 1 Then
        v = -v
    End If
    
    For i = 0 To u
        d = Trim$(p(i))

        If IsDate(d) Then
            R(i) = Format(DateAdd(n, v, d), f)
        Else
            R(i) = p(i)
        End If

    Next
    
    frmMain.SetResult Join(R, vbCrLf)
End Sub

Private Sub Form_Load()

    With cboInterval
        .AddItem "年"
        .AddItem "季"
        .AddItem "月"
        .AddItem "一年的日数"
        .AddItem "日"
        .AddItem "一周的日数"
        .AddItem "周"
        .AddItem "时"
        .AddItem "分钟"
        .AddItem "秒"
        
        .ListIndex = 4
    End With
        
    With cboFormat
        .AddItem "yyyy-MM-dd"
        .AddItem "yyyy-MM-dd HH:mm:ss"
        .AddItem "yyyyMMddHHmmss"
        .AddItem "yyyy年M月d日"
        .AddItem "yyyy年M月d日 HH:mm:ss"
        .AddItem "yyyy年M月d日 H时m分s秒"
        .AddItem "yyyyMMdd_HHmmss"
        .ListIndex = 0
    End With

End Sub

Private Function GetSelectedDateTimeInterval() As String

    Select Case cboInterval.ListIndex

    Case 0: GetSelectedDateTimeInterval = "yyyy"

    Case 1: GetSelectedDateTimeInterval = "q"

    Case 2: GetSelectedDateTimeInterval = "m"

    Case 3: GetSelectedDateTimeInterval = "y"

    Case 4: GetSelectedDateTimeInterval = "d"

    Case 5: GetSelectedDateTimeInterval = "w"

    Case 6: GetSelectedDateTimeInterval = "ww"

    Case 7: GetSelectedDateTimeInterval = "h"

    Case 8: GetSelectedDateTimeInterval = "n"

    Case 9: GetSelectedDateTimeInterval = "s"
    End Select

End Function
