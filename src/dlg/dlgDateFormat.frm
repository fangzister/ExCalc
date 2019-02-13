VERSION 5.00
Begin VB.Form dlgDateFormat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ڸ�ʽ��"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dlgDateFormat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   333
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDelFormat 
      Cancel          =   -1  'True
      Caption         =   "ɾ��(&D)"
      Height          =   360
      Left            =   1740
      TabIndex        =   5
      Top             =   5460
      Width           =   1080
   End
   Begin VB.CommandButton cmdAddFormat 
      Caption         =   "��Ӹ�ʽ(&A)..."
      Height          =   360
      Left            =   180
      TabIndex        =   4
      Top             =   5460
      Width           =   1500
   End
   Begin VB.ListBox lstFormat 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   180
      TabIndex        =   3
      Top             =   480
      Width           =   4635
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�ر�(&C)"
      Height          =   360
      Left            =   3780
      TabIndex        =   2
      Top             =   5460
      Width           =   1080
   End
   Begin VB.CommandButton cmdFormatDate 
      Caption         =   "ִ��"
      Height          =   465
      Left            =   5340
      TabIndex        =   0
      Top             =   0
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ѡ�������ʽ"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "dlgDateFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const INI_SECTION As String = "DateTimeFormats"

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdAddFormat_Click()
    Dim s As String
    Dim ini As INIProfile
    Dim p() As String
    Dim LastKey As String
    Dim nextKey As String
    
    s = InputBox("�����ʽ", "��Ӹ�ʽ")
    
    If Len(s) > 0 Then
        Set ini = GetIniProfile
        nextKey = ini.GetNextKey(INI_SECTION)
        
        ini.SetString INI_SECTION, nextKey, s
        lstFormat.AddItem s
    End If
End Sub

Private Sub cmdDelFormat_Click()
    MsgBox "TODO"
End Sub

Private Sub lstFormat_Click()
    Dim t1   As String
    Dim ts() As String
    Dim i    As Long
    Dim fmt  As String
    Dim ret  As String
    
    t1 = frmMain!txtExpression.Text
    fmt = lstFormat.List(lstFormat.ListIndex)
    t1 = Replace$(t1, ".", "-")
    ts = Split(t1, vbCrLf)
    
    For i = 0 To UBound(ts)
        If IsDate(ts(i)) Then
            ts(i) = Format$(ts(i), fmt)
        End If
    Next
    
    ret = Join(ts, vbCrLf)
    frmMain.SetResult ret
End Sub

Private Sub Form_Load()
    Dim ini As INIProfile
    Dim p() As String
    Dim i As Long
    Dim idx As Long
    
    Set ini = GetIniProfile
    
    p = ini.GetAllKeys("DateTimeFormats")
    idx = ini.GetLong("DateTime", "LastIndex")
    
        
    For i = 0 To UBound(p)
        lstFormat.AddItem ini.GetString("DateTimeFormats", p(i))
    Next
    
    Set ini = Nothing
End Sub
