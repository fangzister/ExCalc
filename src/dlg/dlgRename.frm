VERSION 5.00
Begin VB.Form dlgRename 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "解决重命名冲突"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7095
   ControlBox      =   0   'False
   Icon            =   "dlgRename.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtDest 
      Height          =   330
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   5940
   End
   Begin VB.TextBox txtSource 
      Height          =   330
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   5940
   End
   Begin VB.CommandButton cmdDelSource 
      Caption         =   "保留已有文件，删除原文件(&R)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   300
      TabIndex        =   3
      Top             =   2340
      Width           =   4635
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "什么都不做(&N)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   3000
      Width           =   1995
   End
   Begin VB.CommandButton cmdOverwrite 
      Caption         =   "用原文件覆盖已有文件(&O)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   300
      TabIndex        =   2
      Top             =   1740
      Width           =   4635
   End
   Begin VB.CheckBox chkSameAs 
      Caption         =   "之后x个冲突执行此操作(&D)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00993300&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   4275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已存在"
      Height          =   180
      Left            =   300
      TabIndex        =   8
      Top             =   1260
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原文件"
      Height          =   180
      Left            =   300
      TabIndex        =   7
      Top             =   780
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "此位置已经包含同名文件，请选择如何进行下一步操作。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00993300&
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   120
      Width           =   6000
   End
End
Attribute VB_Name = "dlgRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum ConflictEnum

    OverwriteSource = 1&
    Ignore = 2&
    DelSource = 3&

End Enum

Dim mOperation  As ConflictEnum

Dim mDealAsSame As Boolean

Public Sub Conflict(ByVal Source As String, _
                    ByVal Dest As String, _
                    ByVal FileCount As Long, _
                    ByRef Operation As ConflictEnum, _
                    ByRef DealAsSame As Boolean)
    txtSource.Text = Source
    txtDest.Text = Dest

    If FileCount = 0 Then   '没有更多冲突了
        chkSameAs.Visible = False
        chkSameAs.Value = 0
    Else
        chkSameAs.Caption = Replace(chkSameAs.Caption, "x", FileCount)
    End If
    
    Me.Show vbModal, frmMain
    Operation = mOperation
    DealAsSame = mDealAsSame
End Sub

Private Sub cmdDelSource_Click()
    mDealAsSame = (chkSameAs.Value = 1)
    mOperation = DelSource
    Unload Me
End Sub

Private Sub cmdIgnore_Click()
    mDealAsSame = (chkSameAs.Value = 1)
    mOperation = Ignore
    Unload Me
End Sub

Private Sub cmdOverwrite_Click()
    mDealAsSame = (chkSameAs.Value = 1)
    mOperation = OverwriteSource
    Unload Me
End Sub

