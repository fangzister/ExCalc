VERSION 5.00
Begin VB.Form dlgPropertyShow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "显示文件属性"
   ClientHeight    =   4530
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check8 
      Caption         =   "MD5值"
      Height          =   315
      Left            =   180
      TabIndex        =   10
      Top             =   3300
      Width           =   1215
   End
   Begin VB.CheckBox Check7 
      Caption         =   "MD5值"
      Height          =   315
      Left            =   180
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CheckBox Check6 
      Caption         =   "文件大小"
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      Caption         =   "访问时间"
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "修改时间"
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "创建时间"
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "文件名"
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   780
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   232
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Caption         =   "序号"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "dlgPropertyShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()

End Sub
