VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form dlgStringGenerator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "序列生成器"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   627
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   621
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5835
      Left            =   540
      TabIndex        =   15
      Top             =   2700
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   10292
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "字典"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "日期序列"
      Height          =   915
      Left            =   120
      TabIndex        =   10
      Top             =   1140
      Width           =   5175
      Begin VB.TextBox txtDateStart 
         Height          =   330
         Left            =   600
         TabIndex        =   4
         Top             =   345
         Width           =   1200
      End
      Begin VB.TextBox txtDateEnd 
         Height          =   330
         Left            =   2520
         TabIndex        =   5
         Top             =   345
         Width           =   1200
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "生成"
         Height          =   420
         Left            =   4080
         TabIndex        =   6
         Top             =   300
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "从"
         Height          =   180
         Left            =   300
         TabIndex        =   12
         Top             =   420
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到"
         Height          =   180
         Left            =   2220
         TabIndex        =   11
         Top             =   420
         Width           =   180
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数字序列"
      Height          =   915
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtNumStep 
         Height          =   330
         Left            =   3180
         TabIndex        =   13
         Text            =   "1"
         Top             =   345
         Width           =   420
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "生成"
         Height          =   420
         Left            =   5280
         TabIndex        =   3
         Top             =   300
         Width           =   930
      End
      Begin VB.CheckBox chkNumZeroize 
         Caption         =   "补零"
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtNumEnd 
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         Top             =   345
         Width           =   720
      End
      Begin VB.TextBox txtNumStart 
         Height          =   330
         Left            =   600
         TabIndex        =   0
         Top             =   345
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "步长"
         Height          =   180
         Index           =   1
         Left            =   2760
         TabIndex        =   14
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到"
         Height          =   180
         Index           =   0
         Left            =   1500
         TabIndex        =   9
         Top             =   420
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "从"
         Height          =   180
         Left            =   300
         TabIndex        =   8
         Top             =   420
         Width           =   180
      End
   End
End
Attribute VB_Name = "dlgStringGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CheckNumber(ByVal Msg As String, Text As TextBox) As Boolean
    If IsNumeric(Text.Text) Then
        CheckNumber = True
        Exit Function
    End If
    
    MsgBox Msg, vbCritical, "提示"
    Text.SelStart = 0
    Text.SelLength = Len(Text.Text)
    Text.SetFocus
End Function

Private Sub cmdNumber_Click()
    Dim i As Long
    Dim j As Long
    Dim nStart As Long
    Dim nEnd As Long
    Dim nMax As Long
    Dim nStep As Long
    Dim nZero As Long
    Dim sFmt As String
    Dim p() As String
    
    If Not CheckNumber("请输入起始数值", txtNumStart) Then
        Exit Sub
    End If
    
    If Not CheckNumber("请输入结束数值", txtNumEnd) Then
        Exit Sub
    End If
    
    If Not CheckNumber("请输入步长值", txtNumStep) Then
        Exit Sub
    End If
    
    nStart = CLng(txtNumStart.Text)
    nEnd = CLng(txtNumEnd.Text)
    nStep = CLng(txtNumStep.Text)
    
    If nStart = nEnd Then
        frmMain!txtExpression.Text = nStart
        Exit Sub
    End If
    
    If nStart < nEnd Then
        ReDim p(0 To nEnd - nStart) As String
    Else
        ReDim p(0 To nStart - nEnd) As String
        nStep = -nStep
    End If
    
    If chkNumZeroize.Value = 1 Then
        nMax = Max(nStart, nEnd)
        
        If nMax > 10 Then
            nZero = Len(CStr(nMax))
            sFmt = String(nZero, "0")
        End If
    
        For i = nStart To nEnd Step nStep
            p(j) = Format(i, sFmt)
            j = j + 1
        Next
    Else
        For i = nStart To nEnd Step nStep
            p(j) = i
            j = j + 1
        Next
    End If
    
    If Abs(nStep) <> 1 Then
        ReDim Preserve p(0 To j - 1) As String
    End If
    frmMain!txtExpression.Text = Join(p, vbCrLf)
End Sub

