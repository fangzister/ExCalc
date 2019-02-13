VERSION 5.00
Begin VB.Form frmReleaseHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReleaseHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   493
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtContent 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   1230
   End
End
Attribute VB_Name = "frmReleaseHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetReleaseHistory() As String
    GetReleaseHistory = modStrings.LoadText(App.Path & "\releasehistory.txt")
End Function

Private Sub Form_Load()
    Me.Caption = App.title & " 更新历史"
    txtContent.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    txtContent.Text = GetReleaseHistory
End Sub
