VERSION 5.00
Begin VB.Form dlgMesurement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "单位换算"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dlgMesurement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdGo 
      Caption         =   "转换"
      Default         =   -1  'True
      Height          =   480
      Left            =   120
      TabIndex        =   14
      Top             =   3300
      Width           =   2790
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   1560
      ScaleHeight     =   2655
      ScaleWidth      =   1275
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   1335
      Begin VB.OptionButton optTo 
         Caption         =   "缇"
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   11
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optTo 
         Caption         =   "英尺"
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   10
         Top             =   1620
         Width           =   1215
      End
      Begin VB.OptionButton optTo 
         Caption         =   "厘米"
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optTo 
         Caption         =   "磅"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optTo 
         Caption         =   "像素"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      ScaleHeight     =   2655
      ScaleWidth      =   1275
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   420
      Width           =   1335
      Begin VB.OptionButton optFrom 
         Caption         =   "缇"
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optFrom 
         Caption         =   "英尺"
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   1620
         Width           =   1215
      End
      Begin VB.OptionButton optFrom 
         Caption         =   "厘米"
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optFrom 
         Caption         =   "磅"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optFrom 
         Caption         =   "像素"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "目标单位"
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "源单位"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "dlgMesurement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetUnitFrom() As String

    Dim i As Long
    
    For i = 0 To optFrom.UBound

        If optFrom(i).Value Then
            GetUnitFrom = optFrom(i).Tag

            Exit Function

        End If

    Next

End Function

Private Function GetUnitTo() As String

    Dim i As Long
    
    For i = 0 To optTo.UBound

        If optTo(i).Value Then
            GetUnitTo = optTo(i).Tag

            Exit Function

        End If

    Next

End Function

Private Sub cmdGo_Click()

    Dim sFrom As String

    Dim sTo   As String
    
    sFrom = GetUnitFrom

    If Len(sFrom) = 0 Then
        MsgBox "请选择源单位", vbExclamation

        Exit Sub

    End If
    
    sTo = GetUnitTo
    
    If Len(sTo) = 0 Then
        MsgBox "请选择目标单位", vbExclamation

        Exit Sub

    End If
    
    If sFrom = sTo Then
        MsgBox "待转换的单位必须不同", vbExclamation

        Exit Sub

    End If
    
    MesurementConv sFrom, sTo
End Sub

Private Sub Form_Load()

    Dim arrUnit(0 To 4) As String

    Dim i               As Long
    
    arrUnit(0) = "Pixel"
    arrUnit(1) = "Pound"
    arrUnit(2) = "Centimeter"
    arrUnit(3) = "Inch"
    arrUnit(4) = "Twip"
    
    For i = 0 To UBound(arrUnit)
        optFrom(i).Tag = arrUnit(i)
        optTo(i).Tag = arrUnit(i)
    Next
    
End Sub

Private Sub MesurementConv(ByVal UnitFrom As String, ByVal UnitTo As String)

    Dim v   As Double

    Dim f   As Long

    Dim fnc As String

    Dim mes As Mesurement
    
    Set mes = New Mesurement
    fnc = UnitFrom & "To" & UnitTo
     
    v = Val(frmMain!txtExpression.Text)

    On Error GoTo H

    f = CallByName(mes, fnc, VbMethod, v)
    frmMain.SetResult f
    Set mes = Nothing

    Exit Sub

H:
    frmMain.SetResult Err.Description & vbCrLf & fnc
End Sub
