VERSION 5.00
Begin VB.Form dlgRedix 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "进制转换"
   ClientHeight    =   2700
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3030
   Icon            =   "dlgRedix.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   1275
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Width           =   1335
      Begin VB.OptionButton optFrom 
         Caption         =   "十进制"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton optFrom 
         Caption         =   "十六进制"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optFrom 
         Caption         =   "二进制"
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optFrom 
         Caption         =   "八进制"
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   8
         Top             =   1620
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   1560
      ScaleHeight     =   2115
      ScaleWidth      =   1275
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   1335
      Begin VB.OptionButton optTo 
         Caption         =   "十进制"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton optTo 
         Caption         =   "十六进制"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optTo 
         Caption         =   "二进制"
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optTo 
         Caption         =   "八进制"
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   9
         Top             =   1620
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "源进制"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "目标进制"
      Height          =   180
      Left            =   1560
      TabIndex        =   10
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "dlgRedix"
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

Private Sub GoTrans()

    Dim sFrom As String

    Dim sTo   As String
    
    sFrom = GetUnitFrom

    If Len(sFrom) = 0 Then

        Exit Sub

    End If
    
    sTo = GetUnitTo
    
    If Len(sTo) = 0 Then

        Exit Sub

    End If
    
    If sFrom = sTo Then
        frmMain.SetResult frmMain.txtExpression.Text

        Exit Sub

    End If
    
    MesurementConv sFrom, sTo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    Dim arrUnit(0 To 3) As String

    Dim i               As Long
    
    arrUnit(0) = "Dec"
    arrUnit(1) = "Hex"
    arrUnit(2) = "Bin"
    arrUnit(3) = "Oct"
    
    For i = 0 To UBound(arrUnit)
        optFrom(i).Tag = arrUnit(i)
        optTo(i).Tag = arrUnit(i)
    Next
    
End Sub

Private Sub MesurementConv(ByVal UnitFrom As String, ByVal UnitTo As String)

    Dim v    As String

    Dim f    As String

    Dim fnc  As String

    Dim rx   As Redix

    Dim p()  As String

    Dim i    As Long

    Dim k    As Long

    Dim row  As String

    Dim ns() As String
    
    v = frmMain!txtExpression.Text
    p = Split(v, vbCrLf)
    
    Set rx = New Redix
    fnc = UnitFrom & "2" & UnitTo
    
    For i = 0 To UBound(p)

        On Error Resume Next

        row = p(i)
        ns = Split(p(i), " ")

        For k = 0 To UBound(ns)
            ns(k) = CallByName(rx, fnc, VbMethod, ns(k))

            If Err Then

                Resume Next

            End If

        Next

        p(i) = Join(ns, " ")
    Next

    frmMain.SetResult Join(p, vbCrLf)
    Set rx = Nothing
End Sub

Private Sub optFrom_Click(Index As Integer)
    GoTrans
End Sub

Private Sub optTo_Click(Index As Integer)
    GoTrans
End Sub
