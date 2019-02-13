VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplitPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------
'上下分割面板类

'1.放置所需控件
'用PictureBox做Root，或者直接用Form
'在Root中放两个将被分割的控件,panel1,panel2
'在root中放一个PictureBox,bar作为分割条
'在root中放一个PictureBox,shadow作为拖动分割条的阴影

'2.代码
'在Form中定义
'dim sp as SplitPane
'
'在Form_Load中写
'set sp = new SplitPane
'sp.SetSplitPane root, panel1, panel2, bar, shadow, 3000, 3000
'
'在Form_Unload中写
'Set sp = Nothing

'在root的parent的Resize事件中写
'sp.DoLayout
'
'在splitBar的Mouse事件中写

Option Explicit

Dim mbMoving        As Boolean
Dim root            As Object
Dim panel1          As Object
Dim panel2          As Object
Dim shadow          As PictureBox
Dim mMin1           As Single
Dim mMin2           As Single
Dim mVerticalLayout As Boolean
Dim mSplitPos       As Single
Dim mBarSize        As Single

Private WithEvents bar  As PictureBox
Attribute bar.VB_VarHelpID = -1

Public Property Get SplitPos() As Single
    SplitPos = mSplitPos
End Property

Public Property Let SplitPos(ByVal RHS As Single)
    mSplitPos = RHS
    Me.DoLayout mSplitPos
End Property

Public Property Get VerticalLayout() As Boolean
    VerticalLayout = mVerticalLayout
End Property

Public Property Let VerticalLayout(ByVal RHS As Boolean)
    If mVerticalLayout = RHS Then
        Exit Property
    End If
    
    mVerticalLayout = RHS

    If mSplitPos = 0 Then
        DoLayout 0
        Exit Property
    End If
    
    If RHS Then
        bar.MousePointer = MousePointerConstants.vbSizeNS
        mSplitPos = mSplitPos / root.ScaleWidth
        bar.Move 0, (root.ScaleHeight - mBarSize) * mSplitPos, root.ScaleWidth, mBarSize

        DoLayoutVertical mSplitPos
    Else
        bar.MousePointer = MousePointerConstants.vbSizeWE
        mSplitPos = mSplitPos / root.ScaleHeight
        bar.Move (root.ScaleWidth - mBarSize) * mSplitPos, 0, mBarSize, root.ScaleHeight
        
        DoLayoutHorizontal mSplitPos
    End If
End Property

Public Sub StartDrag()
    With bar
        shadow.Move .Left, .Top, .Width, .Height
    End With

    shadow.Visible = True
    mbMoving = True
End Sub

Public Sub OnDrag(ByVal p As Single)
    Dim sPos As Single
    
    If mbMoving Then
        If mVerticalLayout Then
            sPos = p + bar.Top

            If sPos < mMin1 Then
                shadow.Move 0, mMin1
            ElseIf sPos > root.ScaleHeight - mMin2 Then
                shadow.Move 0, root.ScaleHeight - mMin2
            Else
                shadow.Move 0, sPos
            End If
        Else
            sPos = p + bar.Left

            If sPos < mMin1 Then
                shadow.Move mMin1
            ElseIf sPos > root.ScaleWidth - mMin2 Then
                shadow.Move root.ScaleWidth - mMin2
            Else
                shadow.Move sPos
            End If
        End If
    End If
End Sub

Public Sub Drop()
    If mVerticalLayout Then
        DoLayoutVertical shadow.Top
    Else
        DoLayoutHorizontal shadow.Left
    End If

    shadow.Visible = False
    mbMoving = False
End Sub

Public Sub Init(rootControl As Object, _
                topControl As Object, _
                bottomControl As Object, _
                imgBar As PictureBox, _
                picShadow As PictureBox, _
                Optional SplitterPos As Single = 0.5, _
                Optional Min1 As Single = 1500, _
                Optional Min2 As Single = 1500, _
                Optional BarSize As Single = 60)
    Set root = rootControl
    Set panel1 = topControl
    Set panel2 = bottomControl
    Set bar = imgBar
    Set shadow = picShadow
    
    mBarSize = BarSize
        
    With shadow
        .BackColor = &H808080
        .BorderStyle = 0
        .Visible = False
        .ZOrder 0
    End With
    
    With bar
        .BorderStyle = 0
        .BackColor = root.BackColor
        .ZOrder 0
        
        If mVerticalLayout Then
            .MousePointer = MousePointerConstants.vbSizeNS

            If SplitterPos < 1 Then
                .Move 0, (root.ScaleHeight - BarSize) * SplitterPos, root.ScaleWidth, BarSize
            Else
                .Move 0, SplitterPos, root.ScaleWidth, BarSize
            End If
        Else
            .MousePointer = MousePointerConstants.vbSizeWE

            If SplitterPos < 1 Then
                .Move (root.ScaleWidth - BarSize) * SplitterPos, 0, BarSize, root.ScaleHeight
            Else
                .Move SplitterPos, 0, BarSize, root.ScaleHeight
            End If
        End If
    End With
    
    mMin1 = Min1
    mMin2 = Min2
End Sub

Private Sub DoLayoutVertical(ByVal Y As Single)
    On Error Resume Next

    If Y = 0 Then
        Y = bar.Top
    ElseIf Y < 1 Then
        Y = (root.ScaleHeight - bar.Height) * 0.5
    End If
    
    If Y < mMin1 Then Y = mMin1
    If Y > (root.ScaleHeight - mMin2) Then Y = root.ScaleHeight - mMin2
    
    panel1.Move 0, 0, root.ScaleWidth, Y
    panel2.Move 0, Y + bar.Height, root.ScaleWidth, root.ScaleHeight - Y - bar.Height
    bar.Move 0, Y, root.ScaleWidth
    mSplitPos = Y
End Sub

Private Sub DoLayoutHorizontal(ByVal X As Single)
    On Error Resume Next

    If X = 0 Then
        X = bar.Left
    ElseIf X < 1 Then
        X = (root.ScaleWidth - bar.Width) * 0.5
    End If
    
    If X < mMin1 Then X = mMin1
    If X > (root.ScaleWidth - mMin2) Then X = root.ScaleWidth - mMin2
    
    panel1.Move 0, 0, X, root.ScaleHeight
    panel2.Move X + bar.Width, 0, root.ScaleWidth - X - bar.Width, root.ScaleHeight
    bar.Move X, 0, bar.Width, root.ScaleHeight
    mSplitPos = X
End Sub

Public Sub DoLayout(Optional ByVal p As Single)
    If mVerticalLayout Then
        DoLayoutVertical p
    Else
        DoLayoutHorizontal p
    End If
End Sub

Private Sub bar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        StartDrag
    End If
End Sub

Private Sub bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If mVerticalLayout Then
            OnDrag Y
        Else
            OnDrag X
        End If
    End If
End Sub

Private Sub bar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Drop
    End If
End Sub

Private Sub Class_Initialize()
    mMin1 = 1500
    mMin2 = 1500
End Sub

Private Sub Class_Terminate()
    Set root = Nothing
    Set panel1 = Nothing
    Set panel2 = Nothing
    Set bar = Nothing
    Set shadow = Nothing
End Sub
