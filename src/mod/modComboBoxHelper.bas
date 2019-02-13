Attribute VB_Name = "modComboBoxHelper"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Const CB_GETCURSEL = &H147
Private Const CB_SETCURSEL = &H14E
Private Const CB_FINDSTRING = &H14C
Private Const LB_FINDSTRING = &H18F
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_SHOWDROPDOWN  As Long = &H14F
Private Const CB_ERR = -1

'用记录集填充
Public Sub FillComboByRecordset(rs As Object, cbo As Object, Optional SelectedIndex As Long = 0)
    cbo.Clear

    If rs Is Nothing Then Exit Sub

    While Not rs.EOF
        cbo.AddItem rs(0)
        rs.MoveNext
    Wend

    If cbo.ListCount = 0 Then Exit Sub

    On Error Resume Next

    cbo.ListIndex = SelectedIndex
    rs.Close
    Set rs = Nothing
End Sub

'函数:
'取得 Combo 下拉的宽度
'可以利用该函数比例放大或缩小宽度
'单位为 pixels
Public Function GetDropdownWidth(cboHwnd As Long) As Long
    Dim lRetVal As Long
    
    lRetVal = SendMessage(cboHwnd, CB_GETDROPPEDWIDTH, 0, 0)
    
    If lRetVal <> CB_ERR Then
        GetDropdownWidth = lRetVal
    Else
        GetDropdownWidth = 0
    End If
End Function

'设置列表框的宽度,单位为 pixels
Public Function SetDropdownWidth(cboHwnd As Long, NewWidthPixel As Long) As Boolean
    Dim lRetVal As Long
    
    lRetVal = SendMessage(cboHwnd, CB_SETDROPPEDWIDTH, NewWidthPixel, 0)
    
    If lRetVal <> CB_ERR Then
        SetDropdownWidth = True
    Else
        SetDropdownWidth = False
    End If
End Function

'设置列表框的高度,单位为 pixels
Public Sub SetComboHeight(oComboBox As Object, lNewHeight As Long)
    MoveWindow oComboBox.hwnd, oComboBox.Left, oComboBox.Top, oComboBox.Width, lNewHeight, 1
End Sub

'设置Combobox选定项
Public Function SetComboIndex(cbo As Object, Optional ByVal NewIndex As Long) As Long
    Call SendMessage(cbo.hwnd, CB_SETCURSEL, NewIndex, 0&)
    SetComboIndex = SendMessage(cbo.hwnd, CB_GETCURSEL, NewIndex, 0&)
End Function

'设置当前列表项为指定文本
Public Sub SetListIndexByText(txt As String, cbo As Object, Optional CauseClick As Boolean = False)
    Dim Index As Long
    
    Index = SendMessage(cbo.hwnd, CB_FINDSTRING, -1&, ByVal txt)

    If CauseClick Then
        cbo.ListIndex = Index
    Else
        SetComboIndex cbo, Index
    End If
End Sub

'获取列表文本索引
Public Function GetListIndexByText(txt As String, lst As Object) As Long
    Select Case TypeName(lst)
    Case "ComboBox"
        GetListIndexByText = SendMessage(lst.hwnd, CB_FINDSTRING, -1&, ByVal txt)
    Case "ListBox"
        GetListIndexByText = SendMessage(lst.hwnd, LB_FINDSTRING, -1&, ByVal txt)
    Case Else
        Debug.Print "输入控件类型必须为combobox或listbox"
    End Select
End Function

'用数组填充列表
Public Sub FillArray(vArray As Variant, cbo As Object, Optional DefaultIndex As Long = 0, Optional CauseClick As Boolean = False)
    Dim i As Long
    Dim l As Long
    Dim u As Long
    
    If Not IsArray(vArray) Then
        Err.Raise -1, "fzCore.ComboBoxHelper.FillArray", "参数无效，必须传入包含数组的Variant"
    End If
    
    cbo.Clear
    l = LBound(vArray)
    u = UBound(vArray)

    For i = l To u
        cbo.AddItem vArray(i)
    Next
    
    If DefaultIndex >= 0 Then
        If cbo.ListCount > 0 Then
            If DefaultIndex < cbo.ListCount Then
                If CauseClick Then
                    cbo.ListIndex = DefaultIndex
                Else
                    SetComboIndex cbo, DefaultIndex
                End If
            End If
        End If
    End If
End Sub

Public Sub ShowDrowdownList(cbo As Object)
    SendMessage cbo.hwnd, CB_SHOWDROPDOWN, 1, 0
End Sub

Public Sub CloseDropdownList(cbo As Object)
    SendMessage cbo.hwnd, CB_SHOWDROPDOWN, 0, 0
End Sub
