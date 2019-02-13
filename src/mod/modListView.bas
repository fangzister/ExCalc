Attribute VB_Name = "modListView"

Option Explicit

'****************** 获取 Lvw 当前鼠标下的 列 Index ******************
Private Declare Function ClientToScreen _
                Lib "user32" (ByVal hwnd As Long, _
                              lpPoint As POINTAPI) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function GetScrollPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal nBar As Long) As Long

Private Declare Function SendMessageLong _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lparam As Long) As Long

Private Declare Sub CopyMem _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Private Const SB_HORZ = 0

Private Type POINTAPI

    X As Long
    Y As Long

End Type


Private Type POINT

    X As Long
    Y As Long

End Type

Private Const LVM_FIRST    As Long = &H1000

Private Const LVM_GETITEM  As Long = LVM_FIRST + 5

Private Const LVM_FINDITEM As Long = LVM_FIRST + 13

Private Const LVM_ENSUREVISIBLE = LVM_FIRST + 19

Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30

Private Const LVM_GETTOPINDEX = LVM_FIRST + 39

Private Const LVM_SETITEMSTATE             As Long = LVM_FIRST + 43

Private Const LVM_GETITEMSTATE             As Long = LVM_FIRST + 44

Private Const LVM_GETITEMTEXT              As Long = LVM_FIRST + 45

Private Const LVM_SORTITEMS                As Long = LVM_FIRST + 48

Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 54

Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 55

Private Const LVM_SETCOLUMNORDERARRAY = LVM_FIRST + 58

Private Const LVM_GETCOLUMNORDERARRAY = LVM_FIRST + 59

Private Const LVS_EX_GRIDLINES      As Long = &H1

Private Const LVS_EX_SUBITEMIMAGES  As Long = &H2

Private Const LVS_EX_CHECKBOXES     As Long = &H4

Private Const LVS_EX_TRACKSELECT    As Long = &H8

Private Const LVS_EX_HEADERDRAGDROP As Long = &H10

Private Const LVS_EX_FULLROWSELECT  As Long = &H20

Private Const LVFI_PARAM            As Long = 1

Private Const LVIF_TEXT             As Long = 1

Private Const LVIF_IMAGE            As Long = 2

Private Const LVIF_PARAM            As Long = 4

Private Const LVIF_STATE            As Long = 8

Private Const LVIF_INDENT           As Long = &H10

Private Const LVIF_NORECOMPUTE      As Long = &H800

Private Const LVIS_STATEIMAGEMASK   As Long = &HF000&

Private Const GWL_STYLE = (-16)

Private Const LVM_GETHEADER = (LVM_FIRST + 31)

Private Const HDS_BUTTONS = &H2

Private Type LV_ITEM

    Mask As Long
    Index As Long
    SubItem As Long
    State As Long
    StateMask As Long
    Text As String
    TextMax As Long
    Icon As Long
    Param As Long
    Indent As Long

End Type

Private Type LV_FINDINFO

    Flags As Long
    pSz As String
    lparam As Long
    pt As POINT
    vkDirection As Long

End Type

'--- Array used to speed custom sorts ---'
Private m_lvSortData() As LV_ITEM

Private m_lvSortColl   As Collection

Private m_lvSortColumn As Long

Private m_lvHWnd       As Long

Private m_lvSortType   As LVItemTypes

'--- ListView Set Column Width Messages ---'
Public Enum LVSCW_Styles

    LVSCW_AUTOSIZE = -1
    LVSCW_AUTOSIZE_USEHEADER = -2

End Enum

'LVS_EX_CHECKBOXES: Enables items in a list view control to be displayed
'   as check boxes. This style uses item state images to produce the check
'   box effect.
'LVS_EX_FULLROWSELECT: Specifies that when an item is selected, the item
'   and all its subitems are highlighted. This style is available only in
'   conjunction with the LVS_REPORT style.
'LVS_EX_GRIDLINES: Displays gridlines around items and subitems. This style
'   is available only in conjunction with the LVS_REPORT style.
'LVS_EX_HEADERDRAGDROP: Enables drag-and-drop reordering of columns in a
'   list view control. This style is only available to list view controls
'   that use the LVS_REPORT style.
'LVS_EX_SUBITEMIMAGES: Allows images to be displayed for subitems. This
'   style is available only in conjunction with the LVS_REPORT style.
Public Enum LVStylesEx

    Checkboxes = LVS_EX_CHECKBOXES
    FullRowSelect = LVS_EX_FULLROWSELECT
    GridLines = LVS_EX_GRIDLINES
    HeaderDragDrop = LVS_EX_HEADERDRAGDROP
    SubItemImages = LVS_EX_SUBITEMIMAGES
    TrackSelect = LVS_EX_TRACKSELECT

End Enum

Public Enum LVHeaderStyles

    HeaderFlat = 0
    Header3D = 1

End Enum

'--- Sorting Variables ---'
Public Enum LVItemTypes

    lvDate = 0
    lvNumber = 1
    lvBinary = 2
    lvAlphabetic = 3

End Enum

Public Enum LVSortTypes

    lvAscending = 0
    lvDescending = 1

End Enum

Private SetLvwItemCol As Long '要修改 lvw 的行

'*************************************************************************
'**功能描述：添加 ListView 控件头列表
'**函 数 名：AddLvwHeads
'**输    入：objLvw(ListView) -
'**        ：sHead(String)    -
'**输    出：无
'**例    子：AddLvwItem ListView控件, "序号=500=[关键字]|名称=800=[关键字]"
'**作    者：格式化 QQ:65464145
'**日    期：2009-05-09 13:31:26
'*************************************************************************
Public Sub AddLvwHeads(objLvw As Object, ByVal sHead As String)

    Dim i       As Long

    Dim k       As Long

    Dim str1    As String

    Dim str2    As String

    Dim str3    As String

    Dim pItem() As String

    Dim p()     As String

    pItem = Split(sHead, "|")

    If UBound(pItem) = -1 Then Exit Sub

    With objLvw
        .ColumnHeaders.Clear
        .ListItems.Clear
        .HideColumnHeaders = False
        .View = lvwReport
        .BorderStyle = ccNone
        .GridLines = True
        .FullRowSelect = True

        For i = 0 To UBound(pItem)
            p = Split(pItem(i), "=")
            
            Select Case UBound(p)

            Case 1
                str1 = p(0)
                str2 = p(1)
                str3 = ""
                
            Case 2
                str1 = p(0)
                str2 = p(1)
                str3 = p(2)
                
            Case Else
                str1 = pItem(i)
                str2 = ""
                str3 = ""
            
            End Select
            
            str3 = Trim$(str3)

            If Len(str3) > 0 Then
                .ColumnHeaders.Add , str3, str1
            Else
                .ColumnHeaders.Add , , str1
            End If

            If Len(str2) > 0 Then .ColumnHeaders(i + 1).Width = Val(str2)
        Next

    End With

End Sub

'*************************************************************************
'**功能描述：添加 ListView 项目
'**函 数 名：AddLvwItem
'**输    入：objLvw(ListView)        -
'**        ： Optional sItem(String)  -
'**        ： Optional bBold(Boolean) -
'**        ： Optional lColor(Long)   -
'**输    出：(String) -
'**例    子：AddLvwItem ListView控件, [行首列关键词], ["序号=500|名称=800"], [是否加粗], [颜色]
'**作    者：格式化 QQ:65464145
'**日    期：2009-05-09 13:44:40
'*************************************************************************
Public Function AddLvwItem(ByRef objLvw As Object, _
                           Optional ByVal sItem As String, _
                           Optional ByVal sKey As String, _
                           Optional ByVal bBold As Boolean, _
                           Optional ByVal lColor As Long, _
                           Optional ByVal EnsureVisible As Boolean = True, _
                           Optional ByVal lIconColumnIndex As Variant = "", _
                           Optional ByVal lIcon As Long = 0) As String
                           
    Dim i          As Long, k As String

    Dim colName    As String

    Dim colText    As String

    Dim pItem()    As String

    Dim str1       As String

    Dim objColName As New Collection

    Dim bSorted    As Boolean

    Dim bOK        As Boolean

    With objLvw.ListItems
        bSorted = objLvw.Sorted
        objLvw.Sorted = False
        
        .Add , sKey

        .Item(.Count).Bold = bBold
        .Item(.Count).ForeColor = lColor

        For i = 0 To objLvw.ColumnHeaders.Count - 1

            If i = 0 Then
                If Len(sKey) > 0 Then
                    objColName.Add CStr(i), CStr(sKey)
                End If
                
            Else
                .Item(.Count).ListSubItems.Add , objLvw.ColumnHeaders(i + 1).Key
                .Item(.Count).ListSubItems(i).Bold = bBold
                .Item(.Count).ListSubItems(i).ForeColor = lColor
            End If
            
            objColName.Add CStr(i), CStr(objLvw.ColumnHeaders(i + 1).Text)
        Next

        On Error Resume Next

        pItem = Split(sItem, Chr$(0))

        For i = 0 To UBound(pItem)
            k = InStr(pItem(i), "=")

            If k > 0 Then
                colName = Left$(pItem(i), k - 1)
                colText = Right$(pItem(i), Len(pItem(i)) - k)

                Err.Clear
                bOK = True
                .Item(.Count).ListSubItems(colName).Text = colText '按id来输入失败
                
                If Err.Number = 0 Then
                    .Item(.Count).ListSubItems(colName).ToolTipText = colText
                    
                Else
                    Err.Clear
                    k = objColName(colName)

                    If Err Then
                        Err.Clear
                        bOK = False

                    Else

                        If k = 0 Then
                            .Item(.Count).Text = colText
                            .Item(.Count).ToolTipText = colText

                        Else
                            .Item(.Count).ListSubItems(Val(k)).Text = colText
                            .Item(.Count).ListSubItems(Val(k)).ToolTipText = colText
                        End If
                    End If
                End If
            End If

        Next

        On Error GoTo 0
        
        If lIconColumnIndex <> "" Then
            If lIcon > 0 Then
                .Item(.Count).ListSubItems(lIconColumnIndex).ReportIcon = lIcon
            End If
        End If
        
        If EnsureVisible = True Then .Item(.Count).EnsureVisible
        objLvw.Sorted = bSorted
    End With

    Set objColName = Nothing
End Function

Public Function AddLvwItemEmpty(ByRef objLvw As Object) As Boolean

    Dim i    As Long

    Dim sKey As String

    If objLvw.ColumnHeaders.Count = 0 Then Exit Function

    With objLvw.ListItems
        .Add , , ""
        
        For i = 2 To objLvw.ColumnHeaders.Count
            sKey = objLvw.ColumnHeaders(i).Key

            If Len(sKey) > 0 Then
                .Item(.Count).ListSubItems.Add , sKey, ""
            Else
                .Item(.Count).ListSubItems.Add , , ""
            End If

        Next

    End With

End Function

'------------勾选列表项------------
Public Sub CheckAll(objLvw As Object)

    Dim itm As ListItem
    
    For Each itm In objLvw.ListItems

        itm.Checked = True
    Next

    Set itm = Nothing
End Sub

'勾选指定颜色的项
Public Sub CheckByColor(objLvw As Object, _
                        ByVal Color As OLE_COLOR, _
                        Optional ByVal Index As Variant = "")

    Dim itm As ListItem
    
    If Index = "" Then

        For Each itm In objLvw.ListItems

            itm.Checked = (itm.ForeColor = Color)
        Next

    Else

        For Each itm In objLvw.ListItems

            itm.Checked = (itm.ListSubItems(Index).ForeColor = Color)
        Next

    End If

    Set itm = Nothing
End Sub

'反勾选
Public Sub CheckReverse(objLvw As Object)

    Dim itm As ListItem
    
    For Each itm In objLvw.ListItems

        itm.Checked = Not itm.Checked
    Next

    Set itm = Nothing
End Sub

'获取勾选项的数量
Public Function GetCheckedCount(objLvw As Object) As Long

    Dim c   As Long

    Dim itm As ListItem
    
    For Each itm In objLvw.ListItems

        If itm.Checked Then
            c = c + 1
        End If

    Next

    GetCheckedCount = c
    Set itm = Nothing
End Function

'通过单元格文本获取列表项，如果要通过第一列的值获取Item，则将ColIndex设为空字符串
Public Function GetItemByText(objLvw As Object, _
                              ByVal ColIndex As String, _
                              ByVal Text As String) As ListItem

    Dim itm As ListItem
    
    If ColIndex = "" Then

        For Each itm In objLvw.ListItems

            If itm.Text = Text Then
                Set GetItemByText = itm

                Exit For

            End If

        Next

    Else

        For Each itm In objLvw.ListItems

            If itm.ListSubItems(ColIndex).Text = Text Then
                Set GetItemByText = itm

                Exit For

            End If

        Next

    End If
    
    Set itm = Nothing
End Function

'*************************************************************************
'**功能描述：获取 LivwList 列值
'**函 数 名：GetLvwItem
'**输    入：objLvw(ListView)                   -
'**        ： sItem(String)                      -
'**        ： Optional Col(Long = -1)            -
'**        ： Optional SplitStr(String = vbCrLf) -
'**输    出：(String) -
'**例    子：GetLvwItem ListView控件, "序号|名称|值", [要获取的值 不填则是当前选中的行], [用于分隔列值的分隔符号]
'**作    者：格式化 QQ:65464145
'**日    期：2009-05-10 19:32:17
'*************************************************************************
Public Function GetLvwItem(ByVal objLvw As Object, _
                           ByVal sItem As String, _
                           Optional ByVal Col As Long = -1, _
                           Optional ByVal SplitStr As String = vbCrLf) As String

    Dim i       As Long

    Dim str1    As String

    Dim k       As Long

    Dim p()     As String

    Dim colHead As Collection

    If objLvw.ListItems.Count = 0 Then Exit Function
    If Col = -1 Then Col = objLvw.SelectedItem.Index

    k = objLvw.ColumnHeaders.Count
    Set colHead = New Collection

    For i = 1 To k

        If i = 1 Then
            str1 = objLvw.ListItems(Col).Text
        Else
            str1 = objLvw.ListItems(Col).SubItems(i - 1)
        End If

        colHead.Add str1, objLvw.ColumnHeaders.Item(i).Text
    Next
    
    On Error Resume Next

    p = Split(sItem, "|")

    For i = 0 To UBound(p)

        If Len(Trim(p(i))) > 0 Then
            p(i) = colHead(p(i))
        End If

    Next

    On Error GoTo 0
    
    GetLvwItem = Join(p, SplitStr)
    Set colHead = Nothing
End Function

'获取 Lvw 当前鼠标下的 列 Index
Public Function GetLvwSubItemIndex(ByVal objLvw As Object) As Long

    Dim i           As Long, X As Long

    Dim ScrollWidth As Long

    Dim LvwPT       As POINTAPI, MousePT As POINTAPI

    With objLvw
        ScrollWidth = GetScrollPos(.hwnd, SB_HORZ)
        
        ClientToScreen .hwnd, LvwPT
        GetCursorPos MousePT
        
        X = MousePT.X - LvwPT.X

        If X < 0 Then X = 0
        X = (ScrollWidth + X) * 15

        For i = 1 To .ColumnHeaders.Count

            If X >= .ColumnHeaders(i).Left And X <= .ColumnHeaders(i).Left + .ColumnHeaders(i).Width Then
                GetLvwSubItemIndex = i

                Exit For

            End If

        Next

    End With

End Function

'获取选定项的数量
Public Function GetSelectedCount(objLvw As Object) As Long

    Dim c   As Long

    Dim itm As ListItem
    
    For Each itm In objLvw.ListItems

        If itm.Selected Then
            c = c + 1
        End If

    Next

    GetSelectedCount = c
    Set itm = Nothing
End Function

'获取包含选定项的值的数组
'objLvw:ListView对象引用
'ColumnIndices:数据列的索引，多列以逗号分隔
'Separator:每行各列之间的分隔符，默认vbTab
'WithListItemText:是否获取ListItem.Text，若指定该参数为True，ListItem.Text将作为每行的第一个数据
Public Function GetSelectedItemsValue(objLvw As Object, _
                                      Optional ByVal ColumnIndices As String, _
                                      Optional ByVal Separator As String = vbTab, _
                                      Optional ByVal WithListItemText As Boolean = False) As Variant

    Dim itm    As ListItem

    Dim arrs() As String

    Dim i      As Long

    Dim j      As Long

    Dim row()  As String

    Dim cols() As String

    Dim u      As Long
    
    ReDim arrs(0 To objLvw.ListItems.Count - 1) As String
    cols = Split(ColumnIndices, ",")
    u = UBound(cols)
    
    If WithListItemText Then
        ReDim row(-1 To UBound(cols)) As String
        
        For Each itm In objLvw.ListItems

            If itm.Selected Then
                row(-1) = itm.Text

                For j = 0 To u
                    row(j) = itm.ListSubItems(cols(j)).Text
                Next

                arrs(i) = Join(row, Separator)
                i = i + 1
            End If

        Next

    Else
        ReDim row(0 To UBound(cols)) As String

        For Each itm In objLvw.ListItems

            If itm.Selected Then

                For j = 0 To u
                    row(j) = itm.ListSubItems(cols(j)).Text
                Next

                arrs(i) = Join(row, Separator)
                i = i + 1
            End If

        Next

    End If

    ReDim Preserve arrs(0 To i - 1) As String
    GetSelectedItemsValue = arrs
End Function

'删除勾选项
Public Sub RemoveCheckedItems(objLvw As Object)

    Dim i As Long

    Dim c As Long
    
    c = objLvw.ListItems.Count
    
    For i = c To 1 Step -1

        If objLvw.ListItems.Item(i).Checked Then
            objLvw.ListItems.Remove i
        End If

    Next

End Sub

'删除Item之前的项目
Public Sub RemovePreviousItems(objLvw As Object, _
                               Item As ListItem, _
                               Optional ByVal Contain As Boolean = True)

    Dim i As Long

    Dim c As Long
    
    If Contain Then
        c = Item.Index
    Else
        c = Item.Index - 1

        If c <= 0 Then Exit Sub
    End If
    
    For i = c To 1 Step -1
        objLvw.ListItems.Remove i
    Next

End Sub

'删除选定项
Public Sub RemoveSelectedItems(objLvw As Object)

    Dim i As Long

    Dim c As Long
    
    c = objLvw.ListItems.Count
    
    For i = c To 1 Step -1

        If objLvw.ListItems.Item(i).Selected Then
            objLvw.ListItems.Remove i
        End If

    Next

End Sub

'*************************************************************************
'**功能描述：保存 ListView 清单
'**函 数 名：SaveLvwList
'**输    入：ByVal objLvw(ListView) -
'**        ： ByVal SavePath(String) -
'**输    出：无
'**例    子：
'**作    者：格式化 QQ:65464145
'**日    期：2010-01-31 18:49:42
'*************************************************************************
Public Function SaveLvwList(ByVal objLvw As Object, ByVal SavePath As String)

    Dim i   As Long, i2 As Long

    Dim p() As String

    With objLvw.ListItems
        ReDim p(.Count)

        For i = 1 To objLvw.ColumnHeaders.Count
            p(0) = p(0) & objLvw.ColumnHeaders(i) & vbTab
        Next

        For i = 1 To .Count
            p(i) = .Item(i).Text & vbTab

            For i2 = 1 To .Item(i).ListSubItems.Count
                p(i) = p(i) & .Item(i).SubItems(i2) & vbTab
            Next
        Next

        Open SavePath For Output As #1
        Print #1, Join(p, vbCrLf)
        Close
    End With

End Function

'------------选择列表项------------
Public Sub SelectAll(objLvw As Object)

    Dim itm As ListItem
    
    For Each itm In objLvw.ListItems

        itm.Selected = True
    Next

    Set itm = Nothing
End Sub

'选择指定颜色的项目
Public Sub SelectByColor(objLvw As Object, _
                         ByVal Color As OLE_COLOR, _
                         Optional ByVal Index As Variant = "")

    Dim itm As ListItem
    
    If Index = "" Then

        For Each itm In objLvw.ListItems

            itm.Selected = (itm.ForeColor = Color)
        Next

    Else

        For Each itm In objLvw.ListItems

            itm.Selected = (itm.ListSubItems(Index).ForeColor = Color)
        Next

    End If

    Set itm = Nothing
End Sub

'反选列表项
Public Sub SelectReverse(objLvw As Object)

    Dim itm As ListItem
    
    For Each itm In objLvw.ListItems

        itm.Selected = Not itm.Selected
    Next

    Set itm = Nothing
End Sub

'设置列表项前景色
Public Sub SetItemForecolor(itm As ListItem, ByVal Color As OLE_COLOR)

    Dim sb As ListSubItem

    For Each sb In itm.ListSubItems

        sb.ForeColor = Color
    Next

    Set sb = Nothing
End Sub

'自动设置某一栏宽度
'1.12 修正父容器scalemode为pixels时的错误
Public Function SetLvwHeadsAutoWidth(objLvw As Object, _
                                     ByVal HeadName As String) As Boolean

    Dim k        As Long

    Dim i        As Long

    Dim vscWidth As Long
    
    HeadName = Trim$(HeadName)

    If Len(HeadName) = 0 Then Exit Function
    
    On Error GoTo ToExit

    With objLvw.ColumnHeaders

        For i = 1 To .Count

            If .Item(i).Key <> HeadName Then
                k = k + .Item(i).Width
                Debug.Print k
            End If

        Next
        
        Select Case objLvw.Container.ScaleMode

        Case ScaleModeConstants.vbTwips
            vscWidth = 24 * Screen.TwipsPerPixelX

        Case ScaleModeConstants.vbPixels
            vscWidth = 24
        End Select

        .Item(HeadName).Width = objLvw.Width - k - vscWidth
    End With

ToExit:
End Function

'设置 ListView Item 背景色
Public Function SetLvwItemBackColor(ByVal objLvw As Object, _
                                    ByVal objPicBox As Object, _
                                    Optional ByVal color1 As OLE_COLOR = vbWhite, _
                                    Optional ByVal color2 As OLE_COLOR = &HF3EEEA)

    Dim i          As Integer

    Dim ItemHeight As Single

    Dim ItemCount  As Long
    
    If objLvw.ListItems.Count = 0 Then
        ItemHeight = 209.7738
    Else
        ItemHeight = objLvw.ListItems(1).Height
    End If
    
    ItemCount = 100

    With objPicBox
        .BackColor = objLvw.BackColor
        .ScaleMode = vbTwips
        .BorderStyle = vbBSNone
        .AutoRedraw = True
        .Visible = False
        .Width = Screen.Width   '因为LISTVIEW会自动调整大小的，所以直接用屏幕的宽度
        .Height = ItemHeight * ItemCount '取得要填充的高度
        .ScaleHeight = ItemCount
        .ScaleWidth = 1
        .DrawWidth = 1
    End With

    For i = 1 To ItemCount

        If i / 2 = Int(i / 2) Then
            objPicBox.Line (0, i - 1)-(1, i), color2, BF
        Else
            objPicBox.Line (0, i - 1)-(1, i), color1, BF
        End If

    Next

    objLvw.Picture = objPicBox.Image
End Function

'点击ColumnHeader时排序
'SortMode = 0 字符
'SortMode = 1 数字
Public Sub SortByHead(objLvw As Object, _
                      ByVal ColumnHeader As MSComctlLib.ColumnHeader, _
                      Optional ByVal SortMode As Long = 0)

    With objLvw
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1

        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If

    End With

End Sub

' *********************************************************
'  IListItem-based Sorting Routines
' *********************************************************
Public Function LVSortI(lv As ListView, _
                        ByVal Index As Long, _
                        ByVal ItemType As LVItemTypes, _
                        ByVal SortOrder As LVSortTypes) As Boolean
    'Dim tmr As New CStopWatch
   
    ' turn off the default sorting of the control
    With lv
        .Sorted = False
        .SortKey = Index
        .SortOrder = SortOrder
    End With

    ' no lookups used by this method
    'BuildLookup = 0
   
    ' need to use module variables to let compare routine know
    ' which column and method to use
    m_lvSortColumn = Index
    m_lvSortType = ItemType
   
    ' fire off sorting
    'tmr.Reset
    Call SendMessageLong(lv.hwnd, LVM_SORTITEMS, SortOrder, AddressOf LVCompareI)
    'PerformSort = tmr.Elapsed
   
    ' delete collection of sort data
    Set m_lvSortColl = Nothing
End Function

Private Function LVCompareI(ByVal lParam1 As Long, _
                            ByVal lParam2 As Long, _
                            ByVal SortOrder As Long) As Long

    Static ListItem1 As ListItem

    Static ListItem2 As ListItem

    Static sItem1    As String

    Static sItem2    As String
   
    ' WARNING: This method *will* likely break in the future!
    ' Glom references to internal ListItem class using magic number
    CopyMem ListItem1, lParam1 + 84, 4
    CopyMem ListItem2, lParam2 + 84, 4
   
    ' Grab text items of interest
    If m_lvSortColumn = 0 Then
        sItem1 = ListItem1.Text
        sItem2 = ListItem2.Text
    Else
        sItem1 = ListItem1.SubItems(m_lvSortColumn)
        sItem2 = ListItem2.SubItems(m_lvSortColumn)
    End If
   
    ' Clean up hacked reference
    CopyMem ListItem1, Nothing, 4
    CopyMem ListItem2, Nothing, 4
   
    ' Perform ascending comparison
    On Error GoTo Failure

    Select Case m_lvSortType

    Case lvDate
        LVCompareI = Sgn(CDate(sItem1) - CDate(sItem2))

    Case lvNumber
        LVCompareI = Sgn(CDbl(sItem1) - CDbl(sItem2))

    Case lvBinary
        LVCompareI = StrComp(sItem1, sItem2, vbBinaryCompare)

    Case lvAlphabetic
        LVCompareI = StrComp(sItem1, sItem2, vbTextCompare)

    Case Else ' default ascending text
        LVCompareI = StrComp(sItem1, sItem2, vbTextCompare)
    End Select

    On Error GoTo 0
   
    ' Negate if descending
    If SortOrder = lvDescending Then
        LVCompareI = -LVCompareI
    End If

    Exit Function
   
Failure:

    ' Bail with 0 for failed comparison, because it's "just a visual sort" <g>
    ' Might want to return failure code in real app by setting flag here.
    Exit Function

End Function

