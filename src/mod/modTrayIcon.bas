Attribute VB_Name = "modTrayIcon"
'添加到托盘   AddTrayIcon 窗体对象, [提示文本], [左击菜单], [右击菜单]
'移除托盘图标 RemoveTrayIcon
'修改左击图标菜单 ChangeLClickMenu 菜单对象
'修改右击图标菜单 ChangeRClickMenu 菜单对象
'修改托盘图标 SetTrayIcon 图标句柄
'设置托盘提示 SetTrayTip [提示文本], [气泡标题], [气泡内容], [气泡类型], [气泡时间]
'激活窗口到前台 ActivationWindow hwnd

Option Explicit

Private Declare Function SetWindowLong _
                Lib "user32.dll" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc _
                Lib "user32.dll" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lparam As Long) As Long

Private Declare Function Shell_NotifyIcon _
                Lib "shell32.dll" _
                Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                           lpData As NOTIFYICONDATA) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function GetWindowRect _
                Lib "user32" (ByVal hwnd As Long, _
                              lpRect As Rect) As Long

Private Declare Function FindWindow _
                Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long

Private Declare Function PostMessage _
                Lib "user32" _
                Alias "PostMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lparam As Long) As Long

Private Declare Function GetModuleHandle _
                Lib "kernel32" _
                Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function ExtractIcon _
                Lib "shell32.dll" _
                Alias "ExtractIconA" (ByVal hInst As Long, _
                                      ByVal lpszExeFileName As String, _
                                      ByVal nIconIndex As Long) As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lparam As Any) As Long

Private Type Rect

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type

Private Type POINTAPI

    x As Long
    y As Long

End Type

'=== 激活窗口
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ShowWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal nCmdShow As Long) As Long

Private Type NOTIFYICONDATA

    cbSize As Long              ' 结构大小(字节)
    hwnd As Long                ' 处理消息的窗口的句柄
    uID As Long                 ' 唯一的标识符
    uFlags As Long              ' Flags
    uCallbackMessage As Long    ' 处理消息的窗口接收的消息
    hIcon As Long               ' 托盘图标句柄
    szTip As String * 128       ' Tooltip 提示文本
    dwState As Long             ' 托盘图标状态
    dwStateMask As Long         ' 状态掩码
    szInfo As String * 256      ' 气球提示文本
    uTimeoutOrVersion As Long   ' 气球提示消失时间或版本 uTimeout - 气球提示消失时间(单位:ms, 10000 -- 30000)  uVersion - 版本(0 for V4, 3 for V5)
    szInfoTitle As String * 64  ' 气球提示标题
    dwInfoFlags As Long         ' 气球提示图标

End Type

'dwState to NOTIFYICONDATA structure
Private Const NIS_HIDDEN = &H1              ' 隐藏图标

Private Const NIS_SHAREDICON = &H2          ' 共享图标

'dwInfoFlags to NOTIFIICONDATA structure
Private Const NIIF_NONE = &H0               ' 无图标

Private Const NIIF_INFO = &H1               ' "消息"图标

Private Const NIIF_WARNING = &H2            ' "警告"图标

Private Const NIIF_ERROR = &H3              ' "错误"图标

Private Const NIIF_GUID = &H4               ' 使用托盘图标

Private Const NIIF_ICON_MASK = &HF          ' 6.0版本保留

Private Const NIIF_NOSOUND = &H10           ' 限6.0版本，不播放对应的声音

'uFlags to NOTIFYICONDATA structure
Private Const NIF_ICON       As Long = &H2

Private Const NIF_INFO       As Long = &H10

Private Const NIF_MESSAGE    As Long = &H1

Private Const NIF_STATE      As Long = &H8

Private Const NIF_TIP        As Long = &H4

'dwMessage to Shell_NotifyIcon
Private Const NIM_ADD        As Long = &H0         '增加一个图标到托盘区

Private Const NIM_DELETE     As Long = &H2      '从托盘区删除一个图标

Private Const NIM_MODIFY     As Long = &H1      '修改图标

Private Const NIM_SETFOCUS   As Long = &H3    '将焦点（Focus）返回托盘区。这个消息通常在托盘区图标完成了用户界面下的操作后发出。比如一个托盘图标显示了一个快捷菜单，然后用户按下ESC键了操作，这时使用NIM_SETFOCUS将焦点继续保留在托盘区。该项仅在系统外壳与常用控制DLL（ Shlwapi.dll与Comctl32.dll）5.0以上版本才可用

Private Const NIM_SETVERSION As Long = &H4  '指定使用特定版本的系统外壳与常用控制DLL。缺省值为0，表示使用Win95方式。该项仅在系统外壳与常用控制DLL 5.0以上版本才可用。lpdata：[输入参数] 一个指向NOTIFYICONDATA结构的指针。

Private Const WM_RBUTTONUP = &H205

Private Const WM_USER = &H400

Private Const WM_NOTIFYICON = WM_USER + 1            ' 自定义消息

Private Const WM_LBUTTONDBLCLK = &H203

Private Const WM_LBUTTONUP = &H202

Private Const GWL_WNDPROC = (-4)

'关于气球提示的自定义消息, 2000下不产生这些消息
Private Const NIN_BALLOONSHOW = (WM_USER + &H2)      ' 当 Balloon Tips 弹出时执行

Private Const NIN_BALLOONHIDE = (WM_USER + &H3)      ' 当 Balloon Tips 消失时执行(如 SysTrayIcon 被删除) 但指定的 TimeOut 时间到或鼠标点击 Balloon Tips 后的消失不发送此消息

Private Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)   ' 当 Balloon Tips 的 TimeOut 时间到时执行

Private Const NIN_BALLOONUSERCLICK = (WM_USER + &H5) ' 当鼠标点击 Balloon Tips 时执行。 注意:在XP下执行时 Balloon Tips 上有个关闭按钮, 如果鼠标点在按钮上将接收到 NIN_BALLOONTIMEOUT 消息。

Private Const WM_RBUTTONDOWN = &H204

Private Const WM_LBUTTONDOWN = &H201

Private Const WM_NCDESTROY = &H82 '组件被销毁时的消息,在按下VB IDE的停止按钮也会产生

Private Const WM_CLOSE = &H10

Private Const WM_SIZE = &H5

Private Const SC_MINIMIZE = &HF020

Private Const WM_SYSCOMMAND = &H112

Private Const WM_SETICON = &H80

Private preWndProc As Long

Private IconData   As NOTIFYICONDATA

Private objFrm     As Form '窗体对像

Private LClickMenu As Menu '左击菜单

Private RClickMenu As Menu '右击菜单

Public Enum TrayToolTipType

    NoICO = NIIF_NONE             '无图标
    InformationICO = NIIF_INFO    '提示图标
    ExclamationICO = NIIF_WARNING '警告图标
    CriticalICO = NIIF_ERROR      '错误图标
    TrayICO = NIIF_GUID           '使用托盘图标

End Enum

'****************** 结束进程 ******************
Private Declare Function GetWindowThreadProcessId _
                Lib "user32" (ByVal hwnd As Long, _
                              lpdwProcessId As Long) As Long

Private Declare Function OpenProcess _
                Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                ByVal bInheritHandle As Long, _
                                ByVal dwProcessId As Long) As Long

Private Declare Function TerminateProcess _
                Lib "kernel32" (ByVal hProcess As Long, _
                                ByVal uExitCode As Long) As Long

Private Const PROCESS_TERMINATE = &H1

'子类
Private Function WindowProc(ByVal hwnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lparam As Long) As Long

    Dim pt As POINTAPI

    Dim k  As Long, TaskRect As Rect

    '拦截 WM_NOTIFYICON 消息
    If uMsg = WM_NOTIFYICON Then

        Select Case lparam

            '单击 托盘图标
        Case WM_LBUTTONUP

            If LClickMenu Is Nothing Then '没有左击菜单 则直接激活窗口到前台
                ActivationWindow objFrm.hwnd

            Else '显示左击菜单
                objFrm.PopupMenu LClickMenu
            End If

            '右击 托盘图标
        Case WM_RBUTTONUP '右击托盘图标

            If RClickMenu Is Nothing Then Exit Function
            objFrm.PopupMenu RClickMenu

            '双击 托盘图标
        Case WM_LBUTTONDBLCLK
            ActivationWindow objFrm.hwnd

            '显示气球提示
        Case NIN_BALLOONSHOW

            '删除托盘图标
        Case NIN_BALLOONHIDE

            '气球提示消失
        Case NIN_BALLOONTIMEOUT

            '单击气球提示
        Case NIN_BALLOONUSERCLICK

            '鼠标点击弹出菜单之外时 关闭弹出菜单
        Case WM_RBUTTONDOWN, WM_LBUTTONDOWN
            SetForegroundWindow hwnd '关键的一步

        End Select

    Else

        Select Case uMsg

        Case WM_SIZE

            On Error Resume Next

            If objFrm.WindowState = 1 Then
                k = FindWindow("Shell_traywnd", "")
                GetWindowRect k, TaskRect

                GetCursorPos pt
           
                If pt.y < TaskRect.Top Then
                    objFrm.Hide
                End If
            End If

            On Error GoTo 0

        Case WM_NCDESTROY, WM_CLOSE
            RemoveTrayIcon

        End Select

    End If

    WindowProc = CallWindowProc(preWndProc, hwnd, uMsg, wParam, lparam)
End Function

'激活窗口到前台
Public Function ActivationWindow(hwnd As Long) As Long

    Const SW_SHOWNORMAL = &H1 '置前

    Const SW_INVALIDATE = &H2 '置后
    
    SetForegroundWindow hwnd
    ShowWindow hwnd, &H3
End Function

'添加托盘图标
Public Function AddTrayIcon(objForm As Form, _
                            Optional ToolTipText As String = "[ToolTipText]", _
                            Optional LeftClickMenu As Menu, _
                            Optional RightClickMenu As Menu) As Boolean
    
    If App.LogMode = 0 Then Exit Function
    If Not objFrm Is Nothing Then Exit Function '如果已经有图标则退出

    Set objFrm = objForm
    Set LClickMenu = LeftClickMenu
    Set RClickMenu = RightClickMenu

    If ToolTipText = "[ToolTipText]" Then ToolTipText = objFrm.Caption

    With IconData
        .cbSize = Len(IconData)
        .hwnd = objFrm.hwnd
        .uID = 0
        .uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE Or NIF_INFO Or NIF_STATE
        .uCallbackMessage = WM_NOTIFYICON
        .szTip = ToolTipText & vbNullChar  'ToolTipText
        .hIcon = objFrm.Icon.handle
        .dwState = 0
        .dwStateMask = 0
        .szInfo = "" & vbNullChar      '提示信息
        .szInfoTitle = "" & vbNullChar '标题
        .dwInfoFlags = 1               '提示类型
        .uTimeoutOrVersion = 0         '提示框显示时间
    End With

    Shell_NotifyIcon NIM_ADD, IconData
    preWndProc = SetWindowLong(objFrm.hwnd, GWL_WNDPROC, AddressOf WindowProc)
    AddTrayIcon = True
End Function

'改变 左击菜单
Public Sub ChangeLClickMenu(objMenu As Menu)
    Set LClickMenu = objMenu
End Sub

'改变 右击菜单
Public Sub ChangeRClickMenu(objMenu As Menu)
    Set RClickMenu = objMenu
End Sub

'设置托盘提示
Public Function SetTrayTip(Optional ToolTipText As String = "[ToolTipText]", _
                           Optional InfoTitle As String, _
                           Optional InfoText As String, _
                           Optional InfoType As TrayToolTipType = InformationICO, _
                           Optional InfoOutTime As Long) As Long

    If objFrm Is Nothing Then Exit Function '如果没有图标则退出

    With IconData

        If ToolTipText <> "[ToolTipText]" Then .szTip = ToolTipText & vbNullChar

        .szInfoTitle = InfoTitle & vbNullChar
        .szInfo = InfoText & vbNullChar
        .uTimeoutOrVersion = InfoOutTime
        .dwInfoFlags = InfoType
        '.uFlags = NIF_TIP '指明要对浮动提示进行设置
    End With

    Shell_NotifyIcon NIM_MODIFY, IconData '根据前面定义NIM_MODIFY，设置为“修改模式”
    SetTrayTip = IconData.hIcon
End Function

'修改托盘图标
Public Function SetTrayIcon(IcoHandle As Long) As Long

    If objFrm Is Nothing Then Exit Function '如果没有图标则退出

    With IconData
        .hIcon = IcoHandle
        '.uFlags = NIF_ICON
    End With

    Shell_NotifyIcon NIM_MODIFY, IconData
    SetTrayIcon = IconData.hIcon
End Function

'删除托盘图标
Public Sub RemoveTrayIcon(Optional bUnloadForm As Boolean)

    Const WM_CLOSE = &H10

    If objFrm Is Nothing Then Exit Sub  '如果没有图标则退出

    With IconData
        .cbSize = Len(IconData)
        .hwnd = objFrm.hwnd
        .uID = 0
        .uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
        .uCallbackMessage = WM_NOTIFYICON
        .szTip = ""
        .hIcon = objFrm.Icon.handle
    End With

    Shell_NotifyIcon NIM_DELETE, IconData
    SetWindowLong objFrm.hwnd, GWL_WNDPROC, preWndProc

    If bUnloadForm = True Then
        'Call PostMessage(objFrm.hwnd, WM_CLOSE, 0&, 0&)
        ExitProcess objFrm.hwnd
    End If

    Set objFrm = Nothing
End Sub

'结束进程
Private Sub ExitProcess(ByVal hwnd As Long)

    Dim pid  As Long

    Dim hand As Long

    GetWindowThreadProcessId hwnd, pid
    hand = OpenProcess(PROCESS_TERMINATE, True, pid)  '获取进程句柄
    TerminateProcess hand, 0 '关闭进程
End Sub

'设置窗体图标 从文件加载    窗口句柄, [图标路径 不填则用当前程序路径], [图标索引]
Public Function SetIconFromFile(ByVal hwnd As Long, _
                                Optional ByVal FileName As String, _
                                Optional ByVal IconIndex As Integer) As Long

    Dim m_Icon  As Long

    Dim hmodule As Long

    Dim MyPath  As String

    If Len(FileName) = 0 Or Len(Dir(FileName, vbHidden)) = 0 Then
        MyPath = App.Path

        If Right$(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
        FileName = MyPath & App.EXEName & ".exe"
    End If

    hmodule = GetModuleHandle(FileName)
    m_Icon = ExtractIcon(hmodule, FileName, IconIndex)
    SendMessage hwnd, WM_SETICON, 0, ByVal m_Icon
    SendMessage hwnd, WM_SETICON, 1, ByVal m_Icon
    SetIconFromFile = m_Icon
End Function

