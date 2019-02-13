Attribute VB_Name = "modTrayIcon"
'��ӵ�����   AddTrayIcon �������, [��ʾ�ı�], [����˵�], [�һ��˵�]
'�Ƴ�����ͼ�� RemoveTrayIcon
'�޸����ͼ��˵� ChangeLClickMenu �˵�����
'�޸��һ�ͼ��˵� ChangeRClickMenu �˵�����
'�޸�����ͼ�� SetTrayIcon ͼ����
'����������ʾ SetTrayTip [��ʾ�ı�], [���ݱ���], [��������], [��������], [����ʱ��]
'����ڵ�ǰ̨ ActivationWindow hwnd

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

'=== �����
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ShowWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal nCmdShow As Long) As Long

Private Type NOTIFYICONDATA

    cbSize As Long              ' �ṹ��С(�ֽ�)
    hwnd As Long                ' ������Ϣ�Ĵ��ڵľ��
    uID As Long                 ' Ψһ�ı�ʶ��
    uFlags As Long              ' Flags
    uCallbackMessage As Long    ' ������Ϣ�Ĵ��ڽ��յ���Ϣ
    hIcon As Long               ' ����ͼ����
    szTip As String * 128       ' Tooltip ��ʾ�ı�
    dwState As Long             ' ����ͼ��״̬
    dwStateMask As Long         ' ״̬����
    szInfo As String * 256      ' ������ʾ�ı�
    uTimeoutOrVersion As Long   ' ������ʾ��ʧʱ���汾 uTimeout - ������ʾ��ʧʱ��(��λ:ms, 10000 -- 30000)  uVersion - �汾(0 for V4, 3 for V5)
    szInfoTitle As String * 64  ' ������ʾ����
    dwInfoFlags As Long         ' ������ʾͼ��

End Type

'dwState to NOTIFYICONDATA structure
Private Const NIS_HIDDEN = &H1              ' ����ͼ��

Private Const NIS_SHAREDICON = &H2          ' ����ͼ��

'dwInfoFlags to NOTIFIICONDATA structure
Private Const NIIF_NONE = &H0               ' ��ͼ��

Private Const NIIF_INFO = &H1               ' "��Ϣ"ͼ��

Private Const NIIF_WARNING = &H2            ' "����"ͼ��

Private Const NIIF_ERROR = &H3              ' "����"ͼ��

Private Const NIIF_GUID = &H4               ' ʹ������ͼ��

Private Const NIIF_ICON_MASK = &HF          ' 6.0�汾����

Private Const NIIF_NOSOUND = &H10           ' ��6.0�汾�������Ŷ�Ӧ������

'uFlags to NOTIFYICONDATA structure
Private Const NIF_ICON       As Long = &H2

Private Const NIF_INFO       As Long = &H10

Private Const NIF_MESSAGE    As Long = &H1

Private Const NIF_STATE      As Long = &H8

Private Const NIF_TIP        As Long = &H4

'dwMessage to Shell_NotifyIcon
Private Const NIM_ADD        As Long = &H0         '����һ��ͼ�굽������

Private Const NIM_DELETE     As Long = &H2      '��������ɾ��һ��ͼ��

Private Const NIM_MODIFY     As Long = &H1      '�޸�ͼ��

Private Const NIM_SETFOCUS   As Long = &H3    '�����㣨Focus�������������������Ϣͨ����������ͼ��������û������µĲ����󷢳�������һ������ͼ����ʾ��һ����ݲ˵���Ȼ���û�����ESC���˲�������ʱʹ��NIM_SETFOCUS������������������������������ϵͳ����볣�ÿ���DLL�� Shlwapi.dll��Comctl32.dll��5.0���ϰ汾�ſ���

Private Const NIM_SETVERSION As Long = &H4  'ָ��ʹ���ض��汾��ϵͳ����볣�ÿ���DLL��ȱʡֵΪ0����ʾʹ��Win95��ʽ���������ϵͳ����볣�ÿ���DLL 5.0���ϰ汾�ſ��á�lpdata��[�������] һ��ָ��NOTIFYICONDATA�ṹ��ָ�롣

Private Const WM_RBUTTONUP = &H205

Private Const WM_USER = &H400

Private Const WM_NOTIFYICON = WM_USER + 1            ' �Զ�����Ϣ

Private Const WM_LBUTTONDBLCLK = &H203

Private Const WM_LBUTTONUP = &H202

Private Const GWL_WNDPROC = (-4)

'����������ʾ���Զ�����Ϣ, 2000�²�������Щ��Ϣ
Private Const NIN_BALLOONSHOW = (WM_USER + &H2)      ' �� Balloon Tips ����ʱִ��

Private Const NIN_BALLOONHIDE = (WM_USER + &H3)      ' �� Balloon Tips ��ʧʱִ��(�� SysTrayIcon ��ɾ��) ��ָ���� TimeOut ʱ�䵽������� Balloon Tips �����ʧ�����ʹ���Ϣ

Private Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)   ' �� Balloon Tips �� TimeOut ʱ�䵽ʱִ��

Private Const NIN_BALLOONUSERCLICK = (WM_USER + &H5) ' ������� Balloon Tips ʱִ�С� ע��:��XP��ִ��ʱ Balloon Tips ���и��رհ�ť, ��������ڰ�ť�Ͻ����յ� NIN_BALLOONTIMEOUT ��Ϣ��

Private Const WM_RBUTTONDOWN = &H204

Private Const WM_LBUTTONDOWN = &H201

Private Const WM_NCDESTROY = &H82 '���������ʱ����Ϣ,�ڰ���VB IDE��ֹͣ��ťҲ�����

Private Const WM_CLOSE = &H10

Private Const WM_SIZE = &H5

Private Const SC_MINIMIZE = &HF020

Private Const WM_SYSCOMMAND = &H112

Private Const WM_SETICON = &H80

Private preWndProc As Long

Private IconData   As NOTIFYICONDATA

Private objFrm     As Form '�������

Private LClickMenu As Menu '����˵�

Private RClickMenu As Menu '�һ��˵�

Public Enum TrayToolTipType

    NoICO = NIIF_NONE             '��ͼ��
    InformationICO = NIIF_INFO    '��ʾͼ��
    ExclamationICO = NIIF_WARNING '����ͼ��
    CriticalICO = NIIF_ERROR      '����ͼ��
    TrayICO = NIIF_GUID           'ʹ������ͼ��

End Enum

'****************** �������� ******************
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

'����
Private Function WindowProc(ByVal hwnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lparam As Long) As Long

    Dim pt As POINTAPI

    Dim k  As Long, TaskRect As Rect

    '���� WM_NOTIFYICON ��Ϣ
    If uMsg = WM_NOTIFYICON Then

        Select Case lparam

            '���� ����ͼ��
        Case WM_LBUTTONUP

            If LClickMenu Is Nothing Then 'û������˵� ��ֱ�Ӽ���ڵ�ǰ̨
                ActivationWindow objFrm.hwnd

            Else '��ʾ����˵�
                objFrm.PopupMenu LClickMenu
            End If

            '�һ� ����ͼ��
        Case WM_RBUTTONUP '�һ�����ͼ��

            If RClickMenu Is Nothing Then Exit Function
            objFrm.PopupMenu RClickMenu

            '˫�� ����ͼ��
        Case WM_LBUTTONDBLCLK
            ActivationWindow objFrm.hwnd

            '��ʾ������ʾ
        Case NIN_BALLOONSHOW

            'ɾ������ͼ��
        Case NIN_BALLOONHIDE

            '������ʾ��ʧ
        Case NIN_BALLOONTIMEOUT

            '����������ʾ
        Case NIN_BALLOONUSERCLICK

            '����������˵�֮��ʱ �رյ����˵�
        Case WM_RBUTTONDOWN, WM_LBUTTONDOWN
            SetForegroundWindow hwnd '�ؼ���һ��

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

'����ڵ�ǰ̨
Public Function ActivationWindow(hwnd As Long) As Long

    Const SW_SHOWNORMAL = &H1 '��ǰ

    Const SW_INVALIDATE = &H2 '�ú�
    
    SetForegroundWindow hwnd
    ShowWindow hwnd, &H3
End Function

'�������ͼ��
Public Function AddTrayIcon(objForm As Form, _
                            Optional ToolTipText As String = "[ToolTipText]", _
                            Optional LeftClickMenu As Menu, _
                            Optional RightClickMenu As Menu) As Boolean
    
    If App.LogMode = 0 Then Exit Function
    If Not objFrm Is Nothing Then Exit Function '����Ѿ���ͼ�����˳�

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
        .szInfo = "" & vbNullChar      '��ʾ��Ϣ
        .szInfoTitle = "" & vbNullChar '����
        .dwInfoFlags = 1               '��ʾ����
        .uTimeoutOrVersion = 0         '��ʾ����ʾʱ��
    End With

    Shell_NotifyIcon NIM_ADD, IconData
    preWndProc = SetWindowLong(objFrm.hwnd, GWL_WNDPROC, AddressOf WindowProc)
    AddTrayIcon = True
End Function

'�ı� ����˵�
Public Sub ChangeLClickMenu(objMenu As Menu)
    Set LClickMenu = objMenu
End Sub

'�ı� �һ��˵�
Public Sub ChangeRClickMenu(objMenu As Menu)
    Set RClickMenu = objMenu
End Sub

'����������ʾ
Public Function SetTrayTip(Optional ToolTipText As String = "[ToolTipText]", _
                           Optional InfoTitle As String, _
                           Optional InfoText As String, _
                           Optional InfoType As TrayToolTipType = InformationICO, _
                           Optional InfoOutTime As Long) As Long

    If objFrm Is Nothing Then Exit Function '���û��ͼ�����˳�

    With IconData

        If ToolTipText <> "[ToolTipText]" Then .szTip = ToolTipText & vbNullChar

        .szInfoTitle = InfoTitle & vbNullChar
        .szInfo = InfoText & vbNullChar
        .uTimeoutOrVersion = InfoOutTime
        .dwInfoFlags = InfoType
        '.uFlags = NIF_TIP 'ָ��Ҫ�Ը�����ʾ��������
    End With

    Shell_NotifyIcon NIM_MODIFY, IconData '����ǰ�涨��NIM_MODIFY������Ϊ���޸�ģʽ��
    SetTrayTip = IconData.hIcon
End Function

'�޸�����ͼ��
Public Function SetTrayIcon(IcoHandle As Long) As Long

    If objFrm Is Nothing Then Exit Function '���û��ͼ�����˳�

    With IconData
        .hIcon = IcoHandle
        '.uFlags = NIF_ICON
    End With

    Shell_NotifyIcon NIM_MODIFY, IconData
    SetTrayIcon = IconData.hIcon
End Function

'ɾ������ͼ��
Public Sub RemoveTrayIcon(Optional bUnloadForm As Boolean)

    Const WM_CLOSE = &H10

    If objFrm Is Nothing Then Exit Sub  '���û��ͼ�����˳�

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

'��������
Private Sub ExitProcess(ByVal hwnd As Long)

    Dim pid  As Long

    Dim hand As Long

    GetWindowThreadProcessId hwnd, pid
    hand = OpenProcess(PROCESS_TERMINATE, True, pid)  '��ȡ���̾��
    TerminateProcess hand, 0 '�رս���
End Sub

'���ô���ͼ�� ���ļ�����    ���ھ��, [ͼ��·�� �������õ�ǰ����·��], [ͼ������]
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

