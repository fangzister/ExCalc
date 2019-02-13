Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function inet_addr Lib "ws2_32" (ByVal lpszAddress As String) As Long

Public Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOW = 5
Private Const SW_RESTORE = 9
Private Const GW_HWNDNEXT = 2
Private Const INADDR_NONE = &HFFFFFFFF

Private Function IsIP(ByVal strAddress As String) As Boolean
    IsIP = (inet_addr(strAddress) <> INADDR_NONE)
End Function

Public Function IsURL(ByVal str As String) As Boolean

End Function

Sub Test()
    '
End Sub

Public Function GetAppVersion() As String
    GetAppVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function


Public Function GetIniProfile() As INIProfile
    Dim ini As INIProfile
    Dim fs  As Long
    Dim fn  As String

    Set ini = New INIProfile

    With ini
        .ExeFolderPath = App.Path
        .Name = App.title
    End With
    Set GetIniProfile = ini
    Set ini = Nothing
End Function

Public Function ActivationWindow(hwnd As Long) As Long
    Const SW_SHOWNORMAL = &H1 '÷√«∞
    Const SW_INVALIDATE = &H2 '÷√∫Û
    
    SetForegroundWindow hwnd
    ShowWindow hwnd, &H4
End Function

Sub Main()
    Dim sTitle As String
    Dim h      As Long
    
    sTitle = App.title & " - V" & GetAppVersion
    
    If App.PrevInstance Then
        App.title = ""
        
        h = FindWindow("ThunderRT6FormDC", sTitle)
        ActivationWindow h
    Else
        frmMain.Show
    End If
End Sub
