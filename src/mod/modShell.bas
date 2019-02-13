Attribute VB_Name = "modShell"
Option Explicit

Private Const SYNCHRONIZE = &H100000

Private Const INFINITE = &HFFFFFFFF
 
Private Declare Function OpenProcess _
                Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                ByVal bInheritHandle As Long, _
                                ByVal dwProcessId As Long) As Long

Private Declare Function WaitForSingleObject _
                Lib "kernel32" (ByVal hHandle As Long, _
                                ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 
Public Sub OpenPlugin(ByVal cmd As String)

    Dim lngPId     As Long

    Dim lngPHandle As Long
     
    lngPId = Shell(cmd, vbHide)
    lngPHandle = OpenProcess(SYNCHRONIZE, 0, lngPId)

    If lngPHandle <> 0 Then
        Call WaitForSingleObject(lngPHandle, INFINITE) '无限等待，直到程式结束
        Call CloseHandle(lngPHandle)
    End If

End Sub
