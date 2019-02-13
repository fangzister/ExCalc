Attribute VB_Name = "modUtil"
Option Explicit

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Public Sub ShellOpen(Path As String)
    ShellExecute 0, "Open", Path, vbNullString, vbNullString, vbMaximizedFocus
End Sub

