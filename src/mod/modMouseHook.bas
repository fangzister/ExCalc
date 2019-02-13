Attribute VB_Name = "modMouseHook"
'mMouseWheel
'鼠标滚轮的事件检测
'***********************************************************
Option Explicit

Private Declare Function CallWindowProc _
                Lib "user32" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lparam As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = -4

Private Const WM_MOUSEWHEEL = &H20A

Global lpPrevWndProcA As Long

Private mShadowObj    As Object '鼠标事件激活标志

Public Sub HookMouse(obj As Object)
    Set mShadowObj = obj
    lpPrevWndProcA = SetWindowLong(obj.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookMouse(obj As Object)
    Set mShadowObj = Nothing
    SetWindowLong obj.hwnd, GWL_WNDPROC, lpPrevWndProcA
End Sub

Private Function WindowProc(ByVal hw As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lparam As Long) As Long

    Select Case uMsg

    Case WM_MOUSEWHEEL '滚动

        Dim wzDelta, wKeys As Integer, FontSize As Integer

        'wzDelta传递滚轮滚动的快慢，该值小于零表示滚轮向后滚动（朝用户方向），
        '大于零表示滚轮向前滚动（朝显示器方向）
        wzDelta = HIWORD(wParam)
        'wKeys指出是否有CTRL=8、SHIFT=4、鼠标键(左=2、中=16、右=2、附加)按下，允许复合
        wKeys = LOWORD(wParam)

        If wKeys = 8 Then
            FontSize = mShadowObj.Font.Size

            '--------------------------------------------------
            If wzDelta < 0 Then '朝用户方向

                '缩小字体
                If FontSize > 2 Then
                    mShadowObj.Font.Size = FontSize - 1
                End If

            Else '朝显示器方向
                '放大字体
                mShadowObj.Font.Size = FontSize + 1
            End If

        Else
            WindowProc = CallWindowProc(lpPrevWndProcA, hw, uMsg, wParam, lparam)
        End If

        '--------------------------------------------------
    Case Else
        WindowProc = CallWindowProc(lpPrevWndProcA, hw, uMsg, wParam, lparam)
    End Select

End Function

Private Function HIWORD(LongIn As Long) As Integer
    HIWORD = (LongIn And &HFFFF0000) \ &H10000 '取出32位值的高16位
End Function

Private Function LOWORD(LongIn As Long) As Integer
    LOWORD = LongIn And &HFFFF& '取出32位值的低16位
End Function
