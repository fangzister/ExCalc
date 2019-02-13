Attribute VB_Name = "modMouseHook"
'mMouseWheel
'�����ֵ��¼����
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

Private mShadowObj    As Object '����¼������־

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

    Case WM_MOUSEWHEEL '����

        Dim wzDelta, wKeys As Integer, FontSize As Integer

        'wzDelta���ݹ��ֹ����Ŀ�������ֵС�����ʾ���������������û����򣩣�
        '�������ʾ������ǰ����������ʾ������
        wzDelta = HIWORD(wParam)
        'wKeysָ���Ƿ���CTRL=8��SHIFT=4������(��=2����=16����=2������)���£�������
        wKeys = LOWORD(wParam)

        If wKeys = 8 Then
            FontSize = mShadowObj.Font.Size

            '--------------------------------------------------
            If wzDelta < 0 Then '���û�����

                '��С����
                If FontSize > 2 Then
                    mShadowObj.Font.Size = FontSize - 1
                End If

            Else '����ʾ������
                '�Ŵ�����
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
    HIWORD = (LongIn And &HFFFF0000) \ &H10000 'ȡ��32λֵ�ĸ�16λ
End Function

Private Function LOWORD(LongIn As Long) As Integer
    LOWORD = LongIn And &HFFFF& 'ȡ��32λֵ�ĵ�16λ
End Function
