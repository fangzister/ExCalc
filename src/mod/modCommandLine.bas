Attribute VB_Name = "modCommandLine"
Option Explicit

Private Declare Function GetCommandLineW Lib "kernel32" () As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function CommandLineToArgvW _
                Lib "shell32" (ByVal lpCmdLine As Long, _
                               pnNumArgs As Long) As Long

Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Public Sub GetCommandLine(ByRef Argc As Long, ByRef Argv() As String)

    Dim nNumArgs    As Long     '//命令行参数个数

    Dim lpszArglist As Long     '//命令行参数数组地址

    Dim lpszArg     As Long     '//命令行各参数地址

    Dim nArgLength  As Long     '//命令行各参数长度

    Dim szArg()     As Byte     '//命令行各参数

    Dim i           As Long
    
    lpszArglist = CommandLineToArgvW(GetCommandLineW(), nNumArgs)

    If lpszArglist Then
        Argc = nNumArgs   '//输出总个数
        ReDim Argv(nNumArgs - 1)
        CopyMemory ByVal VarPtr(lpszArg), ByVal lpszArglist, 4   '//得到argv(0)的地址
      
        For i = 0 To nNumArgs - 1
            nArgLength = lstrlenW(lpszArg)
            ReDim szArg(nArgLength * 2 - 1)
            CopyMemory ByVal VarPtr(szArg(0)), ByVal lpszArg, nArgLength * 2
            Argv(i) = CStr(szArg)
            lpszArg = lpszArg + nArgLength * 2 + 2
        Next
      
        Erase szArg
        Call LocalFree(lpszArglist)
        Argc = Argc - 1
    End If

End Sub

