Attribute VB_Name = "modDownloadURL"
Option Explicit

Public Function DownloadFile(ByVal URL As String, _
                             Optional ByVal saveDir As String, _
                             Optional ByVal FileName As String, _
                             Optional ByVal TryCount As Long = 3) As String

    Dim oWinHttp  As WinHttp.WinHttpRequest

    Dim BinBuff() As Byte

    Dim p()       As String

    Dim str1      As String

    Dim SavePath  As String

    Dim k         As Long, DownCount As Long

    On Error GoTo hErr

ToBegin:
    Set oWinHttp = New WinHttp.WinHttpRequest

    With oWinHttp
        URL = Trim$(URL)

        '============== 执行 OPEN
        .Open "GET", URL, True

        .SetRequestHeader "Accept", "text/html, application/xhtml+xml, */*"
        .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36"

        '============= 发送数据
        .Send

        '============= 等待返回
        .WaitForResponse

        If TypeName(.ResponseBody) <> "Empty" Then
            BinBuff = .ResponseBody
            str1 = .GetAllResponseHeaders

            If Len(FileName) = 0 Then
                p = Split(str1, " filename=")

                If UBound(p) <= 0 Then '是直接下载的
                    str1 = .Option(1)
                    k = InStrRev(str1, "/")

                    If k = 0 Then GoTo hErr
    
                    str1 = Right$(str1, Len(str1) - k)
    
                Else
                    str1 = p(1)
                    k = InStr(str1, vbCrLf)

                    If k > 0 Then str1 = Left$(str1, k - 1)
    
                    k = InStr(str1, ";")

                    If k > 0 Then str1 = Left$(str1, k - 1)
                End If
    
                FileName = str1
            End If

            'If LCase$(Right$(FileName, 4)) <> ".exe" Then GoTo hErr

            If Len(saveDir) = 0 Then
                saveDir = Environ$("temp")
            End If

            SavePath = saveDir & "\" & FileName
            SavePath = Replace$(SavePath, "\\", "\")

            Open SavePath For Binary As #1
            Put #1, , BinBuff
            Close

            DownloadFile = SavePath
        End If

hErr:
    End With

    Set oWinHttp = Nothing

    DownCount = DownCount + 1

    If Len(DownloadFile) = 0 And DownCount < TryCount Then
        GoTo ToBegin
    End If

End Function

