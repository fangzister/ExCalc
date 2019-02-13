Attribute VB_Name = "modStrings"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Private Const CP_UTF8 = 65001

'ÕýÔòÌæ»»
Public Function RegReplace(ByVal Source As String, ByVal Find As String, ByVal ReplaceText As String, Optional IgnoreCase As Boolean = True) As String
    Dim reg As RegExp
    
    Set reg = New RegExp
    reg.Pattern = Find
    reg.Global = True
    reg.IgnoreCase = IgnoreCase
    
    If reg.Test(Source) Then
        RegReplace = reg.Replace(Source, ReplaceText)
    End If
End Function

'Ö§³Ö½«×î³¤ÎåÎ»ÊýµÄºº×ÖÊýÖµ×ª»»ÎªÊý×ÖÐÎÊ½
Public Function CNSerial2Albert(ByVal CN As String) As String
    Dim n As Long
    Dim s As String
    Dim l As String
    Dim R As String
    Dim v As Long
    Dim tIndex As Long
    Dim hIndex As Long
    Dim mIndex As Long
    
    n = Len(CN)
    
    If n = 0 Then Exit Function
    If n > 5 Then Exit Function
    
    CNSerial2Albert = CN
    
    s = "Ò»¶þÈýËÄÎåÁùÆß°Ë¾Å"
    
    'Ò»Î»ÊýµÄÇé¿ö
    If n = 1 Then
        If CN = "Áã" Or CN = "©–" Then
            CNSerial2Albert = "0"
        ElseIf CN = "Ê®" Then
            CNSerial2Albert = 10
        Else
            v = InStr(s, CN)
            If v < 1 Then Exit Function
            CNSerial2Albert = v
        End If
        
        Exit Function
    End If

    l = Left$(CN, 1)
    hIndex = InStr(s, l)
    
    If n = 2 Then
        If l = "Ê®" Then
            R = Right$(CN, 1)
            tIndex = InStr(s, R)
            
            'Ê×Î»ÊÇ10£¬Ä©Î»±ØÐëÊÇ1-9
            If tIndex > 0 And tIndex < 10 Then
                CNSerial2Albert = 10 + tIndex
            End If
        Else
            
            If hIndex < 1 Then Exit Function
            R = Right(CN, 1)
            'Ê×Î»ÊÇ1-9£¬Ä©Î»±ØÐëÊÇÊ®»òÕß°Ù
            If R = "Ê®" Then
                CNSerial2Albert = hIndex * 10
            ElseIf R = "°Ù" Then
                CNSerial2Albert = hIndex * 100
            End If
        End If
        
        Exit Function
    End If
    
    If n = 3 Then
        'Ê×Î»±ØÐëÊÇ2-9
        If hIndex < 2 Or hIndex > 9 Then Exit Function
        
        R = Mid$(CN, 2, 1)
        
        'ÖÐ¼ä±ØÐëÊÇÊ®
        If R <> "Ê®" Then Exit Function
        
        'Ä©Î»±ØÐëÊÇ1-9
        R = Right$(CN, 1)
        tIndex = InStr(s, R)
        If tIndex > 0 Then
            CNSerial2Albert = hIndex * 10 + tIndex
        End If
        
        Exit Function
    End If
        
    '³¬¹ýÈýÎ»Êý£¬Ê×Î»±ØÐëÊÇ1-9
    If hIndex < 1 Or hIndex > 9 Then Exit Function
    
    R = Mid$(CN, 2, 1)
    
    'µÚ¶þÎ»±ØÐëÊÇ°Ù
    If R <> "°Ù" Then Exit Function
    
    'µÚËÄÎ»±ØÐëÊÇÊ®
    l = Mid$(CN, 4, 1)
    If l <> "Ê®" Then Exit Function
    
    If n = 4 Then
        R = Mid$(CN, 3, 1)
        If R = "Áã" Or R = "©–" Then
            'µÚÈýÎ»ÊÇÁã
            mIndex = 0
            'µÚËÄÎ»±ØÐëÊÇ1-9
            R = Right$(CN, 1)
            tIndex = InStr(s, R)
            If tIndex > 0 Then
                CNSerial2Albert = hIndex * 100 + tIndex
            End If
        Else
            '·ñÔòµÚÈýÎ»±ØÐëÊÇ0-9
            mIndex = InStr(s, R)
            If mIndex > 0 Then
                R = Right$(CN, 1)
                tIndex = InStr(s, R)
                CNSerial2Albert = hIndex * 100 + mIndex * 10 + tIndex
            End If
        End If
                        
        Exit Function
    End If
    
    'Èý°ÙËÄÊ®Îå,Ò»°ÙÁùÊ®°Ë,Èý°ÙÁã¶þÊ®
    If n = 5 Then
        l = Mid$(CN, 3, 1)
        mIndex = InStr(s, l)
        'µÚÈýÎ»±ØÐëÊÇ1-9
        If mIndex < 1 Then Exit Function
        
        'µÚÎåÎ»±ØÐëÊÇ1-9
        R = Right$(CN, 1)
        tIndex = InStr(s, R)
        If tIndex > 0 Then
            CNSerial2Albert = hIndex * 100 + mIndex * 10 + tIndex
        End If
    End If
End Function


'--------------------------------------------------
' Description : ½øÖÆ×ª»»Îª UTF8/GB2312 Ô´Âë
'--------------------------------------------------
Public Function ByteTosHTML(ByRef BinBuff() As Byte) As String
    Dim objStream As Object

    Set objStream = CreateObject("adodb.stream")

    With objStream
        .Type = 1
        .Mode = 3
        .Open
        .Write BinBuff
        .Position = 0
        .Type = 2
        .Charset = IIf(IsUTF8(BinBuff) = True, "UTF-8", "GB2312")
        ByteTosHTML = .ReadText
        .Close
    End With

    Set objStream = Nothing
End Function

'--------------------------------------------------
' Procedure   : CharAt
' Description : ·µ»ØÔÚstrÖÐµÄµÚindex¸ö×Ö·û
' CreateTime  : 2010-11-16-09:43:58
'
' ParamList   : Str (String)    Ä¿±êString
'               Index (Long)    Î»ÖÃ
' Return      : ·µ»ØStrÖÐµÚIndex¸ö×Ö·û£¬ÈôÎÞ£¬·µ»Ø""
'--------------------------------------------------
Public Function CharAt(ByVal str As String, ByVal Index As Long) As String
    On Error GoTo Err

    CharAt = Mid$(str, Index + 1, 1)
Err:
End Function

'--------------------------------------------------
' Procedure   : Contains
' Description : ²âÊÔStrÖÐÊÇ·ñ°üº¬chars
' CreateTime  : 2010-11-16-09:44:19
'
' ParamList   : Str (String)
'               chars (String)
' Return      : µ±ÇÒ½öµ±str°üº¬charsÊ±£¬²Å·µ»Ø true
'--------------------------------------------------
Public Function Contains(ByVal str As String, ByVal Chars As String) As Boolean
    Contains = (InStr(1, str, Chars, vbBinaryCompare) > 0)
End Function

'--------------------------------------------------
' Procedure   : ContainString
' Description : ¼ì²âStrÖÐÊÇ·ñ°üº¬CheckÖÐµÄÈÎÒ»×Ö·û
' CreateTime  : 2010-04-03 21:26:29
'
' ParamList   : Str (String)                    ±»¼ì²âµÄ×Ö·û´®
'               Check (String)                  Òª¼ì²âµÄ×Ö·û
'               [ByRef] FirstContainedString (String)   Èç¹û°üº¬£¬Ôò½«µÚÒ»¸ö°üº¬µÄ×Ö·û¸³Öµ£¬·ñÔò·µ»ØFalse
' Return      : °üº¬£¬·µ»ØTrue£¬²¢½«°üº¬µÄ×Ö·û´®¸³Öµµ½ FirstContainedString
'--------------------------------------------------
Public Function ContainString(str As String, Check As String, Optional ByRef FirstContainedString As String = "") As Boolean
    Dim i As Long
    Dim c As String
    Dim l As Long
    
    l = Len(Check)
    
    If Len(str) = 0 Or l = 0 Then
        Exit Function
    End If
    
    For i = 1 To l
        c = Mid$(Check, i, 1)

        If InStr(1, str, c) > 0 Then
            FirstContainedString = c
            ContainString = True

            Exit Function
        End If
    Next
End Function

'--------------------------------------------------
' Procedure   : DeleteBlankLines
' Description : É¾³ýÄ¿±ê×Ö·û´®ÖÐµÄ¿Õ°×ÐÐ
' CreateTime  : 2016-01-12 11:22
'
' ParamList   : Source (String)    Ä¿±ê×Ö·û´®
'             : District (Boolean) ÉèÎªFalseÊ±£¬½«½öº¬¿Õ¸ñºÍtabµÄÐÐÒ²×÷Îª¿ÕÐÐ´¦Àí
' Return      : ·µ»ØÐÂµÄ×Ö·û´®
'--------------------------------------------------
Public Function DeleteBlankLines(Source As String, Optional ByVal District As Boolean = True) As String
    Dim s   As String
    Dim p() As String
    Dim R() As String
    Dim i   As Long
    Dim j   As Long
    Dim u   As Long
    
    DeleteBlankLines = Source
    
    If Len(Source) = 0 Then
        Exit Function
    End If
    
    If Len(Trim$(Source)) = 0 Then
        Exit Function
    End If
    
    p = Split(Source, vbCrLf)
    u = UBound(p)
    ReDim R(0 To u) As String
    
    If District Then
        For i = 0 To u
            If Len(p(i)) > 0 Then
                R(j) = p(i)
                j = j + 1
            End If
        Next
    Else
        For i = 0 To u
            If Len(p(i)) > 0 Then
                If Len(Trim$(Replace$(p(i), vbTab, ""))) > 0 Then
                    R(j) = p(i)
                    j = j + 1
                End If
            End If
        Next
    End If
    
    If j > 0 Then
        ReDim Preserve R(0 To j - 1) As String
        DeleteBlankLines = Join(R, vbCrLf)
    End If
End Function

'--------------------------------------------------
' Procedure   : DeleteBlankLines
' Description : É¾³ýÄ¿±ê×Ö·û´®ÖÐµÄÖØ¸´ÐÐ£¨²»º¬¿ÕÐÐ£©
' CreateTime  : 2016-01-12 11:22
'
' ParamList   : Source (String)    Ä¿±ê×Ö·û´®
' Return      : ·µ»ØÐÂµÄ×Ö·û´®
'--------------------------------------------------
Public Function DeleteDuplicateLines(Source As String) As String
    Dim p() As String
    Dim t   As String
    Dim i   As Long
    Dim u   As Long
    Dim s   As String
    
    If InStr(Source, vbCrLf) = 0 Then
        DeleteDuplicateLines = Source

        Exit Function
    End If
    
    t = Replace$(Source, vbCrLf, Chr$(0))
    p = Split(t, Chr$(0))
    
    u = UBound(p)

    For i = 0 To u
        If InStr(t, p(i)) > 0 Then
            t = Replace$(t, p(i), "")
            s = s & p(i) & vbCrLf
        End If
    Next

    If Right$(s, 2) = vbCrLf Then
        s = Left$(s, Len(s) - 2)
    End If

    DeleteDuplicateLines = s
End Function

'--------------------------------------------------
' Procedure   : DeleteFrom
' Description : ½«Str´ÓµÚpos¸öÎ»ÖÃ¿ªÊ¼É¾³ýlength¸ö×Ö·û
' CreateTime  : 2010-11-16-09:46:03
'
' ParamList   : Str (String)    Ä¿±ê×Ö·û´®
'               Pos (Long)      ÆðÊ¼Î»ÖÃ
'               Length (Long)   É¾³ýµÄ³¤¶È
' Return      : ·µ»ØÐÂµÄ×Ö·û´®
'--------------------------------------------------
Public Sub DeleteFrom(ByRef str As String, ByVal Pos As Long, ByVal Length As Long)
    Dim l  As String
    Dim ln As Long
    
    ln = Len(str)

    If Pos > ln Or Pos < 0 Then
        Exit Sub
    End If
    
    l = Left$(str, Pos)

    If (ln - Pos - Length) <= 0 Then
        str = l
    Else
        str = l & Right$(str, ln - Pos - Length)
    End If
End Sub

'--------------------------------------------------
' Procedure   : DeleteFromReverse
' Description : ½«str´Óµ¹ÊýµÚpos¸öÎ»ÖÃ¿ªÊ¼É¾³ýlength¸ö×Ö·û
' CreateTime  : 2010-11-16-09:47:04
'
' ParamList   : [ByRef] str (String)
'               Pos (Long)
'               Length (Long)
'--------------------------------------------------
Public Sub DeleteFromReverse(ByRef str As String, ByVal Pos As Long, ByVal Length As Long)
    Dim l  As String
    Dim R  As String
    Dim ln As Long
    
    ln = Len(str)

    If Pos > ln Or Pos < 0 Then
        Exit Sub
    End If
    
    l = Left$(str, ln - Pos)
    
    If (Pos - Length) <= 0 Then
        str = l
    Else
        R = Right$(str, Pos - Length)
        str = l & R
    End If
End Sub

'--------------------------------------------------
' Procedure   : EndsWith
' Description : ²âÊÔstrÊÇ·ñÒÔsuffix½áÊø
' CreateTime  : 2010-11-16-09:48:06
'
' ParamList   : Str (String)        Ä¿±ê×Ö·û´®
'               suffix (String)     Òª²âÊÔµÄºó×º
'               ignoreCase (Boolean = True)  ÊÇ·ñºöÂÔ´óÐ¡Ð´
' Return      : ÒÔSuffix½áÊø ·µ»Øtrue£¬·ñÔò·µ»ØFalse
'--------------------------------------------------
Public Function EndsWith(ByVal str As String, ByVal Suffix As String, Optional IgnoreCase As Boolean = True) As Boolean
    On Error GoTo Err

    If IgnoreCase Then
        EndsWith = (StrComp(Right$(str, Len(Suffix)), Suffix, vbTextCompare) = 0)
    Else
        EndsWith = (Right$(str, Len(Suffix)) = Suffix)
    End If

Err:
End Function

'--------------------------------------------------
' Procedure   : FillSpace
' Description : Ïò×Ö·û´®Ìî³ä¿Õ¸ñ
' CreateTime  : 2010-11-16-09:49:05
'
' ParamList   : Str (String)
'               Length (Integer)
'               FillAtLeft (Boolean = False)    ÊÇ·ñ½«¿Õ¸ñÌî³äµ½×ó±ß
' Return      : ·µ»Ø½«Str¸ñÊ½»¯Îª³¤¶ÈÎªlength¸ö×Ö½Ú¡¢ÒÔ¿Õ¸ñÏòÇ°»òÏòºóÌî³äµÄ×Ö·û´®
'--------------------------------------------------
Public Function FillSpace(ByVal str As String, ByVal Length As Integer, Optional ByVal FillAtLeft As Boolean = False) As String
    FillSpace = IIf(FillAtLeft, Space$(Length - StrLen(str)) & str, str & Space$(Length - StrLen(str)))
End Function

'--------------------------------------------------
' Procedure   : FormatLong
' Description : ¸ñÊ½»¯Ò»¸ö³¤ÕûÐÎ
' CreateTime  : 2010-11-16-09:51:24
'
' ParamList   : lng (Long)
'               Length (Integer)
'               FillAtLeft (Boolean = True)
' Return      : ·µ»Ø½«lng¸ñÊ½»¯Îª³¤¶ÈÎªlength¸ö×Ö½Ú¡¢ÒÔ¿Õ¸ñÏòÇ°»òÏòºóÌî³äµÄ×Ö·û´®
'--------------------------------------------------
Public Function FormatLong(ByVal LongValue As Long, ByVal Length As Integer, Optional ByVal FillAtLeft As Boolean = True) As String
    FormatLong = IIf(FillAtLeft, Space$(Length - StrLen(LongValue)) & LongValue, LongValue & Space$(Length - StrLen(LongValue)))
End Function

'--------------------------------------------------
' Procedure   : FormatString
' Description : ·µ»Ø¸ñÊ½»¯×Ö·û´®£¬Ô´×Ö·û´®ÖÐµÄ{x}½«±»Ìæ»»ÎªFormats(x)¶ÔÓ¦µÄÖµ
' CreateTime  : 2010-11-16-10:38:11
'
' ParamList   : Str (String)
'               Formats() (Variant)
' Return      :
'--------------------------------------------------
Public Function FormatString(str As String, ParamArray Formats() As Variant) As String
    Dim s As String
    Dim i As Long
    Dim u As Long
    
    If Len(str) = 0 Then
        Exit Function
    End If
    
    s = str
    u = UBound(Formats)

    For i = 0 To u
        s = Replace(s, "{" & i & "}", Formats(i))
    Next
    
    FormatString = s
End Function

'½«ÍøÒ³ÉÏÏÔÊ¾µÄÊ±¼ä×ª»»Îª±ê×¼¸ñÊ½
Public Function FormatWebTime(ByVal sTime As String, ByVal CurrentTime As Date) As String
    Dim s As String
    Dim v As Long
    Dim w As Date
    Dim t As String
    
    If InStr(1, sTime, "ÃëÇ°") > 0 Then
        s = Replace$(sTime, "ÃëÇ°", "")

        FormatWebTime = Format$(CurrentTime, "yyyy-MM-dd HH:mm")

        Exit Function
    ElseIf InStr(1, sTime, "·ÖÖÓÇ°") > 0 Then
        s = Replace$(sTime, "·ÖÖÓÇ°", "")
        v = Trim$(s)
        w = DateAdd("n", -v, CurrentTime)
    ElseIf InStr(1, sTime, "Ð¡Ê±Ç°") > 0 Then
        s = Replace$(sTime, "Ð¡Ê±Ç°", "")
        v = Trim$(s)
        w = DateAdd("h", -v, CurrentTime)
    ElseIf InStr(1, sTime, "½ñÌì") > 0 Then
        s = Replace$(sTime, "½ñÌì", "")
        t = Trim$(s)

        FormatWebTime = Format$(CurrentTime, "yyyy-MM-dd") & " " & t

        Exit Function
    ElseIf InStr(1, sTime, "×òÌì") > 0 Then
        s = Replace$(sTime, "×òÌì", "")
        t = Trim$(s)
        w = DateAdd("d", -1, CurrentTime)

        FormatWebTime = Format$(w, "yyyy-MM-dd") & " " & t

        Exit Function
    ElseIf InStr(1, sTime, "Ç°Ìì") > 0 Then
        s = Replace$(sTime, "Ç°Ìì", "")
        t = Trim$(s)
        w = DateAdd("d", -2, CurrentTime)

        FormatWebTime = Format$(w, "yyyy-MM-dd") & " " & t

        Exit Function
    Else
        FormatWebTime = sTime

        Exit Function
    End If
    
    FormatWebTime = Format$(w, "yyyy-MM-dd HH:mm")
End Function

'--------------------------------------------------
' Procedure   : GetAgeByBirthday
' Description : Í¨¹ýÉúÈÕµÃµ½ÄêÁä¡£²ÎÊýÈô·ÇÈÕÆÚÔò·µ»Ø0£¬·ñÔò·µ»Ø²ÎÊýÈÕÆÚ±íÊ¾µÄÄêÁä
' CreateTime  : 2010-11-16-09:52:13
'
' ParamList   : Birthday (String)
' Return      : ²ÎÊý²»ÊÇÈÕÆÚ£¬·µ»Ø0£¬·ñÔò·µ»Ø²ÎÊýÈÕÆÚ±íÊ¾µÄÄêÁä
'--------------------------------------------------
Public Function GetAgeByBirthday(Birthday As String) As Long
    If IsDate(Birthday) Then
        If CDate(Birthday) > Now Then
            Exit Function '--³öÉúÈÕÆÚ´óÓÚÏÖÔÚ ³ö´í
        End If
        
        GetAgeByBirthday = DateDiff("yyyy", Birthday, Now) + 1
    End If
End Function

'--------------------------------------------------
' Procedure   : GetBirthdayByIdCard
' Description : ´ÓÉí·ÝÖ¤ºÅÂëÖÐ¶ÁÈ¡ÉúÈÕ
' CreateTime  : 2010-11-16-09:53:01
'
' ParamList   : IdCardNumber (String)
' Return      :
'--------------------------------------------------
Public Function GetBirthdayByIdCard(IdCardNumber As String) As String
    GetBirthdayByIdCard = Mid$(IdCardNumber, 7, 4) & "-" & Mid$(IdCardNumber, 11, 2) & "-" & Mid$(IdCardNumber, 13, 2)
End Function

'--------------------------------------------------
' Procedure   : GetCNSpell
' Description : »ñÈ¡ºº×ÖÆ´ÒôÊ××ÖÄ¸
' CreateTime  : 2010-11-16-09:53:11
'
' ParamList   : Str (String)
' Return      :
'--------------------------------------------------
Public Function GetCNSpell(ByVal str As String) As String
    If Len(str) = 0 Then
        GetCNSpell = ""
        Exit Function
    End If
    
    If Asc(str) < 0 Then
        If Asc(Left$(str, 1)) < Asc("°¡") Then
            GetCNSpell = "0"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("°¡") And Asc(Left$(str, 1)) < Asc("°Å") Then
            GetCNSpell = "A"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("°Å") And Asc(Left$(str, 1)) < Asc("²Á") Then
            GetCNSpell = "B"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("²Á") And Asc(Left$(str, 1)) < Asc("´î") Then
            GetCNSpell = "C"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("´î") And Asc(Left$(str, 1)) < Asc("¶ê") Then
            GetCNSpell = "D"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("¶ê") And Asc(Left$(str, 1)) < Asc("·¢") Then
            GetCNSpell = "E"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("·¢") And Asc(Left$(str, 1)) < Asc("¸Á") Then
            GetCNSpell = "F"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("¸Á") And Asc(Left$(str, 1)) < Asc("¹þ") Then
            GetCNSpell = "G"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("¹þ") And Asc(Left$(str, 1)) < Asc("»÷") Then
            GetCNSpell = "H"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("»÷") And Asc(Left$(str, 1)) < Asc("¿¦") Then
            GetCNSpell = "J"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("¿¦") And Asc(Left$(str, 1)) < Asc("À¬") Then
            GetCNSpell = "K"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("À¬") And Asc(Left$(str, 1)) < Asc("Âè") Then
            GetCNSpell = "L"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("Âè") And Asc(Left$(str, 1)) < Asc("ÄÃ") Then
            GetCNSpell = "M"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("ÄÃ") And Asc(Left$(str, 1)) < Asc("Å¶") Then
            GetCNSpell = "N"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("Å¶") And Asc(Left$(str, 1)) < Asc("Å¾") Then
            GetCNSpell = "O"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("Å¾") And Asc(Left$(str, 1)) < Asc("ÆÚ") Then
            GetCNSpell = "P"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("ÆÚ") And Asc(Left$(str, 1)) < Asc("È»") Then
            GetCNSpell = "Q"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("È»") And Asc(Left$(str, 1)) < Asc("Èö") Then
            GetCNSpell = "R"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("Èö") And Asc(Left$(str, 1)) < Asc("Ëú") Then
            GetCNSpell = "S"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("Ëú") And Asc(Left$(str, 1)) < Asc("ÍÚ") Then
            GetCNSpell = "T"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("ÍÚ") And Asc(Left$(str, 1)) < Asc("Îô") Then
            GetCNSpell = "W"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("Îô") And Asc(Left$(str, 1)) < Asc("Ñ¹") Then
            GetCNSpell = "X"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("Ñ¹") And Asc(Left$(str, 1)) < Asc("ÔÑ") Then
            GetCNSpell = "Y"
            Exit Function
        End If
        If Asc(Left$(str, 1)) >= Asc("ÔÑ") Then
            GetCNSpell = "Z"
            Exit Function
        End If
    ElseIf UCase$(str) <= "Z" And UCase$(str) >= "A" Then
        GetCNSpell = UCase$(Left$(str, 1))
    Else
        GetCNSpell = str
    End If
End Function

'--------------------------------------------------
' Procedure   : GetCNSpells
' Description : »ñÈ¡×Ö·û´®µÄËùÓÐ×Ö·û´®Æ´Òô
' CreateTime  : 2010-11-16-09:53:22
'
' ParamList   : mystr (String)
' Return      :
'--------------------------------------------------
Public Function GetCNSpells(ByVal str As String) As String
    Dim i As Integer
    Dim a As Integer
    Dim l As Long
    Dim R As String
    
    l = Len(str)

    If l = 0 Then
        GetCNSpells = ""
        Exit Function
    End If
    
    R = Space$(l)
    
    For i = 1 To l
        a = Asc(Mid$(str, i, 1))

        If a < 0 Then
            If a < Asc("°¡") Then
                Mid$(R, i, 1) = "0"
            ElseIf a >= Asc("°¡") And a < Asc("°Å") Then
                Mid$(R, i, 1) = "A"
            ElseIf a >= Asc("°Å") And a < Asc("²Á") Then
                Mid$(R, i, 1) = "B"
            ElseIf a >= Asc("²Á") And a < Asc("´î") Then
                Mid$(R, i, 1) = "C"
            ElseIf a >= Asc("´î") And a < Asc("¶ê") Then
                Mid$(R, i, 1) = "D"
            ElseIf a >= Asc("¶ê") And a < Asc("·¢") Then
                Mid$(R, i, 1) = "E"
            ElseIf a >= Asc("·¢") And a < Asc("¸Á") Then
                Mid$(R, i, 1) = "F"
            ElseIf a >= Asc("¸Á") And a < Asc("¹þ") Then
                Mid$(R, i, 1) = "G"
            ElseIf a >= Asc("¹þ") And a < Asc("»÷") Then
                Mid$(R, i, 1) = "H"
            ElseIf a >= Asc("»÷") And a < Asc("¿¦") Then
                Mid$(R, i, 1) = "J"
            ElseIf a >= Asc("¿¦") And a < Asc("À¬") Then
                Mid$(R, i, 1) = "K"
            ElseIf a >= Asc("À¬") And a < Asc("Âè") Then
                Mid$(R, i, 1) = "L"
            ElseIf a >= Asc("Âè") And a < Asc("ÄÃ") Then
                Mid$(R, i, 1) = "M"
            ElseIf a >= Asc("ÄÃ") And a < Asc("Å¶") Then
                Mid$(R, i, 1) = "N"
            ElseIf a >= Asc("Å¶") And a < Asc("Å¾") Then
                Mid$(R, i, 1) = "O"
            ElseIf a >= Asc("Å¾") And a < Asc("ÆÚ") Then
                Mid$(R, i, 1) = "P"
            ElseIf a >= Asc("ÆÚ") And a < Asc("È»") Then
                Mid$(R, i, 1) = "Q"
            ElseIf a >= Asc("È»") And a < Asc("Èö") Then
                Mid$(R, i, 1) = "R"
            ElseIf a >= Asc("Èö") And a < Asc("Ëú") Then
                Mid$(R, i, 1) = "S"
            ElseIf a >= Asc("Ëú") And a < Asc("ÍÚ") Then
                Mid$(R, i, 1) = "T"
            ElseIf a >= Asc("ÍÚ") And a < Asc("Îô") Then
                Mid$(R, i, 1) = "W"
            ElseIf a >= Asc("Îô") And a < Asc("Ñ¹") Then
                Mid$(R, i, 1) = "X"
            ElseIf a >= Asc("Ñ¹") And a < Asc("ÔÑ") Then
                Mid$(R, i, 1) = "Y"
            ElseIf a >= Asc("ÔÑ") And a <= Asc("×ù") Then
                Mid$(R, i, 1) = "Z"
            Else
                On Error Resume Next
                Mid$(R, i, 1) = Mid$(GetCNSpellSecondEng, InStr(1, GetCNSpellSecondChn, Mid$(str, i, 1), vbBinaryCompare), 1)
            End If
        ElseIf UCase$(str) <= "Z" And UCase$(str) >= "A" Then
            Mid$(R, i, 1) = UCase$(Left$(str, 1))
        Else
            Mid$(R, i, 1) = Mid$(str, i, 1)
        End If
    Next
    
    GetCNSpells = R
End Function

'--------------------------------------------------
' Procedure   : GetGenderByIdCard
' Description : ´ÓÉí·ÝÖ¤ºÅÂëÅÐ¶ÏÐÔ±ð
' CreateTime  : 2010-11-16-09:57:11
'
' ParamList   : IdCardNumber (String)
' Return      : ·µ»Ø1±íÊ¾ÄÐ ·µ»Ø2±íÊ¾Å® ·µ»Ø0±íÊ¾½âÎö´íÎó
'--------------------------------------------------
Public Function GetGenderByIdCard(IdCardNumber As String) As Long
    On Error GoTo ErrHandle

    If Mid$(IdCardNumber, 17, 1) Mod 2 = 1 Then
        GetGenderByIdCard = 1   '--ÆæÊýÎªÄÐ
    Else
        GetGenderByIdCard = 2   '--Å¼ÊýÎªÅ®
    End If

ErrHandle:  '--´íÎóÎª0
End Function

'»ñÈ¡TextÖÐ³¤¶È´óÓÚµÈÓÚLengthµÄÁ¬ÐøÊý×Ö£¬Êý×ÖÐòÁÐÖÐ°üº¬µÄIgnoreCharsÖÐµÄÈÎÒâ×Ö·û½«±»ºöÂÔ
'È«½Ç½«×Ô¶¯×ª»»³É°ë½Ç
'Èô»ñÈ¡µ½¶à¸öÊý×ÖÐòÁÐ£¬ÔòÓÃ|·Ö¸ôºó·µ»Ø
Public Function GetNumberSerial(ByVal Text As String, Optional ByVal Length As Long = 5, Optional ByVal IgnoreChars As String = " -/\_:;,.", Optional ByVal JoinBy As String = "|") As String
    Dim p()    As String
    Dim i      As Long
    Dim lStart As Long
    Dim lLen   As Long
    Dim str1   As String
    Dim n      As Long
    Dim t      As Long
    
    p = Split("")
    Text = StrConv(Text, vbNarrow)
    
    n = Len(IgnoreChars)

    For i = 1 To n
        Text = Replace$(Text, Mid$(IgnoreChars, i, 1), "")
    Next
    
    t = Len(Text)

    For i = 1 To t
        str1 = Mid$(Text, i, 1)

        If InStr("0123456789", str1) > 0 Then
            If lStart = 0 Then
                lStart = i
                lLen = 1
            Else
                lLen = lLen + 1
            End If
        Else
            If lLen >= Length Then
                ReDim Preserve p(UBound(p) + 1)
                p(UBound(p)) = Mid$(Text, lStart, lLen)
            End If

            lStart = 0
            lLen = 0
        End If
    Next

    If lLen >= Length Then
        ReDim Preserve p(UBound(p) + 1)
        p(UBound(p)) = Mid$(Text, lStart, lLen)
    End If

    GetNumberSerial = Join(p, JoinBy)
End Function

'»ñÈ¡×Ö·û´®ÖÐ×Ó×Ö·û´®
Public Function GetSubString(ByVal SourceString As String, _
                             Optional ByVal ParentStart As String = "", _
                             Optional ByVal ParentCenter As String = "", _
                             Optional ByVal ParentEnd As String = "", _
                             Optional ByVal SubStart As String = "", _
                             Optional ByVal SubCenter As String = "", _
                             Optional ByVal SubEnd As String = "", _
                             Optional ByVal ArrayIndex As Long = 0) As String
    Dim sText     As String
    Dim str1      As String
    Dim OKCount   As Long
    Dim lBegin    As Long
    Dim CurBegin  As Long
    Dim pResult() As String
    Dim k         As Long

    If Len(SourceString) = 0 Then Exit Function
    If ArrayIndex < -1 Then ArrayIndex = -1

    sText = LCase$(SourceString)

    ParentStart = LCase$(ParentStart)
    ParentCenter = LCase$(ParentCenter)
    ParentEnd = LCase$(ParentEnd)
    
    SubStart = LCase$(SubStart)
    SubCenter = LCase$(SubCenter)
    SubEnd = LCase$(SubEnd)

    pResult = Split("")

    If Len(ParentStart) = 0 Then
        ParentStart = SubStart
        SubStart = ""
    End If

    Do
        lBegin = InStr(lBegin + 1, sText, ParentStart)

        If lBegin = 0 Then Exit Do

        lBegin = lBegin + Len(ParentStart)
        CurBegin = lBegin

        str1 = Right$(sText, Len(sText) - lBegin + 1)

        If Len(ParentEnd) > 0 Then 'ÅÐ¶ÏÎ²²¿×Ö·û
            k = InStr(str1, ParentEnd)

            If k = 0 Then Exit Do

            str1 = Left$(str1, k - 1)
            lBegin = lBegin + k + Len(ParentEnd) - 2
        End If
        
        If Len(ParentCenter) > 0 Then 'ÅÐ¶ÏÖÐ¼äÊÇ·ñ´æÔÚ¹Ø¼ü´Ê
            k = InStr(str1, ParentCenter)

            If k = 0 Then GoTo ToNext
        End If

        If Len(SubStart) > 0 Then 'ÅÐ¶Ï¿ªÊ¼×Ö·û
            k = InStr(str1, SubStart)

            If k = 0 Then GoTo ToNext

            CurBegin = CurBegin + k + Len(SubStart) - 1
            str1 = Right$(str1, Len(str1) - k - Len(SubStart) + 1)
        End If

        If Len(SubEnd) > 0 Then 'ÅÐ¶Ï½áÊø×Ö·û
            k = InStr(str1, SubEnd)

            If k = 0 Then GoTo ToNext

            str1 = Left$(str1, k - 1)
        End If

        If Len(SubCenter) > 0 Then 'ÅÐ¶ÏÖÐ¼äÊÇ·ñ´æÔÚ¹Ø¼ü´Ê
            k = InStr(str1, SubCenter)

            If k = 0 Then GoTo ToNext
        End If

        If OKCount = ArrayIndex Or ArrayIndex = -1 Then
            If Len(str1) > 0 Then
                str1 = Mid$(SourceString, CurBegin, Len(str1))

                If ArrayIndex = -1 Then
                    ReDim Preserve pResult(UBound(pResult) + 1)
                    pResult(UBound(pResult)) = str1

                Else
                    ReDim pResult(0)
                    pResult(0) = str1

                    Exit Do

                End If
            End If
        End If

        OKCount = OKCount + 1

ToNext:
        DoEvents
    Loop

    str1 = Join(pResult, vbCrLf)
    GetSubString = str1
End Function

'HTML½âÃÜ
Public Function HTMLDecode(ByVal sText As String)
    Dim i As Long
    
    sText = Replace$(sText, "&quot;", Chr$(34))
    sText = Replace$(sText, "&lt;", Chr$(60))
    sText = Replace$(sText, "&gt;", Chr$(62))
    sText = Replace$(sText, "&amp;", Chr$(38))
    sText = Replace$(sText, "&nbsp;", Chr$(32))
    
    For i = 1 To 255
        sText = Replace$(sText, "&#" & i & ";", Chr$(i))
    Next

    HTMLDecode = sText
End Function

'--------------------------------------------------
' Procedure   : IndexOf
' Description : ·µ»ØcharÔÚstrÖÐµÚÒ»´Î³öÏÖ´¦µÄË÷Òý
' CreateTime  : 2010-11-16-09:57:45
'
' ParamList   : Str (String)
'               char (String)
' Return      :
'--------------------------------------------------
Public Function IndexOf(ByVal str As String, ByVal Chars As String) As Long
    IndexOf = InStr(1, str, Chars, vbBinaryCompare) - 1
End Function

'--------------------------------------------------
' Procedure   : InsertTo
' Description : ²åÈëÒ»¸ö×Ö·û´®µ½Ä¿±ê×Ö·û´®ÖÐ
' CreateTime  : 2010-11-16-09:58:00
'
' ParamList   : [ByRef] Str (String)
'               Pos (Long)
'               Insert (String)
' Return      : ÔÚstrÖÐposÎ»ÖÃ²åÈëinsert
'--------------------------------------------------
Public Sub InsertTo(ByRef str As String, ByVal Pos As Long, ByVal Insert As String)
    Dim sl As String
    Dim sr As String
    
    sl = Left$(str, Pos)
    sr = Right$(str, Len(str) - Pos)
    
    str = sl & Insert & sr
End Sub

'--------------------------------------------------
' Procedure   : InsertToReverse
' Description : ²åÈëÒ»¸ö×Ö·û´®µ½Ä¿±ê×Ö·û´®ÖÐ
' CreateTime  : 2010-11-16-09:58:30
'
' ParamList   : Str (String)
'               Pos (Long)
'               Insert (String)
' Return      : ÔÚstrÖÐµ¹ÊýposÎ»ÖÃ²åÈëinsert
'--------------------------------------------------
Public Sub InsertToReverse(ByRef str As String, ByVal Pos As Long, ByVal Insert As String)
    Dim sl As String
    Dim sr As String

    sl = Left$(str, Len(str) - Pos)
    sr = Right$(str, Pos)
    
    str = sl & Insert & sr
End Sub

'--------------------------------------------------
' Procedure   : IsChineseChar
' Description : ²âÊÔAscII´ú±íµÄ×Ö·ûÊÇ·ñÎªºº×Ö
' CreateTime  : 2010-11-16-09:58:48
'
' ParamList   : AscII (Integer)
' Return      :
'--------------------------------------------------
Public Function IsChineseChar(ByVal AscII As Integer) As Boolean
    If Len(Hex$(AscII)) > 2 Then
        IsChineseChar = True
    Else
        IsChineseChar = False
    End If
End Function

'--------------------------------------------------
' Procedure   : IsIdentifier
' Description : ²âÊÔ×Ö·û´®ÊÇ·ñÊÇºÏ·¨µÄ±êÊ¾·û
' CreateTime  : 2010-11-16-09:59:31
'
' ParamList   : Str (String)
' Return      :
'--------------------------------------------------
Public Function IsIdentifier(str As String) As Boolean
    Dim i As Long
    Dim s As Boolean
    Dim l As Long
    
    If str = "" Then Exit Function
    
    str = LCase$(str)
    l = Len(str)

    For i = 1 To l
        If Not s Then
            If InStr(i, "abcdefghijklmnopqrstuvwxyz_", Mid$(str, i, 1), VbCompareMethod.vbBinaryCompare) > 0 Then
                s = True
            Else
                '"¿ªÍ·²»ºÏ·¨"
                Exit Function
            End If
        End If
        
        If s Then
            If InStr(1, "abcdefghijklmnopqrstuvwxyz0123456789_", Mid$(str, i, 1), vbBinaryCompare) = 0 Then
                '"ÖÐ¼ä²»ºÏ·¨"
                Exit Function
            End If
        End If
    Next

    IsIdentifier = True
End Function


'--------------------------------------------------
' Procedure   : IsLetter
' Description : ²âÊÔ×Ö·ûÊÇ·ñÎªA~Z
' CreateTime  : 2010-11-16-10:00:41
'
' ParamList   : Str (String)
' Return      :
'--------------------------------------------------
Public Function IsLetter(str As String) As Boolean
    Dim c As Integer

    c = Asc(LCase$(str))

    If c >= Asc("a") And c <= Asc("z") Then IsLetter = True
End Function

'ÅÐ¶Ï¶þ½øÖÆÍøÒ³±àÂëÊÇ·ñutf8
Public Function IsUTF8(ByRef Bytes() As Byte) As Boolean
    Dim i As Long
    Dim AscN As Long
    Dim Length As Long

    Length = UBound(Bytes) + 1

    If Length < 3 Then
        IsUTF8 = False
        Exit Function
    ElseIf Bytes(0) = &HEF And Bytes(1) = &HBB And Bytes(2) = &HBF Then
        IsUTF8 = True
        Exit Function
    End If

    Do While i <= Length - 1
        If Bytes(i) < 128 Then
            i = i + 1
            AscN = AscN + 1
        ElseIf (Bytes(i) And &HE0) = &HC0 And (Bytes(i + 1) And &HC0) = &H80 Then
            i = i + 2
        ElseIf i + 2 < Length Then
            If (Bytes(i) And &HF0) = &HE0 And (Bytes(i + 1) And &HC0) = &H80 And (Bytes(i + 2) And &HC0) = &H80 Then
                i = i + 3
            Else
                IsUTF8 = False
                Exit Function
            End If
        Else
            IsUTF8 = False
            Exit Function
        End If
    Loop

    IsUTF8 = IIf(AscN = Length, False, True)
End Function

'--------------------------------------------------
' Procedure   : LastIndexOf
' Description : ·µ»ØcharÔÚstrÖÐ×îºóÒ»´Î³öÏÖ´¦µÄË÷Òý
' CreateTime  : 2010-11-16-10:00:54
'
' ParamList   : Str (String)
'               char (String)
' Return      :
'--------------------------------------------------
Public Function LastIndexOf(ByVal str As String, ByVal Chars As String) As Long
    LastIndexOf = InStrRev(str, Chars, , vbBinaryCompare) - 1
End Function

'¿ìËÙ¶ÁÈ¡ÎÄ¼þÎÄ¼þ
Public Function LoadText(ByVal FilePath As String) As String
    Dim str1 As String
    Dim fn   As Long
    Dim k    As Long

    On Error GoTo IsErr

    If Mid$(FilePath, 2, 2) <> ":\" Then FilePath = Replace$(App.Path & "\", "\\", "\") & FilePath
    If Len(Dir$(FilePath, vbHidden Or vbNormal)) = 0 Then Exit Function

    fn = FreeFile
    Open FilePath For Binary As #fn
    str1 = Space$(LOF(fn))
    Get #fn, , str1
    Close #fn
    k = InStr(str1, Chr$(0))

    If k > 0 Then str1 = Left$(str1, k - 1)

    LoadText = str1
IsErr:

    If Err Then Err.Clear
End Function

'»ñÈ¡Ô´×Ö·û´®ÖÐÁ½¸ö×Ö·û´®ÖÐ¼äµÄ×Ó×Ö·û´®
Public Function MidOf(ByVal Source As String, ByVal Prefix As String, ByVal Suffix As String) As String
    Dim nLft As Long
    Dim nRgt As Long
    Dim nLen As Long
    
    nLen = Len(Prefix)

    If nLen > 0 Then
        nLft = InStr(1, Source, Prefix)

        If nLft > 0 Then
            nLft = nLft + nLen
        End If
    End If
    
    If nLft > 0 Then
        nRgt = InStr(nLft, Source, Suffix)
        
        If nRgt > nLft Then
            MidOf = Mid$(Source, nLft, nRgt - nLft)
        End If
    End If
End Function

'Êä³öÎÄ±¾ÎÄ¼þ
Public Function OutText(ByVal Path As String, ByVal Text As String, Optional ByVal bAppend As Boolean = False) As Boolean
    Dim k As Long

    On Error GoTo ToExit

    If Mid$(Path, 2, 2) <> ":\" Then Path = Replace$(App.Path & "\", "\\", "\") & Path
    k = FreeFile

    If bAppend = True Then
        Open Path For Append As #k
    Else
        Open Path For Output As #k
    End If

    If bAppend = True Then
        Print #k, Text
    Else
        Print #k, Text;
    End If

    Close #k

    OutText = True

    Exit Function
ToExit:
    On Error GoTo 0
End Function

'--------------------------------------------------
' Procedure   : RemoveQuotedString
' Description : É¾³ý³É¶Ô³öÏÖµÄ·ûºÅ¼°ÆäÖÐ°üº¬µÄ×Ö·û´®
' CreateTime  : 2010-11-16-10:01:23
'
' ParamList   : sString (String)                ÊäÈëµÄ×Ö·û´®
'               QuotationMark (String = "'")    Òª¼ì²âµÄ·ûºÅ
' Return      : ²Ù×÷ºóµÃµ½µÄ×Ö·û´®
'--------------------------------------------------
Public Function RemoveQuotedString(sString As String, Optional QuotationMark As String = "'") As String
    Dim i       As Long
    Dim j       As Long
    Dim buf     As String
    Dim c       As String
    Dim Length  As Long
    Dim inQuote As Boolean
    
    Length = Len(sString)

    If Length = 0 Or Len(QuotationMark) = 0 Then
        RemoveQuotedString = sString
        Exit Function
    End If
    
    If InStr(1, sString, QuotationMark) = 0 Then
        RemoveQuotedString = sString
        Exit Function
    End If
    
    buf = String$(Length, vbNullChar)
    
    For i = 1 To Length
        c = Mid$(sString, i, 1)
        
        If inQuote Then
            If c = QuotationMark Then
                inQuote = False
            End If
        Else
            If c = QuotationMark Then
                inQuote = True
            Else
                j = j + 1
                Mid$(buf, j, 1) = c
            End If
        End If
    Next
    
    RemoveQuotedString = Left$(buf, InStr(buf, vbNullChar))
End Function

'*************************************************************************
'¹¦ÄÜÃèÊö£ºÌæ»»Á½¸ö¹Ø¼ü×ÖÖ®¼äµÄ×Ö·û
'º¯ Êý Ãû£ºReplaceMidKey
'sText(String) - ÒªÇå³ýµÄ×Ö·û´®
'Optional ByVal BeginKey(String) - [¿ªÊ¼×Ö·û]
'Optional ByVal MidKey(String) - [ÖÐ¼ä°üº¬×Ö·û]
'Optional ByVal EndKey(String) - [½áÊø×Ö·û]
'Optional ByVal KeepKey(Boolean = False) - [ÊÇ·ñ±£Áô ¿ªÊ¼/½áÊø ×Ö·û]
'Êä ³ö£º(String) - È¥³ýºóµÄ×Ö·û´®
'Àý ×Ó£º ReplaceMidKey("<a>1</a><a>2</a><a>3</a>","<a>","1","</a>") ·µ»Ø "<a>2</a><a>3</a>"
'×÷ Õß£º¸ñÊ½»¯ QQ:65464145
'ÈÕ ÆÚ£º2010-04-23 15:39:56
'*************************************************************************
Public Function ReplaceMidKey(ByVal Text As String, Optional ByVal BeginKey As String = "", Optional ByVal MidKey As String = "", Optional ByVal EndKey As String = "", Optional ByVal KeepKey As Boolean = False) As String
    Dim pResult()    As String
    Dim k            As Long
    Dim k1           As Long
    Dim k2           As Long
    Dim lText        As String
    Dim str1         As String
    Dim bInstrMidKey As Boolean
    Dim lBeginKey    As Long
    Dim lEndKey      As Long
    Dim lCount       As Long
    Dim lMax         As Long
    
    '========= Ð¡Ð´´¦Àí
    lText = LCase$(Text)
    BeginKey = LCase$(BeginKey)
    MidKey = LCase$(MidKey)

    EndKey = LCase$(EndKey)
    
    '========= ³õÊ¼»¯
    If Len(MidKey) > 0 Then
        bInstrMidKey = True
    End If
    
    lBeginKey = Len(BeginKey)
    lEndKey = Len(EndKey)
    
    lMax = 1000
    ReDim pResult(lMax)
    lCount = -1
    
    If Len(BeginKey) > 0 Or Len(EndKey) > 0 Then
        k = 1
        
        Do
            '===== ¿ªÊ¼Î»ÖÃ
            k1 = InStr(k, lText, BeginKey)

            If k1 = 0 Then Exit Do
            
            '===== ½áÊøÎ»ÖÃ
            k2 = InStr(k1 + lBeginKey, lText, EndKey)

            If k2 = 0 Then Exit Do
            
            k2 = k2 + lEndKey ' ¼ÓÉÏ½áÊø×Ö·û³¤¶È
            
            '===== ÊÇ·ñ²éÕÒ°üº¬×Ö·û
            If bInstrMidKey = True Then
                str1 = Mid$(lText, k1 + lBeginKey, k2 - k1 - lBeginKey - lEndKey)
                
                If InStr(str1, MidKey) = 0 Then
                    k1 = k2
                End If
            End If
            
            '===== ½á¹û
            str1 = Mid$(Text, k, k1 - k)

            If KeepKey = True And k1 <> k2 Then
                str1 = str1 & BeginKey & EndKey
            End If
            
            lCount = lCount + 1

            If lCount >= lMax Then
                lMax = lMax + 1000
                ReDim Preserve pResult(lMax)
            End If
            
            pResult(lCount) = str1
            
            k = k2

            DoEvents
        Loop
    
        '===== É¨Î²
        lCount = lCount + 1
        pResult(lCount) = pResult(lCount) & Mid$(Text, k, Len(Text) - k + 1)
    End If
    
    If lCount >= 0 Then
        ReDim Preserve pResult(lCount)
        ReplaceMidKey = Join(pResult, "")
    End If
End Function

'Ìæ»»HTML×Ö·ûÊµÌå
Public Function ReplaceHTMLEntity(ByVal Text As String) As String
    Dim i As Long
    Dim p() As String
    Dim en As Long
    Dim c As String
    Dim f As Long
    Dim m As Long
    Dim s As String
    Dim d As Long
    Dim w As String
    
    f = InStr(1, Text, "&#x")
    
    If f > 0 Then
        m = Len(Text)
        For i = 1 To m
            en = InStr(i, Text, "&#x")
            If en = 0 Then Exit For
            If i <> en Then
                s = s & Mid$(Text, i, en - i)
            End If
            d = InStr(en, Text, ";")
            c = Mid$(Text, en + 3, d - en - 3)
            w = ChrW("&H" & c)
            s = s & w
            
            i = d
            
            If m - i < 3 Then
                s = s & Right$(Text, m - d)
                Exit For
            End If
            
        Next
        ReplaceHTMLEntity = s
        Exit Function
    Else
        f = InStr(1, Text, "&#")
        If f > 0 Then
            m = Len(Text)
            For i = 1 To m
                en = InStr(i, Text, "&#")
                If en = 0 Then Exit For
                If i <> en Then
                    s = s & Mid$(Text, i, en - i)
                End If
                
                d = InStr(en, Text, ";")
                c = Mid$(Text, en + 2, d - en - 2)
                w = ChrW(c)
                s = s & w
                i = d
                If m - i < 2 Then
                    s = s & Right$(Text, m - d)
                    Exit For
                End If
            Next
            
            ReplaceHTMLEntity = s
            Exit Function
        End If
    End If
    
    ReplaceHTMLEntity = Text
End Function

'×Ö·û´®Êý×éÈ¥ÖØ
Public Function DistinctStringArray(StringArray As Variant) As String()
    Dim i    As Long
    Dim j    As Long
    Dim c    As Long
    Dim u    As Long
    Dim k    As Long
    Dim n    As Long
    Dim arr2 As Variant
    
    If IsArray(StringArray) = False Then Exit Function
    
    c = LBound(StringArray)
    u = UBound(StringArray)
    ReDim arr2(c To u) As String
    k = c
    
    For i = 0 To u
        n = i - 1

        For j = 0 To n
            If StringArray(i) = StringArray(j) Then Exit For
        Next

        If j = i Then
            arr2(k) = StringArray(i)
            k = k + 1
        End If
    Next
    
    ReDim Preserve arr2(c To k - 1) As String
    DistinctStringArray = arr2
End Function

'½«ºÁÃëÊý¸ñÊ½»¯ÎªÃëÊý,¾«È·µ½Ð¡ÊýµãºóPointÎ»
Public Function FormatTime(ByVal MsCount As Long, Optional ByVal POINT As Long = 2) As String
    Dim c As Double
    Dim m As Long
    
    c = MsCount / 1000
    c = Math.Round(c, POINT)
    
    If c < 60 Then
        FormatTime = c
    Else
        m = c \ 60
        c = c - (m * 60)
        c = Math.Round(c, 2)

        FormatTime = m & "·Ö" & c
    End If
End Function

'--------------------------------------------------
' Procedure   : SaveAs
' Description : ½«×Ö·û´®Áí´æÎªÎÄ¼þ
' CreateTime  : 2010-11-16-10:02:03
'
' ParamList   : SourceString (String)
'               Path (String)
' Return      :
'--------------------------------------------------
Public Function SaveAs(SourceString As String, Path As String, Optional Append As Boolean = False) As Boolean
    Dim k As Long

    On Error GoTo ToExit

    If Mid$(Path, 2, 2) <> ":\" Then Path = Replace$(App.Path & "\", "\\", "\") & Path
    k = FreeFile

    If Append Then
        Open Path For Append As #k
    Else
        Open Path For Output As #k
    End If

    If Append Then
        Print #k, SourceString
    Else
        Print #k, SourceString;
    End If

    Close #k

    SaveAs = True

    Exit Function
ToExit:
    On Error GoTo 0
End Function

'--------------------------------------------------
' Procedure   : SplitEx
' Description : ¿ìËÙ·Ö¸î×Ö·û´®
' CreateTime  : 2016-01-21 11:24
'
' ParamList   : Source (String)
'               Delimiter (String)
' Return      :
'--------------------------------------------------
Public Function SplitEx(ByVal Source As String, Optional ByVal Delimiter As String = vbCrLf) As String()
    Dim k             As Long
    Dim k1            As Long
    Dim Count         As Long
    Dim str1          As String
    Dim bExit         As Boolean
    Dim pResult()     As String
    Dim lLenDelimiter As Long

    If Len(Delimiter) > 1 Then
        lLenDelimiter = Len(Delimiter) - 1
    End If

    pResult = Split("")

    Do
        k = k + 1
        k1 = InStr(k, Source, Delimiter)

        If k1 = 0 Then
            k1 = Len(Source) - k + 1

            If k1 <= 0 Then
                str1 = ""
            Else
                str1 = Mid$(Source, k, Len(Source) - k + 1)
            End If

            bExit = True
        Else
            str1 = Mid$(Source, k, k1 - k)
        End If

        If Len(str1) > 0 Then
            ReDim Preserve pResult(Count)
            pResult(Count) = str1

            Count = Count + 1
        End If

        If bExit = True Then Exit Do
        k = k1 + lLenDelimiter

        DoEvents
    Loop

    SplitEx = pResult
End Function

'--------------------------------------------------
' Procedure   : StartsWith
' Description : ²âÊÔstrÊÇ·ñÒÔprefix¿ªÊ¼
' CreateTime  : 2010-11-16-10:02:09
'
' ParamList   : Str (String)
'               prefix (String)
'               ignoreCase (Boolean = True)
' Return      :
'--------------------------------------------------
Public Function StartsWith(ByVal str As String, ByVal Prefix As String, Optional IgnoreCase As Boolean = True) As Boolean
    On Error GoTo Err

    If IgnoreCase Then
        StartsWith = (StrComp(Left$(str, Len(Prefix)), Prefix, vbTextCompare) = 0)
    Else
        StartsWith = (Left$(str, Len(Prefix)) = Prefix)
    End If
Err:
End Function

'--------------------------------------------------
' Procedure   : StrLeft
' Description : ·µ»Ø×Ö·û´®×ó±ßµÄLength¸ö×Ö½Ú
' CreateTime  : 2010-11-16-10:02:21
'
' ParamList   : str5 (String)
'               len5 (Long)
' Return      :
'--------------------------------------------------
Public Function StrLeft(ByVal str As String, ByVal Length As Long) As String
    Dim tmpstr As String

    tmpstr = StrConv(str, vbFromUnicode)
    tmpstr = LeftB$(tmpstr, Length)
    StrLeft = StrConv(tmpstr, vbUnicode)
End Function

'--------------------------------------------------
' Procedure   : Strlen
' Description : ·µ»ØStrµÄ×Ö½ÚÊý
' CreateTime  : 2010-11-16-10:02:39
'
' ParamList   : tstr (String)
' Return      :
'--------------------------------------------------
Public Function StrLen(ByVal str As String) As Integer
    StrLen = LenB(StrConv(str, vbFromUnicode))
End Function

'--------------------------------------------------
' Procedure   : StrRight
' Description : ·µ»Ø×Ö·û´®ÓÒ±ßµÄlength¸ö×Ö½Ú
' CreateTime  : 2010-11-16-10:03:07
'
' ParamList   : str5 (String)
'               len5 (Long)
' Return      :
'--------------------------------------------------
Public Function StrRight(ByVal str As String, ByVal Length As Long) As String
    Dim tmpstr As String

    tmpstr = StrConv(str, vbFromUnicode)
    tmpstr = RightB$(tmpstr, Length)
    StrRight = StrConv(tmpstr, vbUnicode)
End Function

'--------------------------------------------------------------------------------
' Procedure  :       SubStr
' Description:       ·µ»Ø×Ö·û´®ÖÐ´ÓStartÖ®ºóµÄlength¸ö×Ö·û
' Created by :       fangzi
' Date-Time  :       2010-11-16-10:03:29
'
' Parameters :       Str (String)
'                    Start (Long)
'                    Length (Long = 0)
'--------------------------------------------------------------------------------
Public Function SubStr(ByVal str As String, ByVal Start As Long, Optional Length As Long = 0) As String
    Dim tmpstr As String

    If Length = 0 Then
        tmpstr = StrConv(MidB$(StrConv(str, vbFromUnicode), Start), vbUnicode)
    Else
        tmpstr = StrConv(MidB$(StrConv(str, vbFromUnicode), Start, Length), vbUnicode)
    End If

    SubStr = tmpstr
End Function

'--------------------------------------------------
' Procedure   : ToBinaryString
' Description : ×ª»»×Ö·û´®Îª×Ö½ÚÊý×é
' CreateTime  : 2010-11-16-10:04:25
'
' ParamList   : txt (String)
'               ret (String)
' Return      : ½«²ÎÊý×ª»»Îª¶þ½øÖÆÐÎÊ½£¬´æ·Åµ½retÖÐµÄ×Ö·û´®²¢·µ»Ø³¤¶È
'--------------------------------------------------
Public Function ToBinaryString(ByVal txt As String, ret As String) As Long
    Dim b()    As Byte         '´æ·Å²ÎÊýµÄ×Ö½ÚÊý×é
    Dim Length As Long      '×Ö½ÚÊý
        
    '½«²ÎÊý×ª»»Îª×Ö½Ú,´æÈë×Ö½ÚÊý×éÊý×é
    b = StrConv(txt, vbFromUnicode)
    
    'µÃµ½×Ö½ÚÊý
    Length = UBound(b) + 1
    
    'Éú³ÉÓÉ×Ö½ÚÊý¸ö¿Õ¸ñ×é³ÉµÄ×Ö·û´®
    ret = Space$(Length)
        
    '×Ö½ÚÊý×é³¤¶È+1
    ReDim Preserve b(0 To Length + 1) As Byte
    
    'ÔÚÄ©Î²²¹ÉÏ&H0
    b(Length) = &H0
        
    '½«×Ö½ÚÊý×é¸´ÖÆµ½ret
    CopyMemory ret, b(0), Length + 1
    
    'Êä³öret
    ToBinaryString = Length + 1
End Function

'--------------------------------------------------
' Procedure   : ToSQL
' Description : ½«×Ö·û´®×ª»»ÎªSQLÓï¾ä¸ñÊ½
' CreateTime  : 2010-11-16-10:04:45
'
' ParamList   : Str (String)
' Return      : ¸ñÊ½»¯ºóµÄ×Ö·û´®
'--------------------------------------------------
Public Function ToSQL(str As String) As Variant
    Dim s     As String
    Dim tks   As Variant
    Dim First As String
    Dim i     As Long
    Dim j     As Long
    Dim heads As Variant
    Dim keys  As Variant
    Dim isKey As Boolean
    Dim u     As Long
    Dim uk    As Long
    Dim ret   As String
    
    heads = Array("SELECT", "UPDATE", "INSERT", "DELETE", "ALTER", "DROP", "CREATE", "PROCEDURE")
    keys = Array("SELECT", "UPDATE", "INSERT", "DELETE", "ALTER", "DROP", "CREATE", "PROCEDURE", "AS", "INTO", "FROM", "GROUP", "BY", "ORDER", "ASC", "DESC", "SET", "JOIN", "LEFT", "INNER", "RIGHT", "UNION", "TOP", "DESTINCT", "MAX", "MIN", "AVG", "COUNT", "HAVING", "NOT", "IN", "AND", "OR", "IS", "NULL", "IIF", "CASE", "WHEN", "WHERE")
    s = Trim$(str)
    s = Replace(str, vbTab, " ")
    s = Replace(str, vbCrLf, " " & vbCrLf)
    tks = Split(s, " ")
    
    First = UCase$(tks(0))
    
    u = UBound(heads)

    For i = 0 To u
        If First = heads(i) Then
            Exit For
        End If
    Next
    
    If i > UBound(heads) Then
        ToSQL = False
        Debug.Print "²»ÊÇsqlÓï¾ä"
        Exit Function
    End If
    
    ret = First
    
    u = UBound(tks)
    uk = UBound(keys)

    For i = 1 To u
        isKey = False
        
        If Left$(tks(i), 1) = vbCrLf Then
            
            For j = 0 To uk

                If StrComp(tks(i), vbCrLf & keys(j), vbTextCompare) = 0 Then
                    ret = ret & " " & vbCrLf & keys(j)
                    isKey = True

                    Exit For
                End If
            Next
        Else
            For j = 0 To uk
                If StrComp(tks(i), keys(j), vbTextCompare) = 0 Then
                    ret = ret & " " & keys(j)
                    isKey = True

                    Exit For
                End If
            Next
        End If

        If Not isKey Then
            If ret <> " " Then
                ret = ret & " " & tks(i)
            End If
        End If
    Next
    
    ToSQL = ret
End Function

'--------------------------------------------------
' Procedure   : TrimEx
' Description : È¥³ýÇ°ºóµÄ²»¿É¼û×Ö·û
' CreateTime  : 2011-04-08-14:33:35
'
' ParamList   : Source (String)
' Return      :
'--------------------------------------------------
Public Function TrimEx(Source As String) As String
    Dim s As String
    Dim c As String
    
    s = Trim$(Source)
      
    c = Right$(s, 1)

    While c = vbCr Or c = vbLf Or c = vbTab Or c = " "
        s = Left$(s, Len(s) - 1)
        c = Right$(s, 1)
    Wend
    
    c = Left$(s, 1)
    
    While c = vbCr Or c = vbLf Or c = vbTab Or c = " "
        s = Right$(s, Len(s) - 1)
        c = Left$(s, 1)
    Wend
    
    TrimEx = s
End Function

'Unicode½âÂë
Public Function UnicodeDecode(ByVal sText As String) As String
    Dim i    As Long
    Dim p()  As String
    Dim k    As Long
    Dim str1 As String
    Dim str2 As String
    
    p = Split(sText, "\u")
    k = UBound(p)

    For i = 1 To k
        str1 = p(i)

        If Len(str1) >= 4 Then
            str2 = Right$(str1, Len(str1) - 4)
            str1 = Left$(str1, 4)
            p(i) = ChrW$("&H" & str1) & str2
        Else
            p(i) = "\u" & p(i)
        End If
    Next

    UnicodeDecode = Join(p, "")
End Function

'Unicode±àÂë
Public Function UnicodeEncode(ByVal sText As String) As String
    Dim i    As Long
    Dim k    As Long
    Dim str1 As String
    Dim u    As Long
    
    If Len(sText) = 0 Then Exit Function
    ReDim p(Len(sText) - 1)
    
    u = Len(sText)

    For i = 1 To u
        str1 = Mid$(sText, i, 1)
        k = Asc(str1)

        If k >= 0 And k <= 254 Then

        Else
            str1 = "\u" & Hex$(AscW(str1))
        End If

        p(i - 1) = str1
    Next

    UnicodeEncode = Join(p, "")
End Function

'Unicode×ªUFT8
Public Function UnicodeToUTF8(ByVal sText As String) As Byte()
    Dim lLength     As Long
    Dim lBufferSize As Long
    Dim lResult     As Long
    Dim abUTF8()    As Byte

    lLength = Len(sText)
    abUTF8 = ""

    If lLength = 0 Then GoTo ToExit

    lBufferSize = lLength * 3 + 1

    ReDim abUTF8(lBufferSize - 1)
    lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), lLength, abUTF8(0), lBufferSize, vbNullString, 0)

    If lResult <> 0 Then
        lResult = lResult - 1
        ReDim Preserve abUTF8(lResult)
    Else
        abUTF8 = ""
    End If

ToExit:
    If Err Then Err.Clear
    UnicodeToUTF8 = abUTF8
End Function

'URL½âÂë
Public Function URLDecode(ByVal sText As String, Optional ByVal bUTF8 As Boolean = True) As String
    Dim i         As Long
    Dim k         As Long
    Dim str1      As String
    Dim str2      As String
    Dim UtfB      As Integer
    Dim UtfB1     As String
    Dim UtfB2     As String
    Dim UtfB3     As String   'Utf-8µ¥¸ö×Ö½Ú 1-3×Ö½Ú
    Dim En1       As String
    Dim En2       As String
    Dim En3       As String
    Dim sResult   As String
    Dim pResult() As String
    Dim bOK       As Boolean
    Dim u         As Long
    
    On Error Resume Next

    k = UBound(Split(sText, "%"))

    If k Mod 2 <> 0 Then bUTF8 = True

    pResult = Split("")
    
    If bUTF8 = False Then
        i = 1

        Do Until i > Len(sText)
            sResult = Mid$(sText, i, 1)
            i = i + 1

            Select Case sResult
            Case "+"
                sResult = " "
            Case "%"
                UtfB1 = Mid$(sText, i, 2)
                k = Val("&H" & UtfB1)

                Select Case k
                Case Is > 128
                    UtfB2 = Mid$(sText, i + 3, 2)
                    k = k * 256 + Val("&H" & UtfB2)
                    i = i + 5
                Case Else
                    i = i + 2
                End Select

                sResult = Chr$(k)
            End Select

            ReDim Preserve pResult(UBound(pResult) + 1)
            pResult(UBound(pResult)) = sResult
        Loop
        'ÕâÊÇ½âÃÜµÄ
    Else 'UTF8
        u = Len(sText)

        For i = 1 To u
            sResult = Mid$(sText, i, 1)

            Select Case sResult
            Case "+"
                sResult = " "
            Case "%"
                str2 = Mid$(sText, i + 1, 2)
                UtfB = CInt("&H" & str2)

                If UtfB < 128 Then
                    bOK = True

                    Select Case UtfB
                    Case 38 '&
                        '======= ÅÐ¶Ï "
                        If bOK = True Then
                            str1 = LCase$(Mid$(sText, i, 10))

                            If str1 = "%26quot%3b" Then
                                i = i + 9
                                bOK = False
                                sResult = """"
                            End If
                        End If

                        '======= ÅÐ¶Ï <
                        If bOK = True Then
                            str1 = LCase$(Mid$(sText, i, 8))

                            If str1 = "%26lt%3b" Then
                                i = i + 7
                                bOK = False
                                sResult = "<"
                            End If
                        End If
                        
                        '======= ÅÐ¶Ï >
                        If bOK = True Then
                            str1 = LCase$(Mid$(sText, i, 8))

                            If str1 = "%26gt%3b" Then
                                i = i + 7
                                bOK = False
                                sResult = ">"
                            End If
                        End If
                        
                        '======= ÅÐ¶Ï '
                        If bOK = True Then
                            str1 = LCase$(Mid$(sText, i, 11))

                            If str1 = "%26%2339%3b" Then
                                i = i + 10
                                bOK = False
                                sResult = "'"
                            End If
                        End If
                    Case Else
                    End Select

                    If bOK = True Then
                        i = i + 2
                        sResult = ChrW$(UtfB)
                    End If
                Else
                    En1 = Hex$(UtfB)
                    En2 = Mid$(sText, i + 4, 2)
                    En3 = Mid$(sText, i + 7, 2)

                    If Len(En3) = 0 Then En3 = 0

                    'UtfB1 = (UtfB And &HF) * &H1000   'È¡µÚ1¸öUtf-8×Ö½ÚµÄ¶þ½øÖÆºó4Î»
                    UtfB1 = (UtfB And &HF) * 4096!   'È¡µÚ1¸öUtf-8×Ö½ÚµÄ¶þ½øÖÆºó4Î»

                    UtfB2 = (Val("&H" & En2) And &H3F) * &H40      'È¡µÚ2¸öUtf-8×Ö½ÚµÄ¶þ½øÖÆºó6Î»
                    UtfB3 = Val("&H" & En3) And &H3F      'È¡µÚ3¸öUtf-8×Ö½ÚµÄ¶þ½øÖÆºó6Î»

                    If Err Then
                        Debug.Assert False
                        bUTF8 = False
                        sResult = URLDecode(sText, bUTF8)

                        Exit For
                    End If

                    sResult = ChrW$(UtfB1 Or UtfB2 Or UtfB3)

                    If i = 1 Then
                        str2 = URLEncode(sResult, True)
                        str2 = Replace$(str2, "%", "")

                        If LCase$(str2) <> LCase$(En1 & En2 & En3) Then
                            bUTF8 = False
                            URLDecode = URLDecode(sText, False)

                            Exit Function
                        End If
                    End If

                    i = i + 8
                End If
            Case Else 'AsciiÂë
            End Select

            ReDim Preserve pResult(UBound(pResult) + 1)
            pResult(UBound(pResult)) = sResult
        Next
    End If

    URLDecode = Join(pResult, "")
End Function

'ÎÄ±¾×ªURL±àÂë
'EnMode = 0  encodeURIComponent ´«µÝ²ÎÊýÊ±Ê¹ÓÃ      ²»±àÂë×Ö·ûÓÐ71¸ö(!'()*-._~0-9a-zA-Z)
'EnMode = 1  encodeURI          ½øÐÐ URL Ìø×ªÊ±Ê¹ÓÃ ²»±àÂë×Ö·ûÓÐ82¸ö(!#$&'()*+-./:;=?@_~0-9a-zA-Z)
'EnMode = 2  escape             Ê¹ÓÃÊý¾ÝÊ±          ²»±àÂë×Ö·ûÓÐ69¸ö(*+-./@_0-9a-zA-Z)
Public Function URLEncode(ByVal sText As String, _
                          Optional ByVal bUTF8 As Boolean = False, _
                          Optional ByVal EnMode As Long = 0) As String

    Dim szChar   As String

    Dim szTemp   As String

    Dim szCode   As String

    Dim szHex    As String

    Dim szBin    As String

    Dim iCount1  As Long

    Dim iCount2  As Long

    Dim iStrLen1 As Long

    Dim iStrLen2 As Long

    Dim lResult  As Long

    Dim lAscVal  As Long

    Dim k        As Long

    Dim pCode()  As String

    sText = Trim$(sText)
    iStrLen1 = Len(sText)
    pCode = Split("")

    For iCount1 = 1 To iStrLen1
        szChar = Mid$(sText, iCount1, 1)
        lAscVal = IIf(bUTF8 = True, AscW(szChar), Asc(szChar))

        If lAscVal >= Asc("0") And lAscVal <= Asc("9") Or lAscVal >= Asc("a") And lAscVal <= Asc("z") Or lAscVal >= Asc("A") And lAscVal <= Asc("Z") Or (EnMode = 0 And InStr("!'()*-._~", szChar) > 0) Or (EnMode = 1 And InStr("!#$&'()*+-./:;=?@_~", szChar) > 0) Or (EnMode = 2 And InStr("*+-./@_", szChar) > 0) Then
            szCode = szCode & "%" & Right$("00" & Hex$(lAscVal), 2)
        ElseIf szChar = " " Then
            szChar = "%20"
        ElseIf bUTF8 = False Or InStr(""",", szChar) > 0 Then
            szHex = Hex$(Asc(szChar))

            If Len(szHex) = 4 Then
                szChar = "%" & Left$(szHex, 2) & "%" & Right$(szHex, 2)
            Else
                szChar = "%" & szHex
            End If

        Else
            szHex = Hex$(AscW(szChar))
            iStrLen2 = Len(szHex)

            For iCount2 = 1 To iStrLen2
                szChar = Mid$(szHex, iCount2, 1)

                Select Case szChar

                Case Is = "0"
                    szBin = szBin & "0000"

                Case Is = "1"
                    szBin = szBin & "0001"

                Case Is = "2"
                    szBin = szBin & "0010"

                Case Is = "3"
                    szBin = szBin & "0011"

                Case Is = "4"
                    szBin = szBin & "0100"

                Case Is = "5"
                    szBin = szBin & "0101"

                Case Is = "6"
                    szBin = szBin & "0110"

                Case Is = "7"
                    szBin = szBin & "0111"

                Case Is = "8"
                    szBin = szBin & "1000"

                Case Is = "9"
                    szBin = szBin & "1001"

                Case Is = "A"
                    szBin = szBin & "1010"

                Case Is = "B"
                    szBin = szBin & "1011"

                Case Is = "C"
                    szBin = szBin & "1100"

                Case Is = "D"
                    szBin = szBin & "1101"

                Case Is = "E"
                    szBin = szBin & "1110"

                Case Is = "F"
                    szBin = szBin & "1111"

                Case Else
                End Select

            Next

            szTemp = "1110" & Left$(szBin, 4) & "10" & Mid$(szBin, 5, 6) & "10" & Right$(szBin, 6)
            k = Len(szTemp)

            For iCount2 = 1 To k

                If Mid$(szTemp, iCount2, 1) = "1" Then
                    lResult = lResult + 1 * 2 ^ (k - iCount2)
                Else
                    lResult = lResult + 0 * 2 ^ (k - iCount2)
                End If

            Next iCount2

            szTemp = Hex$(lResult)
            szChar = "%" & Left$(szTemp, 2) & "%" & Mid$(szTemp, 3, 2) & "%" & Right$(szTemp, 2)
        End If

        ReDim Preserve pCode(UBound(pCode) + 1)
        pCode(UBound(pCode)) = szChar

        szBin = vbNullString
        lResult = 0
    Next

    URLEncode = Join(pCode, "")
End Function

'UTF8×ªUnicode
Public Function UTF8ToUnicode(ByRef BinBuff() As Byte) As String

    Dim lRet        As Long

    Dim lLength     As Long

    Dim lBufferSize As Long

    On Error GoTo ToExit

    lLength = UBound(BinBuff) - LBound(BinBuff) + 1

    If lLength <= 0 Then Exit Function

    lBufferSize = lLength * 2
    UTF8ToUnicode = String$(lBufferSize, Chr$(0))
    lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(BinBuff(0)), lLength, StrPtr(UTF8ToUnicode), lBufferSize)

    If lRet <> 0 Then
        UTF8ToUnicode = Left$(UTF8ToUnicode, lRet)
    End If

ToExit:

    If Err Then Err.Clear
End Function

Private Function GetCNSpellSecondChn() As String
    GetCNSpellSecondChn = "Ø¡Ø¢Ø£Ø¤Ø¥Ø¦Ø§Ø¨Ø©ØªØ«Ø¬Ø­Ø®Ø¯Ø°Ø±Ø²Ø³Ø´ØµØ¶Ø·Ø¸Ø¹ØºØ»Ø¼Ø½" & _
       "Ø¾Ø¿ØÀØÁØÂØÃØÄØÅØÆØÇØÈØÉØÊØËØÌØÍØÎØÏØÐØÑØÒØÓØÔØÕØÖØ×ØØØÙØÚØÛØÜØÝØÞØßØàØáØâØãØäØåØæØçØèØéØêØëØìØíØîØïØðØñØòØóØôØõØöØ÷ØøØùØúØûØüØýØþÙ¡Ù¢Ù£Ù¤Ù¥Ù¦Ù§Ù¨Ù©ÙªÙ«Ù¬Ù­Ù®Ù¯Ù°Ù±Ù²Ù³Ù´ÙµÙ¶Ù·Ù¸Ù¹ÙºÙ»Ù¼Ù½Ù¾Ù¿ÙÀÙÁÙÂÙÃÙÄÙÅÙÆÙÇÙÈÙÉÙÊÙËÙÌÙÍÙÎÙÏÙÐÙÑÙÒÙÓÙÔÙÕÙÖÙ×ÙØÙÙÙÚÙÛÙÜÙÝÙÞÙßÙàÙáÙâÙãÙäÙåÙæÙçÙèÙéÙêÙëÙìÙíÙîÙïÙðÙñÙòÙóÙôÙõÙöÙ÷ÙøÙùÙúÙûÙüÙýÙþÚ¡Ú¢Ú£Ú¤Ú¥Ú¦Ú§Ú¨Ú©ÚªÚ«Ú¬Ú­Ú®Ú¯Ú°Ú±Ú²Ú³Ú´ÚµÚ¶Ú·Ú¸Ú¹ÚºÚ»Ú¼Ú½Ú¾Ú¿ÚÀÚÁÚÂÚÃÚÄÚÅÚÆÚÇÚÈÚÉÚÊÚËÚÌÚÍÚÎÚÏÚÐÚÑÚÒÚÓÚÔÚÕÚÖÚ×ÚØÚÙÚÚÚÛÚÜÚÝÚÞÚßÚàÚáÚâÚãÚäÚåÚæÚçÚè" & _
       "ÚéÚêÚëÚìÚíÚîÚïÚðÚñÚòÚóÚôÚõÚöÚ÷ÚøÚùÚúÚûÚüÚýÚþÛ¡Û¢Û£Û¤Û¥Û¦Û§Û¨Û©ÛªÛ«Û¬Û­Û®Û¯Û°Û±Û²Û³Û´ÛµÛ¶Û·Û¸Û¹ÛºÛ»Û¼Û½Û¾Û¿ÛÀÛÁÛÂÛÃÛÄÛÅÛÆÛÇÛÈÛÉÛÊÛËÛÌÛÍÛÎÛÏÛÐÛÑÛÒÛÓÛÔÛÕÛÖÛ×ÛØÛÙÛÚÛÛÛÜÛÝÛÞÛßÛàÛáÛâÛãÛäÛåÛæÛçÛèÛéÛêÛëÛìÛíÛîÛïÛðÛñÛòÛóÛôÛõÛöÛ÷ÛøÛùÛúÛûÛüÛýÛþÜ¡Ü¢Ü£Ü¤Ü¥Ü¦Ü§Ü¨Ü©ÜªÜ«Ü¬Ü­Ü®Ü¯Ü°Ü±Ü²Ü³Ü´ÜµÜ¶Ü·Ü¸Ü¹ÜºÜ»Ü¼Ü½Ü¾Ü¿ÜÀÜÁÜÂÜÃÜÄÜÅÜÆÜÇÜÈÜÉÜÊÜËÜÌÜÍÜÎÜÏÜÐÜÑÜÒÜÓÜÔÜÕÜÖÜ×ÜØÜÙÜÚÜÛÜÜÜÝÜÞÜßÜàÜáÜâÜãÜäÜåÜæÜçÜèÜéÜêÜëÜìÜíÜîÜïÜðÜñÜòÜóÜôÜõÜöÜ÷ÜøÜùÜúÜûÜüÜýÜþÝ¡Ý¢Ý£Ý¤Ý¥Ý¦Ý§Ý¨Ý©ÝªÝ«Ý¬Ý­Ý®Ý¯Ý°Ý±Ý²Ý³Ý´ÝµÝ¶" & _
       "Ý·Ý¸Ý¹ÝºÝ»Ý¼Ý½Ý¾Ý¿ÝÀÝÁÝÂÝÃÝÄÝÅÝÆÝÇÝÈÝÉÝÊÝËÝÌÝÍÝÎÝÏÝÐÝÑÝÒÝÓÝÔÝÕÝÖÝ×ÝØÝÙÝÚÝÛÝÜÝÝÝÞÝßÝàÝáÝâÝãÝäÝåÝæÝçÝèÝéÝêÝëÝìÝíÝîÝïÝðÝñÝòÝóÝôÝõÝöÝ÷ÝøÝùÝúÝûÝüÝýÝþÞ¡Þ¢Þ£Þ¤Þ¥Þ¦Þ§Þ¨Þ©ÞªÞ«Þ¬Þ­Þ®Þ¯Þ°Þ±Þ²Þ³Þ´ÞµÞ¶Þ·Þ¸Þ¹ÞºÞ»Þ¼Þ½Þ¾Þ¿ÞÀÞÁÞÂÞÃÞÄÞÅÞÆÞÇÞÈÞÉÞÊÞËÞÌÞÍÞÎÞÏÞÐÞÑÞÒÞÓÞÔÞÕÞÖÞ×ÞØÞÙÞÚÞÛÞÜÞÝÞÞÞßÞàÞáÞâÞãÞäÞåÞæÞçÞèÞéÞêÞëÞìÞíÞîÞïÞðÞñÞòÞóÞôÞõÞöÞ÷ÞøÞùÞúÞûÞüÞýÞþß¡ß¢ß£ß¤ß¥ß¦ß§ß¨ß©ßªß«ß¬ß­ß®ß¯ß°ß±ß²ß³ß´ßµß¶ß·ß¸ß¹ßºß»ß¼ß½ß¾ß¿ßÀßÁßÂßÃßÄßÅßÆßÇßÈßÉßÊßËßÌßÍßÎßÏßÐßÑßÒßÓßÔßÕßÖß×ßØßÙßÚßÛßÜßÝßÞßßßàßáßâßã" & _
       "ßäßåßæßçßèßéßêßëßìßíßîßïßðßñßòßóßôßõßöß÷ßøßùßúßûßüßýßþà¡à¢à£à¤à¥à¦à§à¨à©àªà«à¬à­à®à¯à°à±à²à³à´àµà¶à·à¸à¹àºà»à¼à½à¾à¿àÀàÁàÂàÃàÄàÅàÆàÇàÈàÉàÊàËàÌàÍàÎàÏàÐàÑàÒàÓàÔàÕàÖà×àØàÙàÚàÛàÜàÝàÞàßàààáàâàãàäàåàæàçàèàéàêàëàìàíàîàïàðàñàòàóàôàõàöà÷àøàùàúàûàüàýàþá¡á¢á£á¤á¥á¦á§á¨á©áªá«á¬á­á®á¯á°á±á²á³á´áµá¶á·á¸á¹áºá»á¼á½á¾á¿áÀáÁáÂáÃáÄáÅáÆáÇáÈáÉáÊáËáÌáÍáÎáÏáÐáÑáÒáÓáÔáÕáÖá×áØáÙáÚáÛáÜáÝáÞáßáàáááâáãáäáåáæáçáèáéáêáëáìáíáîáïáðáñáòáóáôáõáöá÷áøáùáúáûáüáýáþâ¡â¢â£â¤â¥â¦â§â¨â©âªâ«â¬â­â®â¯â°â±â²â³â´âµ" & _
       "â¶â·â¸â¹âºâ»â¼â½â¾â¿âÀâÁâÂâÃâÄâÅâÆâÇâÈâÉâÊâËâÌâÍâÎâÏâÐâÑâÒâÓâÔâÕâÖâ×âØâÙâÚâÛâÜâÝâÞâßâàâáâââãâäâåâæâçâèâéâêâëâìâíâîâïâðâñâòâóâôâõâöâ÷âøâùâúâûâüâýâþã¡ã¢ã£ã¤ã¥ã¦ã§ã¨ã©ãªã«ã¬ã­ã®ã¯ã°ã±ã²ã³ã´ãµã¶ã·ã¸ã¹ãºã»ã¼ã½ã¾ã¿ãÀãÁãÂãÃãÄãÅãÆãÇãÈãÉãÊãËãÌãÍãÎãÏãÐãÑãÒãÓãÔãÕãÖã×ãØãÙãÚãÛãÜãÝãÞãßãàãáãâãããäãåãæãçãèãéãêãëãìãíãîãïãðãñãòãóãôãõãöã÷ãøãùãúãûãüãýãþä¡ä¢ä£ä¤ä¥ä¦ä§ä¨ä©äªä«ä¬ä­ä®ä¯ä°ä±ä²ä³ä´äµä¶ä·ä¸ä¹äºä»ä¼ä½ä¾ä¿äÀäÁäÂäÃäÄäÅäÆäÇäÈäÉäÊäËäÌäÍäÎäÏäÐäÑäÒäÓäÔäÕäÖä×äØäÙäÚäÛäÜäÝäÞäßäàäáäâäãäääåäæäçäè" & _
       "äéäêäëäìäíäîäïäðäñäòäóäôäõäöä÷äøäùäúäûäüäýäþå¡å¢å£å¤å¥å¦å§å¨å©åªå«å¬å­å®å¯å°å±å²å³å´åµå¶å·å¸å¹åºå»å¼å½å¾å¿åÀåÁåÂåÃåÄåÅåÆåÇåÈåÉåÊåËåÌåÍåÎåÏåÐåÑåÒåÓåÔåÕåÖå×åØåÙåÚåÛåÜåÝåÞåßåàåáåâåãåäåååæåçåèåéåêåëåìåíåîåïåðåñåòåóåôåõåöå÷åøåùåúåûåüåýåþæ¡æ¢æ£æ¤æ¥æ¦æ§æ¨æ©æªæ«æ¬æ­æ®æ¯æ°æ±æ²æ³æ´æµæ¶æ·æ¸æ¹æºæ»æ¼æ½æ¾æ¿æÀæÁæÂæÃæÄæÅæÆæÇæÈæÉæÊæËæÌæÍæÎæÏæÐæÑæÒæÓæÔæÕæÖæ×æØæÙæÚæÛæÜæÝæÞæßæàæáæâæãæäæåæææçæèæéæêæëæìæíæîæïæðæñæòæóæôæõæöæ÷æøæùæúæûæüæýæþç¡ç¢ç£ç¤ç¥ç¦ç§ç¨ç©çªç«ç¬ç­ç®ç¯ç°ç±ç²ç³ç´çµç¶ç·ç¸ç¹çºç»ç¼ç½" & _
       "ç¾ç¿çÀçÁçÂçÃçÄçÅçÆçÇçÈçÉçÊçËçÌçÍçÎçÏçÐçÑçÒçÓçÔçÕçÖç×çØçÙçÚçÛçÜçÝçÞçßçàçáçâçãçäçåçæçççèçéçêçëçìçíçîçïçðçñçòçóçôçõçöç÷çøçùçúçûçüçýçþè¡è¢è£è¤è¥è¦è§è¨è©èªè«è¬è­è®è¯è°è±è²è³è´èµè¶è·è¸è¹èºè»è¼è½è¾è¿èÀèÁèÂèÃèÄèÅèÆèÇèÈèÉèÊèËèÌèÍèÎèÏèÐèÑèÒèÓèÔèÕèÖè×èØèÙèÚèÛèÜèÝèÞèßèàèáèâèãèäèåèæèçèèèéèêèëèìèíèîèïèðèñèòèóèôèõèöè÷èøèùèúèûèüèýèþé¡é¢é£é¤é¥é¦é§é¨é©éªé«é¬é­é®é¯é°é±é²é³é´éµé¶é·é¸é¹éºé»é¼é½é¾é¿éÀéÁéÂéÃéÄéÅéÆéÇéÈéÉéÊéËéÌéÍéÎéÏéÐéÑéÒéÓéÔéÕéÖé×éØéÙéÚéÛéÜéÝéÞéßéàéáéâéãéäéåéæéçéèéééêéëéìéíéîéïéðéñéòéó" & _
       "éôéõéöé÷éøéùéúéûéüéýéþê¡ê¢ê£ê¤ê¥ê¦ê§ê¨ê©êªê«ê¬ê­ê®ê¯ê°ê±ê²ê³ê´êµê¶ê·ê¸ê¹êºê»ê¼ê½ê¾ê¿êÀêÁêÂêÃêÄêÅêÆêÇêÈêÉêÊêËêÌêÍêÎêÏêÐêÑêÒêÓêÔêÕêÖê×êØêÙêÚêÛêÜêÝêÞêßêàêáêâêãêäêåêæêçêèêéêêêëêìêíêîêïêðêñêòêóêôêõêöê÷êøêùêúêûêüêýêþë¡ë¢ë£ë¤ë¥ë¦ë§ë¨ë©ëªë«ë¬ë­ë®ë¯ë°ë±ë²ë³ë´ëµë¶ë·ë¸ë¹ëºë»ë¼ë½ë¾ë¿ëÀëÁëÂëÃëÄëÅëÆëÇëÈëÉëÊëËëÌëÍëÎëÏëÐëÑëÒëÓëÔëÕëÖë×ëØëÙëÚëÛëÜëÝëÞëßëàëáëâëãëäëåëæëçëèëéëêëëëìëíëîëïëðëñëòëóëôëõëöë÷ëøëùëúëûëüëýëþì¡ì¢ì£ì¤ì¥ì¦ì§ì¨ì©ìªì«ì¬ì­ì®ì¯ì°ì±ì²ì³ì´ìµì¶ì·ì¸ì¹ìºì»ì¼ì½ì¾ì¿ìÀìÁìÂìÃìÄìÅìÆìÇìÈìÉìÊìËìÌìÍ" & _
       "ìÎìÏìÐìÑìÒìÓìÔìÕìÖì×ìØìÙìÚìÛìÜìÝìÞìßìàìáìâìãìäìåìæìçìèìéìêìëìììíìîìïìðìñìòìóìôìõìöì÷ìøìùìúìûìüìýìþí¡í¢í£í¤í¥í¦í§í¨í©íªí«í¬í­í®í¯í°í±í²í³í´íµí¶í·í¸í¹íºí»í¼í½í¾í¿íÀíÁíÂíÃíÄíÅíÆíÇíÈíÉíÊíËíÌíÍíÎíÏíÐíÑíÒíÓíÔíÕíÖí×íØíÙíÚíÛíÜíÝíÞíßíàíáíâíãíäíåíæíçíèíéíêíëíìíííîíïíðíñíòíóíôíõíöí÷íøíùíúíûíüíýíþî¡î¢î£î¤î¥î¦î§î¨î©îªî«î¬î­î®î¯î°î±î²î³î´îµî¶î·î¸î¹îºî»î¼î½î¾î¿îÀîÁîÂîÃîÄîÅîÆîÇîÈîÉîÊîËîÌîÍîÎîÏîÐîÑîÒîÓîÔîÕîÖî×îØîÙîÚîÛîÜîÝîÞîßîàîáîâîãîäîåîæîçîèîéîêîëîìîíîîîïîðîñîòîóîôîõîöî÷îøîùîúîûîüîýîþï¡ï¢ï£ï¤ï¥ï¦ï§ï¨ï©ïª" & _
       "ï«ï¬ï­ï®ï¯ï°ï±ï²ï³ï´ïµï¶ï·ï¸ï¹ïºï»ï¼ï½ï¾ï¿ïÀïÁïÂïÃïÄïÅïÆïÇïÈïÉïÊïËïÌïÍïÎïÏïÐïÑïÒïÓïÔïÕïÖï×ïØïÙïÚïÛïÜïÝïÞïßïàïáïâïãïäïåïæïçïèïéïêïëïìïíïîïïïðïñïòïóïôïõïöï÷ïøïùïúïûïüïýïþð¡ð¢ð£ð¤ð¥ð¦ð§ð¨ð©ðªð«ð¬ð­ð®ð¯ð°ð±ð²ð³ð´ðµð¶ð·ð¸ð¹ðºð»ð¼ð½ð¾ð¿ðÀðÁðÂðÃðÄðÅðÆðÇðÈðÉðÊðËðÌðÍðÎðÏðÐðÑðÒðÓðÔðÕðÖð×ðØðÙðÚðÛðÜðÝðÞðßðàðáðâðãðäðåðæðçðèðéðêðëðìðíðîðïðððñðòðóðôðõðöð÷ðøðùðúðûðüðýðþñ¡ñ¢ñ£ñ¤ñ¥ñ¦ñ§ñ¨ñ©ñªñ«ñ¬ñ­ñ®ñ¯ñ°ñ±ñ²ñ³ñ´ñµñ¶ñ·ñ¸ñ¹ñºñ»ñ¼ñ½ñ¾ñ¿ñÀñÁñÂñÃñÄñÅñÆñÇñÉñÊñËñÌñÍñÎñÏñÐñÑñÒñÓñÔñÕñÖñ×ñØñÙñÚñÛñÜñÝñÞñßñàñâñãñäñåñæñç" & _
       "ñèñéñêñëñìñíñîñïñðñññòñóñôñõñöñ÷ñøñùñúñûñüñýñþò¡ò¢ò£ò¤ò¥ò¦ò§ò¨ò©òªò«ò¬ò­ò®ò¯ò°ò±ò²ò³ò´òµò¶ò·ò¸ò¹òºò»ò¼ò½ò¾ò¿òÀòÁòÂòÃòÄòÅòÆòÇòÈòÉòÊòËòÌòÍòÎòÏòÐòÑòÒòÓòÔòÕòÖò×òØòÙòÚòÛòÜòÝòÞòßòàòáòâòãòäòåòæòçòèòéòêòëòìòíòîòïòðòñòòòóòôòõòöò÷òøòùòúòûòüòýòþó¡ó¢ó£ó¤ó¥ó¦ó§ó¨ó©óªó«ó¬ó­ó®ó¯ó°ó±ó²ó³ó´óµó¶ó·ó¸ó¹óºó»ó¼ó½ó¾ó¿óÀóÁóÂóÃóÄóÅóÆóÇóÈóÉóÊóËóÌóÍóÎóÏóÐóÑóÒóÓóÔóÕóÖó×óØóÙóÚóÛóÜóÝóÞóßóàóáóâóãóäóåóæóçóèóéóêóëóìóíóîóïóðóñóòóóóôóõóöó÷óøóùóúóûóüóýóþô¡ô¢ô£ô¤ô¥ô¦ô§ô¨ô©ôªô«ô¬ô­ô®ô¯ô°ô±ô²ô³ô´ôµô¶ô·ô¸ô¹ôºô»ô¼ô½ô¾ô¿ôÀôÁôÂôÃôÄôÅôÆôÇ" & _
       "ôÈôÉôÊôËôÌôÍôÎôÏôÐôÑôÒôÓôÔôÕôÖô×ôØôÙôÚôÛôÜôÝôÞôßôàôáôâôãôäôåôæôçôèôéôêôëôìôíôîôïôðôñôòôóôôôõôöô÷ôøôùôúôûôüôýôþõ¡õ¢õ£õ¤õ¥õ¦õ§õ¨õ©õªõ«õ¬õ­õ®õ¯õ°õ±õ²õ³õ´õµõ¶õ·õ¸õ¹õºõ»õ¼õ½õ¾õ¿õÀõÁõÂõÃõÄõÅõÆõÇõÈõÉõÊõËõÌõÍõÎõÏõÐõÑõÒõÓõÔõÕõÖõ×õØõÙõÚõÛõÜõÝõÞõßõàõáõâõãõäõåõæõçõèõéõêõëõìõíõîõïõðõñõòõóõôõõõöõ÷õøõùõúõûõüõýõþö¡ö¢ö£ö¤ö¥ö¦ö§ö¨ö©öªö«ö¬ö­ö®ö¯ö°ö±ö²ö³ö´öµö¶ö·ö¸ö¹öºö»ö¼ö½ö¾ö¿öÀöÁöÂöÃöÄöÅöÆöÇöÈöÉöÊöËöÌöÍöÎöÏöÐöÑöÒöÓöÔöÕöÖö×öØöÙöÚöÛöÜöÝöÞößöàöáöâöãöäöåöæöçöèöéöêöëöìöíöîöïöðöñöòöóöôöõööö÷öøöùöúöûöüöýöþ÷¡÷¢÷£÷¤÷¥÷¦÷§" & _
       "÷¨÷©÷ª÷«÷¬÷­÷®÷¯÷°÷±÷²÷³÷´÷µ÷¶÷·÷¸÷¹÷º÷»÷¼÷½÷¾÷¿÷À÷Á÷Â÷Ã÷Ä÷Å÷Æ÷Ç÷È÷É÷Ê÷Ë÷Ì÷Í÷Î÷Ï÷Ð÷Ñ÷Ò÷Ó÷Ô÷Õ÷Ö÷×÷Ø÷Ù÷Ú÷Û÷Ü÷Ý÷Þ÷ß÷à÷á÷â÷ã÷ä÷å÷æ÷ç÷è÷é÷ê÷ë÷ì÷í÷î÷ï÷ð÷ñ÷ò÷ó÷ô÷õ÷ö÷÷÷ø÷ù÷ú÷û÷ü÷ý÷þîÚ"
End Function

Private Function GetCNSpellSecondEng() As String
    GetCNSpellSecondEng = "CJWGNSPGCGNESYPBTYYZDXYKYGTDJNNJQMBSGZSCYJSYYQPGKBZGYCYWJKGKLJSWKPJQHYTWDDZLSGMRYPYWWCCKZNKYDGTTNGJEYKKZYTCJNMCYLQLYPYQFQRPZSLWBTGKJFYXJWZLTBNCXJJJJZXDTTSQZYCDXXHGCKBPHFFSSWYBGMXLPBYLLLHLXSPZMYJHSOJNGHDZQYKLGJHSGQZHXQGKEZZWYSCSCJXYEYXADZPMDSSMZJZQJYZCDJZWQJBDZBXGZNZCPWHKXHQKMWFBPBYDTJZZKQHYLYGXFPTYJYYZPSZLFCHMQSHGMXXSXJJSDCSBBQBEFSJYHWWGZKPYLQBGLDLCCTNMAYDDKSSNGYCSGXLYZAYBNPTSDKDYLHGYMYLCXPYCJNDQJWXQXFYYFJLEJBZRXCCQWQQSBNKYMGPLBMJRQCFLNYMYQMSQTRBCJTHZTQFRXQ" & _
       "HXMJJCJLXQGJMSHZKBSWYEMYLTXFSYDSGLYCJQXSJNQBSCTYHBFTDCYZDJWYGHQFRXWCKQKXEBPTLPXJZSRMEBWHJLBJSLYYSMDXLCLQKXLHXJRZJMFQHXHWYWSBHTRXXGLHQHFNMNYKLDYXZPWLGGTMTCFPAJJZYLJTYANJGBJPLQGDZYQYAXBKYSECJSZNSLYZHZXLZCGHPXZHZNYTDSBCJKDLZAYFMYDLEBBGQYZKXGLDNDNYSKJSHDLYXBCGHXYPKDJMMZNGMMCLGWZSZXZJFZNMLZZTHCSYDBDLLSCDDNLKJYKJSYCJLKOHQASDKNHCSGANHDAASHTCPLCPQYBSDMPJLPCJOQLCDHJJYSPRCHNWJNLHLYYQYYWZPTCZGWWMZFFJQQQQYXACLBHKDJXDGMMYDJXZLLSYGXGKJRYWZWYCLZMSSJZLDBYDCFCXYHLXCHYZJQSFQAGMNYXPFRKSSB" & _
       "JLYXYSYGLNSCMHCWWMNZJJLXXHCHSYDSTTXRYCYXBYHCSMXJSZNPWGPXXTAYBGAJCXLYSDCCWZOCWKCCSBNHCPDYZNFCYYTYCKXKYBSQKKYTQQXFCWCHCYKELZQBSQYJQCCLMTHSYWHMKTLKJLYCXWHEQQHTQHZPQSQSCFYMMDMGBWHWLGSSLYSDLMLXPTHMJHWLJZYHZJXHTXJLHXRSWLWZJCBXMHZQXSDZPMGFCSGLSXYMJSHXPJXWMYQKSMYPLRTHBXFTPMHYXLCHLHLZYLXGSSSSTCLSLDCLRPBHZHXYYFHBBGDMYCNQQWLQHJJZYWJZYEJJDHPBLQXTQKWHLCHQXAGTLXLJXMSLXHTZKZJECXJCJNMFBYCSFYWYBJZGNYSDZSQYRSLJPCLPWXSDWEJBJCBCNAYTWGMPAPCLYQPCLZXSBNMSGGFNZJJBZSFZYNDXHPLQKZCZWALSBCCJXJYZGWKYP" & _
       "SGXFZFCDKHJGXDLQFSGDSLQWZKXTMHSBGZMJZRGLYJBPMLMSXLZJQQHZYJCZYDJWBMJKLDDPMJEGXYHYLXHLQYQHKYCWCJMYYXNATJHYCCXZPCQLBZWWYTWBQCMLPMYRJCCCXFPZNZZLJPLXXYZTZLGDLDCKLYRZZGQTGJHHHJLJAXFGFJZSLCFDQZLCLGJDJCSNCLLJPJQDCCLCJXMYZFTSXGCGSBRZXJQQCTZHGYQTJQQLZXJYLYLBCYAMCSTYLPDJBYREGKLZYZHLYSZQLZNWCZCLLWJQJJJKDGJZOLBBZPPGLGHTGZXYGHZMYCNQSYCYHBHGXKAMTXYXNBSKYZZGJZLQJDFCJXDYGJQJJPMGWGJJJPKQSBGBMMCJSSCLPQPDXCDYYKYFCJDDYYGYWRHJRTGZNYQLDKLJSZZGZQZJGDYKSHPZMTLCPWNJAFYZDJCNMWESCYGLBTZCGMSSLLYXQSXSBSJS" & _
       "BBSGGHFJLWPMZJNLYYWDQSHZXTYYWHMCYHYWDBXBTLMSYYYFSXJCSDXXLHJHFSSXZQHFZMZCZTQCXZXRTTDJHNNYZQQMNQDMMGYYDXMJGDHCDYZBFFALLZTDLTFXMXQZDNGWQDBDCZJDXBZGSQQDDJCMBKZFFXMKDMDSYYSZCMLJDSYNSPRSKMKMPCKLGDBQTFZSWTFGGLYPLLJZHGJJGYPZLTCSMCNBTJBQFKTHBYZGKPBBYMTTSSXTBNPDKLEYCJNYCDYKZDDHQHSDZSCTARLLTKZLGECLLKJLQJAQNBDKKGHPJTZQKSECSHALQFMMGJNLYJBBTMLYZXDCJPLDLPCQDHZYCBZSCZBZMSLJFLKRZJSNFRGJHXPDHYJYBZGDLQCSEZGXLBLGYXTWMABCHECMWYJYZLLJJYHLGBDJLSLYGKDZPZXJYYZLWCXSZFGWYYDLYHCLJSCMBJHBLYZLYCBLYDPDQYSXQZB" & _
       "YTDKYXJYYCNRJMDJGKLCLJBCTBJDDBBLBLCZQRPXJCGLZCSHLTOLJNMDDDLNGKAQHQHJGYKHEZNMSHRPHQQJCHGMFPRXHJGDYCHGHLYRZQLCYQJNZSQTKQJYMSZSWLCFQQQXYFGGYPTQWLMCRNFKKFSYYLQBMQAMMMYXCTPSHCPTXXZZSMPHPSHMCLMLDQFYQXSZYJDJJZZHQPDSZGLSTJBCKBXYQZJSGPSXQZQZRQTBDKYXZKHHGFLBCSMDLDGDZDBLZYYCXNNCSYBZBFGLZZXSWMSCCMQNJQSBDQSJTXXMBLTXZCLZSHZCXRQJGJYLXZFJPHYMZQQYDFQJJLZZNZJCDGZYGCTXMZYSCTLKPHTXHTLBJXJLXSCDQXCBBTJFQZFSLTJBTKQBXXJJLJCHCZDBZJDCZJDCPRNPQCJPFCZLCLZXZDMXMPHJSGZGSZZQJYLWTJPFSYASMCJBTZKYCWMYTCSJJLJCQLWZM" & _
       "ALBXYFBPNLSFHTGJWEJJXXGLLJSTGSHJQLZFKCGNNDSZFDEQFHBSAQTGLLBXMMYGSZLDYDQMJJRGBJTKGDHGKBLQKBDMBYLXWCXYTTYBKMRTJZXQJBHLMHMJJZMQASLDCYXYQDLQCAFYWYXQHZY"
End Function

'--------------------------------------------------
' Procedure   : SubCount
' Description : ¼ÆËãSourceTextÖÐSubText³öÏÖµÄ´ÎÊý
' CreateTime  : 2010-11-16-09:43:58
'
' ParamList   : Str (String)    Ä¿±êString
'               Index (Long)    Î»ÖÃ
' Return      : ·µ»ØStrÖÐµÚIndex¸ö×Ö·û£¬ÈôÎÞ£¬·µ»Ø""
'--------------------------------------------------
Public Function SubCount(ByVal SourceText As String, _
                         ByVal SubText As String, _
                         Optional ByVal IgnoreCase As Boolean = True) As Long

    Dim c           As Long

    Dim i           As Long

    Dim CompareMode As VbCompareMethod

    Dim n           As Long
    
    n = Len(SubText)

    If Len(n) = 0 Then Exit Function
    If Len(SourceText) = 0 Then Exit Function
    
    CompareMode = IIf(IgnoreCase, vbTextCompare, vbBinaryCompare)
    
    i = 1

    Do
        i = InStr(i, SourceText, SubText, CompareMode)

        If i = 0 Then
            SubCount = c

            Exit Function

        Else
            c = c + 1
        End If

        i = i + n
    Loop

End Function

'--------------------------------------------------------------------------------
' Procedure  :  ContainsCharOf
' Description:  ²âÊÔSourceTextÊÇ·ñ°üº¬CharListÖÐµÄÈÎÒ»×Ö·û
' Created by :  fangzi
' Date-Time  :  6/14/2017-08:56:28
'
' Parameters :  SourceText (String)
'               CharList (String)
'               IgnoreCase (Boolean = True)
'--------------------------------------------------------------------------------
Public Function ContainsCharOf(ByVal SourceText As String, _
                               ByVal CharList As String, _
                               Optional ByVal IgnoreCase As Boolean = True) As Boolean

    Dim i           As Long

    Dim c           As String

    Dim uSource     As Long

    Dim uList       As Long

    Dim CompareMode As VbCompareMethod
    
    uSource = Len(SourceText)
    uList = Len(CharList)
    
    CompareMode = IIf(IgnoreCase, vbTextCompare, vbBinaryCompare)
    
    'ÁÐ±í±È½Ï¶ÌÊ±£¬¿´ÁÐ±íÖÐµÄ×Ö·ûÊÇ·ñÔÚÔ´×Ö·û´®ÖÐ³öÏÖ
    If uSource > uList Then

        For i = 1 To uList
            c = Mid$(CharList, i, 1)

            If InStr(SourceText, c, CompareMode) > 0 Then
                ContainsCharOf = True

                Exit Function

            End If

        Next

    Else    'Ô´×Ö·û´®±ÈÁÐ±í¶ÌÊ±£¬¿´Ô´×Ö·û´®ÖÐµÄ×Ö·ûÊÇ·ñÔÚÁÐ±íÖÐ³öÏÖ

        For i = 1 To uSource
            c = Mid$(SourceText, i, 1)

            If InStr(CharList, c, CompareMode) > 0 Then
                ContainsCharOf = True

                Exit Function

            End If

        Next

    End If

End Function
