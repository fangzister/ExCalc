Attribute VB_Name = "modDateTime"
Option Explicit

Public Function TimeStampToDate(ByVal ts As Variant, _
                                Optional ByVal TimeFormat As String = "yyyy-MM-dd HH:mm:ss", _
                                Optional ByVal InitTime As String = "1970-1-1") As String

    Dim d As Date

    Dim s As String

    Dim n As Long
    
    s = ts

    If Len(s) = 0 Then Exit Function
    
    If Not IsNumeric(s) Then
        TimeStampToDate = ts

        Exit Function

    End If
    
    If Len(s) = 13 Then
        n = ts / 1000
    Else
        n = ts
    End If

    d = DateAdd("s", n, InitTime)
    TimeStampToDate = Format(d, TimeFormat)

    If Len(s) = 13 Then
        TimeStampToDate = TimeStampToDate & Right$(ts, 3)
    End If

End Function

Public Function DateToTimeStamp(ByVal DateValue As String, _
                                Optional ByVal TimeFormat As String = "yyyy-MM-dd HH:mm:ss", _
                                Optional ByVal InitTime As String = "1970-1-1", _
                                Optional ByVal OutputMS As Boolean = False) As String

    Dim d As Date

    Dim s As String

    Dim n As Long
    
    If Not IsDate(DateValue) Then
        DateToTimeStamp = DateValue

        Exit Function

    End If

    DateToTimeStamp = DateDiff("s", DateValue, InitTime)

    If OutputMS Then
        DateToTimeStamp = DateToTimeStamp * 1000
    End If

End Function

