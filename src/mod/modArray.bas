Attribute VB_Name = "modArray"
Option Explicit

Public Function Array_unique(Arr() As String) As Variant
    Dim k As Integer
    Dim i As Integer
    Dim NewArr() As Variant
    
    ShellSortAsc Arr
    
    For i = 0 To UBound(Arr)
        If Arr(k) <> Arr(i) Then
            Arr(k + 1) = Arr(i)
            k = k + 1
        End If
    Next

    ReDim NewArr(k)

    For i = 0 To k
        NewArr(i) = Arr(i)
    Next

    Array_unique = NewArr
End Function

Public Function RandomSort(SourceArray() As String) As Variant
    Dim i As Long
    Dim u As Long
    Dim l As Long
    Dim d As Long
    
    l = LBound(SourceArray)
    u = UBound(SourceArray)
    
    ReDim retArray(l, u) As Variant
    
    For i = l To u
        d = RandLong(i, u)
    Next
    
    RandomSort = retArray
End Function

'返回小于等于Upper且大于等于Lower的长整数
Private Function RandLong(ByVal Upper As Long, Optional ByVal Lower As Long = 0) As Long
    Randomize DateDiff("s", Now(), "1970-1-1")
    RandLong = Int((Upper - Lower + 1) * Rnd + Lower)
End Function

Public Function IsArrayInitialized(ByVal sArray As Variant) As Boolean '判断数组是否为空
    Dim i As Long
    
    On Error GoTo lerr:

    i = UBound(sArray)
    IsArrayInitialized = True
lerr:
End Function
