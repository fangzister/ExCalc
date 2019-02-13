Attribute VB_Name = "modSort"
Option Explicit

'VB对数字/字符数组的快速排序.查找
'数值与数组操作'

'数值快速排序(从小到大)
'函数:NumSortAZ
'参数:Myarray Double数组,nLft 数组的左边界,nRgt 数组右边界.
'返回值:无
'例子:
Public Sub NumSortAZ(ByRef Myarray, nLft As Long, nRgt As Long)
    Dim i    As Long
    Dim j    As Long
    Dim tmpX As Variant
    Dim tmpA As Variant

    i = nLft
    j = nRgt
    tmpX = Val(Myarray((nLft + nRgt) / 2))

    While (i <= j)
        While (Val(Myarray(i)) < tmpX And i < nRgt)
            i = i + 1
        Wend

        While (tmpX < Val(Myarray(j)) And j > nLft)
            j = j - 1
        Wend

        If (i <= j) Then
            tmpA = Val(Myarray(i))
            Myarray(i) = Val(Myarray(j))
            Myarray(j) = tmpA
            i = i + 1
            j = j - 1
        End If
    Wend

    If (nLft < j) Then Call NumSortAZ(Myarray, nLft, j)
    If (i < nRgt) Then Call NumSortAZ(Myarray, i, nRgt)
End Sub

'
'数值快速排序(从大到小)
'函数:NumSortZA
'参数:Myarray Double数组,nLft 数组的左边界,nRgt 数组右边界.
'返回值:无
'例子:
Public Sub NumSortZA(ByRef Myarray, nLft As Long, nRgt As Long)
    Dim i    As Long
    Dim j    As Long
    Dim LT   As Long
    Dim RT   As Long
    Dim tmpX As Variant
    Dim tmpA As Variant

    i = nLft
    j = nRgt
    tmpX = Val(Myarray((nLft + nRgt) / 2))
    
    While (i <= j)

        While (Val(Myarray(i)) > tmpX And i < nRgt)
            i = i + 1
        Wend

        While (tmpX > Val(Myarray(j)) And j > nLft)
            j = j - 1
        Wend

        If (i <= j) Then
            tmpA = Val(Myarray(i))
            Myarray(i) = Val(Myarray(j))
            Myarray(j) = tmpA
            i = i + 1
            j = j - 1
        End If
    Wend

    If (nLft < j) Then Call NumSortZA(Myarray, nLft, j)
    If (i < nRgt) Then Call NumSortZA(Myarray, i, nRgt)
End Sub

'
'字符串快速排序(从大到小)
'函数:StrSortZA
'参数:sArr String数组,L 数组的左边界,R 数组右边界.
'返回值:无
'例子:
Public Sub StrSortZA(ByRef sArr() As String, First As Long, Last As Long)
    Dim vSplit As String
    Dim vT     As String
    Dim i      As Long
    Dim j      As Long
    Dim iRand  As Long

    If First < Last Then
        If Last - First = 1 Then
            If sArr(First) < sArr(Last) Then
                vT = sArr(First): sArr(First) = sArr(Last): sArr(Last) = vT
            End If
        Else
            iRand = Int(First + (Rnd * (Last - First + 1)))
            vT = sArr(Last): sArr(Last) = sArr(iRand): sArr(iRand) = vT
            vSplit = sArr(Last)

            Do
                i = First: j = Last

                Do While (i < j) And (sArr(i) >= vSplit)
                    i = i + 1
                Loop

                Do While (j > i) And (sArr(j) <= vSplit)
                    j = j - 1
                Loop

                If i < j Then
                    vT = sArr(i): sArr(i) = sArr(j): sArr(j) = vT
                End If
            Loop While i < j

            vT = sArr(i): sArr(i) = sArr(Last): sArr(Last) = vT

            If (i - First) < (Last - i) Then
                StrSortZA sArr(), First, i - 1
                StrSortZA sArr(), i + 1, Last
            Else
                StrSortZA sArr(), i + 1, Last
                StrSortZA sArr(), First, i - 1
            End If
        End If
    End If

End Sub

'
'字符串快速排序(从小到大)
'函数:StrSortAZ
'参数:sArr String数组,First 数组的左边界,Last 数组右边界.
'返回值:无
'例子:
Public Sub StrSortAZ(ByRef sArr() As String, First As Long, Last As Long)
    Dim vSplit As String
    Dim vT     As String
    Dim i      As Long
    Dim j      As Long
    Dim iRand  As Long

    If First < Last Then
        If Last - First = 1 Then
            If sArr(First) > sArr(Last) Then
                vT = sArr(First): sArr(First) = sArr(Last): sArr(Last) = vT
            End If
        Else
            iRand = Int(First + (Rnd * (Last - First + 1)))
            vT = sArr(Last): sArr(Last) = sArr(iRand): sArr(iRand) = vT
            vSplit = sArr(Last)

            Do
                i = First
                j = Last

                Do While (i < j) And (sArr(i) <= vSplit)
                    i = i + 1
                Loop

                Do While (j > i) And (sArr(j) >= vSplit)
                    j = j - 1
                Loop

                If i < j Then
                    vT = sArr(i): sArr(i) = sArr(j): sArr(j) = vT
                End If

            Loop While i < j

            vT = sArr(i): sArr(i) = sArr(Last): sArr(Last) = vT

            If (i - First) < (Last - i) Then
                StrSortAZ sArr(), First, i - 1
                StrSortAZ sArr(), i + 1, Last
            Else
                StrSortAZ sArr(), i + 1, Last
                StrSortAZ sArr(), First, i - 1
            End If
        End If
    End If
End Sub

'对MyArray进行希尔排序，可指定排序方向
Public Sub ShellSortAsc(ByRef Myarray() As String)
    Dim Distance    As Long
    Dim nSize       As Long
    Dim Index       As Long
    Dim NextElement As Long
    Dim temp        As String

    nSize = UBound(Myarray) - LBound(Myarray) + 1
    Distance = 1

    While (Distance <= nSize)
        Distance = 2 * Distance
    Wend

    Distance = (Distance / 2) - 1
    
    While (Distance > 0)
        NextElement = LBound(Myarray) + Distance
    
        While (NextElement <= UBound(Myarray))
            Index = NextElement

            Do
                If Index >= (LBound(Myarray) + Distance) Then
                    If Myarray(Index) < Myarray(Index - Distance) Then
                        temp = Myarray(Index)
                        Myarray(Index) = Myarray(Index - Distance)
                        Myarray(Index - Distance) = temp
                        Index = Index - Distance
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            NextElement = NextElement + 1
        Wend

        Distance = (Distance - 1) / 2
    Wend
End Sub

Public Sub ShellSortDesc(ByRef Myarray() As String)
    Dim Distance    As Long
    Dim nSize       As Long
    Dim Index       As Long
    Dim NextElement As Long
    Dim temp        As String

    nSize = UBound(Myarray) - LBound(Myarray) + 1
    Distance = 1

    While (Distance <= nSize)
        Distance = 2 * Distance
    Wend

    Distance = (Distance / 2) - 1
    
    While (Distance > 0)
        NextElement = LBound(Myarray) + Distance
    
        While (NextElement <= UBound(Myarray))
            Index = NextElement
            
            Do
                If Index >= (LBound(Myarray) + Distance) Then
                    If Myarray(Index) >= Myarray(Index - Distance) Then
                        temp = Myarray(Index)
                        Myarray(Index) = Myarray(Index - Distance)
                        Myarray(Index - Distance) = temp
                        Index = Index - Distance
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            NextElement = NextElement + 1
        Wend

        Distance = (Distance - 1) / 2
    Wend
End Sub

'对MyArray进行随机希尔排序
Public Sub ShellSortRandom(ByRef Myarray() As String)
    Dim Distance    As Long
    Dim Size        As Long
    Dim Index       As Long
    Dim NextElement As Long
    Dim temp        As String

    Size = UBound(Myarray) - LBound(Myarray) + 1
    Distance = 1

    While (Distance <= Size)
        Distance = 2 * Distance
    Wend

    Distance = (Distance / 2) - 1
    Randomize
    
    While (Distance > 0)
        NextElement = LBound(Myarray) + Distance
    
        While (NextElement <= UBound(Myarray))
            Index = NextElement

            Do
                If Index >= (LBound(Myarray) + Distance) Then
                    If Rnd() > 0.5 Then
                        temp = Myarray(Index)
                        Myarray(Index) = Myarray(Index - Distance)
                        Myarray(Index - Distance) = temp
                        Index = Index - Distance
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            NextElement = NextElement + 1
        Wend

        Distance = (Distance - 1) / 2
    Wend
End Sub
