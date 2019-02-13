Attribute VB_Name = "modMath"
Option Explicit

'除法运算，返回商到Quotient，返回余数到Remainder
Public Sub DivideLong(ByVal Dividend As Long, _
                      ByVal Divisor As Long, _
                      ByRef Quotient As Long, _
                      ByRef Remainder As Long)

    If Divisor = 0 Then Exit Sub
    
    Quotient = Dividend \ Divisor
    Remainder = Dividend - Divisor * Quotient
End Sub

'返回两数中较大的值
Public Function Max(ByVal Value1 As Double, ByVal Value2 As Double) As Variant
    Max = IIf((Value1 > Value2), Value1, Value2)
End Function

Public Function MaxLong(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    MaxLong = IIf((Value1 > Value2), Value1, Value2)
End Function

'返回两数中较小的值
Public Function Min(ByVal Value1 As Double, ByVal Value2 As Double) As Variant
    Min = IIf((Value1 < Value2), Value1, Value2)
End Function

Public Function MinLong(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    MinLong = IIf((Value1 < Value2), Value1, Value2)
End Function

'--------------------------------------------------
' Procedure   : LCM
' Description : 计算A和B的最小公倍数
' CreateTime  : 2011-04-11-11:09:41
'
' ParamList   : A (Long)
'               B (Long)
' Return      :
'--------------------------------------------------
Public Function LCM(ByVal a As Long, ByVal b As Long) As Long

    Dim m  As Long

    Dim n  As Long

    Dim mn As Long

    Dim R  As Long
    
    If a = 0 Or b = 0 Then Exit Function
    
    m = Max(a, b)
    n = Min(a, b)
    mn = n * m

    R = m Mod n

    Do While (R <> 0)
        m = n
        n = R
        R = m Mod n
    Loop

    LCM = mn / n
End Function

'--------------------------------------------------
' Procedure   : GCD
' Description : 计算A和B的最大公约数
' CreateTime  : 2011-04-11-11:09:48
'
' ParamList   : A (Long)
'               B (Long)
' Return      :
'--------------------------------------------------
Public Function GCD(ByVal a As Long, ByVal b As Long) As Long

    Dim m As Long

    Dim n As Long

    Dim R As Long
    
    If a = 0 Or b = 0 Then Exit Function
    
    m = Max(a, b)
    n = Min(a, b)

    R = m Mod n

    Do While (R <> 0)
        m = n
        n = R
        R = m Mod n
    Loop
    
    GCD = n
End Function

'--------------------------------------------------
' Procedure   : Factorial
' Description : 计算Operator的阶乘
' CreateTime  : 2011-04-11-14:17:07
'
' ParamList   : Operator (Long)
' Return      :
'--------------------------------------------------
Public Function Factorial(ByVal Operator As Long) As Long

    Dim i As Long

    Dim k As Long
    
    k = 1

    For i = 1 To Operator
        k = i * k
    Next

    Factorial = k
End Function

'--------------------------------------------------
' Procedure   : Arrangement
' Description : 计算m和n的排列
' CreateTime  : 2011-04-11-14:17:53
'
' ParamList   : m (Long)
'               n (Long)
' Return      :
'--------------------------------------------------
Public Function Arrangement(ByVal m As Long, ByVal n As Long) As Long
    Arrangement = Factorial(n) / Factorial(n - m)
End Function

'--------------------------------------------------
' Procedure   : Combination
' Description : 计算m和n的组合
' CreateTime  : 2011-04-11-14:18:01
'
' ParamList   : m (Long)
'               n (Long)
' Return      :
'--------------------------------------------------
Public Function Combination(ByVal m As Long, ByVal n As Long) As Long
    Combination = Factorial(n) / (Factorial(m) * Factorial(n - m))
End Function

