VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Redix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Function IsBinValue(ByVal RHS As String) As Boolean
    Dim i As Long
    Dim c As String
    
    RHS = Trim$(RHS)
    
    For i = 1 To Len(RHS)
        c = Mid$(RHS, i, 1)

        If c = "0" Or c = "1" Then
            'pass
        Else
            Exit Function
        End If
    Next
    
    IsBinValue = True
End Function

Private Function IsDecValue(ByVal RHS As String) As Boolean
    Dim i As Long
    Dim c As String
    Dim a As Long
    
    RHS = Trim$(RHS)
    
    For i = 1 To Len(RHS)
        c = Mid$(RHS, i, 1)
        a = Asc(c)

        If a < 48 Or a > 57 Then
            Exit Function
        End If
    Next
    
    IsDecValue = True
End Function

Private Function IsOctValue(ByVal RHS As String) As Boolean
    Dim i As Long
    Dim c As String
    Dim a As Long
    
    RHS = Trim$(RHS)
    
    For i = 1 To Len(RHS)
        c = Mid$(RHS, i, 1)
        a = Asc(c)

        If a < 48 Or a > 55 Then
            Exit Function
        End If
    Next
    
    IsOctValue = True
End Function

Private Function IsHexValue(ByVal RHS As String) As Boolean
    Dim i As Long
    Dim c As String
    Dim a As Long
    
    RHS = UCase$(Trim$(RHS))
    
    For i = 1 To Len(RHS)
        c = Mid$(RHS, i, 1)
        a = Asc(c)

        If a < 48 Or a > 70 Or (a > 57 And a < 65) Then
            Exit Function
        End If
    Next
    
    IsHexValue = True
End Function

' 用途：将二进制转化为十进制
' 输入：Bin(二进制数)
' 输入数据类型：String
' 输出：BIN_to_DEC(十进制数)
' 输出数据类型：Long
' 输入的最大数为1111111111111111111111111111111(31个1),输出最大数为2147483647
Public Function Bin2Dec(ByVal RHS As String) As Variant
    Dim i As Long
    Dim u As Long
    
    RHS = Trim$(RHS)

    If IsBinValue(RHS) = False Then
        Bin2Dec = "输入类型错误"

        Exit Function
    End If
    
    u = Len(RHS)

    For i = 1 To u
        Bin2Dec = Bin2Dec * 2 + Val(Mid$(RHS, i, 1))
    Next i
End Function

' 用途：将二进制转化为十六进制
' 输入：Bin(二进制数)
' 输入数据类型：String
' 输出：Bin2Hex(十六进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function Bin2Hex(ByVal RHS As String) As Variant
    Dim i As Long
    Dim h As String
    Dim u As Long
   
    RHS = Trim$(RHS)

    If IsBinValue(RHS) = False Then
        Bin2Hex = "输入类型错误"

        Exit Function

    End If
    
    If Len(RHS) Mod 4 <> 0 Then
        RHS = String$(4 - Len(RHS) Mod 4, "0") & RHS
    End If
    
    u = Len(RHS)

    For i = 1 To u Step 4

        Select Case Mid$(RHS, i, 4)

        Case "0000": h = h & "0"

        Case "0001": h = h & "1"

        Case "0010": h = h & "2"

        Case "0011": h = h & "3"

        Case "0100": h = h & "4"

        Case "0101": h = h & "5"

        Case "0110": h = h & "6"

        Case "0111": h = h & "7"

        Case "1000": h = h & "8"

        Case "1001": h = h & "9"

        Case "1010": h = h & "A"

        Case "1011": h = h & "B"

        Case "1100": h = h & "C"

        Case "1101": h = h & "D"

        Case "1110": h = h & "E"

        Case "1111": h = h & "F"
        End Select

    Next i

    While Left$(h, 1) = "0"

        h = Right$(h, Len(h) - 1)

    Wend

    Bin2Hex = h
End Function

' 用途：将二进制转化为八进制
' 输入：Bin(二进制数)
' 输入数据类型：String
' 输出：Bin2Oct(八进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function Bin2Oct(ByVal RHS As String) As String
    Dim i As Long
    Dim h As String
    Dim u As Long
    
    RHS = Trim$(RHS)

    If IsBinValue(RHS) = False Then
        Bin2Oct = "输入类型错误"

        Exit Function
    End If
    
    If Len(RHS) Mod 3 <> 0 Then
        RHS = String$(3 - Len(RHS) Mod 3, "0") & RHS
    End If
    
    u = Len(RHS)

    For i = 1 To u Step 3
        Select Case Mid$(RHS, i, 3)
        Case "000": h = h & "0"
        Case "001": h = h & "1"
        Case "010": h = h & "2"
        Case "011": h = h & "3"
        Case "100": h = h & "4"
        Case "101": h = h & "5"
        Case "110": h = h & "6"
        Case "111": h = h & "7"
        End Select
    Next i

    While Left$(h, 1) = "0"
        h = Right$(h, Len(h) - 1)
    Wend

    Bin2Oct = h
End Function

' 用途：将十进制转化为二进制
' 输入：Dec(十进制数)
' 输入数据类型：Long
' 输出：DEC_to_BIN(二进制数)
' 输出数据类型：String
' 输入的最大数为2147483647,输出最大数为1111111111111111111111111111111(31个1)
Public Function Dec2Bin(ByVal RHS As String) As Variant
    RHS = Trim$(RHS)

    If IsDecValue(RHS) = False Then
        Dec2Bin = "输入类型错误"

        Exit Function
    End If
    
    Dec2Bin = ""
    
    Do While RHS > 0
        Dec2Bin = RHS Mod 2 & Dec2Bin
        RHS = RHS \ 2
    Loop
End Function

' 用途：将十进制转化为十六进制
' 输入：Dec(十进制数)
' 输入数据类型：Long
' 输出：Dec2Hex(十六进制数)
' 输出数据类型：String
' 输入的最大数为2147483647,输出最大数为7FFFFFFF
Public Function Dec2Hex(ByVal RHS As String) As Variant
    Dim a As String
    Dim v As Long
    
    RHS = Trim$(RHS)

    If IsDecValue(RHS) = False Then
        Dec2Hex = "输入类型错误"

        Exit Function
    End If
    
    Dec2Hex = ""
    v = RHS

    Do While v > 0
        a = CStr(v Mod 16)

        Select Case a
        Case "10": a = "A"
        Case "11": a = "B"
        Case "12": a = "C"
        Case "13": a = "D"
        Case "14": a = "E"
        Case "15": a = "F"
        End Select

        v = v \ 16
        Dec2Hex = a & Dec2Hex
    Loop
End Function

' 用途：将十进制转化为八进制
' 输入：Dec(十进制数)
' 输入数据类型：Long
' 输出：Dec2Oct(八进制数)
' 输出数据类型：String
' 输入的最大数为2147483647,输出最大数为17777777777
Public Function Dec2Oct(ByVal RHS As String) As Variant
    RHS = Trim$(RHS)

    If IsDecValue(RHS) = False Then
        Dec2Oct = "输入类型错误"

        Exit Function
    End If

    Dec2Oct = ""

    Do While RHS > 0
        Dec2Oct = RHS Mod 8 & Dec2Oct
        RHS = RHS \ 8
    Loop
End Function

' 用途：将十六进制转化为二进制
' 输入：Hex(十六进制数)
' 输入数据类型：String
' 输出：Hex2Bin(二进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function Hex2Bin(ByVal RHS As String) As Variant
    Dim i As Long
    Dim b As String
    Dim l As Long
    
    RHS = UCase$(Trim$(RHS))

    If IsHexValue(RHS) = False Then
        Hex2Bin = "输入类型错误"

        Exit Function
    End If
    
    l = Len(RHS)

    For i = 1 To l
        Select Case Mid$(RHS, i, 1)
        Case "0": b = b & "0000"
        Case "1": b = b & "0001"
        Case "2": b = b & "0010"
        Case "3": b = b & "0011"
        Case "4": b = b & "0100"
        Case "5": b = b & "0101"
        Case "6": b = b & "0110"
        Case "7": b = b & "0111"
        Case "8": b = b & "1000"
        Case "9": b = b & "1001"
        Case "A": b = b & "1010"
        Case "B": b = b & "1011"
        Case "C": b = b & "1100"
        Case "D": b = b & "1101"
        Case "E": b = b & "1110"
        Case "F": b = b & "1111"
        Case Else
            Exit Function
        End Select
    Next

    While Left$(b, 1) = "0"
        b = Right$(b, Len(b) - 1)
    Wend

    Hex2Bin = b
End Function

' 用途：将十六进制转化为十进制
' 输入：Hex(十六进制数)
' 输入数据类型：String
' 输出：Hex2Dec(十进制数)
' 输出数据类型：Long
' 输入的最大数为7FFFFFFF,输出的最大数为2147483647
Public Function Hex2Dec(ByVal RHS As String) As Variant
    Dim i As Long
    Dim b As Long
    Dim u As Long
    
    RHS = UCase$(Trim$(RHS))

    If IsHexValue(RHS) = False Then
        Hex2Dec = "输入类型错误"

        Exit Function
    End If

    u = Len(RHS)

    On Error Resume Next

    For i = 1 To u
        Select Case Mid$(RHS, Len(RHS) - i + 1, 1)
        Case "0": b = b + 16 ^ (i - 1) * 0
        Case "1": b = b + 16 ^ (i - 1) * 1
        Case "2": b = b + 16 ^ (i - 1) * 2
        Case "3": b = b + 16 ^ (i - 1) * 3
        Case "4": b = b + 16 ^ (i - 1) * 4
        Case "5": b = b + 16 ^ (i - 1) * 5
        Case "6": b = b + 16 ^ (i - 1) * 6
        Case "7": b = b + 16 ^ (i - 1) * 7
        Case "8": b = b + 16 ^ (i - 1) * 8
        Case "9": b = b + 16 ^ (i - 1) * 9
        Case "A": b = b + 16 ^ (i - 1) * 10
        Case "B": b = b + 16 ^ (i - 1) * 11
        Case "C": b = b + 16 ^ (i - 1) * 12
        Case "D": b = b + 16 ^ (i - 1) * 13
        Case "E": b = b + 16 ^ (i - 1) * 14
        Case "F": b = b + 16 ^ (i - 1) * 15
        Case Else
            Exit Function
        End Select
    Next
    
    If Err Then
        Hex2Dec = Err.Description

        Exit Function
    End If
    
    Hex2Dec = b
End Function

' 用途：将十六进制转化为八进制
' 输入：Hex(十六进制数)
' 输入数据类型：String
' 输出：Hex2Oct(八进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function Hex2Oct(ByVal RHS As String) As Variant
    Dim Bin As String
    
    RHS = UCase$(Trim$(RHS))

    If IsHexValue(RHS) = False Then
        Hex2Oct = "输入类型错误"

        Exit Function
    End If
    
    Bin = Hex2Bin(RHS)
    Hex2Oct = Bin2Oct(Bin)
End Function

' 用途：将八进制转化为二进制
' 输入：Oct(八进制数)
' 输入数据类型：String
' 输出：Oct2Bin(二进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function Oct2Bin(ByVal RHS As String) As Variant
    Dim i As Long
    Dim b As String
    Dim u As Long
    
    RHS = UCase$(Trim$(RHS))

    If IsOctValue(RHS) = False Then
        Oct2Bin = "输入类型错误"

        Exit Function
    End If

    u = Len(RHS)

    For i = 1 To u
        Select Case Mid$(RHS, i, 1)
        Case "0": b = b & "000"
        Case "1": b = b & "001"
        Case "2": b = b & "010"
        Case "3": b = b & "011"
        Case "4": b = b & "100"
        Case "5": b = b & "101"
        Case "6": b = b & "110"
        Case "7": b = b & "111"
        End Select
    Next

    While Left$(b, 1) = "0"
        b = Right$(b, Len(b) - 1)
    Wend

    Oct2Bin = b
End Function

' 用途：将八进制转化为十进制
' 输入：Oct(八进制数)
' 输入数据类型：String
' 输出：Oct2Dec(十进制数)
' 输出数据类型：Long
' 输入的最大数为17777777777,输出的最大数为2147483647
Public Function Oct2Dec(ByVal RHS As String) As Variant
    Dim i As Long
    Dim b As Long
    Dim l As Long

    RHS = UCase$(Trim$(RHS))

    If IsOctValue(RHS) = False Then
        Oct2Dec = "输入类型错误"

        Exit Function
    End If
    
    l = Len(RHS)

    For i = 1 To l
        Select Case Mid$(RHS, Len(RHS) - i + 1, 1)
        Case "0": b = b + 8 ^ (i - 1) * 0
        Case "1": b = b + 8 ^ (i - 1) * 1
        Case "2": b = b + 8 ^ (i - 1) * 2
        Case "3": b = b + 8 ^ (i - 1) * 3
        Case "4": b = b + 8 ^ (i - 1) * 4
        Case "5": b = b + 8 ^ (i - 1) * 5
        Case "6": b = b + 8 ^ (i - 1) * 6
        Case "7": b = b + 8 ^ (i - 1) * 7
        End Select
    Next

    Oct2Dec = b
End Function

' 用途：将八进制转化为十六进制
' 输入：Oct(八进制数)
' 输入数据类型：String
' 输出：Oct2Hex(十六进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function Oct2Hex(ByVal RHS As String) As Variant
    Dim sBin As String
    
    RHS = UCase$(Trim$(RHS))

    If IsOctValue(RHS) = False Then
        Oct2Hex = "输入类型错误"

        Exit Function
    End If
    
    sBin = Oct2Bin(RHS)
    Oct2Hex = Bin2Hex(sBin)
End Function
