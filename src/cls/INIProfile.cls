VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : INIProfile
'    Project    : fzCore
'
'    Description: Ini文件操作类
'
'    Modified   : 2012.7.4
'--------------------------------------------------------------------------------
Option Explicit

Private mProfileName   As String
Private mExeFolderPath As String

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Const STRING_BUFFER_SIZE = 99999 '字符串缓存大小

'ini文件名
Public Property Get Name() As String
Attribute Name.VB_Description = "返回/设置ini文件名"
    Name = mProfileName
End Property

Public Property Let Name(ByVal Value As String)
    mProfileName = Value
End Property

'exe文件的所在目录路径
Public Property Get ExeFolderPath() As String
Attribute ExeFolderPath.VB_Description = "返回/设置exe文件的所在目录路径"
    ExeFolderPath = mExeFolderPath
End Property

Public Property Let ExeFolderPath(ByVal Value As String)
    mExeFolderPath = Value
End Property

'返回配置文件路径
Public Property Get ProfilePath() As String
Attribute ProfilePath.VB_Description = "返回配置文件路径"
    ProfilePath = mExeFolderPath & "\" & mProfileName & ".ini"
End Property

'测试当前Section是否存在某个值
Public Function IsValueExists(ByVal Section As String, ByVal Value As Variant, Optional ByVal CompareMode As VbCompareMethod = VbCompareMethod.vbTextCompare) As Boolean
    Dim p() As String
    Dim i As Long
    Dim v As String
    Dim u As Long
    
    p = GetAllKeys(Section)
    
    u = UBound(p)
    If u = -1 Then
        Exit Function
    End If
    
    Select Case TypeName(Value)
    Case "String"
        For i = 0 To u
            If StrComp(Value, GetString(Section, p(i)), CompareMode) = 0 Then
                IsValueExists = True
                Exit Function
            End If
        Next
    Case Else
        For i = 0 To u
            If Value = GetLong(Section, p(i)) Then
                IsValueExists = True
                Exit Function
            End If
        Next
    End Select
    
End Function

'获取值对应的Key
Public Function GetKeyByValue(ByVal Section As String, ByVal Value As Variant, Optional ByVal CompareMode As VbCompareMethod = vbTextCompare) As String
    Dim s As String
    Dim p() As String
    Dim u As Long
    Dim i As Long
    
    p = GetAllKeys(Section)
    u = UBound(p)
    
    If u = -1 Then
        Exit Function
    End If
    
    Select Case TypeName(Value)
    Case "String"
        For i = 0 To u
            If StrComp(Value, GetString(Section, p(i)), CompareMode) = 0 Then
                GetKeyByValue = p(i)
                Exit Function
            End If
        Next
    Case Else
        For i = 0 To u
            If Value = GetLong(Section, p(i)) Then
                GetKeyByValue = p(i)
                Exit Function
            End If
        Next
    End Select
End Function

'获取当前Section中的下一个键值
Public Function GetNextKey(ByVal Section As String) As String
    Dim c As Long
    Dim k As Long
    Dim p() As String
    Dim sLastKey As String
    Dim u As Long
    
    p = GetAllKeys(Section)
    u = UBound(p)
    If u = -1 Then
        GetNextKey = "a"
        Exit Function
    End If
    
    sLastKey = p(UBound(p))
    c = Asc(sLastKey)
    If c < 122 Then 'Asc("z")
        GetNextKey = Chr(c + 1)
    Else
        MsgBox "TODO"
    End If
End Function

'获取Ini字符串值
Public Function GetString(ByVal SectionName As String, ByVal KeyWord As String, Optional ByVal DefString As String = "") As String
Attribute GetString.VB_Description = "获取Ini字符串值"
    Dim ret  As String * 144
    Dim temp As Long
    Dim s    As String
    Dim i    As Integer

    temp = GetPrivateProfileString(SectionName, KeyWord, "", ret, 144, ProfilePath)

    '检索关键词的值
    If temp > 0 Then '关键词的值不为空
        s = ""

        For i = 1 To 144
            If Asc(Mid$(ret, i, 1)) = 0 Then
                Exit For
            Else
                s = s & Mid$(ret, i, 1)
            End If
        Next
    Else
        temp = WritePrivateProfileString(SectionName, KeyWord, DefString, ProfilePath)
        '将缺省值写入INI文件
        s = DefString
    End If

    GetString = s
End Function

'获取Ini数值
Public Function GetLong(ByVal SectionName As String, ByVal KeyWord As String, Optional ByVal DefaultValue As Long = 0) As Long
Attribute GetLong.VB_Description = "获取Ini数值"
    Dim d As Long
    Dim s As String
    
    d = DefaultValue
    GetLong = GetPrivateProfileInt(SectionName, KeyWord, DefaultValue, ProfilePath)

    If d <> DefaultValue Then
        s = "" & d
        d = WritePrivateProfileString(SectionName, KeyWord, s, ProfilePath)
    End If
End Function

'写入字符串值
Public Sub SetString(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)
Attribute SetString.VB_Description = "写入字符串值"
    Dim ret As Long

    ret = WritePrivateProfileString(SectionName, KeyWord, ValStr, ProfilePath)
End Sub

'写入数值
Public Sub SetLong(ByVal SectionName As String, ByVal KeyWord As String, ByVal Value As Long)
Attribute SetLong.VB_Description = "写入数值"
    Dim ret As Long
    Dim s   As String

    s = str$(Value)
    ret = WritePrivateProfileString(SectionName, KeyWord, s, ProfilePath)
End Sub

'清除KeyWord键
Public Sub DeleteKey(ByVal SectionName As String, ByVal KeyWord As String)
Attribute DeleteKey.VB_Description = "清除KeyWord键"
    Dim ret As Long

    ret = WritePrivateProfileString(SectionName, KeyWord, 0&, ProfilePath)
End Sub

'清除 Section"段"(Sub)，如果是清除section就少写一个Key多一个""。
Public Sub DeleteSection(ByVal SectionName As String)
Attribute DeleteSection.VB_Description = "清除 Section""段""(Sub)，如果是清除section就少写一个Key多一个""""。"
    Dim ret As Long

    ret = WritePrivateProfileString(SectionName, 0&, "", ProfilePath)
End Sub


'读取 小节 下所有 子键名
Public Function GetAllKeys(ByVal AppName As String) As String()
    Dim str1 As String, str2 As String
    Dim i As Long, k As Long
    Dim p() As String
    Dim Rec As Long

    str1 = Space$(STRING_BUFFER_SIZE)
    Rec = GetPrivateProfileString(AppName, 0&, "", str1, STRING_BUFFER_SIZE, ProfilePath)

    k = InStr(str1 & Chr$(0), String$(2, Chr$(0)))
    If k > 0 Then
        str1 = Left$(str1, k - 1)
    Else
        str1 = ""
    End If
    p = Split(str1, Chr$(0))
    GetAllKeys = p
End Function

'读取所有 小节名
Public Function GetAllSections() As String()
    Dim str1 As String, str2 As String
    Dim i As Long, k As Long, Rec As Long
    Dim p() As String

    str1 = Space$(STRING_BUFFER_SIZE)
    Rec = GetPrivateProfileString(vbNullString, 0&, "", str1, STRING_BUFFER_SIZE, ProfilePath)

    k = InStr(str1 & Chr$(0), String$(2, Chr$(0)))
    If k > 0 Then
        str1 = Left$(str1, k - 1)
    Else
        str1 = ""
    End If
    
    p = Split(str1, Chr$(0))
    
    GetAllSections = p
End Function

