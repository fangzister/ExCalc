VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EvidenceIMBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'EvidenceIMBase
'电子数据取证时，从取证软件复制的IM软件账号基本信息

Dim m_AccountInfo As String    '所有内容
Public AccountType As String    '账号类型：微信,QQ
Public Account As String  '账号
Public Nickname As String
Public BindPhone As String
Public BindEmail As String
Public BindQQ As String
Public WeixinID As String

Private Function GetColon(ByVal Item As String) As Long
    GetColon = InStr(1, Item, ":")
End Function

Private Function GetValue(ByVal Item As String, ByVal Colon As Long) As String
    GetValue = Trim$(Mid$(Item, Colon + 2))
End Function


Public Sub Init(ByVal AccountInfo As String)
    Dim i As Long
    Dim j As Long
    Dim p() As String
    Dim m As String
    Dim nColon As Long
    Dim sValue As String
    Dim sType As String
    
    m_AccountInfo = DeleteBlankLines(AccountInfo)
    
    p = Split(m_AccountInfo, vbCrLf)
    
    m = p(0)
    nColon = GetColon(m)
    AccountType = Left$(m, nColon - 1)
    If AccountType = "微信号" Then
        'pass
        Account = GetValue(m, nColon)
    ElseIf AccountType = "QQ号码" Then
        AccountType = "QQ号"
        Account = GetValue(m, nColon)
    Else
        AccountType = ""
        Exit Sub
    End If
    
    For i = 1 To UBound(p)
        m = p(i)
        nColon = GetColon(m)
        If nColon > 0 Then
            sType = Left$(m, nColon - 1)
            
            Select Case sType
            Case "QQ号码"
                BindQQ = GetValue(m, nColon)
            Case "昵称"
                Nickname = GetValue(m, nColon)
            Case "手机号码"
                BindPhone = GetValue(m, nColon)
            Case "微信ID"
                WeixinID = GetValue(m, nColon)
            Case "邮箱地址"
                BindEmail = GetValue(m, nColon)
            Case "手机号"
                BindPhone = GetValue(m, nColon)
            End Select
        End If
    Next
End Sub

Public Function ToString() As String
    Dim s As String
    If Len(AccountType) = 0 Then Exit Function
    
    ToString = "该手机登录" & AccountType & "“" & Account & "”"
    
    If Len(Nickname) > 0 Then
        s = s & "昵称：“" & Nickname & "”，"
    End If
    If Len(WeixinID) > 0 Then
        s = s & "微信ID：" & WeixinID & "，"
    End If
    If Len(BindPhone) > 0 Then
        s = s & "绑定手机号：" & BindPhone & "，"
    End If
    If Len(BindQQ) > 0 Then
        s = s & "绑定QQ号：" & BindQQ & "，"
    End If
    If Len(BindEmail) > 0 Then
        s = s & "绑定邮箱：" & BindEmail & "，"
    End If
    
    If Len(s) > 0 Then
        ToString = ToString & "（" & Left$(s, Len(s) - 1) & "）"
    End If
End Function
