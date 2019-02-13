VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------
' Component  :  StringBuffer
' Project    :  fzCore
'
' Description:  字符串缓存类
'
' CreateBy   :  fangzi
' Modified   :
'--------------------------------------------------------

Option Explicit

Public Size          As Long

Private Buffer       As String
Private InitCapacity As Long

'--------------------------------------------------
' Procedure  :  Add
' Description:  添加参数到缓存
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :  Values() (Variant)
'--------------------------------------------------
Public Sub Add(ParamArray Values() As Variant)
Attribute Add.VB_Description = "添加参数到缓存"
    Dim i As Long
    Dim l As Long
    
    l = UBound(Values)

    For i = 0 To l
        AddText Values(i)
    Next
End Sub

'--------------------------------------------------
' Procedure  :  AddLine
' Description:  添加文本到缓存，并插入换行符
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :  Value (Variant)
'--------------------------------------------------
Public Sub AddLine(Value)
Attribute AddLine.VB_Description = "添加文本到缓存，并插入换行符"
    AddText Value
    AddText vbCrLf
End Sub

'--------------------------------------------------
' Procedure  :  AddLines
' Description:  添加多行文本到缓存，并插入换行符
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :  Values() (Variant)
'--------------------------------------------------
Public Sub AddLines(ParamArray Values() As Variant)
Attribute AddLines.VB_Description = "添加多行文本到缓存，并插入换行符"
    Dim i As Long
    Dim u As Long
    
    u = UBound(Values)

    For i = 0 To u
        AddText Values(i)
    Next

    AddText vbCrLf
End Sub

'--------------------------------------------------
' Procedure  :  AddText
' Description:  添加文本到缓存
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :  Value (Variant)
'--------------------------------------------------
Public Sub AddText(Value)
Attribute AddText.VB_Description = "添加文本到缓存"
    Dim nLen As Long

    nLen = Len(Value)

    If Size + nLen >= Len(Buffer) Then
        Buffer = Buffer & String$(GetInitCapacity, Chr$(0))
    End If
    
    Mid$(Buffer, Size + 1, nLen) = Value
    Size = Size + nLen
End Sub

'--------------------------------------------------
' Procedure  :  BufferSize
' Description:  返回StringBuffer的缓冲区大小
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :
'--------------------------------------------------
Public Function BufferSize() As Long
Attribute BufferSize.VB_Description = "返回StringBuffer的缓冲区大小"
    BufferSize = Len(Buffer)
End Function

'--------------------------------------------------
' Procedure  :  CharAt
' Description:  返回Index位置上的字符
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :  Index (Long)
'--------------------------------------------------
Public Function CharAt(ByVal Index As Long) As String
Attribute CharAt.VB_Description = "返回Index位置上的字符"
    CharAt = Mid$(Buffer, Index, 1)
End Function

'--------------------------------------------------
' Procedure  :  NewLine
' Description:  在缓存末尾追加一个回车换行
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :
'--------------------------------------------------
Public Sub NewLine()
Attribute NewLine.VB_Description = "在缓存末尾追加一个回车换行"
    AddText vbCrLf
End Sub

'--------------------------------------------------
' Procedure  :  Remove
' Description:  从startSize开始删除length个字符
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :  StartSize (Long)
'               Length (Long)
'--------------------------------------------------
Public Sub Remove(StartSize As Long, Length As Long)
Attribute Remove.VB_Description = "从startSize开始删除length个字符"
    Mid$(Buffer, StartSize + 1, Length) = Chr$(0)
    Size = Size - Length
End Sub

'--------------------------------------------------
' Procedure  :  RemoveReverse
' Description:  从当前Size开始，向前删除length个字符
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :  Length (Long)
'--------------------------------------------------
Public Sub RemoveReverse(Length As Long)
Attribute RemoveReverse.VB_Description = "从当前Size开始，向前删除length个字符"
    Mid$(Buffer, Size - Length + 1, Length) = Chr$(0)
    Size = Size - Length
End Sub

'--------------------------------------------------
' Procedure  :  ReplaceFrom
' Description:  从SizeFrom开始，用replacement替换后续字符
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :  SizeFrom (Long)
'               Replacement (String)
'--------------------------------------------------
Public Sub ReplaceFrom(SizeFrom As Long, Replacement As String)
    Mid$(Buffer, SizeFrom + 1, Len(Replacement)) = Replacement
    Size = SizeFrom + Len(Replacement)
End Sub

'--------------------------------------------------
' Procedure  :  ReplaceLast
' Description:  用Replacement替换最后一个字符
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :  Replacement (String)
'--------------------------------------------------
Public Sub ReplaceLast(Replacement As String)
Attribute ReplaceLast.VB_Description = "用Replacement替换最后一个字符"
    Mid$(Buffer, Size, Len(Replacement)) = Replacement
    Size = Size + Len(Replacement) - 1
End Sub

'--------------------------------------------------
' Procedure  :  ReplaceReverse
' Description:  用Replacement倒数Length个字符
' Date-Time  :  2010-11-14-21:00:18
'
' Parameters :  Replacement (String)
'               Length (Long)
'--------------------------------------------------
Public Sub ReplaceReverse(Replacement As String, Length As Long)
Attribute ReplaceReverse.VB_Description = "用Replacement倒数Length个字符"
    Mid$(Buffer, Size - Length + 1, Len(Replacement)) = Replacement
    Size = Size + Len(Replacement) - Length
End Sub

'--------------------------------------------------
' Procedure  :  Reset
' Description:  清空缓存，将缓存大小重置为Capacity
' Date-Time  :  2010-11-14-21:00:17
'
' Parameters :  Capacity (Long)
'--------------------------------------------------
Public Sub Reset(ByVal Capacity As Long)
Attribute Reset.VB_Description = "清空缓存，将缓存大小重置为Capacity"
    InitCapacity = Capacity
    Buffer = String$(Capacity, Chr$(0))
    Size = 0
End Sub

'--------------------------------------------------
' Procedure  :  SetCapacity
' Description:  设置缓存大小
' Date-Time  :  2010-11-14-21:00:17
'
' Parameters :  Value (Long = 1024&)
'--------------------------------------------------
Public Sub SetCapacity(Optional Value As Long = 1024&)
Attribute SetCapacity.VB_Description = "设置缓存大小"
    InitCapacity = Value
    Buffer = String$(Value, Chr$(0))
End Sub

'--------------------------------------------------
' Procedure  :  ToString
' Description:  返回缓存中的字符串
' Date-Time  :  2010-11-14-21:00:17
'
' Parameters :
'--------------------------------------------------
Public Function ToString() As String
Attribute ToString.VB_Description = "返回缓存中的字符串"
    ToString = Left$(Buffer, Size)
End Function

'获取初始缓存大小
Private Function GetInitCapacity() As Long
    GetInitCapacity = IIf((InitCapacity = 0), 1024, InitCapacity)
End Function

Private Sub Class_Initialize()
    Buffer = String$(GetInitCapacity, Chr$(0))
End Sub

Private Sub Class_Terminate()
    Buffer = ""
End Sub
