VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Base64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************
'人人为我，我为人人
'枕善居
'发布日期：05/05/30
'描  述：Base64 编码/解码测试程序
'网  站：http://www.mndsoft.com/
'e-mail：mnd@mndsoft.com
'OICQ  : 88382850
'****************************************************************************
Option Explicit

Public Event BlockRead(ByVal lngCurrentBlock As Long, ByVal lngTotalBlocks As Long, ByVal lBlockMode As enBlockMode, ByRef bCancel As Boolean)
Public Event FileDecodeComplete(strOriginalFile As String, strDecodedFile As String)
Public Event FileEncodeComplete(strOriginalFile As String, strEncodedFile As String)
Public Event ErrorOccured(ByVal lngCode As Long, ByVal strDescription As String)
Public Event BeforeFileOpenIn(ByVal strFileName As String, ByRef bCancel As Boolean)
Public Event BeforeFileOpenOut(ByVal strFileName As String, ByRef bCancel As Boolean)
Public Event AfterFileCloseIn(ByVal strFileName As String)
Public Event AfterFileCloseOut(ByVal strFileName As String)

Private m_oXMLDoc          As Object
Private m_oXMLElement      As Object
Private m_bXMLObjectOK     As Boolean
Private m_lngLastErrorCode As Long
Private m_strLastErrorDesc As String
Private m_lngEncBlockSize  As Long
Private m_lngDecBlockSize  As Long
Private m_bCancel          As Boolean

Public Enum enBlockMode
    b64Encode
    b64Decode
End Enum

Private Const cMSXMLDom = "MSXML.DomDocument"
Private Const cDefBlockSize = 1048576 '默认块大小: 1Mb

Private Sub Class_Initialize()
    Set m_oXMLDoc = fCreateObject(cMSXMLDom)
    ClearError
    m_bXMLObjectOK = Not m_oXMLDoc Is Nothing

    If Not m_bXMLObjectOK Then
        RaiseError "90001", "错误,建立必须 '" & cMSXMLDom & "' 对象"
    Else
        Set m_oXMLElement = m_oXMLDoc.createElement("TmpBase64")
        m_oXMLElement.DataType = "bin.base64"
    End If

    '初始化块大小大约 1Mb
    Me.BlockSize = cDefBlockSize
End Sub

'释放对象
Private Sub Class_Terminate()
    Set m_oXMLElement = Nothing
    Set m_oXMLDoc = Nothing
End Sub

'编码字符串
Public Function Encode(ByVal strIn As String) As String
    Dim oElement As Object
    Dim sData    As String
    Dim aData()  As Byte
    
    On Error GoTo errEncode:

    If m_bXMLObjectOK And Len(strIn) > 0 Then
        aData = StrConv(strIn, vbFromUnicode)
        m_oXMLElement.nodetypedvalue = aData
        Encode = m_oXMLElement.Text
        m_oXMLElement.Text = "" '释放内存
    End If

    Exit Function
errEncode:
    RaiseError Err.Number, Err.Description
End Function

'解码字符串
Public Function Decode(ByVal strIn As String) As String
    Dim oElement As Object
    
    On Error GoTo errDecode:
    
    If m_bXMLObjectOK And Len(strIn) > 0 Then
        m_oXMLElement.Text = strIn
        Decode = StrConv(m_oXMLElement.nodetypedvalue, vbUnicode)
        m_oXMLElement.Text = "" '释放内存
    End If

    Exit Function
errDecode:
    RaiseError Err.Number, Err.Description
End Function

Public Function EncodeFile(ByRef strInFile As String, ByRef strOutFile As String) As Boolean
    EncodeFile = EncDecFile(strInFile, strOutFile, True)
End Function

Public Function DecodeFile(ByRef strInFile As String, ByRef strOutFile As String) As Boolean
    DecodeFile = EncDecFile(strInFile, strOutFile, False)
End Function

Public Property Get LastErrorCode() As Long
    LastErrorCode = m_lngLastErrorCode
End Property

Public Property Get LastErrorDescription() As String
    LastErrorDescription = m_strLastErrorDesc
End Property

Public Property Let BlockSize(ByVal lngBytes As Long)
    m_lngEncBlockSize = ((lngBytes \ 54)) * 54 + IIf(IsOnBoundary(lngBytes, 54), 0, 54)
    m_lngDecBlockSize = ((lngBytes \ 584)) * 584 + IIf(IsOnBoundary(lngBytes, 584), 0, 584)
End Property

Public Sub ClearError()
    m_lngLastErrorCode = 0
    m_strLastErrorDesc = ""
End Sub

Private Function IsOnBoundary(lLen As Long, lBoundary As Long) As Boolean
    IsOnBoundary = lLen Mod lBoundary = 0
End Function

Private Sub RaiseError(ByVal lngCode As Long, ByVal strDescription As String)
    m_lngLastErrorCode = lngCode
    m_strLastErrorDesc = strDescription
    RaiseEvent ErrorOccured(lngCode, strDescription)
End Sub

Private Function fCreateObject(sObject As String) As Object
    On Error GoTo errCreateObject

    Set fCreateObject = CreateObject(sObject)

    On Error GoTo 0
    
    Exit Function
errCreateObject:
    Set fCreateObject = Nothing
End Function

'编码/解码文件
Private Function EncDecFile(strInFile As String, strOutFile As String, bEncode As Boolean)
    Dim lBlockCount As Long
    Dim lSize       As Long
    Dim lMaxBlocks  As Long
    Dim lRest       As Long
    Dim strTMP      As String
    Dim FFIn        As Integer
    Dim FFOut       As Integer
    Dim lBlockSize  As Long
    Dim strOut      As String
    
    On Error GoTo errEncDecFile

    If bEncode Then lBlockSize = m_lngEncBlockSize Else lBlockSize = m_lngDecBlockSize
    
    m_bCancel = False
    lBlockCount = 0

    If Len(Dir(strInFile)) > 0 Then '文件是否存在
        RaiseEvent BeforeFileOpenIn(strInFile, m_bCancel)

        If m_bCancel Then Exit Function
        
        FFIn = FreeFile             '获取空闲文件句柄
        Open strInFile For Binary As #FFIn
        
        RaiseEvent BeforeFileOpenOut(strOutFile, m_bCancel)

        If m_bCancel Then
            Close #FFIn '关闭文件

            Exit Function
        End If
        
        FFOut = FreeFile
        Open strOutFile For Binary As #FFOut
        
        lSize = LOF(FFIn)                '获取文件大小
        lMaxBlocks = lSize \ lBlockSize  '定义读取的块数
        lRest = lSize - (lMaxBlocks * lBlockSize)
        strTMP = String(lBlockSize, Chr(0))             '初始化缓冲区

        While Not EOF(FFIn) And (lBlockCount < lMaxBlocks) And (Not m_bCancel) '读取文件
            Get #FFIn, , strTMP                               '读取块

            If bEncode Then strOut = Me.Encode(strTMP) & IIf(lRest > 0, vbLf, "") Else strOut = Me.Decode(strTMP)
            Put #FFOut, , strOut
            lBlockCount = lBlockCount + 1                       '统计块
            RaiseEvent BlockRead(lBlockCount, lMaxBlocks + IIf(lRest > 0, 1, 0), IIf(bEncode, b64Encode, b64Decode), m_bCancel)
        Wend

        If (lRest > 0) And (Not EOF(FFIn)) And (Not m_bCancel) Then     '是否结束
            strTMP = String(lRest, Chr(0))                      '初始化缓冲区
            Get #FFIn, , strTMP                                 '读取块

            If bEncode Then strOut = Me.Encode(strTMP) Else strOut = Me.Decode(strTMP)
            Put #FFOut, , strOut
            RaiseEvent BlockRead(lBlockCount + 1, lMaxBlocks + 1, IIf(bEncode, b64Encode, b64Decode), m_bCancel)
        End If

        Close #FFOut
        RaiseEvent AfterFileCloseOut(strOutFile)
        
        Close #FFIn
        RaiseEvent AfterFileCloseIn(strInFile)

        If bEncode Then
            RaiseEvent FileEncodeComplete(strInFile, strOutFile)
        Else
            RaiseEvent FileDecodeComplete(strInFile, strOutFile)
        End If

        EncDecFile = True
    Else
        RaiseError 90002, "文件不存在 / 打开文件错误:" & strInFile
    End If

    Exit Function
errEncDecFile:
    RaiseError Err.Number, Err.Description
End Function
