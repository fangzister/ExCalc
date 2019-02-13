VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_DONTADDTORECENT = &H2000000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLEINCLUDENOTIFY = &H400000
Private Const OFN_ENABLESIZING = &H800000
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EX_NOPLACESBAR = &H1
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_FORCESHOWHIDDEN = &H10000000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10
Private Const OFN_USEMONIKERS = &H1000000

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type OPENFILENAME
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    Flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type

Private ofn              As OPENFILENAME
Private Filter           As String
Private Const MAX_PATH   As Long = 260&
Private Const MULTI_PATH As Long = 1024&

Public SelectedFilePath  As String

Public Sub SetDialogTitle(Optional Title As String = "")
    ofn.lpstrTitle = Title
End Sub

Public Sub AddFilter(Optional Description As String = "所有文件", Optional Extension As String = "*")
    ofn.lpstrFilter = ofn.lpstrFilter & Description & "(" & Extension & ")" & Chr$(0) & Extension & Chr$(0)
End Sub

Public Sub SetOwner(ByVal hwnd As Long)
    ofn.hwndOwner = hwnd
End Sub

Public Sub SetBufferSize(BufferSize As Long)
    ofn.nMaxFile = BufferSize
End Sub

Public Sub SetInitDir(Optional initDir As String = "")
    ofn.lpstrInitialDir = initDir
End Sub

Public Sub SetSelectedFileName(Optional FileName As String = "")
    ofn.lpstrFile = FileName & String$(MAX_PATH, 0)
End Sub

Public Sub Setup(ByVal hwnd As Long, _
                 Optional Title As String, _
                 Optional initDir As String = "", _
                 Optional SelectedFileName As String)
    ofn.hwndOwner = hwnd
    ofn.lpstrTitle = Title
    ofn.lpstrInitialDir = initDir
    ofn.lpstrFile = SelectedFileName & String$(MAX_PATH, 0)
End Sub

'返回目录路径，填充文件名数组
'Public Function OpenFiles(path As String, filenames() As String) As Boolean
'    Dim ret As Long
'    Dim nullPos As Long
'    Dim files As String
'
'    With ofn
'        .lStructSize = Len(ofn)
'        .nFilterIndex = 1
'        .lpstrFilter = .lpstrFilter & Chr$(0)
'        .flags = .flags Or OFN_ALLOWMULTISELECT Or OFN_EXPLORER
'        .nMaxFile = MULTI_PATH
'    End With
'
'    ret = GetOpenFileName(ofn)
'
'    If ret <> 0 Then
'        nullPos = InStr(ofn.lpstrFile, Chr$(0))
'
'        If nullPos > 0 Then
'            Stop
'            Dim r As String
'
'            '路径
'            path = Left$(ofn.lpstrFile, nullPos - 1)
'            Debug.Print path
'            files = Right$(ofn.lpstrFile, Len(ofn.lpstrFile) - ofn.nFileOffset)
'            PS files
'            filenames = Split(files, Chr$(0),)
''            PS files
'            PrintStringArray filenames
''            files = Mid$(ofn.lpstrFile, nullPos, Len(ofn.lpstrFile) - nullPos)
''            PS files
''            PS Trim(ofn.lpstrFile)
''            PS path
''            filenames = Split(files, Chr$(0))
'            OpenFiles = True
'        End If
'    End If
'End Function

Public Sub OpenFile()
    Dim ret    As Long
    Dim iDelim As Long
    
    With ofn
        .lStructSize = Len(ofn)
        .nFilterIndex = 1
        .lpstrFilter = .lpstrFilter & Chr$(0)
        .Flags = .Flags Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        .nMaxFile = MAX_PATH  '显示文件名的长度
    End With
    
    ret = GetOpenFileName(ofn) '取得文件名
    
    If ret <> 0 Then
        iDelim = InStr(ofn.lpstrFile, Chr$(0))
        
        If iDelim > 0 Then
            SelectedFilePath = Left$(ofn.lpstrFile, iDelim - 1)
        End If
    End If
End Sub

Public Sub SaveFile()
    Dim ret    As Long
    Dim iDelim As Long
   
    With ofn
        .lStructSize = Len(ofn)
        .nFilterIndex = 1
        .lpstrFilter = .lpstrFilter & Chr$(0)
        .nMaxFile = MAX_PATH
    End With
    
    ret = GetSaveFileName(ofn) '取得文件名
    
    If ret <> 0 Then
        iDelim = InStr(ofn.lpstrFile, Chr$(0))
        
        If iDelim > 0 Then
            SelectedFilePath = Left$(ofn.lpstrFile, iDelim - 1)
        End If
    End If
End Sub

Private Sub Class_Initialize()
    '    Stop
    'Initialize
    With ofn
        .lStructSize = Len(ofn)
        .Flags = OFN_HIDEREADONLY
        .hInstance = App.hInstance
        .lpstrFileTitle = vbNullString
    End With
End Sub

Private Sub Class_Terminate()
    'Terminate
End Sub
