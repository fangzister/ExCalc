Attribute VB_Name = "modBrowseDirectory"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHGetPathFromIdlist Lib "shell32" Alias "SHGetPathFromIDList" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As browseinfo) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Sub SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nfolder As Long, pidl As ItemIdList)

Private Const WM_USER = &H400
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_INITIALIZED   As Long = 1
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const lPtr = (LMEM_FIXED + LMEM_ZEROINIT)
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_USENEWUI = &H40
Private Const BIF_BROWSEINCLUDEFILES = &H4000
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_PATHMUSTEXIST = &H800
Private Const MAX_PATH As Long = 260&

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

Private ofn As OPENFILENAME

Private Type browseinfo
    hOwner As Long
    pidlroot As Long
    pszdisplayname As String
    lpsztitle As String
    uiFlags As Long
    lpfnCallBack As Long
    lparam As Long
    iimage As Long
End Type

Private Type ShiteMid
    cb As Long
    abid As Byte
End Type

Private Type ItemIdList
    mkid As ShiteMid
End Type

Enum SpecialFolderTypeConstants
    Desktop = &H0
    Programs = &H2
    Controls = &H3
    Printers = &H4
    Persional = &H5
    Favorites = &H6
    Startup = &H7
    Recent = &H8
    SendTo = &H9
    BitBucket = &HA
    StartMenu = &HB
    DesktopDrectory = &H10
    Dirves = &H11
    Network = &H12
    Nethood = &H13
    Fonts = &H14
    Templates = &H15
End Enum

Public Function GetSpecialFolder(FolderType As SpecialFolderTypeConstants) As String
    Dim nRet    As Long
    Dim idl     As ItemIdList
    Dim sPath   As String

    'Get the special folder
    Call SHGetSpecialFolderLocation(100, FolderType, idl)
    
    'Create a buffer
    sPath = Space$(512)
    'Get the path from the IDList
    nRet = SHGetPathFromIdlist(ByVal idl.mkid.cb, ByVal sPath)
    'Remove the unnecessary chr$(0)'s
    GetSpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
End Function

Public Function OpenFile( _
    Optional OwnerHwnd As Long = 0, _
    Optional title As String = "打开", _
    Optional initDir As String = "", _
    Optional SelectedFileName As String = "", _
    Optional Filter As String = "", _
    Optional SaveMode As Boolean = False) As String

    Dim i As Long
    Dim u As Long
    Dim s() As String
    Dim nRet As Long
    Dim iDelim As Long
    Dim filters() As String
    
    If Len(Filter) > 0 Then
        filters = Split(Filter, ";")
        u = UBound(filters)

        For i = 0 To u
            s = Split(filters(i), "|")
            ofn.lpstrFilter = ofn.lpstrFilter & s(0) & "(" & s(1) & ")" & Chr$(0) & s(1) & Chr$(0)

            If i = 0 Then
                ofn.lpstrDefExt = s(1)
            End If
        Next
    End If
    
    With ofn
        .hwndOwner = OwnerHwnd
        .hInstance = App.hInstance
        .lpstrTitle = title
        .lpstrFileTitle = vbNullString
        .lpstrInitialDir = initDir
        .lpstrFile = SelectedFileName & String$(MAX_PATH, 0)
        .nFilterIndex = 1
        .lpstrFilter = .lpstrFilter & Chr$(0)
        .Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        .nMaxFile = MAX_PATH  '显示文件名的长度
        .lStructSize = Len(ofn)
    End With
    
    If SaveMode Then
        nRet = GetSaveFileName(ofn) '取得文件名
    Else
        nRet = GetOpenFileName(ofn) '取得文件名
    End If
    
    If nRet <> 0 Then
        iDelim = InStr(ofn.lpstrFile, Chr$(0))
        
        If iDelim > 0 Then
            OpenFile = Left$(ofn.lpstrFile, iDelim - 1)
        End If
    End If

End Function

Public Function BrowseDirectory(ByVal hwnd As Long, _
                                strPrompt As String, _
                                Optional initDir As String, _
                                Optional ByVal dirOnly As Boolean = True, _
                                Optional ByVal newUI As Boolean = True) As String

    On Error GoTo ehBrowseForFolder

    Dim intNull   As Integer
    Dim lngIDList As Long
    Dim lngResult As Long
    Dim strPath   As String
    Dim udtBI     As browseinfo
    Dim ret       As Long

    With udtBI
        .hOwner = hwnd
        .lpsztitle = strPrompt

        If Len(initDir) > 0 Then
            .lpfnCallBack = MyAddressOf(AddressOf BrowseForFolders_CallbackProc)
            ret = LocalAlloc(lPtr, VBA.Len(initDir) + 1)
            CopyMemory ByVal ret, ByVal initDir, Len(initDir) + 1
            .lparam = ret
        End If
        
        If newUI Then
            .uiFlags = .uiFlags Or BIF_USENEWUI
        End If
        
        If dirOnly Then
            .uiFlags = .uiFlags Or BIF_RETURNONLYFSDIRS
        Else
            .uiFlags = .uiFlags Or BIF_BROWSEINCLUDEFILES
        End If
        
        .uiFlags = .uiFlags Or BIF_RETURNFSANCESTORS
    End With

    lngIDList = SHBrowseForFolder(udtBI)

    If lngIDList <> 0 Then
        strPath = String$(MAX_PATH, Chr$(0))
        lngResult = SHGetPathFromIdlist(lngIDList, strPath)
        Call CoTaskMemFree(lngIDList)
        intNull = InStr(strPath, vbNullChar)

        If intNull > 0 Then
            strPath = Left$(strPath, intNull - 1)
        End If
    End If

    BrowseDirectory = strPath

    Exit Function

ehBrowseForFolder:
    BrowseDirectory = Empty
End Function

Private Function MyAddressOf(AddressOfX) As Long
    MyAddressOf = AddressOfX
End Function

Private Function BrowseForFolders_CallbackProc(ByVal hwnd As Long, _
                                               ByVal uMsg As Long, _
                                               ByVal lparam As Long, _
                                               ByVal lpData As Long) As Long

    If uMsg = BFFM_INITIALIZED Then
        SendMessage hwnd, BFFM_SETSELECTIONA, True, ByVal lpData
    End If
End Function

