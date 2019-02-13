VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExeProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'frmMain
Public FormMain_WindowState As Long
Public FormMain_SplitterPos As Long

'Font
Public FontSize As Long
Public FontName As String

'TimeStamp
Public TimeStamp_Format As String
Public TimeStamp_StartDate As String


Public Function Save() As Boolean
    Dim ini As INIProfile
    
    Set ini = New INIProfile
    
    With ini
        .ExeFolderPath = App.Path
        .Name = App.title
        
        .SetLong "FormMain", "WindowState", FormMain_WindowState
        .SetLong "FormMain", "SplitterPos", FormMain_SplitterPos
        
        .SetLong "Font", "Size", FontSize
        .SetString "Font", "Name", FontName
        
        .SetString "TimeStamp", "Format", TimeStamp_Format
        .SetString "TimeStamp", "StartDate", TimeStamp_StartDate
    End With
    
    Set ini = Nothing
End Function

Public Sub Load()
    Dim ini As INIProfile
    
    Set ini = New INIProfile
    
    With ini
        .ExeFolderPath = App.Path
        .Name = App.title
        
        FormMain_WindowState = .GetLong("FormMain", "WindowState", 0)
        FormMain_SplitterPos = .GetLong("FormMain", "SplitterPos", 0)
        
        FontSize = .GetLong("Font", "Size", frmMain!txtExpression.Font.Size)
        FontName = .GetString("Font", "Name", frmMain!txtExpression.Font.Name)
    End With
    Set ini = Nothing
End Sub
