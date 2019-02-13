Attribute VB_Name = "modURL"
Option Explicit

'将URL保存为mht文件
Public Function SavePageToMHT(ByVal URL As String, ByVal FileName As String) As Boolean

    Dim objMsg    As New CDO.Message

    Dim objStream As ADODB.Stream

    On Error GoTo eH:
    
    objMsg.CreateMHTMLBody URL, cdoSuppressAll, "", ""
    Set objStream = objMsg.GetStream
    objStream.SaveToFile FileName, adSaveCreateOverWrite
    SavePageToMHT = True

    Exit Function

eH:
End Function

Public Function GetTiebaReplyCount(ByVal URL As String) As Long
    
End Function
