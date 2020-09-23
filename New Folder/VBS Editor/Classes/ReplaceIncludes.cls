VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReplaseIncludes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Replacer(Match As String, SubMatch As String, Index As Long, FullText As String) As String
Attribute Replacer.VB_UserMemId = 0
Dim sFileName As String
Dim sTemp As String
Dim File As New CTextFile

sFileName = Trim$(SubMatch)

If File.FileOpen(sFileName, OpenForInput) Then
        sTemp = File.ReadAll
        File.FileClose
        sTemp = Process_VBS_Xtensions(sTemp)

'try looking in Scripts subfolder:
ElseIf File.FileOpen(RemoveSlash(App.Path) & "\Scripts\" & sFileName, OpenForInput) Then
        sTemp = File.ReadAll
        File.FileClose
        sTemp = Process_VBS_Xtensions(sTemp)
Else
        'Couldn't load file, so raise an error
        'NOT here, but in the script! COOL ;-)
        sTemp = "Err.Raise 53" ' 53=File not found
End If

Replacer = vbCrLf & sTemp & vbCrLf  'just in case, we add CRLF
Set File = Nothing

End Function
