VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRXReplaceQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Replacer(Match As String, Index As Long, FullText As String) As String
Attribute Replacer.VB_UserMemId = 0
Dim sTemp As String

sTemp = Match
sTemp = Replace(sTemp, "\\", Chr$(7)) '\\ has higher priority,so hide it
sTemp = Replace$(sTemp, "\q", Chr$(34) & Chr$(34))
sTemp = Replace$(sTemp, "\n", Chr$(34) & " & vbcrlf & " & Chr$(34))
sTemp = Replace$(sTemp, "\t", vbTab)
sTemp = Replace(sTemp, "\", "")          'use \\ to indicate \
sTemp = Replace(sTemp, Chr$(7), "\")

Replacer = sTemp

End Function
