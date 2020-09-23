VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRXReplaceRepeated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Replacer(Match As String, SubMatch1 As String, SubMatch2 As String, Index As Long, FullText As String) As String
Attribute Replacer.VB_UserMemId = 0
Dim sTemp As String
Dim iRepeat As Long
Dim sChar As String

iRepeat = CLng(SubMatch1)
sChar = SubMatch2

If iRepeat > 0 Then
    Replacer = String$(iRepeat, sChar)
Else
    Replacer = ""
End If

End Function

