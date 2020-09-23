VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRX_ReplaceArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Replacer(Match As String, SubMatch1 As String, SubMatch2 As String, Index As Long, FullText As String) As String
Attribute Replacer.VB_UserMemId = 0
Dim idx As Integer, sArrayVar As String, sTemp As String
Dim iNumItems As Long, vTempArray As Variant


sTemp = Trim$(SubMatch1) & " = Array(" & SubMatch2 & ")" & vbCrLf

'vTempArray = Split(SubMatch2, ",")
'iNumItems = UBound(vTempArray) - LBound(vTempArray) + 1
'For idx = LBound(vTempArray) To UBound(vTempArray)
'    sTemp = sTemp & sArrayVar & "(" & Format$(idx) & ")=" & vTempArray(idx) & vbCrLf
'Next
'sTemp = "Dim " & sArrayVar & "(" & Format$(iNumItems - 1) & ")" & vbCrLf & sTemp & vbCrLf

Replacer = sTemp

End Function
