VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRXEntityDecodeASCII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Replacer(Match As String, Index As Long, FullText As String) As String
Attribute Replacer.VB_UserMemId = 0
Dim idx As Long, NumEnt As Long

Replacer = Match ' in case not found

If InStr(Match, "#") > 0 Then
    'Numeric Entity
    NumEnt = Val(Mid$(Match, 3, Len(Match) - 3))
    If NumEnt < 256 Then ' index = char code
                If EntityInfo(NumEnt).Asc127 = "REMOVE" Then
                    Replacer = ""
                ElseIf EntityInfo(NumEnt).Char = "" Then 'Unlisted Entry
                    Replacer = Chr$(NumEnt)
                Else
                    Replacer = EntityInfo(NumEnt).Asc127
                End If
    Else
    For idx = 256 To 300
            If EntityInfo(idx).Code = NumEnt Then
                Replacer = EntityInfo(idx).Asc127
                Exit For
            End If
        Next
    End If

Else
    'Named Entity
    For idx = 0 To 300
        If EntityInfo(idx).Name = Match Then
            Replacer = EntityInfo(idx).Asc127
            Exit For
        End If
    Next

End If


End Function
