Attribute VB_Name = "TextTrigger"
Option Explicit
Public Function HandleTextTrigger(ByVal Text As String, TriggerChars As String) As String

Dim sTemp As String
Dim sResult As String
Dim sTrig As String
Dim sCaption As String

sTemp = Text

Do
    sTrig = FindTrigger(sTemp, TriggerChars)
    If sTrig = "" Then Exit Do
    sCaption = RemoveTriggerChars(sTrig, TriggerChars)
    sResult = InputBox(sCaption, "Enter Data", sCaption)
    sTemp = Replace(sTemp, sTrig, sResult)
Loop

HandleTextTrigger = sTemp

End Function

Private Function FindTrigger(sInputStr As String, sTriggerChar As String) As String

Dim lPos1 As Long
Dim lPos2 As Long

lPos1 = InStr(sInputStr, sTriggerChar)
lPos2 = InStr(lPos1 + 1, sInputStr, sTriggerChar)

If (lPos1 <> 0) And lPos2 <> 0 Then
    FindTrigger = Mid$(sInputStr, lPos1, lPos2 - lPos1 + Len(sTriggerChar))
Else
    FindTrigger = ""
End If

End Function


Private Function RemoveTriggerChars(InputStr As String, CharsToRemove As String) As String
    RemoveTriggerChars = Replace$(InputStr, CharsToRemove, "")
End Function


