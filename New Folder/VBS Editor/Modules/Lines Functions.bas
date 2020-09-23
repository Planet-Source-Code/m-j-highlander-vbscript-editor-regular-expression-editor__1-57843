Attribute VB_Name = "Lines_Functions"
Option Explicit

Public Function LinesExterminateDuplicates(ByVal Text As String) As String
Dim idx As Integer
Dim sArray() As String
Dim oDic As Dictionary
Dim oDic2 As Dictionary

Text2Array Text, sArray

Set oDic = New Dictionary
Set oDic2 = New Dictionary

oDic.CompareMode = TextCompare
oDic2.CompareMode = TextCompare

For idx = LBound(sArray) To UBound(sArray)
    If Not (oDic.Exists(sArray(idx))) Then
        oDic.Add sArray(idx), sArray(idx)
    Else
        oDic.Item(sArray(idx)) = "~KILL_ME~"
    End If
Next

For idx = LBound(oDic.Items) To UBound(oDic.Items)
    If oDic.Items(idx) <> "~KILL_ME~" Then
        oDic2.Add oDic.Items(idx), oDic.Items(idx)
    End If
Next

LinesExterminateDuplicates = Join(oDic2.Items, vbCrLf)

Set oDic = Nothing
Set oDic2 = Nothing

End Function
Public Function LinesRemoveDuplicates(ByVal Text As String) As String
Dim idx As Integer
Dim sArray() As String
Dim oDic As Dictionary


Text2Array Text, sArray

Set oDic = New Dictionary
oDic.CompareMode = TextCompare

For idx = LBound(sArray) To UBound(sArray)
    If oDic.Exists(sArray(idx)) = False Then
        oDic.Add sArray(idx), sArray(idx)
    End If
Next

LinesRemoveDuplicates = Join(oDic.Items, vbCrLf)

Set oDic = Nothing

End Function
Public Function LinesReverseOrder(ByVal Text As String) As String
Dim sTemp As String, idx As Long, lShift As Long
Dim sArray() As String, sRevArray() As String

Text2Array Text, sArray  'the array will be dimensioned here
ReDim sRevArray(LBound(sArray) To UBound(sArray))

lShift = UBound(sArray)
For idx = LBound(sArray) To UBound(sArray)
        sRevArray(idx) = sArray(lShift - idx)
Next

LinesReverseOrder = Join(sRevArray, vbCrLf)

End Function
Public Function LinesDelLeft(ByVal Text As String, ByVal Count As Long) As String
Dim TempArray() As String
Dim idx As Long

Text2Array Text, TempArray

For idx = LBound(TempArray) To UBound(TempArray)
    TempArray(idx) = DelLeft(TempArray(idx), Count)
Next idx

LinesDelLeft = Array2Text(TempArray)

End Function
Public Function linesDelLeftTo(ByVal Text As String, ByVal DelToWhat As String, ByVal MatchCase As Boolean, ByVal Inclusive As Boolean) As String
Dim TempArray() As String
Dim idx As Long
    
Text2Array Text, TempArray
For idx = LBound(TempArray) To UBound(TempArray)
    TempArray(idx) = DelLeftTo(TempArray(idx), DelToWhat, MatchCase, Inclusive)
Next idx

linesDelLeftTo = Array2Text(TempArray)

End Function
Public Function LinesDelRight(ByVal Text As String, ByVal Count As Long) As String
Dim TempArray() As String
Dim idx As Long

Text2Array Text, TempArray

For idx = LBound(TempArray) To UBound(TempArray)
    TempArray(idx) = DelRight(TempArray(idx), Count)
Next idx

LinesDelRight = Array2Text(TempArray)

End Function
Public Function LinesDelRightTo(ByVal Text As String, ByVal DelToWhat As String, ByVal MatchCase As Boolean, ByVal Inclusive As Boolean) As String
Dim TempArray() As String
Dim idx As Long
    
Text2Array Text, TempArray
For idx = LBound(TempArray) To UBound(TempArray)
    TempArray(idx) = DelRightTo(TempArray(idx), DelToWhat, MatchCase, Inclusive)
Next idx

LinesDelRightTo = Array2Text(TempArray)

End Function
Public Function LinesTrim(ByVal Text As String, ByVal TrimLeft As Boolean, ByVal TrimRight As Boolean, ByVal TrimTab As Boolean) As String

Dim idx As Long
ReDim sTempArray(1 To 1) As String

Text2Array Text, sTempArray

For idx = LBound(sTempArray) To UBound(sTempArray)
    If (TrimLeft And TrimRight) Then
        sTempArray(idx) = Trim(sTempArray(idx))
        If TrimTab Then sTempArray(idx) = TrimTabs(sTempArray(idx), True, True)
     ElseIf TrimLeft Then
        sTempArray(idx) = LTrim(sTempArray(idx))
        If TrimTab Then sTempArray(idx) = TrimTabs(sTempArray(idx), True, False)
    ElseIf TrimRight Then
        sTempArray(idx) = RTrim(sTempArray(idx))
       If TrimTab Then sTempArray(idx) = TrimTabs(sTempArray(idx), False, True)
    Else
        'Do Nothing
    End If
Next idx

LinesTrim = Join(sTempArray, vbCrLf)

End Function
Private Function TrimTabs(ByVal TextLine As String, ByVal TrimLeft As Boolean, ByVal TrimRight As Boolean) As String
Dim idx As Long
Dim ch As String * 1

'kinda REDUNDANCY !!!!
If (TrimLeft And TrimRight) Then
    'Trim both
    For idx = 1 To Len(TextLine)
        ch = Mid$(TextLine, idx, 1)
        If ch = Chr$(9) Then 'TAB
               Mid$(TextLine, idx) = Chr$(7)
        Else
                Exit For
        End If
    Next idx
    For idx = Len(TextLine) To 1 Step -1
        ch = Mid$(TextLine, idx, 1)
        If ch = Chr$(9) Then 'TAB
               Mid$(TextLine, idx) = Chr$(7)
        Else
                Exit For
        End If
    Next idx

ElseIf TrimLeft Then
    For idx = 1 To Len(TextLine)
        ch = Mid$(TextLine, idx, 1)
        If ch = Chr$(9) Then 'TAB
               Mid$(TextLine, idx) = Chr$(7)
        Else
                Exit For
        End If
    Next idx

ElseIf TrimRight Then
    For idx = Len(TextLine) To 1 Step -1
        ch = Mid(TextLine, idx, 1)
        If ch = Chr$(9) Then 'TAB
               Mid$(TextLine, idx) = Chr$(7)
        Else
                Exit For
        End If
    Next idx

End If

TrimTabs = Replace$(TextLine, Chr$(7), "")

End Function
Private Function SetLineMaxWidthWrap(ByVal TextLine As String, ByVal MaxWidth As Long) As String
Dim HowManyTimes As Long, idx As Long, PosShift As Long
Dim sTemp As String, sReturn As String

HowManyTimes = Len(TextLine) \ MaxWidth
If Len(TextLine) Mod MaxWidth <> 0 Then 'there is a remainder
    HowManyTimes = HowManyTimes + 1
End If

PosShift = MaxWidth

If MaxWidth < Len(TextLine) Then
    sTemp = TextLine
    For idx = 1 To HowManyTimes
        
        Do While PosShift > 0
            If IsChrAlphaNumeric(Mid$(sTemp, PosShift, 1)) Then
                PosShift = PosShift - 1
            Else
                Exit Do
            End If
        Loop
        
        sTemp = InsertString(sTemp, vbCrLf, PosShift)
        PosShift = PosShift + MaxWidth + 2
    Next idx
    SetLineMaxWidthWrap = sTemp
Else
    SetLineMaxWidthWrap = TextLine
End If

End Function
Private Function SetLineMaxWidth(ByVal TextLine As String, ByVal MaxWidth As Integer) As String
Dim HowManyTimes As Integer, idx As Integer, PosShift As Long
Dim sReturn As String, sTemp As String

HowManyTimes = Len(TextLine) \ MaxWidth
PosShift = MaxWidth

If MaxWidth < Len(TextLine) Then
    sTemp = TextLine
    For idx = 1 To HowManyTimes
        sTemp = InsertString(sTemp, vbCrLf, PosShift)
'        If sReturn <> "" Then sTemp = sReturn
        PosShift = PosShift + MaxWidth + 2
    Next idx
    SetLineMaxWidth = sTemp
Else
    SetLineMaxWidth = TextLine
End If

End Function
Public Function LinesSetMaxWidth(ByVal Text As String, ByVal MaxLength As Integer, ByVal WordWrap As Boolean) As String
Dim idx As Long
ReDim sTempArray(1 To 1) As String

Text2Array Text, sTempArray

If WordWrap Then
    For idx = LBound(sTempArray) To UBound(sTempArray)
            sTempArray(idx) = SetLineMaxWidthWrap(sTempArray(idx), MaxLength)
    Next idx

Else
    
    For idx = LBound(sTempArray) To UBound(sTempArray)
            sTempArray(idx) = SetLineMaxWidth(sTempArray(idx), MaxLength)
    Next idx
End If


LinesSetMaxWidth = Join(sTempArray, vbCrLf)

End Function
Public Function LinesSort(ByVal Text As String, ByVal Ascending As Boolean, ByVal IgnoreCase As Boolean) As String

ReDim sArray(1 To 1) As String

Text2Array Text, sArray

StrSort sArray, Ascending, IgnoreCase

LinesSort = Array2Text(sArray)

End Function
Private Function UNUSED_LinesRemoveBlank(ByVal Text As String) As String
'Dim idx As Long, sTemp As String
'ReDim sArray(1 To 1) As String
'
'Text2Array Text, sArray
'
'For idx = LBound(sArray) To UBound(sArray)
'    If sArray(idx) = "" Then
'            sArray(idx) = Chr$(7)
'    End If
'Next idx
'
'sTemp = Join$(sArray, vbCrLf)
'sTemp = Replace$(sTemp, Chr$(7) & vbCrLf, "")
'
'LinesRemoveBlank = Replace$(sTemp, Chr$(7), "") 'Trim trailing if exists.

End Function
Private Function UNUSED_LinesCompactBlank(ByVal Text As String) As String
'' Convert successive blank lines into one line,
'' also trims leading and trailing CR,LF. ?
'
'Const DOUBLE_CRLF = vbCrLf & vbCrLf
'Const TRIPPLE_CRLF = vbCrLf & vbCrLf & vbCrLf
'
'Dim pos As Long
'Dim sWork As String
'Dim idx As Long
'
'sWork = Trim$(Text)
'
''Keep Removing double CRLF's until none found:
'pos = InStr(sWork, TRIPPLE_CRLF)
'Do While pos > 0
'     sWork = Replace$(sWork, TRIPPLE_CRLF, DOUBLE_CRLF)
'     pos = InStr(sWork, TRIPPLE_CRLF)
'Loop
'
''Trim right and left of Text
'sWork = CrLfTabTrim(sWork) 'could be slow?
'
'LinesCompactBlank = sWork & vbCrLf  'it surely has none

End Function
Public Function LinesAddNumbers(ByVal Text As String, ByVal NumStart As Long, ByVal NumStep As Long, ByVal Delimiter As String, ByVal NumDigits As Long, ByVal IgnoreEmptyLines As Boolean) As String
Dim FormatStr As String
Dim idx As Long
Dim cntr As Long
ReDim sTempArray(1 To 1) As String

FormatStr = String(NumDigits, "0")

Text2Array Text, sTempArray

cntr = NumStart
For idx = LBound(sTempArray) To UBound(sTempArray)
    If IgnoreEmptyLines And sTempArray(idx) = "" Then
        'do nothing
    Else
        sTempArray(idx) = Format(cntr, FormatStr) & Delimiter & sTempArray(idx)
        cntr = cntr + NumStep
    End If
    
Next idx

LinesAddNumbers = Join(sTempArray, vbCrLf)

End Function
Public Function LinesInsert(ByVal Text As String, ByVal InsertWhat As String, ByVal InsertPos As Long, ByVal IgnoreEmptyLines As Boolean) As String
Dim idx As Long
ReDim TempArray(1 To 1) As String

'InsertWhat = stringf(InsertWhat)     ' handle special chars

Text2Array Text, TempArray
For idx = LBound(TempArray) To UBound(TempArray)
    If TempArray(idx) = "" And IgnoreEmptyLines Then
        'do nothing
    Else
        'Insert...
        TempArray(idx) = InsertString(TempArray(idx), InsertWhat, InsertPos)
    End If
Next idx

LinesInsert = Array2Text(TempArray)

End Function
Function LinesAddLeftRight(ByVal Text As String, ByVal AddToLeft As String, ByVal AddToRight As String, ByVal IgnoreEmptyLines As Boolean) As String
Dim idx As Long
ReDim TheArray(1 To 1) As String

Text2Array Text, TheArray

For idx = LBound(TheArray) To UBound(TheArray)
    If IgnoreEmptyLines And TheArray(idx) = "" Then
        'do nothing
    Else
         TheArray(idx) = AddToLeft & TheArray(idx) & AddToRight
    End If
Next idx

LinesAddLeftRight = Array2Text(TheArray)

End Function
