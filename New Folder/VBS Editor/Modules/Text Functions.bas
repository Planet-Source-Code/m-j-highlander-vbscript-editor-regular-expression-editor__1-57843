Attribute VB_Name = "Text_Functions"
Option Explicit

Private Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long

Public Const Quote = """"
Public Const ALPHANUMERIC_CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Public Const ALL_PRINTABLE_CHARS = ALPHANUMERIC_CHARS & Quote & " !#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
Public Const ALL_TEXT_CHARS = vbTab & vbCrLf & ALL_PRINTABLE_CHARS

Public Enum CharRangeConstants
    [AlphaNumeric Only] = 1  ' Starting from 1 , since VB defaults vars
    [All Printable] = 2      ' to zero, which might cause problems
    [All Text Chars] = 3
End Enum

Function UniqueWords(ByVal Text As String) As Collection
'Date: 4/14/2001
'Author: The VB2TheMax Team <fbalena@vb2themax.com>

' Build a list of all the individual words in a string
' returns a collection that contains all the unique words.
' The key for each item is the word itself
' so you can easily use the result collection to both
' enumerate the words and test whether a given word appears
' in the text. Words are inserted in the order they appear
' and are stored as lowercase strings.
'
' Numbers are ignored, but digit characters are preserved
' if they appear in the middle or at the end of a word.

Dim thisWord As String
Dim i As Long
Dim wordStart As Long
Dim varWord As Variant
Dim res As String

' prepare the result collection
Set UniqueWords = New Collection

' ignore duplicate words
On Error Resume Next

' extract all words from the text
For i = 1 To Len(Text)
    Select Case Asc(Mid$(Text, i, 1))
        Case 65 To 90, 97 To 122
            ' an alpha char
            If wordStart = 0 Then wordStart = i
        Case 48 To 57
            ' include digits only if suffix of a word (as in "ABCD23")
        Case Else
            If wordStart Then
                ' extract the word
                thisWord = LCase$(Mid$(Text, wordStart, i - wordStart))
                ' add to the collection, but ignore if already there
                UniqueWords.Add thisWord, thisWord
                ' reset the flag/pointer
                wordStart = 0
            End If
    End Select
Next

' account for the last word
If wordStart Then
    ' extract the word
    thisWord = LCase$(Mid$(Text, wordStart, i - wordStart))
    ' add to the collection, but ignore if already there
    UniqueWords.Add thisWord, thisWord
End If
    
End Function
Public Function WordWrap(ByRef Text As String, ByVal Width As Long, Optional ByRef CountLines As Long) As String
' by Donald, donald@xbeat.net, 20040913
  Dim i As Long
  Dim lenLine As Long
  Dim posBreak As Long
  Dim cntBreakChars As Long
  Dim abText() As Byte
  Dim abTextOut() As Byte
  Dim ubText As Long

  ' no fooling around
  If Width <= 0 Then
    CountLines = 0
    Exit Function
  End If
  If Len(Text) <= Width Then  ' no need to wrap
    CountLines = 1
    WordWrap = Text
    Exit Function
  End If
  
  abText = StrConv(Text, vbFromUnicode)
  ubText = UBound(abText)
  ReDim abTextOut(ubText * 3) 'dim to potential max
  
  For i = 0 To ubText
    Select Case abText(i)
    Case 32, 45 'space, hyphen
      posBreak = i
    Case Else
    End Select
    
    abTextOut(i + cntBreakChars) = abText(i)
    lenLine = lenLine + 1
    
    If lenLine > Width Then
      If posBreak > 0 Then
        ' don't break at the very end
        If posBreak = ubText Then Exit For
        ' wrap after space, hyphen
        abTextOut(posBreak + cntBreakChars + 1) = 13  'CR
        abTextOut(posBreak + cntBreakChars + 2) = 10  'LF
        i = posBreak
        posBreak = 0
      Else
        ' cut word
        abTextOut(i + cntBreakChars) = 13     'CR
        abTextOut(i + cntBreakChars + 1) = 10 'LF
        i = i - 1
      End If
      cntBreakChars = cntBreakChars + 2
      lenLine = 0
    End If
  Next
  
  CountLines = cntBreakChars \ 2 + 1
  
  ReDim Preserve abTextOut(ubText + cntBreakChars)
  WordWrap = StrConv(abTextOut, vbUnicode)
  
End Function
Public Function Chop(ByVal Text As String) As String

Text = Replace(Text, Chr(0), "")
Text = Replace(Text, vbCr, "")
Text = Replace(Text, vbLf, "")
Chop = Trim(Text)

End Function
Public Function SplitEx(ByVal Expression As String, _
                        Optional ByVal Delimiter As String = " ", _
                        Optional ByVal Limit As Long = -1, _
                        Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant

Dim varItems As Variant, i As Long

varItems = Split(Expression, Delimiter, Limit, Compare)

For i = LBound(varItems) To UBound(varItems)

    If Len(varItems(i)) = 0 Then varItems(i) = Delimiter

Next i

SplitEx = Filter(varItems, Delimiter, False)
    
End Function
Public Function CountWords(ByVal Text As String) As Long
'Original Function Source:   VB-World

'Assume a hyphen at the end of a line is part of a full-word, so combine together
Text = Trim(Replace(Text, "-" & vbNewLine, ""))

'Replace new lines with a single space
Text = Trim(Replace(Text, vbNewLine, " "))

'Replace Tab with Space     '//Me'
Text = Trim(Replace(Text, vbTab, " "))

'Collapse multiple spaces into one single space
Do While Text Like "*  *"
    Text = Replace(Text, "  ", " ")
Loop

'Split the string and return counted words
CountWords = 1 + UBound(Split(Text, " "))
    
End Function
Public Function LeftTo(ByVal MainStr As String, ByVal SubStr As String)

' Returns a sub-string from the leftmost of "MainStr" until "SubStr" is found.
' "SubStr" is not included in the returned string.

Dim iPos As Long

iPos = InStr(MainStr, SubStr)
If iPos <> 0 Then
    LeftTo = Left$(MainStr, iPos - 1)
Else
    LeftTo = MainStr
End If

End Function
Public Function RightTo(ByVal MainStr As String, ByVal SubStr As String)

' Returns a sub-string from the rightmost of "MainStr" until "SubStr" is found.
' "SubStr" is not included in the returned string.


Dim iPos As Long

iPos = InStrRev(MainStr, SubStr)
If iPos <> 0 Then
    RightTo = Right$(MainStr, Len(MainStr) - iPos - Len(SubStr) + 1)
Else
    RightTo = MainStr
End If

End Function
Public Function GuessTitle(ByVal Text As String) As String
Dim sTemp As String
Dim ch As String
Dim idx As Integer

sTemp = Left(Text, 1000) ' 1000 chars max
'''Left Trim (Convert Cr,Lf and Tab to Spaces)
For idx = 1 To Len(sTemp)
    ch = CharAt(sTemp, idx)
    If (ch = vbCr Or ch = vbLf Or ch = vbTab Or ch = " ") Then
        ch = " "
        CharAt(sTemp, idx) = ch
    Else
        Exit For 'break at first non Cr/Lf/Tab char
    End If
Next
sTemp = LTrim(sTemp)
sTemp = LeftTo(sTemp, vbCrLf)

GuessTitle = CrLfTabTrim(sTemp)

End Function
Public Function EscapeChars_ForFindReplace(ByVal Text As String) As String
Const ESCAPE_CHAR = "~"

' Chr(7) is a value that surly won't (shouldn't) exist in plain text

Text = Replace(Text, ESCAPE_CHAR & ESCAPE_CHAR, Chr$(7)) 'hide it

Text = Replace(Text, ESCAPE_CHAR & "n", vbCrLf)
Text = Replace(Text, ESCAPE_CHAR & "t", vbTab)
'Text = Replace(Text, ESCAPE_CHAR , "")  '???????????

Text = Replace(Text, Chr$(7), ESCAPE_CHAR)  'unhide

EscapeChars_ForFindReplace = Text

End Function

Public Function FirstCharUpper(ByVal Text As String) As String
Dim sArray() As String
Dim idx As Long

Text2Array Text, sArray
For idx = LBound(sArray) To UBound(sArray)
    If sArray(idx) <> "" Then
        Mid$(sArray(idx), 1, 1) = UCase$(Mid$(sArray(idx), 1, 1))
    End If
Next

FirstCharUpper = Join(sArray, vbCrLf)

End Function
Public Sub SplitAt(ByVal Text As String, ByVal Delimiter As String, ByRef LeftPart As String, ByRef RightPart As String)
Dim pos As Long

pos = InStr(1, Text, Delimiter, vbTextCompare)

If pos = 0 Then
    LeftPart = Text
    RightPart = ""
Else
    LeftPart = Left(Text, pos - 1)
    RightPart = Right(Text, Len(Text) - pos)
End If

End Sub
Public Function StrFill(ByVal Count As Long, ByVal StrToFill As String) As String
Dim objStr As CStrCat
Dim idx As Long

If StrToFill = "" Or Count = 0 Then
    StrFill = ""
    Exit Function
End If


Set objStr = New CStrCat
objStr.MaxLength = Count * Len(StrToFill)
For idx = 1 To Count
    objStr.AddStr StrToFill
Next

StrFill = objStr.StrVal

End Function
Public Function RemoveNonAlphaNum4(ByVal Words As String, ByVal CharsToKeep As CharRangeConstants) As String
Dim ItIsValid  As Boolean
Dim Alpha As String
Dim iPos As Long
Dim sChar As String, sWork As String
Dim cntr As Long, idx As Long
Dim Chars() As Byte
Dim b() As Byte
Dim C() As Byte

Select Case CharsToKeep
    Case [AlphaNumeric Only]
        Alpha = ALPHANUMERIC_CHARS
    Case [All Printable]
        Alpha = ALL_PRINTABLE_CHARS
    Case [All Text Chars]
        Alpha = ALL_TEXT_CHARS
End Select


ReDim b(0 To Len(Words) - 1)
ReDim C(0 To Len(Words) - 1)
b = StrConv(Words, vbFromUnicode) ' VB Strings are Double-Byte Unicode
idx = 0
For cntr = 0 To Len(Words) - 1
    ItIsValid = False
    If InStr(Alpha, Chr$(b(cntr))) Then
        C(idx) = b(cntr)
        'If idx > 0 Then If c(idx) = 10 And c(idx - 1) <> 13 Then c(idx) = 32
        idx = idx + 1
    End If
Next cntr

RemoveNonAlphaNum4 = FixNewLineChars(Left$(StrConv(C, vbUnicode), idx))


End Function
Public Function RemoveNonAlphaNum_Byte(ByRef byteChars() As Byte, ByVal CharsToKeep As CharRangeConstants) As String
Dim ItIsValid  As Boolean
Dim Alpha As String
Dim iPos As Long
Dim sChar As String, sWork As String
Dim cntr As Long, idx As Long
Dim Chars() As Byte
'Dim b() As Byte
Dim C() As Byte

Select Case CharsToKeep
    Case [AlphaNumeric Only]
        Alpha = ALPHANUMERIC_CHARS
    Case [All Printable]
        Alpha = ALL_PRINTABLE_CHARS
    Case [All Text Chars]
        Alpha = ALL_TEXT_CHARS
End Select


'ReDim b(0 To Len(Words) - 1)
ReDim C(LBound(byteChars) To UBound(byteChars))
'b = StrConv(Words, vbFromUnicode) ' VB Strings are Double-Byte Unicode
idx = 0
For cntr = LBound(byteChars) To UBound(byteChars)
    ItIsValid = False
    If InStr(Alpha, Chr$(byteChars(cntr))) Then
        C(idx) = byteChars(cntr)
        'If idx > 0 Then If c(idx) = 10 And c(idx - 1) <> 13 Then c(idx) = 32
        idx = idx + 1
    End If
Next cntr

RemoveNonAlphaNum_Byte = StrConv(C, vbUnicode)
RemoveNonAlphaNum_Byte = FixNewLineChars(Left$(RemoveNonAlphaNum_Byte, idx))


End Function
Public Function HandleTextTrigger(ByVal Text As String, ByVal TriggerChars As String) As String

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
Public Property Get CharAt(ByRef Text As String, ByVal Position As Long) As String

    If Position > 0 Then
        CharAt = Mid$(Text, Position, 1)
    End If

End Property
Public Property Let CharAt(ByRef Text As String, ByVal Position As Long, ByVal sNewValue As String)
    
    If Position > 0 Then
        Mid$(Text, Position, 1) = sNewValue
    End If
    
End Property
Public Function CrLfTabTrim(ByVal Text As String) As String
'Trim leading and trailing Cr , Lf , Tab and Space Chars

Dim sTemp As String
Dim ch As String * 1
Dim idx As Long

sTemp = Trim$(Text)

'''Left Trim (Convert Cr,Lf and Tab to Spaces)
For idx = 1 To Len(sTemp)
    ch = CharAt(sTemp, idx)
    If (ch = vbCr Or ch = vbLf Or ch = vbTab Or ch = " ") Then
        ch = " "
        CharAt(sTemp, idx) = ch
    Else
        Exit For 'break at first non Cr/Lf/Tab/Spc char
    End If
Next idx

'''Right Trim (Convert Cr,Lf and Tab to Spaces)
For idx = Len(sTemp) To 1 Step -1
    ch = CharAt(sTemp, idx)
    If (ch = vbCr Or ch = vbLf Or ch = vbTab Or ch = " ") Then
        ch = " "
        CharAt(sTemp, idx) = ch
    Else
        Exit For 'break at first non-Cr/Lf/Tab/Spc char
    End If
Next idx

'''Trim Spaces
CrLfTabTrim = Trim$(sTemp)

End Function
Public Function DoEllipses(ByVal Text As String, Optional MaxLength As Long = 0) As String

Dim sTemp As String
Dim pos As Long

'Remove Leading and Trailing Cr,Lf,Tab and Space
sTemp = CrLfTabTrim(Text)

'Ignore text after the first Cr char
pos = InStr(1, sTemp, vbCr)
If pos > 0 Then
    sTemp = Left$(sTemp, pos - 1)
End If

'Take only the specified length (if specified!)
If MaxLength > 0 Then
    sTemp = Left$(sTemp, MaxLength)
End If

'Did we truncate the text? if yes add Ellipses (...)
If Len(sTemp) < Len(Text) Then
    sTemp = sTemp & " ..."
End If

DoEllipses = sTemp

End Function
Public Function GetTextFileContents(ByVal FileName As String) As String
Dim objTextFile As New CTextFile


If objTextFile.FileOpen(FileName, OpenForInput) Then
            GetTextFileContents = objTextFile.ReadAll
            objTextFile.FileClose
Else
            GetTextFileContents = ""
End If

Set objTextFile = Nothing

End Function
Public Function RemoveNonAlphaNum3(ByVal Words As String, ByVal CharsToKeep As CharRangeConstants) As String
' Remove all non-alphanumeric characters from the Words
' The fastest method.

Dim Alpha As String
Dim iPos As Long
Dim sChar As String, sWork As String
Dim cntr As Long, idx As Long
Dim ch As String * 1

Select Case CharsToKeep
    Case [AlphaNumeric Only]
        Alpha = ALPHANUMERIC_CHARS
    Case [All Printable]
        Alpha = ALL_PRINTABLE_CHARS
    Case [All Text Chars]
        Alpha = ALL_TEXT_CHARS
End Select

sWork = String(Len(Words), " ")

For cntr = 1 To Len(Words)
    ch = Mid$(Words, cntr, 1)
    If InStr(Alpha, ch) <> 0 Then
        idx = idx + 1
        Mid$(sWork, idx, 1) = ch
    End If
Next cntr

sWork = Left$(sWork, idx)

RemoveNonAlphaNum3 = FixNewLineChars(sWork)

End Function

Public Function RemoveNonAlphaNum2(ByVal Words As String) As String
' Remove all non-alphanumeric characters from the Words

Const Alpha = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz "
Const SNG_SPACE = " "
Const DBL_SPACE = "  "

Dim iPos As Long
Dim sChar As String, sWork As String
Dim lCntr As Long

'Set the working variable
sWork = Trim(Words)

'Remove all charaters that are NOT in the ALPHA const
For lCntr = 0 To 255
    If InStr(Alpha, Chr(lCntr)) = 0 Then
        sWork = Replace(sWork, Chr(lCntr), SNG_SPACE)
    End If
Next lCntr

'Remove all double spaces created
iPos = InStr(sWork, DBL_SPACE)
While iPos > 0
     sWork = Replace(sWork, DBL_SPACE, SNG_SPACE)
    iPos = InStr(sWork, DBL_SPACE)
Wend


RemoveNonAlphaNum2 = sWork
    
End Function

Public Function ReverseStr(ByVal Text As String, Optional ByLine As Boolean = True) As String
ReDim sArray(1 To 1) As String
Dim sTemp As String
Dim idx As Long
Dim cidx As Long

If ByLine = False Then  ' Reverse Entire Text
    sTemp = String$(Len(Text), " ")
    For idx = 0 To Len(Text) - 1
        ' Kinda Cool!
        CharAt(sTemp, idx + 1) = CharAt(Text, Len(Text) - idx)
    Next idx
    ' CrLf will also be reversed, so fix it
    sTemp = Replace$(sTemp, vbLf & vbCr, vbCrLf)

Else                    ' Reverse Each Line Alone
    Text2Array Text, sArray
    For idx = LBound(sArray) To UBound(sArray)
            sTemp = String$(Len(sArray(idx)), " ")
            For cidx = 0 To Len(sTemp) - 1
                ' Kinda Cool!
                CharAt(sTemp, cidx + 1) = CharAt(sArray(idx), Len(sTemp) - cidx)
            Next cidx
            sArray(idx) = sTemp
    Next idx
    sTemp = Join$(sArray, vbCrLf)
End If

ReverseStr = sTemp

End Function

Public Function IsChrAlphaNumeric(ByVal Char As String) As Boolean

Dim cChar As Byte
If Char <> "" Then
    cChar = Asc(Char)
    IsChrAlphaNumeric = CBool(IsCharAlphaNumeric(cChar))
Else
    IsChrAlphaNumeric = False
End If

End Function
Function MultiInstrRev(ByVal Start As Long, ByVal Text As String, ByVal LookFor As String, ByVal Compare As VbCompareMethod) As Long

Dim iLen As Long
Dim chLookFor As String * 1
Dim idx As Long
Dim iPos As Long
Dim iFirstPos As Long


iLen = Len(LookFor)
iFirstPos = 0

For idx = 1 To iLen
    chLookFor = Mid(LookFor, idx, 1)
    iPos = InStrRev(Text, chLookFor, Start, Compare)
    If (iPos <> 0 And iPos > iFirstPos) Then
         iFirstPos = iPos
'         MsgBox iFirstPos
    End If

Next idx


'If iFirstPos = Len(Text) + 1 Then ' value didn't change / nothing found
'    MultiInstrRev = 0
'    MsgBox "X"
'Else
    MultiInstrRev = iFirstPos
'End If

End Function
Public Function MultiInstr(ByVal Start As Long, ByVal Text As String, ByVal LookFor As String, ByVal Compare As VbCompareMethod) As Long

Dim iLen As Long
Dim chLookFor As String * 1
Dim idx As Long
Dim iPos As Long
Dim iFirstPos As Long


iLen = Len(LookFor)
iFirstPos = Len(Text) + 1 ' out of text boundry // an impossible value

For idx = 1 To iLen
    chLookFor = Mid(LookFor, idx, 1)
    iPos = InStr(Start, Text, chLookFor, Compare)
    If (iPos <> 0 And iPos < iFirstPos) Then
         iFirstPos = iPos
    End If
    
Next idx

If iFirstPos = Len(Text) + 1 Then ' value didn't change / nothing found
    MultiInstr = 0
Else
    MultiInstr = iFirstPos
End If

End Function
Function InsertString(ByVal MainStr As String, ByVal SubStr As String, ByVal Position As Long)
Dim sLeft As String, sRight As String

Select Case Position
    Case Is > Len(MainStr)
        sLeft = MainStr
        sRight = ""
    Case Is <= 0
        sLeft = ""
        sRight = MainStr
    Case Is <= Len(MainStr)
        sLeft = Left$(MainStr, Position)
        sRight = Right$(MainStr, Len(MainStr) - Position)
End Select
    
InsertString = sLeft & SubStr & sRight

End Function

Public Function DelLeftTo(ByVal MainStr As String, ByVal SubStr As String, ByVal MatchCase As Boolean, ByVal Inclusive As Boolean) As String

Dim pos As Long
If MatchCase = True Then
        pos = InStr(1, MainStr, SubStr, vbBinaryCompare)
Else
        pos = InStr(1, MainStr, SubStr, vbTextCompare)
End If


If (pos = 0) Then
            DelLeftTo = MainStr
Else
            If Inclusive Then
                        DelLeftTo = Right(MainStr, Len(MainStr) - pos - Len(SubStr) + 1)
            Else
                        DelLeftTo = Right(MainStr, Len(MainStr) - pos + 1)
            End If
End If

End Function
Public Function DelRightTo(ByVal MainStr As String, ByVal SubStr As String, ByVal MatchCase As Boolean, ByVal Inclusive As Boolean) As String

Dim pos As Long
If MatchCase = True Then
        pos = InStrRev(MainStr, SubStr, -1, vbBinaryCompare)
Else
        pos = InStrRev(MainStr, SubStr, -1, vbTextCompare)
End If


If (pos = 0) Then
            DelRightTo = MainStr
Else
            If Inclusive Then
                        DelRightTo = Left(MainStr, pos - 1)
            Else
                        DelRightTo = Left(MainStr, pos + Len(SubStr) - 1)
            End If
End If

End Function
Public Function Array2Text(sArray() As String) As String

Array2Text = Join$(sArray, vbCrLf)

End Function
Public Function DelLeft(ByVal TextLine As String, ByVal Count As Long) As String

If Count <= 0 Then
    DelLeft = TextLine
ElseIf Count > Len(TextLine) Then
    DelLeft = ""
Else
    DelLeft = Right$(TextLine, Len(TextLine) - Count)
End If

End Function
Function DelRight(ByVal TextLine As String, ByVal Count As Long) As String

If Count <= 0 Then
    DelRight = TextLine
ElseIf Count > Len(TextLine) Then
    DelRight = ""
Else
    DelRight = Left$(TextLine, Len(TextLine) - Count)
End If

End Function

Function FixNewLineChars(ByVal Text As String) As String

Dim NonTextChar As String * 1

NonTextChar = Chr$(7) 'any char that does not show up in Text files

Text = Replace$(Text, vbCrLf, NonTextChar)
Text = Replace$(Text, vbLf & vbCr, NonTextChar)
Text = Replace$(Text, vbLf, NonTextChar)
Text = Replace$(Text, vbCr, NonTextChar)

FixNewLineChars = Replace$(Text, NonTextChar, vbCrLf)

End Function
Function InsertString2(ByVal TextLine As String, ByVal StrToInsert As String, ByVal Position As Integer) As String
Dim sLeft As String, sRight As String

If Position > Len(TextLine) Then
    Position = Len(TextLine)
End If

If TextLine = "" Then
    InsertString2 = ""
Else
    
    sLeft = Left(TextLine, Position)
    sRight = Right(TextLine, Len(TextLine) - Position)
    
    InsertString2 = sLeft & StrToInsert & sRight

End If

End Function

Public Function CompactSpaces(ByVal Text As String) As String
' Convert successive spaces into one space,
' also trims leading and trailing spaces. / Really Fast!

Const SNG_SPACE = " "
Const DBL_SPACE = "  "

Dim pos As Long
Dim sWork As String
Dim idx As Long

sWork = Trim$(Text)

'Keep Removing double spaces until none found:
pos = InStr(sWork, DBL_SPACE)
Do While pos > 0
     sWork = Replace$(sWork, DBL_SPACE, SNG_SPACE)
     pos = InStr(sWork, DBL_SPACE)
Loop

'Trim right and left of each line, don't trim tabs:
sWork = LinesTrim(sWork, True, True, False)

CompactSpaces = sWork

End Function
Public Function EscapeChars(ByVal Text As String) As String

' Chr(7) is a value that surly won't (shouldn't) exist in plain text
Text = Replace(Text, "\\", Chr$(7)) 'hide it
Text = Replace(Text, "\q", """")
Text = Replace(Text, "\r\n", vbCrLf)
Text = Replace(Text, "\t", vbTab)
'Text = Replace(Text, "\", "")  ' use \\ to indicate \
Text = Replace(Text, Chr$(7), "\")  'unhide

EscapeChars = Text

End Function
Public Function Tab2Spaces(ByVal Text As String, ByVal NumSpaces As Integer) As String

Dim sTemp As String

If NumSpaces < 1 Then NumSpaces = 1
sTemp = Space$(NumSpaces)

Tab2Spaces = Replace(Text, vbTab, sTemp)

End Function
Public Function RemoveLineBreaks(ByVal Text As String) As String

Dim sTemp As String

sTemp = Replace(Text, vbCrLf, " ")

RemoveLineBreaks = sTemp

End Function
Public Function Text2Array(ByVal Text As String, ByRef LinesArray() As String) As Boolean
' Returns True if successful, False otherwisee

Dim vTmpArray As Variant
Dim idx As Long
    
vTmpArray = Split(Text, vbCrLf)

If UBound(vTmpArray) > -1 Then
    ReDim LinesArray(LBound(vTmpArray) To UBound(vTmpArray))
    
    For idx = LBound(vTmpArray) To UBound(vTmpArray)
        LinesArray(idx) = vTmpArray(idx)
    Next idx
    Text2Array = True
Else
    'not an array
    Text2Array = False
End If

End Function
Public Function UnQuote(ByVal Text As String) As String
Dim sTemp As String

sTemp = Text

If Left$(Text, 1) = Quote And Right$(Text, 1) = Quote Then
'    sTemp = Right$(sTemp, Len(sTemp) - 1)
'    sTemp = Left$(sTemp, Len(sTemp) - 1)
    sTemp = Mid$(sTemp, 2, Len(sTemp) - 2)
End If


UnQuote = sTemp

End Function
