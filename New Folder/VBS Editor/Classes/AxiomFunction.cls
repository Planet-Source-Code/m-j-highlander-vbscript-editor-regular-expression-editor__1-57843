VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAxiomFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Axiom Macro Functions

Option Explicit

Public Function RegExpEscape(ByVal PlainText As String) As String
'Input:  plain text, NO RegExp special chars allowed
'Output: RegExp compatible text
' Chars that need Escaping:  \ ^ $ * + { } ? . : = ! | [ ] - ( ) ,

RegExpEscape = RegExp_Functions.EscapeRegExpChars(PlainText)

End Function
Public Function KeepHTMLBody(ByVal Html As String, ByVal OptRemove As HTML_REMOVAL_OPTIONS) As String

KeepHTMLBody = RegExp_Functions.RX_KeepHTMLBody(Html, OptRemove)

End Function

Public Function RegExpExtractSubMatch(ByVal Text As String, ByVal Pattern As String, Optional ByVal SubMatchIndex As Integer = 0) As String

RegExpExtractSubMatch = RX_GenericExtractSubMatch(Text, Pattern, SubMatchIndex)

End Function
Public Function ANSIToASCII7Remove(ByVal Text As String) As String

ANSIToASCII7Remove = StrToANSI127_Remove(Text)

End Function
Public Function LinesInsert(ByVal Text As String, ByVal InsertWhat As String, ByVal InsertPos As Long, Optional ByVal IgnoreEmptyLines As Boolean = True) As String

LinesInsert = Lines_Functions.LinesInsert(Text, InsertWhat, InsertPos, IgnoreEmptyLines)

End Function
Public Function LinesSetMaxWidth(ByVal Text As String, ByVal MaxLength As Integer, Optional ByVal WordWrap As Boolean = True) As String

LinesSetMaxWidth = Lines_Functions.LinesSetMaxWidth(Text, MaxLength, WordWrap)

End Function
Public Function LinesSort(ByVal Text As String, Optional ByVal Ascending As Boolean = True, Optional ByVal IgnoreCase As Boolean = True) As String

LinesSort = Lines_Functions.LinesSort(Text, Ascending, IgnoreCase)

End Function
Public Function LinesTrim(ByVal Text As String, Optional ByVal TrimLeft As Boolean = True, Optional ByVal TrimRight As Boolean = True, Optional ByVal TrimTab As Boolean = True) As String

LinesTrim = Lines_Functions.LinesTrim(Text, TrimLeft, TrimRight, TrimTab)

End Function
Public Function ReplaceTagKeepContent(ByVal Html As String, ByVal Tag As String, ByVal OpenTag As String, ByVal CloseTag As String) As String

ReplaceTagKeepContent = RX_ReplaceTagKeepContent(Html, Tag, OpenTag, CloseTag)

End Function
Public Function ReplaceTagAndContents(ByVal Html As String, ByVal Tag As String, ByVal TagIsSingle As String, ByVal ReplaceWith As String) As String

ReplaceTagAndContents = RX_ReplaceTagAndContents(Html, Tag, ReplaceWith, CBool(TagIsSingle))

End Function
Public Function LinesExterminateDuplicates(ByVal Text As String) As String

    LinesExterminateDuplicates = Lines_Functions.LinesExterminateDuplicates(Text)

End Function
Public Function HTMLizeTemplate( _
                        ByVal Text As String, _
                        ByVal TemplateFileName As String, _
                        Optional ByVal PageTitle As String = "", _
                        Optional ByVal PicturePath As String = "", _
                        Optional ByVal PreserveSpaces As Boolean = False, _
                        Optional ByVal KeepHTTP As Boolean = True, _
                        Optional ByVal Target As String = "" _
                        ) As String
Dim sTemp As String

sTemp = RemoveSlash(App.Path) & "\HTML Templates\" & TemplateFileName

If FileExists(TemplateFileName) Then
            'do nothing

'try looking in "HTML Templates" subfolder:
ElseIf FileExists(sTemp) Then
            TemplateFileName = sTemp
Else
            'not found!
            TemplateFileName = ""
End If

HTMLizeTemplate = HTML_Functions.HTMLizeTemplate(Text, PageTitle, PicturePath, PreserveSpaces, KeepHTTP, Target, TemplateFileName)

End Function
Public Function InsertFile(ByVal MainText As String, ByVal FileName As String, Optional ByVal AtEnd As Boolean = True) As String
Dim F As CTextFile
Dim sTemp As String


Set F = New CTextFile

F.FileOpen FileName, OpenForInput
sTemp = F.ReadAll
F.FileClose

Set F = Nothing

If AtEnd Then
    InsertFile = MainText & sTemp
Else
    InsertFile = sTemp & MainText
End If

End Function

Public Function LinesRemoveDuplicates(ByVal Text As String) As String

    LinesRemoveDuplicates = Lines_Functions.LinesRemoveDuplicates(Text)

End Function
Public Function LinesReverseOrder(ByVal Text As String) As String

    LinesReverseOrder = Lines_Functions.LinesReverseOrder(Text)

End Function
Public Function SaveFile(ByVal MainText As String, ByVal FileName As String) As String
Dim F As CTextFile


Set F = New CTextFile

F.FileOpen FileName, OpenForOutput
F.WriteStr MainText
F.FileClose

Set F = Nothing

End Function

Public Function LoadFile(ByVal MainText As String, ByVal FileName As String) As String
Dim F As CTextFile

' "MainText" is not used but MUST exist!
' This function is over-ridden in "CVBScriptEx" for use with VBScript.

Set F = New CTextFile

F.FileOpen FileName, OpenForInput
LoadFile = F.ReadAll

F.FileClose

Set F = Nothing

End Function
Public Function DelFromStartTo(ByVal MainStr As String, ByVal SubStr As String, Optional ByVal MatchCase As Boolean = False, Optional ByVal Inclusive As Boolean = False) As String

DelFromStartTo = Text_Functions.DelLeftTo(MainStr, SubStr, MatchCase, Inclusive)

End Function
Public Function DelFromEndTo(ByVal MainStr As String, ByVal SubStr As String, Optional ByVal MatchCase As Boolean = False, Optional ByVal Inclusive As Boolean = False) As String

DelFromEndTo = Text_Functions.DelRightTo(MainStr, SubStr, MatchCase, Inclusive)

End Function
Public Function LinesInsertText(ByVal Text As String, ByVal InsertWhat As String, ByVal InsertPos As Long, Optional ByVal IgnoreEmptyLines As Boolean = True) As String

    LinesInsertText = Lines_Functions.LinesInsert(Text, InsertWhat, InsertPos, IgnoreEmptyLines)

End Function
Public Function ReplaceAll(ByVal Text As String, ByVal FindWhat As String, ByVal ReplaceWith As String, Optional ByVal MatchCase = False) As String

If MatchCase Then
    ReplaceAll = Replace$(Text, FindWhat, ReplaceWith, 1, -1, vbBinaryCompare)
Else
    ReplaceAll = Replace$(Text, FindWhat, ReplaceWith, 1, -1, vbTextCompare)
End If

End Function
Public Function SentenceCase(ByVal Text As String) As String

SentenceCase = FirstCharUpper(Text)

End Function


Public Function EncodeToNumericEntities(ByVal Text As String) As String

EncodeToNumericEntities = HTML_Functions.Encode_HTML_Entities(Text, False)

End Function

Public Function AddLast(ByVal MainText As String, ByVal StringToAdd As String) As String

    AddLast = MainText & StringToAdd

End Function
Public Function AddFirst(ByVal MainText As String, ByVal StringToAdd As String) As String

    AddFirst = StringToAdd & MainText

End Function
Public Function EncodeToNamedEntities(ByVal Text As String) As String

EncodeToNamedEntities = HTML_Functions.Encode_HTML_Entities(Text, True)

End Function

Public Function HTMLEntitiesToASCII(ByVal Html As String) As String

HTMLEntitiesToASCII = RX_HTML_EntitiesASCII(Html)

End Function
Public Function LinesDelRightTo(ByVal Text As String, ByVal DelToWhat As String, Optional ByVal MatchCase As Boolean = False, Optional ByVal Inclusive As Boolean = False) As String
    
    LinesDelRightTo = Lines_Functions.LinesDelRightTo(Text, DelToWhat, MatchCase, Inclusive)

End Function
Public Function DelRight(ByVal Text As String, ByVal sCount As String) As String
Dim lCount As Long
lCount = Val(sCount)

    DelRight = Lines_Functions.LinesDelRight(Text, lCount)

End Function
Public Function CrLfTabTrim(ByVal Text As String) As String

    CrLfTabTrim = Text_Functions.CrLfTabTrim(Text)

End Function
Public Function LinesDelLeft(ByVal Text As String, ByVal lCount As Long) As String

    LinesDelLeft = Lines_Functions.LinesDelLeft(Text, lCount)

End Function
Public Function ExtractURLs(ByVal Html As String) As String

    ExtractURLs = RX_ExtractURLs(Html)

End Function
Public Function linesDelLeftTo(ByVal Text As String, ByVal DelToWhat As String, Optional ByVal MatchCase As Boolean = False, Optional ByVal Inclusive As Boolean = False) As String
    
    linesDelLeftTo = Lines_Functions.linesDelLeftTo(Text, DelToWhat, MatchCase, Inclusive)

End Function
Public Function RemoveTagKeepContent(ByVal Html As String, ByVal Tag As String) As String

RemoveTagKeepContent = RX_RemoveTagKeepContent(Html, Tag)

End Function
Public Function RemoveAllTags(ByVal Html As String)

    RemoveAllTags = RX_RemoveAllTags(Html)

End Function
Public Function PutCSSinHTML(ByVal Html As String) As String

PutCSSinHTML = RX_ProcessLink(Html)

End Function

Public Function HTMLEntitiesToANSI(ByVal Html As String) As String

HTMLEntitiesToANSI = RX_HTML_EntitiesANSI(Html)

End Function
Public Function TextToHtml(ByVal Text As String, Optional ByVal sPre As String = "") As String

Dim sTemp As String
Dim bPre As Boolean

If sPre <> "" Then
    bPre = CBool(sPre)
Else
    bPre = False
End If

sTemp = HTML_Functions.AddBr(Text, bPre)

If bPre Then
    TextToHtml = "<PRE>" & vbCrLf & sTemp & vbCrLf & "</PRE>" & vbCrLf
Else
    TextToHtml = sTemp
End If

End Function
Public Function RemoveLineBreaks(ByVal Text As String) As String

RemoveLineBreaks = Text_Functions.RemoveLineBreaks(Text)

End Function

Public Function RemoveNonAlphaNumChars(ByVal Words As String, ByVal CharsToKeep As String) As String
Dim iChars As Long
    
'    [AlphaNumeric Only] = 1
'    [All Printable] = 2
'    [All Text Chars] = 3

iChars = Val(CharsToKeep)
If iChars > 3 Or iChars < 1 Then
    iChars = 3
End If

RemoveNonAlphaNumChars = RemoveNonAlphaNum4(Words, iChars)

End Function
Public Function RegExpExtract(ByVal Text As String, ByVal Pattern As String) As String


RegExpExtract = RX_GenericExtract(Text, Pattern)

End Function
Public Function RegExpReplace(ByVal Text As String, ByVal Pattern As String, ByVal ReplaceWith As String) As String

RegExpReplace = RX_GenericReplace(Text, Pattern, ReplaceWith)

End Function

Public Function RemoveCommentTag(ByVal Html As String) As String

RemoveCommentTag = RX_RemoveCommentTagAndContent(Html)

End Function
Public Function UnicodeToANSI(ByVal Text As String) As String

UnicodeToANSI = StrConv(Text, vbFromUnicode)

End Function
Public Function ASCIItoANSI(ByVal Text As String) As String

ASCIItoANSI = ASCIItoANSI_API(Text)

End Function

Public Function ANSItoASCII(ByVal Text As String) As String

ANSItoASCII = ANSItoASCII_API(Text)

End Function

Public Function ANSIToASCII7(ByVal Text As String) As String

ANSIToASCII7 = StrToANSI127(Text)

End Function
Public Function ProperCase(ByVal Text As String) As String

ProperCase = StrConv(Text, vbProperCase)

End Function
Public Function LowerCase(ByVal Text As String) As String

LowerCase = StrConv(Text, vbLowerCase)

End Function

Public Function AddBrTag(ByVal Text As String) As String

    AddBrTag = RX_AddBR(Text)

End Function
Public Function LinesAddNumbers(ByVal Text As String, ByVal NumStart As Long, ByVal NumStep As Long, ByVal Delimiter As String, ByVal numDigits As Long, Optional ByVal IgnoreEmptyLines As Boolean = True) As String

    LinesAddNumbers = Lines_Functions.LinesAddNumbers(Text, NumStart, NumStep, Delimiter, numDigits, IgnoreEmptyLines)

End Function
Public Function LinesAddLeftRight(ByVal Text As String, ByVal AddToLeft As String, ByVal AddToRight As String, Optional ByVal IgnoreEmptyLines As Boolean = True) As String

    LinesAddLeftRight = Lines_Functions.LinesAddLeftRight(Text, AddToLeft, AddToRight, IgnoreEmptyLines)

End Function
Public Function LinesCompactBlank(ByVal Text As String) As String
    
    LinesCompactBlank = RegExp_Functions.RX_CompactBlankLines(Text)

End Function
Public Function CompactSpaces(ByVal Text As String) As String
    
    CompactSpaces = Text_Functions.CompactSpaces(Text)
    
End Function

Public Function FixNewLineChars(ByVal Text As String) As String

    FixNewLineChars = Text_Functions.FixNewLineChars(Text)
    
End Function
Public Function HtmlToText(ByVal Text As String)

HtmlToText = Html2Text(Text)

End Function

Public Function LinesRemoveBlank(ByVal Text As String) As String
    
    LinesRemoveBlank = RegExp_Functions.RX_RemoveBlankLines(Text)
    
End Function
Public Function ReverseText(ByVal Text As String, Optional ByVal ByLine As String = "True") As String

ReverseText = ReverseStr(Text, CBool(ByLine))

End Function
Public Function ExtractHREFs(ByVal Html As String) As String

ExtractHREFs = RX_ExtractHREFs(Html)

End Function

Public Function RemoveTagAndContents(ByVal Html As String, ByVal Tag As String, ByVal TagIsSingle As String) As String

RemoveTagAndContents = RX_RemoveTagWithContents(Html, Tag, CBool(TagIsSingle))

End Function
Public Function RemoveTagPath(ByVal Html As String, TagAttr As String, LocalOnly As String) As String

RemoveTagPath = RX_RemoveTagAttrPath(Html, TagAttr, CBool(LocalOnly))

End Function
Public Function ExtractTagWithContents(ByVal Html As String, ByVal Tag As String) As String

ExtractTagWithContents = RX_ExtractTagWithContents(Html, Tag)

End Function

Public Function SetMaxLineWidth(ByVal Text As String, ByVal iMax As String) As String
    
    SetMaxLineWidth = LinesSetMaxWidth(Text, CInt(iMax))
    
End Function

Public Function UpperCase(ByVal Text As String) As String

UpperCase = StrConv(Text, vbUpperCase)

End Function
Public Function ValidateImageTags(ByVal Html As String) As String

ValidateImageTags = RX_ValidateImageTags(Html)

End Function

Public Function TabToSpaces(ByVal Text As String, Optional ByVal NumSpaces As String)

TabToSpaces = Tab2Spaces(Text, Val(NumSpaces))

End Function
