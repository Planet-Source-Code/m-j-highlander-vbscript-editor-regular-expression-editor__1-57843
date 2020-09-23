Attribute VB_Name = "HTML_Functions"
Option Explicit

Public Enum HTML_Operation
    Cancel
    DeleteTagAndContent
    DeleteTagKeepContent
    ExtractTagAndContent
    
    ReplaceTagAndContent
    ReplaceTagKeepContent
    
End Enum


Type sTag
    Href As String
    Text As String
End Type

Public Function AddBrTag(ByVal sText As String) As String
Dim sTemp As String

sTemp = Replace$(sText, vbCrLf, "<br>" & vbCrLf)

If sTemp = sText Then  ' didnt find any CRLFs
    AddBrTag = sText & "<br>" '& vbCrLf
Else
    AddBrTag = sTemp
End If


End Function

Public Sub EncloseTag(ctlEdit As Control, ByVal Before As String, ByVal After As String)

    Dim X As Integer, l As Long
'    X = t.SelStart
'    l = t.SelLength
'    t.SelText = b + t.SelText + E
'    t.SelStart = X + Len(b)
'    t.SelLength = l

ctlEdit.SelText = Before & ctlEdit.SelText & After

End Sub
Public Function HTML_List(ByVal Text As String, ByVal Ordered As Boolean) As String
Dim vTemp As Variant, idx As Integer
Dim sOpenTag As String, sCloseTag As String

If Ordered Then
    sOpenTag = "<ol>"
    sCloseTag = "</ol>"
Else
    sOpenTag = "<ul>"
    sCloseTag = "</ul>"
End If


If Trim(Text) = "" Then

    HTML_List = sOpenTag & vbCrLf & vbTab & "<li>" & vbCrLf & sCloseTag


Else

    vTemp = Split(Text, vbCrLf)
    For idx = LBound(vTemp) To UBound(vTemp)
        vTemp(idx) = vbTab & "<li>" & vTemp(idx)
    Next
    
    HTML_List = sOpenTag & vbCrLf & Join(vTemp, vbCrLf) & vbCrLf & sCloseTag

End If


End Function
Public Function TextAreaAddBR(ByVal sText As String, bPre As Boolean) As String
'Only ENTITIES are converted in TextArea

'sText = Replace(sText, "&", "&amp;")
'sText = Replace(sText, Chr$(34), "&quot;")
sText = Replace(sText, "<", "&lt;")
sText = Replace(sText, ">", "&gt;")

If bPre Then
    'nothing
Else
    sText = Replace$(sText, vbCrLf, "<BR>" & vbCrLf)
End If

TextAreaAddBR = sText

End Function
Public Function HTMLizeTemplate( _
                        ByVal Text As String, _
                        ByVal PageTitle As String, _
                        ByVal PicturePath As String, _
                        ByVal PreserveSpaces As Boolean, _
                        ByVal KeepHTTP As Boolean, _
                        ByVal Target As String, _
                        ByVal TemplateFileName As String _
                        ) As String

Dim sHTML As String
Dim sHead As String
Dim sBody As String
Dim sBGPic As String
Dim sBoldOpen As String, sBoldClose As String
Dim sPreOpen As String, sPreClose As String
Dim sFont As String, sBGColor As String, sTextColor As String
Dim sBGScrollable As String
Dim sTarget As String
Dim sTemplate As String
Dim f As CTextFile

Set f = New CTextFile
f.FileOpen TemplateFileName, OpenForInput
sTemplate = f.ReadAll
f.FileClose


sTemplate = Replace(sTemplate, "%TITLE%", PageTitle, 1, -1, vbBinaryCompare)

'If PicturePath = "" Then
'    sTemplate = Replace(sTemplate, "%BACKGROUND%", PicturePath, 1, -1, vbBinaryCompare)
'Else
'    sTemplate = Replace(sTemplate, "%BACKGROUND%", PicturePath, 1, -1, vbBinaryCompare)
'End If

If PreserveSpaces Then
    sPreOpen = vbCrLf & "<PRE>" & vbCrLf
    sPreClose = vbCrLf & "</PRE>" & vbCrLf
    sHTML = sPreOpen & AddBr(Text, True) & sPreClose
Else
    sPreOpen = ""
    sPreClose = ""
    sHTML = AddBr(Text, False)
End If

sHTML = DoLinks(sHTML, KeepHTTP, Target) ' http://  ftp://  www. (will ad http:// to it | IS IT A BUG?)


sHTML = Replace(sTemplate, "%BODY%", sHTML, 1, -1, vbBinaryCompare)

HTMLizeTemplate = sHTML

End Function
Public Function Encode_HTML_Entities(ByVal Text As String, Optional ByVal NamedEntities As Boolean = False) As String
Dim cntr As Long, idx As Long
Dim sAscArray() As Integer

Dim bText() As Byte
ReDim bText(0 To Len(Text) - 1)
ReDim sAscArray(0 To Len(Text) - 1)

' First hide &
Text = Replace(Text, "&", Chr$(7))

bText = StrConv(Text, vbFromUnicode) ' VB Strings are Double-Byte Unicode

idx = 0
For cntr = 0 To Len(Text) - 1
    Select Case bText(cntr)
        Case Is >= 127, 34, 60, 62  ' 32=SPACE, 38=& : HIDDEN!
            sAscArray(idx) = bText(cntr)
            idx = idx + 1
        Case Else 'do nothing
    End Select
Next cntr

If idx = 0 Then  ' No need for escaping any chars...except maybe &
    Encode_HTML_Entities = Replace(Text, Chr$(7), "&amp;")
    Exit Function
End If

ReDim Preserve sAscArray(0 To idx - 1)

InitEntityInfo

If NamedEntities Then
        For idx = 0 To UBound(sAscArray)
            Text = Replace$(Text, Chr$(sAscArray(idx)), EntityInfo(sAscArray(idx)).Name)
        Next
        Encode_HTML_Entities = Replace(Text, Chr$(7), "&amp;")
Else
        For idx = 0 To UBound(sAscArray)
            Text = Replace$(Text, Chr$(sAscArray(idx)), "&#" & EntityInfo(sAscArray(idx)).Code & ";")
        Next
        Encode_HTML_Entities = Replace(Text, Chr$(7), "&#38;")
End If

    
End Function
Public Function AddBROnly(ByVal Text As String) As String

AddBROnly = Replace$(Text, vbCrLf, "<BR>" & vbCrLf)

End Function
Function Html2Text(ByVal Html As String) As String
'not just strip tags, but tries to conserve formatting...

'<PRE>,<TEXAREA>,<XMP>,(that's all?) must be handled differently!
'NOTE that the following functions "hide"
'Spaces as Chr(2) and Tabs as Chr(3)
'XMP: preserve everything
Html = RX_NormalizeXmp(Html)
'TextArea: only convert Entities
Html = RX_NormalizeTextArea(Html)
'PRE: only preserve Spaces and Tabs and CrLfs
Html = RX_NormalizePre(Html)

Html = Replace$(Html, "</title>", "<br>", 1, -1, vbTextCompare)
'Remove Tags containing HTML-Specific data
Html = RX_RemoveTagWithContents(Html, "style", TagIsSingle:=False)
Html = RX_RemoveTagWithContents(Html, "script", TagIsSingle:=False)

' CrLf in HTML evaluates to a single space
Html = Replace$(Html, vbCrLf, " ", 1, -1, vbTextCompare)

' Preserve formatting (kind-of)
Html = Replace$(Html, "<br>", vbCrLf, 1, -1, vbTextCompare)
Html = Replace$(Html, "<p>", vbCrLf, 1, -1, vbTextCompare)
Html = Replace$(Html, "</p>", vbCrLf, 1, -1, vbTextCompare)
Html = Replace$(Html, "</tr>", vbCrLf, 1, -1, vbTextCompare)
Html = Replace$(Html, "</td>", vbTab, 1, -1, vbTextCompare)
Html = Replace$(Html, "</table>", vbCrLf, 1, -1, vbTextCompare)

Html = RX_RemoveAllTags(Html)
Html = LinesTrim(Html, True, True, True)
Html = CompactSpaces(Html)
Html = RX_CompactBlankLines(Html)

'Unhide Spaces & Tabs
Html = Replace$(Html, Chr$(2), " ")
Html = Replace$(Html, Chr$(3), vbTab)

'Most Common Entities
Html = Replace$(Html, "&nbsp;", " ", 1, -1, vbTextCompare)
Html = Replace$(Html, "&quot;", """", 1, -1, vbTextCompare)
Html = Replace$(Html, "&lt;", "<", 1, -1, vbTextCompare)
Html = Replace$(Html, "&gt;", ">", 1, -1, vbTextCompare)
Html = Replace$(Html, "&amp;", "&", 1, -1, vbTextCompare)


Html2Text = Html

End Function
Public Function IsURLLocal(ByVal Text As String) As Boolean
Dim pos As Long

pos = 0
pos = InStr(1, Text, "http://", vbTextCompare)
pos = pos + InStr(1, Text, "ftp://", vbTextCompare)
pos = pos + InStr(1, Text, "www.", vbTextCompare)
' if after this pos is still 0 , then none was found

If pos = 0 Then
    IsURLLocal = True
Else
    IsURLLocal = False
End If

End Function
Public Function RemovePath(ByVal Attr As String) As String
' INPUT:  Attribute with contents
' Output: Attribute with content after removing file path
' Example - HREF attribute: (Quotes are part of the strings)
' Input   href="file:///cool folder/help.html"
' Output  href="help.html"

Dim sTemp As String
Dim OpenQuote As Long, CloseQuote As Long, LastSlash As Long

CloseQuote = 0

sTemp = Attr
OpenQuote = InStr(1, sTemp, Quote)
CloseQuote = InStr(OpenQuote + 1, sTemp, Quote)

If CloseQuote <> 0 Then
    '''''Quotes Found, Handle Path
    sTemp = Mid$(sTemp, OpenQuote, CloseQuote - OpenQuote)
    sTemp = Replace$(Attr, "\", "/") ' just in case
    LastSlash = InStrRev(sTemp, "/")
    If LastSlash = 0 Then
        'No path info, do nothing.
        RemovePath = sTemp
    Else
        'Remove Path:
        sTemp = Mid$(sTemp, LastSlash + 1, Len(sTemp) - LastSlash - 1)
        RemovePath = Left$(Attr, OpenQuote - 1) & Quote & sTemp & Quote
    End If

Else
    '''''Quotes NOT Found, do nothing
    RemovePath = sTemp
End If

'MsgBox Attr
'MsgBox RemovePath

End Function

Public Function HTML_RemoveImgTags2(ByRef si As String) As String
'<=60  >=62
Dim b() As Byte
Dim c() As Byte
ReDim imgl(0 To 3) As Byte
ReDim imgu(0 To 3) As Byte
Dim InTag  As Boolean, InImg  As Boolean
Dim s100 As String * 100
Dim pos100 As Long
Dim so As String
Dim idx As Long, idxc As Long

ReDim b(0 To Len(si) - 1)
ReDim c(0 To Len(si) - 1)
b = StrConv(si, vbFromUnicode) ' VB Strings are Double-Byte Unicode
imgl = StrConv("<img", vbFromUnicode)
imgu = StrConv("<IMG", vbFromUnicode)

idxc = 0
For idx = 0 To UBound(b)
    If b(idx) = 60 Then
        InTag = True
''''''''''''''''''''''''''''''New Code //Start
'        s100 = Mid$(si, idx + 1, 100) 'read 100 chars ahead
'        pos100 = InStr(s100, ">") 'find first ">"
'        If pos100 > 0 Then idx = idx + pos100      'jump to it
'        ch = Mid$(si, idx, 1)
''''''''''''''''''''''''''''''New Code //End
    End If
    If InTag Then
        If (idx + 1 = InStrB(idx + 1, b, imgl) Or idx + 1 = InStrB(idx + 1, b, imgu)) Then
             InImg = True
             idx = idx + 6  ' 6=length_of[<img ]+1
        End If
    End If
    If b(idx) = 62 Then
        InTag = False
        If InImg Then InImg = False: b(idx) = 0
    End If
    
    If (Not (InImg) And b(idx) <> 0) Then
        c(idxc) = b(idx)
        idxc = idxc + 1
    End If
Next idx


HTML_RemoveImgTags2 = Left$(StrConv(c, vbUnicode), idxc - 1)

End Function

Public Function Make_CSS_Style(ByVal CSS_File_Contents As String) As String
' Convert contents of a CSS file into a
' <STYLE>...</STYLE> tag block

Make_CSS_Style = vbCrLf & "<STYLE type=text/css>" & vbCrLf & _
                 CSS_File_Contents & vbCrLf & "</STYLE>" & vbCrLf

End Function
Public Function HTML_RemoveAllTags3(ByRef si As String) As String
' si is passed by Reference to speed things up a little
' Hybrid algorithm:
' Reads char by char, but also uses InStr() to skip some loops

Dim InTag  As Boolean
Dim ch As String * 1
Dim s100 As String * 100
Dim pos100 As Long
Dim so As String
Dim idx As Long, idx2 As Long

so = String$(Len(si), " ") ' Allocate (more than) enough space

For idx = 1 To Len(si)
    ch = Mid$(si, idx, 1)
    If ch = "<" Then
        InTag = True
        ch = ""
''''''''''''''''''''''''''''''New Code //Start
        s100 = Mid$(si, idx + 1, 100) 'read 100 chars ahead
        pos100 = InStr(s100, ">") 'find first ">"
        If pos100 > 0 Then idx = idx + pos100      'jump to it
        ch = Mid$(si, idx, 1)
''''''''''''''''''''''''''''''New Code //End
    End If
    
    
    If ch = ">" Then
        InTag = False
        ch = ""
    End If
    If Not (InTag) Then
        idx2 = idx2 + 1
        Mid$(so, idx2, 1) = ch
    End If
Next idx

HTML_RemoveAllTags3 = Left$(so, idx2)

End Function

Public Function StripHrefPath(ByVal Href As String) As String
' href  : href="file:///cool folder/help.html"

Dim sTemp As String
Dim LastSlash As Long

If Href = "" Then
    StripHrefPath = ""
    Exit Function
End If

sTemp = Mid$(Href, 7, Len(Href) - 7)

sTemp = Replace$(sTemp, "\", "/") ' just in case
LastSlash = InStrRev(sTemp, "/")

If LastSlash = 0 Then
    'do nothing
Else
    sTemp = Mid$(sTemp, LastSlash + 1, Len(sTemp) - LastSlash + 1)
End If

StripHrefPath = "HREF=" & Quote & sTemp & Quote

End Function
Function UNUSED_DoHref(ByVal Html As String) As String
Dim qpos1 As Integer
Dim qpos2 As Integer
Dim hpos As Integer


Dim xTemp
Dim sTemp As String
Dim StartPos As Long, EndPos As Long
Dim EndChars As String
Dim sTempChars As String
Dim sTempCharsX As String
ReDim CurrentTag(1 To 1) As String
Dim idx As Long
Dim sTarget As String


sTemp = Html
EndPos = 1
idx = 1
StartPos = 0
Do
   
   'Find the HREF attribute
   StartPos = InStr(StartPos + 1, sTemp, "href", vbTextCompare)
   ' the Opening Quote is at: (UNUSED)
'   StartPos = InStr(StartPos + 1, sTemp, Qout, vbTextCompare)
If StartPos = 0 Then Exit Do
    'Find the Closing Quote '''' 6 = len("href") + 2
    EndPos = InStr(StartPos + 6, sTemp, Quote, vbTextCompare)
If EndPos = 0 Then Exit Do
    CurrentTag(idx) = Mid$(sTemp, StartPos, EndPos - StartPos + 1)
    'MsgBox CurrentTag(idx)
    sTempCharsX = String$(idx, Chr$(7))
    sTemp = Replace$(sTemp, CurrentTag(idx), sTempCharsX, 1, 1) 'replace only once
    idx = idx + 1
    ReDim Preserve CurrentTag(1 To idx)
Loop

If idx = 1 Then
    'do nothing
Else
    ReDim Preserve CurrentTag(1 To idx - 1)  ' Kill the extra cell
End If

For idx = LBound(CurrentTag) To UBound(CurrentTag)
    sTempCharsX = String$(idx, Chr$(7))
    CurrentTag(idx) = StripHrefPath(CurrentTag(idx))
    sTemp = Replace$(sTemp, sTempCharsX, CurrentTag(idx), 1, 1)
Next idx
 
 
UNUSED_DoHref = sTemp

End Function
Function GetImgSrc(ByVal sImgTag As String) As String
' FUNCTION: Retrieve the value of the SRC attribute in an IMG tag
' ASSUMPTIONS:
'   - Content of SRC attribute is enclosed in double Quotes
'   - LOWSRC attribute is not present or is after the SRC attribute
' INPUT:   Full IMG tag
' RETURN:  the filename and path in the SRC attribute
'          withoute the Quotes
' Example:
' sImageTag := "<IMG SRC="helpdesk/cool.gif">
' Will Return            helpdesk/cool.gif


Dim SrcPos As Long
Dim FirstQuote As Long
Dim LastQuote As Long

' Find SRC attribute
SrcPos = InStr(1, sImgTag, "src", vbTextCompare)

If SrcPos > 0 Then ' SRC Found, get openinig and closing Quotes
    FirstQuote = InStr(SrcPos + 4, sImgTag, Quote, vbTextCompare) + 1
    LastQuote = InStr(FirstQuote + 1, sImgTag, Quote, vbTextCompare)
End If

If ((FirstQuote > 0) And (LastQuote > 0)) Then
    GetImgSrc = Mid$(sImgTag, FirstQuote, LastQuote - FirstQuote)
Else
    ' No SRC or no Quotes
    GetImgSrc = ""
End If

End Function
Public Function HTML_ValidateImageTags(ByVal sHTML As String) As String
Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long
Dim sImgTag As String
Dim sImgFile As String

iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, "<img", vbTextCompare)

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, ">", vbTextCompare)

    sImgTag = Mid$(sHTML, iOpeningPos, iClosingPos - iOpeningPos + 1)
    sImgFile = CurrentDir & "\" & GetImgSrc(sImgTag)
    If FileExists(sImgFile) = False Then
        sTempChars = String$(iClosingPos - iOpeningPos + 1, Chr$(7))
        Mid$(sHTML, iOpeningPos) = sTempChars
    Else
        ' File exists, so do nothing (keep tag)
    End If

Loop

HTML_ValidateImageTags = Replace(sHTML, Chr$(7), "")

End Function
Public Function HTML_RemoveAllTags2(ByRef si As String) As String
' Handle source Char by Char ,
'  A small improvement is achieved by jumping 2 chars when
'  a "<" is found.

Dim InTag  As Boolean
Dim ch As String * 1
Dim so As String
Dim idx As Long, idx2 As Long

so = String$(Len(si), " ")

For idx = 1 To Len(si)
    ch = Mid$(si, idx, 1)
    If ch = "<" Then
        InTag = True
        ch = ""
        idx = idx + 1  'Here we increment the Loop's control variable'
    End If
    
    If ch = ">" Then
        InTag = False
        ch = ""
    End If
    If Not (InTag) Then
        idx2 = idx2 + 1
        Mid$(so, idx2, 1) = ch
    End If
Next idx

HTML_RemoveAllTags2 = Left$(so, idx2)

End Function
Public Function DoLinks(ByVal Text As String, ByVal KeepHTTP As Boolean, ByVal Target As String) As String
'Supported protocols:
' ftp://
' http://
' www.

Dim sTemp As String

sTemp = Text
sTemp = Replace(sTemp, "http://www.", Chr$(7), 1, -1, vbTextCompare)
sTemp = Replace(sTemp, "www.", "http://www.", 1, -1, vbTextCompare)
sTemp = Replace(sTemp, Chr$(7), "http://www.", 1, -1, vbTextCompare)
sTemp = DoHyperLinks(sTemp, "ftp://", True, "") ' no target for "ftp"
sTemp = DoHyperLinks(sTemp, "http://", KeepHTTP, Target)

DoLinks = sTemp

End Function
Public Function BeautifyLink(ByVal HyperLink As String, ByVal KeepHTTP As Boolean, ByVal SmallCase As Boolean) As String
Dim sTemp As String

If LCase$(Left$(HyperLink, 7)) <> "http://" Then
    sTemp = HyperLink  ' Not HTTP , so do nothing
Else
    If KeepHTTP Then
            sTemp = HyperLink  ' Keep HTTP , so do nothing
    Else
            sTemp = Right$(HyperLink, Len(HyperLink) - 7) ' Remove HTTP://
    End If
End If

If SmallCase Then
    sTemp = LCase$(sTemp)
End If

BeautifyLink = sTemp

End Function
Public Function DoEMails(ByVal Text As String) As String

Dim sTemp As String
Dim StartPos As Long, EndPos As Long, AtPos As Long
Dim EndChars As String
Dim sTempChars As String
Dim sTempCharsX As String
ReDim CurrentTag(1 To 1) As String
Dim idx As Long

'possible delemeters:
EndChars = " ()[],<>" & vbCrLf & vbTab & Quote
sTemp = Text
'++++++++++++++++++++++++++++++++++++++++++++++++++++
EndPos = 1
idx = 1
AtPos = 0
Do
    AtPos = InStr(AtPos + 1, sTemp, "@", vbTextCompare)
    If AtPos = 0 Then Exit Do

    EndPos = MultiInstr(AtPos, sTemp, EndChars, vbTextCompare)
    If EndPos = 0 Then EndPos = Len(sTemp) + 1
    StartPos = MultiInstrRev(AtPos, sTemp, EndChars, vbTextCompare) + 1
    
    CurrentTag(idx) = Mid(sTemp, StartPos, EndPos - StartPos)
   ' MsgBox CurrentTag(idx)
    sTempChars = String(EndPos - StartPos, "X")
    sTempCharsX = String(idx, Chr$(7))
    sTemp = Replace(sTemp, CurrentTag(idx), sTempCharsX, 1, 1) 'replace only once
    idx = idx + 1
    ReDim Preserve CurrentTag(1 To idx)
Loop

If idx = 1 Then
    'do nothing
Else
    ReDim Preserve CurrentTag(1 To idx - 1)  ' Kill the extra cell

    For idx = LBound(CurrentTag) To UBound(CurrentTag)
        sTempCharsX = String(idx, Chr$(7))
        CurrentTag(idx) = "<A HREF=" & Quote & "mailto:" & CurrentTag(idx) & Quote & ">" & CurrentTag(idx) & "</A>"
        sTemp = Replace(sTemp, sTempCharsX, CurrentTag(idx), 1, 1)
    Next idx

End If

''++++++++++++++++++++++++++++++++++++++++++++++++++++
'AtPos = InStr(1, sTemp, "@", vbTextCompare)
'EndPos = MultiInstr(AtPos, sTemp, EndChars, vbTextCompare)
'StartPos = MultiInstrRev(AtPos, sTemp, EndChars, vbTextCompare) + 1
'MsgBox Mid(sTemp, StartPos, EndPos - StartPos)

DoEMails = sTemp

End Function

Function DoHyperLinks(ByVal Text As String, ByVal Protocol As String, ByVal KeepHTTP As Boolean, ByVal Target As String) As String
Dim sTemp As String
Dim StartPos As Long, EndPos As Long
Dim EndChars As String
Dim sTempChars As String
ReDim CurrentTag(1 To 1) As String
Dim idx As Long
Dim sTarget As String

If Target = "" Then
    sTarget = " "
Else
    sTarget = " TARGET=" & Quote & Target & Quote & " "
End If

'Possible Endings:
EndChars = " ,)]<" & vbCrLf & vbTab & Quote

sTemp = Text
EndPos = 1
idx = 1
StartPos = 0
Do
    StartPos = InStr(StartPos + 1, sTemp, Protocol, vbTextCompare)

If StartPos = 0 Then Exit Do

    EndPos = MultiInstr(StartPos, sTemp, EndChars, vbTextCompare)
    CurrentTag(idx) = Mid(sTemp, StartPos, EndPos - StartPos)
    sTempChars = String$(idx, Chr$(7))
    sTemp = Replace$(sTemp, CurrentTag(idx), sTempChars, 1, 1) 'replace only once
    idx = idx + 1
    ReDim Preserve CurrentTag(1 To idx)
Loop

If idx > 1 Then
    ReDim Preserve CurrentTag(1 To idx - 1)  ' Kill the extra cell
End If

For idx = LBound(CurrentTag) To UBound(CurrentTag)
    sTempChars = String$(idx, Chr$(7))
    CurrentTag(idx) = "<A" & sTarget & "HREF=" & Quote & _
                      CurrentTag(idx) & Quote & ">" & _
                      BeautifyLink(CurrentTag(idx), KeepHTTP, False) & _
                      "</A>"
    sTemp = Replace$(sTemp, sTempChars, CurrentTag(idx), 1, 1)
Next idx
 
DoHyperLinks = sTemp

End Function
Public Function AddBr(ByVal sText As String, ByVal bPre As Boolean) As String
    
sText = Replace(sText, "&", "&amp;")
sText = Replace(sText, Chr$(34), "&quot;")
sText = Replace(sText, "<", "&lt;")
sText = Replace(sText, ">", "&gt;")

If bPre Then
    'nothing
Else
    sText = Replace$(sText, vbCrLf, "<BR>" & vbCrLf)
End If

AddBr = sText

End Function
Function RevRGB(ByVal VBHexRGB As String) As String
' VB generated Hex RGB must be reversed to be used in HTML

Dim var1 As String
Dim var2 As String
Dim Var3 As String

var1 = Left$(VBHexRGB, 2)
var2 = Mid$(VBHexRGB, 3, 2)
Var3 = Right$(VBHexRGB, 2)

RevRGB = Var3 & var2 & var1

End Function

Public Function HTMLize(ByVal Text As String, _
                        ByVal PageTitle As String, _
                        ByVal PicturePath As String, _
                        ByVal PageBackColor As String, _
                        ByVal TextFontName As String, _
                        ByVal TextColor As String, _
                        ByVal TextSize As String, _
                        ByVal CopyPicture As Boolean, _
                        ByVal BackScroll As Boolean, _
                        ByVal TextBold As Boolean, _
                        ByVal PreserveSpaces As Boolean, _
                        ByVal KeepHTTP As Boolean, _
                        ByVal Target As String _
                        ) As String

Dim sHTML As String
Dim sHead As String
Dim sBody As String
Dim sBGPic As String
Dim sBoldOpen As String, sBoldClose As String
Dim sPreOpen As String, sPreClose As String
Dim sFont As String, sBGColor As String, sTextColor As String
Dim sBGScrollable As String
Dim sTarget As String

sHead = "<HEAD>" & vbCrLf
sHead = sHead & "<TITLE>" & PageTitle & "</TITLE>" & vbCrLf _
        & "<META content=""text/html; charset=windows-1252"" http-equiv=""Content-Type"">" _
        & vbCrLf & "</HEAD>" & vbCrLf

sFont = "<FONT FACE=" & Chr(34) & TextFontName & Chr(34) & " SIZE=" & TextSize & ">" & vbCrLf

If PicturePath = "" Then
    sBGPic = ""
Else
'    If chkCopy.Value = vbChecked Then
'        sPicFile = ExtractFileName(Trim(txtBGPic.Text))
'        sTgtDir = ExtractDirName(sTgtFile)
'        On Error Resume Next
'        FileCopy Trim(txtBGPic.Text), sTgtDir & sPicFile
'        If Err Then
'            sCopyResult = vbCrLf & "Couldn't copy " & Chr(34) & UCase(sPicFile) & Chr(34)
'        Else
'            sCopyResult = vbCrLf & Chr(34) & UCase(sPicFile) & Chr(34) & " was copied successfully."
'        End If
'        On Error GoTo 0
'    Else
'        sPicFile = Trim(txtBGPic.Text)
'    End If
    sBGPic = " BACKGROUND=" & Chr$(34) & PicturePath & Chr$(34)
End If


If TextBold Then
    sBoldOpen = "<B>" & vbCrLf
    sBoldClose = "</B>" & vbCrLf
Else
    sBoldOpen = ""
    sBoldClose = ""
End If

If PreserveSpaces Then
    sPreOpen = vbCrLf & "<PRE>" & vbCrLf
    sPreClose = vbCrLf & "</PRE>" & vbCrLf
    sHTML = AddBr(Text, True)
Else
    sPreOpen = ""
    sPreClose = ""
    sHTML = AddBr(Text, False)
End If

sHTML = DoLinks(sHTML, KeepHTTP, Target) ' http://  ftp://  www. (will ad http:// to it | IS IT A BUG?)

If BackScroll Then
    sBGScrollable = ""
Else
    sBGScrollable = " BGPROPERTIES = FIXED "
End If

sBody = "<BODY BGCOLOR=" & PageBackColor & " TEXT=" & TextColor & sBGPic & sBGScrollable & ">" & vbCrLf
sBody = sBody & sPreOpen & sFont & sBoldOpen


sHTML = "<HTML>" & vbCrLf & sHead & sBody & sHTML & sBoldClose & "</FONT>" & sPreClose & "</BODY>" & vbCrLf & "</HTML>"

HTMLize = sHTML

End Function
Public Function ColorToHex(ByVal lColor As Long) As String
Dim sTemp As String

sTemp = Hex$(lColor)

If Len(sTemp) < 6 Then sTemp = String$(6 - Len(sTemp), "0") + sTemp
sTemp = "#" & RevRGB(sTemp)

ColorToHex = sTemp

End Function
