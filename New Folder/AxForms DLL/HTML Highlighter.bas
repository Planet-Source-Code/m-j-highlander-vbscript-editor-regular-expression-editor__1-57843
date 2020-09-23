Attribute VB_Name = "HTML_Highlighter"
Option Explicit
' Thanks to Chris Thomas for giving me a push ahead with this
' little module, altering the RTF text has always been my
' solution to syntax highlighting, I just didn't know RTF so
' I had to resort to the old top-down highlight and then
' change color method, which sucked.

'Use:

'    Bubble_KeyDown KeyCode, Shift, rchText
'    Bubble_Change_HTML rchText




Private m_lColor As Long


' Returns a header.
Public Function BeginRTF(l_fntFont As StdFont, ParamArray l_alColors()) As String
    Dim alColors() As Long
    Dim lCounter As Long
    
    ReDim alColors(UBound(l_alColors))
    For lCounter = 0 To UBound(l_alColors)
        alColors(lCounter) = l_alColors(lCounter)
    Next lCounter
    
    m_lColor = -1
    BeginRTF = GenerateHeader(l_fntFont, alColors)
End Function

' Returns a color tag.
Public Function Color(l_lColor As Long)
    ' Create a Color tag.
    If m_lColor = l_lColor Then Exit Function
    Color = "\cf" + CStr(l_lColor) + " "
    m_lColor = l_lColor
End Function

' Creates a header.
Private Function GenerateHeader(l_fntFont As StdFont, l_alColors() As Long) As String
    Dim sTable As String
    Dim alColors() As Long
    Dim lCounter As Long
    
    ReDim alColors(UBound(l_alColors))
    For lCounter = 0 To UBound(l_alColors)
        alColors(lCounter) = l_alColors(lCounter)
    Next lCounter
    
    ' Define the RTF format.
    sTable = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033"
    
    ' Add the font table to it.
    sTable = sTable + GenerateFontTable(l_fntFont)
    ' Add the colors to it.
    sTable = sTable + GenerateColorTable(alColors)
    ' Add the first paragraph to it.
    sTable = sTable + "\lang1033\viewkind4\uc1\pard\cf0\fs" + CStr(Int(l_fntFont.SIZE * 2)) + " "
    ' Return the value.
    GenerateHeader = sTable
End Function

' Creates a font table with a single font.
Private Function GenerateFontTable(l_fntFont As StdFont) As String
    Dim sTable As String
    
    ' Define a color table.
    sTable = "{\fonttbl"
    ' With this font.
    sTable = sTable + "{\f0\fnil\fcharset" + CStr(l_fntFont.Charset) + " "
    sTable = sTable + l_fntFont.Name & "\fs" + Trim(Str(Int(l_fntFont.SIZE * 2)))
    ' Finish off.
    sTable = sTable + ";}}"
    ' Return the value.
    GenerateFontTable = sTable
End Function

' Generates a color table with a varying amount of colors.
Private Function GenerateColorTable(l_alColors() As Long) As String
    
    Dim sTable As String
    Dim lCounter As Long
    Dim sColor As String
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    
    sTable = "{\colortbl ;"

    ' Loop throught each color and get add it to the color table.
    For lCounter = 0 To UBound(l_alColors)
        ' Get the color components.
        lRed = CLng(l_alColors(lCounter)) Mod 256
        lGreen = CLng(l_alColors(lCounter)) \ 256 Mod 256
        lBlue = CLng(l_alColors(lCounter)) \ 256 \ 256 Mod 256
        
        ' Create a color definition.
        sColor = "{\red" + CStr(lRed) + "\green" + CStr(lGreen) + "\blue" + CStr(lBlue) + ";}"
    
        ' Add the color to the table.
        sTable = sTable + sColor
    Next lCounter
    
    sTable = sTable & "}"
    
    GenerateColorTable = sTable
End Function

' Call this from your change event for the right text box to
' Highight for HTML.
Public Sub Bubble_Change_HTML(ByRef TextBox As RichTextBox)
    Dim iPos As String
    iPos = TextBox.SelStart
    TextBox.TextRTF = HTMLHighlight(TextBox.Text, TextBox.Font, vbBlack, _
                                    vbBlue, RGB(128, 0, 0), RGB(0, 0, 128), _
                                    RGB(0, 128, 0))
    TextBox.SelStart = iPos
End Sub

' Call this from your keydown event for the rich text box.
Public Sub Bubble_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer, ByRef TextBox As RichTextBox)
    If KeyCode = vbKeyTab Then
        KeyCode = 0 ' This actually changes the keycode, as if you
                    ' did subclassing, I bet some people are going to
                    ' be knocking thier heads on the desk about this:
                    ' All that subclassing for NOTHING!
        If Shift = 1 Then
            On Error Resume Next
            If Mid(TextBox.Text, TextBox.SelStart - 3, 4) = "    " Then
                TextBox.SelText = "" ' Remove the text that was selected.
                TextBox.SelStart = TextBox.SelStart - 4
                TextBox.SelLength = 4
                TextBox.SelText = ""
            End If
        Else
            TextBox.SelText = Space(4) ' or vbTab if you want, I just prefer spaces...
        End If
    End If
    If KeyCode = vbKeyReturn Then
        ' Keep indent.
        Dim lPos As Long
        Dim lNum As Long
        For lPos = TextBox.SelStart To 1 Step -1
            If Mid(TextBox.Text, lPos, 1) = vbLf Then Exit For
            If Mid(TextBox.Text, lPos, 1) = " " Then lNum = lNum + 1 Else lNum = 0
        Next lPos
        KeyCode = 0
        TextBox.SelText = vbCrLf + Space(lNum)
    End If
End Sub

' Call this to convert text into highlighted text.
Public Function HTMLHighlight(ByVal l_sText As String, fntFont As StdFont, _
        lNormalColor As Long, lBraceColor As Long, lTagColor As Long, _
        lStringColor As Long, lCommentColor As Long) As String
    
    Dim sRtfHeader As String
    Dim sRtfText As String
    
    Dim lCounter As Long
    Dim sChar As String
    Dim bChar As Byte
    Dim psChar As String
    
    Dim bAfter As Boolean ' Do we change the color before or after the character?
    
    ' State Handlers.
    Dim lInStr As Long
    Dim bInTag As Boolean
    Dim bComment As Boolean
    Dim sColor As String
    Dim bAppendPar As Boolean
    
    If Right(l_sText, 2) = vbCrLf Then bAppendPar = True
    
    ' Normal Color, <> Color, Tag Color, String Color, Comment Color.
    sRtfHeader = BeginRTF(fntFont, lNormalColor, lBraceColor, lTagColor, lStringColor, lCommentColor)
    ' Escape the characters so that the RTF parser doesn't missunderstand them.
    l_sText = Replace(l_sText, "\", "\\")
    l_sText = Replace(l_sText, "{", "\{")
    l_sText = Replace(l_sText, "}", "\}")
    
    ' Loop through each letter.
    For lCounter = 1 To Len(l_sText)
        ' Get the character.
        psChar = sChar ' The old character.
        sChar = Mid(l_sText, lCounter, 1)
        bChar = Asc(sChar)
        bAfter = False ' Reset this flag.
        
        sColor = "" ' We have no color yet.
        ' First we have the inside tag color.
        If bInTag Then sColor = Color(3)
        
        ' Next we check for the <> characters.
        If sChar = "<" Then sColor = Color(2): bInTag = True
        If sChar = ">" Then sColor = Color(2): bInTag = False
        If psChar = ">" And Not bInTag Then sColor = Color(1)
        
        ' Now we do strings.
        If bInTag Then
            If sChar = "'" Then
                If lInStr = 0 Then
                    lInStr = 1
                Else
                    lInStr = 0
                End If
            End If
        End If
        
        ' Color the strings.
        If lInStr <> 0 Then sColor = Color(4): bAfter = True
        
        ' Do comment recognition.
        If sChar = "!" And psChar = "<" Then bComment = True
        If sChar = ">" And psChar = "-" And bComment Then bComment = False
        If sChar = vbCr Then sChar = "\par "
        If sChar = vbLf Then sChar = ""
        
        If bComment Then sColor = Color(5)
        
        ' Add a space if there is a color tag.
        'If sChar = " " And sColor <> "" Then sColor = sColor + " "
        If bAfter Then
            sRtfText = sRtfText + sChar + sColor
        Else
            sRtfText = sRtfText + sColor + sChar
        End If
    Next lCounter
    
    HTMLHighlight = sRtfHeader + sRtfText & IIF(bAppendPar, "\par ", "") & "}"
End Function
