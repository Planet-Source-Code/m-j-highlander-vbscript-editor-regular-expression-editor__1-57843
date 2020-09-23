Attribute VB_Name = "VBS_SyntaxColorize"
Option Explicit
'by: M. Schweighauser <jogeli2@yahoo.de> | Date: 24.06. 2000

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_LINEINDEX = &HBB
Private Const EM_GETRECT = &HB2
Private Const WM_GETFONT = &H31

Private Const KeyWords = "|UseEscapes|#INCLUDE|OFF|And|As|Base|ByVal|Call|Case|CBool|CByte|CCur|CDate|CDbl|CDec|CInt|CLng|Close|Compare|Const|CSng|CStr|Currency|CVar|Dim|Do|Each|Else|ElseIf|End|Enum|Error|Exit|Explicit|False|For|Function|Get|GoTo|If|In|Input|Is|LBound|Let|Lib|Like|Line|Lock|Loop|LSet|Name|New|Next|Not|Object|On|Open|Option|Or|Output|Print|Private|Property|Public|Put|Random|Read|ReDim|Resume|Return|RSet|Seek|Select|Set|Single|Spc|Static|String|Stop|Sub|Tab|Then|Then|True|Type|UBound|Unlock|Variant|Wend|While|With|Xor|Nothing|To|"

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type
Public Sub SyntaxColorize(ByVal RTFBox As RichTextBox, _
                        Optional ByVal CommentColor As Long = 32768, _
                        Optional ByVal StringColor As Long = 12806565, _
                        Optional ByVal KeysColor As Long = 16711680)

Dim lTextSelPos As Long, lTextSelLen As Long
Dim i As Long
Dim sBuffer As String, lBufferLen As Long
Dim lSelPos As Long, lSelLen As Long
Dim sTempBuffer As String
Dim sSearchChar As String, lSearchCharLen As Long

'// Save the cursor position
lTextSelPos = RTFBox.SelStart
'lTextSelLen = RTFBox.SelLength


'// Lock the WindowUpdate of the ReichTextBox
LockWindowUpdate RTFBox.hWnd

On Error GoTo SyntaxColorize_ErrHandler

With RTFBox
    sBuffer = .Text & " "
    lBufferLen = Len(sBuffer)
    sTempBuffer = ""

'////Me'
    RTFBox.SelStart = FirstVisibleChar(RTFBox)
    RTFBox.SelLength = LastVisibleChar(RTFBox, lBufferLen) - FirstVisibleChar(RTFBox)
    RTFBox.SelColor = vbBlack
    RTFBox.SelStart = lTextSelPos
    RTFBox.SelLength = lTextSelLen
'///Me'

    For i = FirstVisibleChar(RTFBox) To LastVisibleChar(RTFBox, lBufferLen)

      Select Case Asc(Mid(sBuffer, i, 1))
      
        Case 34 '// Stringtexts -> " ... "
          .SelStart = i - 1
          i = InStr(i + 1, sBuffer, """", 1)
          .SelLength = i - .SelStart
          .SelColor = StringColor
        
        Case 47, 39, 60 '// Comments              Examples:
          
          If Mid(sBuffer, i, 2) = "//" Then       '// C    Comment
            sSearchChar = vbCrLf
            lSearchCharLen = 0
          ElseIf Mid(sBuffer, i, 2) = "/*" Then   '// C++  Comment
            sSearchChar = "*/"
            lSearchCharLen = 2
          ElseIf Mid(sBuffer, i, 4) = "<!--" Then '// HTML Comment
            sSearchChar = "//-->"
            lSearchCharLen = 5
          ElseIf Mid(sBuffer, i, 1) = "'" Then    '// VB   Comment
            sSearchChar = vbCrLf
            lSearchCharLen = 0
          Else                                    '// No   Comment
            GoTo ExitComment
          End If
          
          '// Kill TempBuffer
          sTempBuffer = ""
          
          '// Colorize the comment string
          .SelStart = i - 1
          lSelLen = InStr(i, sBuffer, sSearchChar) + lSearchCharLen
          If lSelLen <> lSearchCharLen Then '// FileEnd ?
            lSelLen = lSelLen - i
          Else
            lSelLen = lBufferLen - i
          End If
          .SelLength = lSelLen
          .SelColor = CommentColor
          i = .SelStart + .SelLength
          
ExitComment:

        Case 97 To 122, 65 To 90, 35
          '// a to  z ,  A to Z , #
          '// Only this char can be colorize
          If sTempBuffer = "" Then lSelPos = i
          sTempBuffer = sTempBuffer & Mid(sBuffer, i, 1)
          
        Case Else
          If Trim(sTempBuffer) <> "" Then
            .SelStart = lSelPos - 1
            .SelLength = Len(sTempBuffer)
            If InStr(1, KeyWords, "|" & sTempBuffer & "|", 1) <> 0 Then
             .SelColor = KeysColor
            End If
          End If
        
          sTempBuffer = ""
        End Select
      Next
End With

SyntaxColorize_ErrHandler:

'// Set the Cursor to the old position
RTFBox.SelStart = lTextSelPos
'RTFBox.SelLength = lTextSelLen

'// Unlock the WindoUpdate-Lock
LockWindowUpdate 0


End Sub
Private Function LastVisibleChar(RTFBox As RichTextBox, LenFile As Long) As Long
Dim rc As RECT
Dim tm As TEXTMETRIC
Dim hdc As Long
Dim lFont As Long
Dim OldFont As Long
Dim di As Long
Dim lc As Long
Dim VisibleLines As Long
Dim LastVisibleLine As Long


  lc = SendMessage(RTFBox.hWnd, EM_GETRECT, 0, rc)
  lFont = SendMessage(RTFBox.hWnd, WM_GETFONT, 0, 0)
  hdc = GetDC(RTFBox.hWnd)
  If lFont <> 0 Then OldFont = SelectObject(hdc, lFont)
  di = GetTextMetrics(hdc, tm)
  If lFont <> 0 Then lFont = SelectObject(hdc, OldFont)
  VisibleLines = (rc.Bottom - rc.Top) / tm.tmHeight
  di = ReleaseDC(RTFBox.hWnd, hdc)
  
  LastVisibleLine = SendMessage(RTFBox.hWnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
  LastVisibleLine = LastVisibleLine + VisibleLines - 1      '// -1 ADDED BY me //'
  
  LastVisibleChar = SendMessageByNum(RTFBox.hWnd, EM_LINEINDEX, LastVisibleLine, 0&)
  If LastVisibleChar = -1 Or LastVisibleChar = 0 Then LastVisibleChar = LenFile
  
End Function

Public Sub SyntaxColorizeAll(ByVal RTFBox As RichTextBox, _
                            ByVal Text As String, _
                            Optional ByVal CommentColor As Long = 32768, _
                            Optional ByVal StringColor As Long = 12806565, _
                            Optional ByVal KeysColor As Long = 16711680)
'MsgBox "COLORIZING"
Dim i As Long
Dim sBuffer As String, lBufferLen As Long
Dim lSelPos As Long, lSelLen As Long
Dim sTempBuffer As String
Dim sSearchChar As String, lSearchCharLen As Long

RTFBox.Tag = "LOCK"
Set go_EditCtl = RTFBox


'// Lock the WindowUpdate of the RichTextBox
LockWindowUpdate RTFBox.hWnd

' Don't remove, or coloring fails
RTFBox.Text = ""
RTFBox.SelColor = vbBlack
RTFBox.Text = Text

'On Error GoTo SyntaxColorizeAll_ErrHandler

With RTFBox
    sBuffer = .Text & " "
    lBufferLen = Len(sBuffer)
    sTempBuffer = ""
    
    For i = 1 To Len(RTFBox.Text)

      Select Case Asc(Mid(sBuffer, i, 1))
      
        Case 34 'String -> " ... "
          .SelStart = i - 1
          i = InStr(i + 1, sBuffer, """", 1)
          .SelLength = i - .SelStart
          .SelColor = StringColor
        
        Case 39 '// Comment
          
          'If Mid(sBuffer, i, 1) = "'" Then    '// VB   Comment
            sSearchChar = vbCrLf
            lSearchCharLen = 0
          'Else                                    '// No   Comment
            'GoTo ExitComment
          'End If
          
          '// Kill TempBuffer
          sTempBuffer = ""
          
          '// Colorize the comment string
          .SelStart = i - 1
          lSelLen = InStr(i, sBuffer, sSearchChar) + lSearchCharLen
          If lSelLen <> lSearchCharLen Then '// FileEnd ?
            lSelLen = lSelLen - i
          Else
            lSelLen = lBufferLen - i
          End If
          .SelLength = lSelLen
          .SelColor = CommentColor
          i = .SelStart + .SelLength
          
ExitComment:

        Case 97 To 122, 65 To 90, 35
          '// a to  z ,  A to Z , #
          '// Only this char can be colorize
          If sTempBuffer = "" Then lSelPos = i
          sTempBuffer = sTempBuffer & Mid(sBuffer, i, 1)
          
        Case Else
          If Trim(sTempBuffer) <> "" Then
            .SelStart = lSelPos - 1
            .SelLength = Len(sTempBuffer)
            If InStr(1, KeyWords, "|" & sTempBuffer & "|", 1) <> 0 Then
             .SelColor = KeysColor
            End If
          End If
        
          sTempBuffer = ""
        End Select
      Next
End With


SyntaxColorizeAll_ErrHandler:

RTFBox.SelStart = 0 'lTextSelPos
RTFBox.SelLength = 0 'lTextSelLen
RTFBox.SelColor = vbBlack

'// Unlock the WindoUpdate-Lock
LockWindowUpdate 0

RTFBox.Tag = ""
Set go_EditCtl = Nothing

End Sub
Private Function FirstVisibleChar(RTFBox As RichTextBox) As Long
Dim FirstVisibleLine As Long

  FirstVisibleLine = SendMessage(RTFBox.hWnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
  FirstVisibleChar = SendMessageByNum(RTFBox.hWnd, EM_LINEINDEX, FirstVisibleLine, 0&)
  If FirstVisibleChar = 0 Then FirstVisibleChar = 1

End Function
