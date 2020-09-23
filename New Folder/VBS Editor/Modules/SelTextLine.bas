Attribute VB_Name = "SelTextLine"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'not same as above, last param is Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const EM_GETSEL  As Long = &HB0
Private Const EM_SETSEL  As Long = &HB1
Private Const EM_GETLINE = &HC4
Private Const EM_GETLINECOUNT  As Long = &HBA
Private Const EM_LINEINDEX  As Long = &HBB
Private Const EM_LINELENGTH  As Long = &HC1
Private Const EM_LINEFROMCHAR  As Long = &HC9
Private Const EM_SCROLLCARET As Long = &HB7
Private Const WM_SETREDRAW  As Long = &HB
Private Const EM_SETRECT As Long = &HB3

Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2
Private Const EC_USEFONTINFO = &HFFFF&
Private Const EM_SETMARGINS = &HD3&
Private Const EM_GETMARGINS = &HD4&

Private Type RECT_API
   l As Long ' left of rectangular region
   t As Long ' top of region
   r As Long ' right of region
   b As Long ' bottom of region
End Type

Public Function GotoRTFLine(ctl As RichTextBox, ByVal LineNum As Long) As Boolean

On Error Resume Next

   Dim copyStart As Long
   Dim copyEnd As Long
   Dim currLine As Long
   Dim lineCount As Long
   Dim success As Long
   Dim currCursorPos As Long

   ctl.SetFocus
   

  'get the number of lines in the textbox
   lineCount = SendMessage(ctl.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)
                               
  'the control lines are 0-based, but we're making it
  'more friendly by allowing 1-based numbers to be passed,
  'so subtract 1 from the start number.
  
  'Nothing is subtracted from the end number
  'because we want the end line + its contents
  '(IOW, the specified line -1 + Len(specified line) )
  'to be selected.
  
  'The If statement below takes care of specifying
  'a line index larger than the actual number of
  'lines available. It is required.
   LineNum = LineNum - 1
  
  'proceeding only if there are lines to work with
   If lineCount > 0 Then
   
     'if the startline greater than 0
      If LineNum > 0 Then
         
        'get the number of chrs up to the
        'end of the desired start line
         copyStart = SendMessage(ctl.hwnd, _
                                 EM_LINEINDEX, _
                                 LineNum, ByVal 0&)
                                     
      Else 'start at the beginning
            'of the textbox
             copyStart = 0
      
      End If
      
     'if the lastline greater than 0 and
     'less then the number of lines in the
     'control..
'      If LineNum > 0 And _
'         LineNum <= lineCount Then
'
'         '..get the number of chrs up to the
'         'end of the desired last line
'          copyEnd = SendMessage(ctl.hwnd, _
'                                EM_LINEINDEX, _
'                                LineNum, ByVal 0&)
'
'      Else 'copy the whole thing
'             copyEnd = Len(ctl)
'
'      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''ME
      'copyEnd = copyEnd - 2  ' CRLF
      copyEnd = copyStart
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''ME
'MsgBox copyStart
'MsgBox copyEnd

     'Set the selection for the returned range.
     'This will return -1 if unsuccessful
      success = SendMessage(ctl.hwnd, _
                            EM_SETSEL, _
                            copyStart, _
                            ByVal copyEnd)
                               
      If success <> -1 Then
        'return the selected text
         GotoRTFLine = True
      Else
         GotoRTFLine = False
      End If
      
   End If
   
     'scroll the selected item into view
      Call SendMessage(ctl.hwnd, EM_SCROLLCARET, 0, ByVal 0)

End Function
Public Sub SetMargins(txtBox As TextBox, ByVal iLeft As Long, ByVal iTop As Long)

Dim X As Long
Dim RECT As RECT_API

RECT.l = iLeft              ' Set left to upper left corner
RECT.t = iTop               ' Set top to upper left corner
RECT.r = CLng(txtBox.Width)        ' Set right of region
RECT.b = CLng(txtBox.Height)       ' Set bottom of region
X = SendMessage(txtBox.hwnd, EM_SETRECT, 0, RECT)


End Sub
Public Function CurrentLine(txtBox As Control) As Long

CurrentLine = SendMessage(txtBox.hwnd, EM_LINEFROMCHAR, ByVal -1, 0&) + 1

End Function
Public Property Let RightMargin(ByVal ctl As TextBox, ByVal lMargin As Long)
   Dim lhWnd As Long
   lhWnd = ctl.hwnd
   If (lhWnd <> 0) Then
      SendMessageLong lhWnd, EM_SETMARGINS, EC_RIGHTMARGIN, lMargin * &H10000
   End If
End Property
Public Property Get RightMargin(ByVal ctl As TextBox) As Long
   Dim lhWnd As Long
   lhWnd = ctl.hwnd
   If (lhWnd <> 0) Then
      RightMargin = SendMessageLong(lhWnd, EM_GETMARGINS, 0, 0) \ &H10000
   End If
End Property
Public Property Let LeftMargin(ByVal ctl As TextBox, ByVal lMargin As Long)
   Dim lhWnd As Long
   lhWnd = ctl.hwnd
   If (lhWnd <> 0) Then
      SendMessageLong lhWnd, EM_SETMARGINS, EC_LEFTMARGIN, lMargin
   End If
End Property
Public Property Get LeftMargin(ByVal ctl As TextBox) As Long
   Dim lhWnd As Long
   lhWnd = ctl.hwnd
   If (lhWnd <> 0) Then
      LeftMargin = (SendMessageLong(lhWnd, EM_GETMARGINS, 0, 0) And &HFFFF&)
   End If
End Property
Public Function SelText(ctl As Control, _
                             ByVal LineStart As Long, _
                             ByVal LineEnd As Long) As String

On Error Resume Next

   Dim copyStart As Long
   Dim copyEnd As Long
   Dim currLine As Long
   Dim lineCount As Long
   Dim success As Long
   Dim currCursorPos As Long

   ctl.SetFocus
   

  'get the number of lines in the textbox
   lineCount = SendMessage(ctl.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)
                               
  'the control lines are 0-based, but we're making it
  'more friendly by allowing 1-based numbers to be passed,
  'so subtract 1 from the start number.
  
  'Nothing is subtracted from the end number
  'because we want the end line + its contents
  '(IOW, the specified line -1 + Len(specified line) )
  'to be selected.
  
  'The If statement below takes care of specifying
  'a line index larger than the actual number of
  'lines available. It is required.
   LineStart = LineStart - 1
  
  'proceeding only if there are lines to work with
   If lineCount > 0 Then
   
     'if the startline greater than 0
      If LineStart > 0 Then
         
        'get the number of chrs up to the
        'end of the desired start line
         copyStart = SendMessage(ctl.hwnd, _
                                 EM_LINEINDEX, _
                                 LineStart, ByVal 0&)
                                     
      Else 'start at the beginning
            'of the textbox
             copyStart = 0
      
      End If
      
     'if the lastline greater than 0 and
     'less then the number of lines in the
     'control..
      If LineEnd > 0 And _
         LineEnd <= lineCount Then
         
         '..get the number of chrs up to the
         'end of the desired last line
          copyEnd = SendMessage(ctl.hwnd, _
                                EM_LINEINDEX, _
                                LineEnd, ByVal 0&)
 
      Else 'copy the whole thing
             copyEnd = Len(ctl)
      
      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''ME
      copyEnd = copyEnd - 2  ' CRLF
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''ME

     'Set the selection for the returned range.
     'This will return -1 if unsuccessful
      success = SendMessage(ctl.hwnd, _
                            EM_SETSEL, _
                            copyStart, _
                            ByVal copyEnd)
                               
      If success <> -1 Then
        'return the selected text
         SelText = ctl.SelText
      End If
      
   End If

   
     'scroll the selected item into view
      Call SendMessage(ctl.hwnd, EM_SCROLLCARET, 0, ByVal 0)


End Function
