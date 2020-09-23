Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const EM_SETMARGINS = &HD3&
Private Const EC_LEFTMARGIN = &H1

Private Const EM_SETSEL = &HB1

Public Sub TextBoxSelectAll(txtBox As TextBox)
'Select all text in a text box

    Call SendMessageLong(txtBox.hWnd, EM_SETSEL, 0, -1)

End Sub

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

Public Sub SetLeftMarginsAll(ByVal frmX As Form, ByVal LeftMargin As Long)
Dim ctl As Control

For Each ctl In frmX.Controls
    If TypeOf ctl Is TextBox Then
        'If ctl.MultiLine = False Then
            SetLeftMargin ctl, LeftMargin
        'End If
    End If
Next

End Sub

Public Sub SetLeftMargin(ByVal ctl As Control, ByVal lMargin As Long)
   
   Dim lhWnd As Long
   lhWnd = ctl.hWnd 'EdithWnd(ctl)
   If (lhWnd <> 0) Then
      SendMessageLong lhWnd, EM_SETMARGINS, EC_LEFTMARGIN, lMargin
   End If

End Sub
Public Sub CButtons(frmX As Form, Optional Identifier As String)
' Button.Style must be GRAPHICAL

Dim ctl As Control

For Each ctl In frmX      'loop trough all the controls on the form
    
    '3 Methods of doing it
    'If LCase(Left(Control.Name, Len(Identifier))) = LCase(Identifier) Then
    'If TypeName(Control) = "CommandButton" Then
    If TypeOf ctl Is CommandButton Then
                SendMessage ctl.hWnd, &HF4&, &H0&, 0&
    End If

Next ctl

End Sub
