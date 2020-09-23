Attribute VB_Name = "ComboHeight"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you can not publish
'               or reproduce this code on any web site,
'               on any online service, or distribute on
'               any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Declare Function MoveWindow Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long

Private Declare Function GetWindowRect Lib "user32" _
  (ByVal hWnd As Long, _
   lpRect As RECT) As Long

Private Declare Function ScreenToClient Lib "user32" _
  (ByVal hWnd As Long, _
   lpPoint As POINTAPI) As Long

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETITEMHEIGHT = &H154

Public Sub SetComboHeight(frmX As Form, ctlCombo As ComboBox, ByVal NumItemsToDisplay As Integer)

Dim pt As POINTAPI
Dim rc As RECT
Dim cWidth As Long
Dim newHeight As Long
Dim oldScaleMode As Long

Dim itemHeight As Long
   
 
 
'Save the current form scalemode, then
'switch to pixels
 oldScaleMode = frmX.ScaleMode
 frmX.ScaleMode = vbPixels
 
'the width of the combo, used below
 cWidth = ctlCombo.Width

'get the system height of a single
'combo box list item
 itemHeight = SendMessage(ctlCombo.hWnd, CB_GETITEMHEIGHT, 0, ByVal 0)
 
'Calculate the new height of the combo box. This
'is the number of items times the item height
'plus two. The 'plus two' is required to allow
'the calculations to take into account the size
'of the edit portion of the combo as it relates
'to item height. In other words, even if the
'combo is only 21 px high (315 twips), if the
'item height is 13 px per item (as it is with
'small fonts), we need to use two items to
'achieve this height.
 newHeight = itemHeight * (NumItemsToDisplay + 2)
 
'Get the co-ordinates of the combo box
'relative to the screen
 Call GetWindowRect(ctlCombo.hWnd, rc)
 pt.x = rc.Left
 pt.y = rc.Top

'Then translate into co-ordinates
'relative to the form.
 Call ScreenToClient(frmX.hWnd, pt)

'Using the values returned and set above,
'call MoveWindow to reposition the combo box

If TypeOf ctlCombo.Container Is comctllib.Toolbar Then
    'a quick fix rather than a real solution!
    Call MoveWindow(ctlCombo.hWnd, pt.x, pt.y, ctlCombo.Width / Screen.TwipsPerPixelX, newHeight, True)
Else
    Call MoveWindow(ctlCombo.hWnd, pt.x, pt.y, ctlCombo.Width, newHeight, True)
End If
 
'Its done, so show the new combo height
' DROP-DOWN the combo:
'   Call SendMessage(Combo1.hWnd, CB_SHOWDROPDOWN, True, ByVal 0)
 
'restore the original form scalemode before leaving
 frmX.ScaleMode = oldScaleMode
   
End Sub

