Attribute VB_Name = "XP_Graphical_Button"
Option Explicit


' ********** API **********

Private Const GWL_WNDPROC = (-4)

Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function GetProp Lib "user32" _
    Alias "GetPropA" ( _
    ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" _
    Alias "SetPropA" ( _
    ByVal hWnd As Long, ByVal lpString As String, _
    ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" _
    Alias "RemovePropA" ( _
    ByVal hWnd As Long, ByVal lpString As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)

Private Const WM_PAINT = &HF
Private Const WM_DESTROY = &H2
Private Const WM_NCPAINT = &H85
Private Const WM_MOUSEHOVER = &H2A1
Private Const WM_MOUSELEAVE = &H2A3
Private Const WM_MOUSEMOVE = &H200
Private Const WM_SETFOCUS = &H7
Private Const WM_KILLFOCUS = &H8
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_ENABLE = &HA
Private Const WM_MOUSEACTIVATE = &H21
Private Const BM_GETSTATE = &HF2

Private Const BST_PUSHED = &H4
Private Const BST_FOCUS = &H8

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type PAINTSTRUCT
   hdc As Long
   fErase As Long
   rcPaint As RECT
   fRestore As Long
   fIncUpdate As Long
   rgbReserved(32) As Byte
End Type

Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" ( _
    ByVal hWnd As Long, _
    lpRect As Any, _
    ByVal bErase As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" ( _
    ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" ( _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal nPlanes As Long, _
    ByVal nBitCount As Long, _
    lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function InflateRect Lib "user32" ( _
    lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const COLOR_BTNTEXT = 18
Private Const COLOR_GRAYTEXT = 17

Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_WORDBREAK = &H10

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" ( _
    ByVal hdc As Long, _
    ByVal lpStr As String, _
    ByVal nCount As Long, _
    lpRect As RECT, _
    ByVal wFormat As Long) As Long

Type TrackMouseEvent
   cbSize As Long
   dwFlags As Long
   hwndTrack As Long
   dwHoverTime As Long
End Type

Private Const TME_HOVER = 1
Private Const TME_LEAVE = 2

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TrackMouseEvent) As Long

Const TRANSPARENT = 1

Private Declare Function TransparentBlt Lib "msimg32" ( _
  ByVal hDCDest As Long, _
  ByVal nXOriginDest As Long, _
  ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, _
  ByVal hHeightDest As Long, _
  ByVal hDCSrc As Long, _
  ByVal nXOriginSrc As Long, _
  ByVal nYOriginSrc As Long, _
  ByVal nWidthSrc As Long, _
  ByVal nHeightSrc As Long, _
  ByVal crTransparent As Long) As Long

Const SM_CXFOCUSBORDER = 83
Const SM_CYFOCUSBORDER = 84

' ********** Theme API **********

Const STAP_ALLOW_CONTROLS = 2

Private Declare Function GetThemeAppProperties Lib "uxtheme" () As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Long

Private Declare Function DrawThemeBackground Lib "uxtheme" ( _
   ByVal hTheme As Long, _
   ByVal hdc As Long, _
   ByVal iPartID As Long, _
   ByVal iStateID As Long, _
   pRect As RECT, _
   pClipRect As RECT) As Long

Private Declare Function DrawThemeText Lib "uxtheme" ( _
   ByVal hTheme As Long, _
   ByVal hdc As Long, _
   ByVal iPartID As Long, _
   ByVal iStateID As Long, _
   ByVal pszText As Long, _
   ByVal iCharCount As Long, _
   ByVal dwTextFlags As Long, _
   ByVal dwTextFlags2 As Long, _
   pRect As RECT) As Long

Private Declare Function DrawThemeEdge Lib "uxtheme" ( _
   ByVal hTheme As Long, _
   ByVal hdc As Long, _
   ByVal iPartID As Long, _
   ByVal iStateID As Long, _
   pDestRect As RECT, _
   ByVal uEdge As Long, _
   ByVal uFlags As Long, _
   pContentRect As Any) As Long

Declare Function GetThemeTextExtent Lib "uxtheme" ( _
   ByVal hTheme As Long, _
   ByVal hdc As Long, _
   ByVal iPartID As Long, _
   ByVal iStateID As Long, _
   ByVal pszText As Long, _
   ByVal iCharCount As Long, _
   ByVal dwTextFlags As Long, _
   pBoundingRect As Any, _
   pExtentRect As RECT) As Long

Private Declare Function IsAppThemed Lib "uxtheme" () As Long

Private Declare Function OpenThemeData Lib "uxtheme" ( _
   ByVal hWnd As Long, _
   ByVal pszClassList As Long) As Long

Private Declare Function CloseThemeData Lib "uxtheme" ( _
   ByVal hTheme As Long) As Long

Private Declare Function GetThemeSysColor Lib "uxtheme" ( _
   ByVal hTheme As Long, _
   ByVal iColorId As Long) As Long

Private Declare Function GetThemeSysSize Lib "uxtheme" ( _
   ByVal hTheme As Long, _
   ByVal iSizeId As Long) As Long
'
' MakeXPButton
'
' Converts a "Graphical" button to XP style
'
Sub MakeXPButton(ByVal Button As Object)
Dim hWnd As Long

   On Error GoTo NoXP

   If IsThemeActive() = 0 Then Exit Sub
   If IsAppThemed() = 0 Then Exit Sub

   ' Check the object class
   If TypeOf Button Is CommandButton Or _
      TypeOf Button Is OptionButton Or _
      TypeOf Button Is CheckBox Then

      ' Only subclass if the style is Graphical
      If Button.style = vbButtonGraphical Then

         ' Store the button object in the
         ' window and subclass it
         hWnd = Button.hWnd
         SetProp hWnd, "Button", ObjPtr(Button)
         SetProp hWnd, "WinProc", SetWindowLong(Button.hWnd, GWL_WNDPROC, AddressOf WinProc_Button)

      End If

   End If

NoXP:

End Sub

'
' DrawButton
'
' Draws a graphical button using the current
' XP visual style
'
Sub DrawButton(ByVal hWnd As Long)
Dim hdc As Long
Dim tPS As PAINTSTRUCT
Dim hTheme As Long, hBR As Long
Dim lState As Long
Dim bChecked As Boolean, bHot As Boolean, bFocused As Boolean
Dim bPushed As Boolean, bNoPicture As Boolean
Dim Button As Object, lFontOld As Long
Dim oPict As IPicture, oFont As IFont
Dim tCR As RECT, tCRText As RECT

   On Error Resume Next

   ' Get the button object
   CopyMemory Button, GetProp(hWnd, "Button"), 4&

   ' Get the button state
   lState = SendMessage(hWnd, BM_GETSTATE, 0&, ByVal 0&)
   bChecked = Button.Value
   bHot = GetProp(hWnd, "Hot")
   bPushed = lState And BST_PUSHED
   bFocused = lState And BST_FOCUS

   ' Get the client rectangle
   GetClientRect hWnd, tCR

   ' Open the theme
   hTheme = OpenThemeData(hWnd, StrPtr("Button"))

   ' Get the button DC
   hdc = BeginPaint(hWnd, tPS)

   ' Fill the background using the
   ' parent window background because
   ' the button can have transparent parts
   hBR = CreateSolidBrush(TranslateColor(Button.Container.BackColor))
   FillRect hdc, tCR, hBR
   DeleteObject hBR

   ' Set the state and picture
   If Button.Enabled = False Then

      lState = 4
      Set oPict = Button.DisabledPicture

      If oPict Is Nothing Then
         Set oPict = Button.Picture
      ElseIf oPict.Handle = 0 Then
         Set oPict = Button.Picture
      End If

   ElseIf bHot And Not bPushed Then

      lState = 2

      If bChecked Then
         Set oPict = Button.DownPicture

         If oPict Is Nothing Then
            Set oPict = Button.Picture
         ElseIf oPict.Handle = 0 Then
            Set oPict = Button.Picture
         End If
      Else
         Set oPict = Button.Picture
      End If

   ElseIf bChecked Or bPushed Then

      lState = 3

      Set oPict = Button.DownPicture

      If oPict Is Nothing Then
         Set oPict = Button.Picture
      ElseIf oPict.Handle = 0 Then
         Set oPict = Button.Picture
      End If

   ElseIf GetProp(hWnd, "Hot") = 1 Then

      lState = 2
      Set oPict = Button.Picture

   ElseIf bFocused Then

      lState = 5
      Set oPict = Button.Picture

   Else

      lState = 1
      Set oPict = Button.Picture

   End If

   If oPict Is Nothing Then
      bNoPicture = True
   ElseIf oPict.Handle = 0 Then
      bNoPicture = True
   End If

   ' Draw the button background
   DrawThemeBackground hTheme, hdc, 1, lState, tCR, tCR

   If bFocused Then

      ' Draw the focus rectangle
      tCRText = tCR
      InflateRect tCRText, -3, -3

      DrawFocusRect hdc, tCRText

   End If

   If Len(Button.Caption) Then

      ' Select the button font
      Set oFont = Button.Font
      lFontOld = SelectObject(hdc, oFont.hFont)

      ' Calculate the text size
      tCRText = tCR
      DrawText hdc, Button.Caption, -1, tCRText, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK

      tCRText.Left = tCR.Left
      tCRText.Right = tCR.Right

      If bNoPicture Then
         tCRText.Top = (tCR.Bottom - tCRText.Bottom) / 2
         tCRText.Bottom = tCRText.Top + tCRText.Bottom
      Else
         tCRText.Top = tCR.Bottom - tCRText.Bottom - 5
         tCRText.Bottom = tCR.Bottom
      End If

      ' Set the text background
      SetBkMode hdc, TRANSPARENT

      ' Set the color
      If Button.Enabled Then
         SetTextColor hdc, GetThemeSysColor(hTheme, COLOR_BTNTEXT)
      Else
         SetTextColor hdc, GetThemeSysColor(hTheme, COLOR_GRAYTEXT)
      End If

      ' Draw the text
      DrawText hdc, Button.Caption, -1, tCRText, DT_CENTER Or DT_WORDBREAK

      ' Restore the original font
      SelectObject hdc, lFontOld

      tCR.Bottom = tCRText.Top

   End If

   If Not bNoPicture Then

      Dim lW As Long, lH As Long

      ' Convert from HIMETRIC to Pixels
      lW = oPict.Width / 2540 * (1440 / Screen.TwipsPerPixelX)
      lH = oPict.Height / 2540 * (1440 / Screen.TwipsPerPixelY)

      If Button.Enabled Then

         If Button.UseMaskColor Then
            ' Draw the image using the mask color
            DrawTransparentPicture oPict, hdc, (tCR.Right - lW) / 2, (tCR.Bottom - lH) / 2, lW, lH, Button.MaskColor
         Else
            ' Draw the image without using the mask color
            oPict.Render hdc, (tCR.Right - lW) / 2, (tCR.Bottom - lH) / 2 + lH, lW, -lH, _
                         0, 0, oPict.Width, oPict.Height, ByVal 0&
         End If

      Else

         ' Draw the image in disabled mode
         DrawDisabledPicture oPict, hdc, (tCR.Right - lW) / 2, (tCR.Bottom - lH) / 2, _
                             lW, lH, Button.MaskColor

      End If

   End If

   ' Release button object
   CopyMemory Button, 0&, 4&

   ' Release the DC
   EndPaint hWnd, tPS

   ' Close the theme
   CloseThemeData hTheme

End Sub

'
' DrawTransparentPicture
'
' Draws a transparent picture
'
Private Sub DrawTransparentPicture( _
   ByVal picSource As Picture, _
   ByVal hDCDest As Long, _
   ByVal xDest As Long, _
   ByVal yDest As Long, _
   ByVal cxDest As Long, _
   ByVal cyDest As Long, _
   ByVal clrMask As Long, _
   Optional ByVal xSrc As Long, _
   Optional ByVal ySrc As Long, _
   Optional ByVal cxSrc As Long, _
   Optional ByVal cySrc As Long)

Dim hDCSrc As Long, hDCScreen As Long
Dim hbmOld As Long

   If picSource Is Nothing Then Exit Sub
   If picSource.Type <> vbPicTypeBitmap Then Exit Sub

   If cxSrc = 0 Then cxSrc = cxDest
   If cySrc = 0 Then cySrc = cyDest

   hDCScreen = GetDC(0&)

   ' Select passed picture into an HDC
   hDCSrc = CreateCompatibleDC(hDCScreen)
   hbmOld = SelectObject(hDCSrc, picSource.Handle)

   ' Draw the bitmap in the destination DC
   TransparentBlt hDCDest, xDest, yDest, cxDest, cyDest, hDCSrc, xSrc, ySrc, cxSrc, cySrc, clrMask

   ' Restore the original bitmap
   SelectObject hDCSrc, hbmOld

   ' Release the DCs
   DeleteDC hDCSrc
   ReleaseDC 0&, hDCScreen

End Sub

'
' DrawDisabledPicture
'
' Draws a picture in B&W
'
Private Sub DrawDisabledPicture( _
   ByVal picSource As Picture, _
   ByVal hDCDest As Long, _
   ByVal xDest As Long, _
   ByVal yDest As Long, _
   ByVal cxDest As Long, _
   ByVal cyDest As Long, _
   ByVal MaskColor As Long)
Dim hDCSrc As Long, hDCScreen As Long, hDCBW As Long
Dim lBMPBW As Long, lBMPOld As Long

   If picSource Is Nothing Then Exit Sub
   If picSource.Type <> vbPicTypeBitmap Then Exit Sub

   hDCScreen = GetDC(0&)

   ' Select passed picture into an HDC
   hDCSrc = CreateCompatibleDC(hDCScreen)
   lBMPOld = SelectObject(hDCSrc, picSource.Handle)

   ' Create a B&W picture
   hDCBW = CreateCompatibleDC(hDCScreen)
   lBMPBW = CreateBitmap(cxDest, cyDest, 1, 1, ByVal 0&)
   DeleteObject SelectObject(hDCBW, lBMPBW)

   ' Set the source background to white
   ' When you use BitBlt to copy from a
   ' color to a B&W bitmap, windows
   ' will convert all pixels matching
   ' the source background color to white
   ' and everything else to black
   SetBkColor hDCSrc, MaskColor

   BitBlt hDCBW, 0, 0, cxDest, cyDest, hDCSrc, 0, 0, vbSrcCopy

   ' Draw the image using white
   ' as the transparent color
   TransparentBlt hDCDest, xDest, yDest, cxDest, cyDest, hDCBW, 0, 0, cxDest, cyDest, vbWhite

   SelectObject hDCSrc, lBMPOld

   DeleteDC hDCBW
   DeleteDC hDCSrc
   ReleaseDC 0&, hDCScreen

End Sub


'
' TranslateColor
'
' Converts an OLE_COLOR to RGB
'
Function TranslateColor(ByVal Clr As OLE_COLOR)

   If (Clr And &H80000000) = &H80000000 Then
      TranslateColor = GetSysColor(Clr And &HFF)
   Else
      TranslateColor = Clr
   End If

End Function

'
' WinProc_Button
'
' Button window procedure
'
Private Function WinProc_Button( _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Dim tTME As TrackMouseEvent
Dim lProc As Long

   ' Get the previous window procedure
   lProc = GetProp(hWnd, "WinProc")

   Select Case Msg

      Case WM_NCPAINT
         ' Do nothing
         Exit Function

      Case WM_PAINT

         ' Draw the button
         DrawButton hWnd
         Exit Function

      Case WM_DESTROY

         ' Unsubclass the window
         SetWindowLong hWnd, GWL_WNDPROC, lProc
         RemoveProp hWnd, "WinProc"
         RemoveProp hWnd, "Button"
   
   End Select

   ' Call the previous window procedure
   WinProc_Button = CallWindowProc(lProc, hWnd, Msg, wParam, lParam)

   Select Case Msg

      Case WM_MOUSEHOVER

         ' Mouse is over the button

         SetProp hWnd, "Hot", 1

         ' Redraw the button
         DrawButton hWnd

      Case WM_MOUSELEAVE

         ' Mouse has left the button

         RemoveProp hWnd, "Hot"
         DrawButton hWnd

      Case WM_MOUSEMOVE

         If GetProp(hWnd, "Hot") = 0 Then

            tTME.cbSize = LenB(tTME)
            tTME.hwndTrack = hWnd
            tTME.dwFlags = TME_HOVER Or TME_LEAVE
            tTME.dwHoverTime = 1

            TrackMouseEvent tTME

         End If

      Case WM_SETFOCUS, WM_KILLFOCUS, _
           WM_LBUTTONDOWN, WM_LBUTTONUP, _
           WM_KEYDOWN, WM_KEYUP, _
           WM_ENABLE, WM_MOUSEACTIVATE

         ' Draw the button
         DrawButton hWnd

   End Select


End Function
