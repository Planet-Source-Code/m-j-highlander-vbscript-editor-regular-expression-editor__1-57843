Attribute VB_Name = "API_Code"
Option Explicit

'--- WinXP Theming API ----------------------------------------------------------------------------------------------
'Init CpmCtl32.dll ver 6 for WinXP styles
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public Declare Function IsThemeActive Lib "UxTheme.dll" () As Boolean
Public Declare Function IsAppThemed Lib "UxTheme.dll" () As Boolean

'Activate XP Theming for a single control (pass ctl.hWnd)
Private Declare Function ActivateWindowTheme Lib "uxtheme" Alias "SetWindowTheme" ( _
    ByVal hwnd As Long, _
    Optional ByVal pszSubAppName As Long = 0, _
    Optional ByVal pszSubIdList As Long = 0) As Long

'Deactivate XP Theming for a single control (pass ctl.hWnd) 'why is it a string? i have no idea!
Private Declare Function DeactivateWindowTheme Lib "uxtheme" Alias "SetWindowTheme" ( _
     ByVal hwnd As Long, _
     Optional ByRef pszSubAppName As String = " ", _
     Optional ByRef pszSubIdList As String = " ") As Long
'------------------------------------------------------------------------------------------------------------------

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    y As Long
End Type

Public Type tagHH_POPUP
    cbStruct As Long    'specifies the size of the structure. You must always fill in this value before passing the structure.
    hinst As Long       'the instance handle for the string resource
    idString As Long    'string resource ID,
    pszText As String   'explicit string to display. To display this string you must set idString to 0 (zero)
    pt As POINTAPI      'POINTAPI structure that specifies the top center of the popup window
    clrForeground As ColorConstants     'foreground (text) color used in the pop-up. Use a VB color constant or a number in the format &HBBGGRR
    clrBackground As ColorConstants     'background color used in the pop-up. Use a VB color constant or a number in the format &HBBGGRR
    rcMargins As RECT   'RECT structure indicating the amount of space between edges of window and text. Use -1 for each member to ignore
    pszFont As String   'font to use in the pop-up. The string is in the following format: facename [, point size[, char set[, BOLD ITALIC UNDERLINE]]] You can skip an attribute by entering a comma for the attribute.
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWndCaller As Long, ByVal pszFileName As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Private Declare Function HtmlHelpPopUp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWndCaller As Long, ByVal pszFileName As String, ByVal uCommand As Long, ByRef dwData As tagHH_POPUP) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByRef Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7
Private Const EM_LINEINDEX = &HBB
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const WM_USER = &H400
Private Const TB_SETSTYLE = WM_USER + 56
Private Const TB_GETSTYLE = WM_USER + 57
Private Const TBSTYLE_FLAT = &H800
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const EM_SETTARGETDEVICE = (WM_USER + 72)
Private Const ES_LOWERCASE = &H10&
Private Const ES_UPPERCASE = &H8&
Private Const ES_NUMBER = &H2000&
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_CLOSE_ALL = &H12
Private Const HH_DISPLAY_TEXT_POPUP = (14 Or &HE)
Private Const MF_BITMAP = &H4&
Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MFT_STRING = &H0&
Private Const EM_SETSEL = &HB1
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_PASTESPECIAL = WM_USER + 64   ' for RTF only
Private Const EM_CANPASTE = WM_USER + 50       ' for RTF only
Private Const CF_TEXT = 1 ' Clipboard Format
Private Const WM_PASTE = &H302
Private Const WM_COPY = &H301
Private Const WM_CUT = &H300
Private Const WM_CLEAR = &H303
Private Const EM_EMPTYUNDOBUFFER = &HCD
Private Const EM_SETTEXTMODE = (WM_USER + 89)
Private Const EM_GETUNDONAME = (WM_USER + 86)
Private Const EM_GETREDONAME = (WM_USER + 87)

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const STILL_ACTIVE = &H103
Const PROCESS_QUERY_INFORMATION = &H400

Public Enum EditOperations 'All are WM_ messages
    Edit_Cut = &H300
    Edit_Copy = &H301
    Edit_Paste = &H302
    Edit_Clear = &H303
    Edit_Uundo = &H304
End Enum

'Private Type MENUITEMINFO
'cbSize As Long
'fMask As Long
'   fType As Long
'   fState As Long
'   wID As Long
'   hSubMenu As Long
'  hbmpChecked As Long
'   hbmpUnchecked As Long
'   dwItemData As Long
'   dwTypeData As String
'   cch As Long
'End Type
Private Const MF_MENUBARBREAK = &H20& ' columns with a separator line
Private Const MF_MENUBREAK = &H40&    ' columns w/o a separator line
Private Const MF_STRING = &H0&
Private Const MF_HELP = &H4000&
Private Const MFS_DEFAULT = &H1000&
'Private Const MIIM_ID = &H2
Private Const MIIM_SUBMENU = &H4
'Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20
'Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function GetMenuItemInfo Lib "user32" _
   Alias "GetMenuItemInfoA" _
   (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, _
   lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" _
   Alias "SetMenuItemInfoA" _
   (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, _
   lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) _
   As Long
'Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, _
   ByVal nPos As Long) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()


'************************************
' Shell from a VB App and open a URL
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal _
    lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public Enum Shell_Window_Style
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
End Enum

'RichText Undo Mode
Public Enum TextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2 ' /* default behavior */
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8 ' /* default behavior */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32 ' /* default behavior */
End Enum

'RichText Undo Operation Name
Public Enum ERECUndoTypeConstants
    ercUID_UNKNOWN = 0
    ercUID_TYPING = 1
    ercUID_DELETE = 2
    ercUID_DRAGDROP = 3
    ercUID_CUT = 4
    ercUID_PASTE = 5
End Enum

Private Const EM_SETMARGINS = &HD3&
Private Const EC_LEFTMARGIN = &H1


'Private Const VK_CONTROL = &H11 '= vbKeyControl
'Private Const VK_SHIFT = &H10   '= vbKeyShift
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const DT_CALCRECT = &H400

Private Declare Function DrawText Lib "user32" Alias _
    "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As RECT, ByVal wFormat _
    As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Const LF_FACESIZE = 32
Private Type CHARFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To LF_FACESIZE - 1) As Byte
    wPad2 As Integer
    wWeight As Integer              '// Font weight (LOGFONT value)
    sSpacing As Integer             '// Amount to space between letters
    crBackColor As Long             '// Background color
    lLCID As Long                   '// Locale ID
    dwReserved As Long              '// Reserved. Must be 0
    sStyle As Integer               '// Style handle
    wKerning As Integer             '// Twip size above which to kern char pair
    bUnderlineType As Byte          '// Underline type
    bAnimation As Byte              '// Animated text like marching ants
    bRevAuthor As Byte              '// Revision author index
    bReserved1 As Byte
End Type
Private Const EM_SETCHARFORMAT = (WM_USER + 68)
'// Font Back Color
Private Const CFM_BACKCOLOR = &H4000000
Private Const CFE_AUTOBACKCOLOR = CFM_BACKCOLOR
'// Selection type
Private Const SCF_SELECTION = &H1&
Private Const SCF_ALL = &H4&


Private Const MIIM_STATE As Long = &H1&

Private Const API_FALSE As Long = 0&
Private Const API_TRUE As Long = 1&

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Const LB_ITEMFROMPOINT = &H1A9
Public Sub ExploreFolder(FolderName As String)
    
'On Error Resume Next

ShellExecute 0&, "Explore", FolderName, 0&, 0&, SW_NORMAL

End Sub
Public Sub Set_ListBox_ToolTip(ListBox As ListBox)
'Use it in ListBox_MouseMove() Event Handler
'by Aaron Young <ajyoung@pressenter.com>,<aarony@redwingsoftware.com>

Dim tPOINT As POINTAPI
Dim Index As Long

'get the Mouse Cursor Position
Call GetCursorPos(tPOINT)

'Convert the Coords to be Relative to the Listbox
Call ScreenToClient(ListBox.hwnd, tPOINT)

'Find which Item the Mouse is Over
Index = SendMessageLong(ListBox.hwnd, LB_ITEMFROMPOINT, 0&, ByVal ((tPOINT.X And &HFF) Or (&H10000 * (tPOINT.y And &HFF))))
If Index >= 0 Then
    'Extract the List Index
    Index = Index And &HFF
    'set the Lists ToolTipText
    ListBox.ToolTipText = ListBox.List(Index)
End If

End Sub
Public Sub SetDefaultMenuItem(ByVal hwnd As Long, ByVal MenuPos As Long)
' by Bryan Stafford of New Vision Software - newvision@mvps.org
' Copyright © 2002  Bryan Stafford/New Vision Software
' Web page: http://www.mvps.org/vbvision/


Dim hMenu&, hSubMenu&, mii As MENUITEMINFO

  ' get the handle to the menubar
  hMenu = GetMenu(hwnd)

  If hMenu Then
        With mii
          .cbSize = Len(mii)
          .fMask = MIIM_STATE
          .fState = MFS_DEFAULT
        End With
        
        SetMenuItemInfo hMenu, MenuPos, API_TRUE, mii
   End If
    
    'Call DrawMenuBar(hWnd)

End Sub
Public Sub RemoveFontBackColor(ByVal RichTextBox As RichTextBox, Optional ByVal RemoveAll As Boolean = False)
Dim lRemoveWhat As Long
Dim tCF2 As CHARFORMAT2

tCF2.dwMask = CFM_BACKCOLOR '// Set BackColor mask
tCF2.dwEffects = CFE_AUTOBACKCOLOR '// Set AutoBackColor Effect
tCF2.crBackColor = -1 '// Set color to autocolor

tCF2.cbSize = Len(tCF2) '// Size of structure

'// Set char format to selection

If RemoveAll Then
    lRemoveWhat = SCF_ALL
Else
    lRemoveWhat = SCF_SELECTION
End If

SendMessageByRef RichTextBox.hwnd, EM_SETCHARFORMAT, lRemoveWhat, tCF2

End Sub
Public Sub SelFontBackColor(ByVal RichTextBox As RichTextBox, ByVal NewSelFontBackColor As OLE_COLOR)
    Dim tCF2 As CHARFORMAT2

    If NewSelFontBackColor = -1 Then '// If value is -1 then
        tCF2.dwMask = CFM_BACKCOLOR '// Set BackColor mask
        tCF2.dwEffects = CFE_AUTOBACKCOLOR '// Set AutoBackColor Effect
        tCF2.crBackColor = -1 '// Set color to autocolor
    Else
        tCF2.dwMask = CFM_BACKCOLOR '// Set BackColor mask
        tCF2.crBackColor = TranslateColor(NewSelFontBackColor) '// Set backcolor to new value
    End If
    tCF2.cbSize = Len(tCF2) '// Size of structure
    '// Set char format to selection
    SendMessageByRef RichTextBox.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, tCF2

End Sub

Private Function TranslateColor(ByVal Clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(Clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function
Public Function AutoSizeDropDownWidth(Combo As Object) As Boolean
'**************************************************************
'PURPOSE: Automatically size the combo box drop down width
'         based on the width of the longest item in the combo box

'PARAMETERS: Combo - ComboBox to size

'RETURNS: True if successful, false otherwise

'ASSUMPTIONS: 1. Form's Scale Mode is vbTwips, which is why
'                conversion from twips to pixels are made.
'                API functions require units in pixels
'
'             2. Combo Box's parent is a form or other
'                container that support the hDC property

'EXAMPLE: AutoSizeDropDownWidth Combo1
'****************************************************************
Dim lRet As Long
Dim bAns As Boolean
Dim lCurrentWidth As Single
Dim rectCboText As RECT
Dim lParentHDC As Long
Dim lListCount As Long
Dim lCtr As Long
Dim lTempWidth As Long
Dim lWidth As Long
Dim sSavedFont As String
Dim sngSavedSize As Single
Dim bSavedBold As Boolean
Dim bSavedItalic As Boolean
Dim bSavedUnderline As Boolean
Dim bFontSaved As Boolean

On Error GoTo ErrorHandler

If Not TypeOf Combo Is ComboBox Then Exit Function
lParentHDC = Combo.Parent.hdc
If lParentHDC = 0 Then Exit Function
lListCount = Combo.ListCount
If lListCount = 0 Then Exit Function


'Change font of parent to combo box's font
'Save first so it can be reverted when finished
'this is necessary for drawtext API Function
'which is used to determine longest string in combo box
With Combo.Parent

    sSavedFont = .FontName
    sngSavedSize = .FontSize
    bSavedBold = .FontBold
    bSavedItalic = .FontItalic
    bSavedUnderline = .FontUnderline
    
    .FontName = Combo.FontName
    .FontSize = Combo.FontSize
    .FontBold = Combo.FontBold
    .FontItalic = Combo.FontItalic
    .FontUnderline = Combo.FontItalic

End With

bFontSaved = True

'Get the width of the largest item
For lCtr = 0 To lListCount
   DrawText lParentHDC, Combo.List(lCtr), -1, rectCboText, _
        DT_CALCRECT
   'adjust the number added (20 in this case to
   'achieve desired right margin
   lTempWidth = rectCboText.Right - rectCboText.Left + 20

   If (lTempWidth > lWidth) Then
      lWidth = lTempWidth
   End If
Next
 
lCurrentWidth = SendMessageLong(Combo.hwnd, CB_GETDROPPEDWIDTH, _
    0, 0)

If lCurrentWidth > lWidth Then 'current drop-down width is
'                               sufficient

    AutoSizeDropDownWidth = True
    GoTo ErrorHandler
    Exit Function
End If
 
'don't allow drop-down width to
'exceed screen.width
 
   If lWidth > Screen.Width \ Screen.TwipsPerPixelX - 20 Then _
    lWidth = Screen.Width \ Screen.TwipsPerPixelX - 20

'//added by me: 5% increase:
lRet = SendMessageLong(Combo.hwnd, CB_SETDROPPEDWIDTH, 1.05 * lWidth, 0)

AutoSizeDropDownWidth = lRet > 0
ErrorHandler:
On Error Resume Next
If bFontSaved Then
'restore parent's font settings
  With Combo.Parent
    .FontName = sSavedFont
    .FontSize = sngSavedSize
    .FontUnderline = bSavedUnderline
    .FontBold = bSavedBold
    .FontItalic = bSavedItalic
 End With
End If
End Function
Public Sub HH_ShowPopUp(ByVal Text As String, Optional ByVal FontName As String = "Verdana", Optional ByVal FontSize As Byte = 8, Optional ByVal FontColor As Long = 0)
'Opens a pop-up window and displays a text string
Dim hhPopUp As tagHH_POPUP
Dim pt As POINTAPI
Dim rct As RECT

'    cbStruct As Long    'specifies the size of the structure. You must always fill in this value before passing the structure.
'    hinst As Long       'the instance handle for the string resource
'    idString As Long    'string resource ID,
'    pszText As String   'explicit string to display. To display this string you must set idString to 0 (zero)
'    pt As POINTAPI      'POINTAPI structure that specifies the top center of the popup window
'    clrForeground As ColorConstants     'foreground (text) color used in the pop-up. Use a VB color constant or a number in the format &HBBGGRR
'    clrBackground As ColorConstants     'background color used in the pop-up. Use a VB color constant or a number in the format &HBBGGRR
'    rcMargins As RECT   'RECT structure indicating the amount of space between edges of window and text. Use -1 for each member to ignore
'    pszFont As String   'font to use in the pop-up. The string is in the following format: facename [, point size[, char set[, BOLD ITALIC UNDERLINE]]] ,skip an attribute by entering a comma for the attribute


'Fill the pt POINTAPI structure
Call GetCursorPos(pt)

'Ignore the rct RECT structure
rct.Bottom = -1
rct.Left = -1
rct.Right = -1
rct.Top = -1

hhPopUp.clrBackground = RGB(255, 255, 225)
hhPopUp.clrForeground = FontColor
hhPopUp.hinst = 0
hhPopUp.idString = 0
hhPopUp.pszFont = FontName & "," & CStr(FontSize)
hhPopUp.pszText = Text
hhPopUp.pt = pt
hhPopUp.rcMargins = rct
hhPopUp.cbStruct = Len(hhPopUp)

Call HtmlHelpPopUp(0, "", HH_DISPLAY_TEXT_POPUP, hhPopUp)

End Sub
Public Function KeyIsPressed(VirtualKey As Integer) As Boolean
'RETURNS: true if the key specified is pressed
'determine if a keys state is down(includes mouse downs)

On Error GoTo KeyIsPressed_Erro_Handler:
  
If GetKeyState(VirtualKey) = -127 Or _
   GetKeyState(VirtualKey) = -128 Then
   
         KeyIsPressed = True
End If
                  

Exit Function
KeyIsPressed_Erro_Handler:
   
   If Err.Number <> 0 Then
       MsgBox Err.Number & vbCrLf & Err.Description
       Err.Clear
   End If
End Function
Public Sub RTF_Paste_Text(ByVal hWndRTF As Long)

    SendMessageLong hWndRTF, EM_PASTESPECIAL, CF_TEXT, 0

End Sub
Private Function EdithWnd(ByVal ctl As Control) As Long
'Used by SetLeftMargin()

   If TypeName(ctl) = "ComboBox" Then
      EdithWnd = FindWindowEx(ctl.hwnd, 0, "EDIT", vbNullString)
   ElseIf TypeName(ctl) = "TextBox" Then
      EdithWnd = ctl.hwnd
   End If

End Function
Public Sub SelectAll(ByVal EditCtlHwnd As Long)

SendMessageLong EditCtlHwnd, EM_SETSEL, 0, -1

End Sub
Public Sub SetLeftMargin(ByVal ctl As Control, ByVal lMargin As Long)
   
   Dim lhWnd As Long
   lhWnd = EdithWnd(ctl)
   If (lhWnd <> 0) Then
      SendMessageLong lhWnd, EM_SETMARGINS, EC_LEFTMARGIN, lMargin
   End If

End Sub
Public Function GetUndoType(ByVal rtfText As RichTextBox) As ERECUndoTypeConstants

    GetUndoType = SendMessageLong(rtfText.hwnd, EM_GETUNDONAME, 0, 0)

End Function
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
Public Function TranslateUndoType(ByVal eType As ERECUndoTypeConstants) As String

Select Case eType

   Case ercUID_UNKNOWN
      TranslateUndoType = "Last Action"

   Case ercUID_TYPING
      TranslateUndoType = "Typing"
   
   Case ercUID_PASTE
      TranslateUndoType = "Paste"
   
   Case ercUID_DRAGDROP
      TranslateUndoType = "Drag Drop"
   
   Case ercUID_DELETE
      TranslateUndoType = "Delete"
   
   Case ercUID_CUT
      TranslateUndoType = "Cut"

End Select

End Function
Public Sub EnableMultipleUndo(ByVal hWndRTF As RichTextBox)
Dim lStyle As Long

'// required to 'reveal' multiple undo
'// set rich text box style
lStyle = TM_RICHTEXT Or TM_MULTILEVELUNDO Or TM_MULTICODEPAGE
SendMessageLong hWndRTF.hwnd, EM_SETTEXTMODE, lStyle, 0

End Sub
Public Sub OpenHyperLink(ByVal sURL As String)
Dim iRet As Long

On Error GoTo URL_Error

iRet = ShellExecute(0, vbNullString, sURL, vbNullString, "c:\", SW_NORMAL)
Exit Sub

'<error handler>
URL_Error:
    MsgBox "Couldn't open URL", vbCritical, "Open URL Error"
    Err = 0
'</error handler>
End Sub


Public Sub RunFile(ByVal mFile As String, Optional ByVal RunStyle As Shell_Window_Style = SW_NORMAL)

Dim temp As Long
Dim msg As String, mFilePath As String
Dim X As Long

mFilePath = ExtractDirName(mFile)
temp = 0 'GetActiveWindow()
X = ShellExecute(temp, "Open", mFile, "", mFilePath, RunStyle)

If X < 32 Then
    Select Case X
        Case 0
            msg = "The file could not be run due to insufficient system memory or a corrupt program file"
        Case 2
            msg = "File Not Found"
        Case 3
            msg = "Invalid Path"
        Case 5
            msg = "Sharing or protection error"
        Case 6
            msg = "Separate data segments are required for each task "
        Case 8
            msg = "Insufficient memory to run the program"
        Case 10
            msg = "Incorrect Windows version"
        Case 11
            msg = "Invalid Program File"
        Case 12
            msg = "Program file requires a different operating System "
        Case 13
            msg = "Program requires MS-DOS 4.0"
        Case 14
            msg = "Unknown program file type"
        Case 15
            msg = "Windows prgram does not support protected memory mode"
        Case 16
            msg = "Invalid use of data segments when loading a second instance of a program"
        Case 19
            msg = "Attempt to run a compressed program file"
        Case 20
            msg = "Invalid dynamic link library"
        Case 21
            msg = "Program requires Windows 32-bit extensions"
        Case 31
            msg = mFilePath
            If Right(msg, 1) <> "\" Then msg = msg + "\"
            msg = msg + mFile
            Shell "rundll32.exe shell32.dll,OpenAs_RunDLL " + msg
    End Select

    If X <> 31 Then MsgBox msg, vbCritical, "Error Message"

End If
    
End Sub

Public Sub DragForm(Theform As Form)
    ReleaseCapture
    Call SendMessage(Theform.hwnd, &HA1, 2, 0&)
End Sub

Public Sub RightAlignMenu(ByVal frmX As Form, ByVal MenuIndex As Integer)

Dim mnuItemInfo As MENUITEMINFO, hMenu As Long
Dim BuffStr As String * 80   ' Define as largest possible menu text.

hMenu = GetMenu(frmX.hwnd)   ' Retrieve the menu handle.
BuffStr = Space(80)
With mnuItemInfo   ' Initialize UDT
      .cbSize = Len(mnuItemInfo)   ' 44
      .dwTypeData = BuffStr & Chr(0)
      .fType = MF_STRING
      .cch = Len(mnuItemInfo.dwTypeData)   ' 80
      .fState = MFS_DEFAULT
      .fMask = MIIM_ID Or MIIM_DATA Or MIIM_TYPE Or MIIM_SUBMENU
End With
' Use the desired item's position for the '3' below (zero-based list).
If GetMenuItemInfo(hMenu, MenuIndex, True, mnuItemInfo) = 0 Then
   MsgBox "GetMenuItemInfo failed. Error: " & Err.LastDllError, , _
          "Error"
Else
   mnuItemInfo.fType = mnuItemInfo.fType Or MF_HELP
   If SetMenuItemInfo(hMenu, MenuIndex, True, mnuItemInfo) = 0 Then
      MsgBox "SetMenuItemInfo failed. Error: " & Err.LastDllError, , _
             "Error"
   End If
End If
DrawMenuBar (frmX.hwnd)   ' Repaint top level Menu

End Sub
Sub MenuCols(frmX As Form, ByVal MenuNumber As Integer, ByVal ItemPerCol As Integer)
' Splitting a menu here demonstrates that this can be done dynamically.
   Dim mnuItemInfo As MENUITEMINFO, hMenu As Long, hSubMenu As Long
   Dim BuffStr As String * 80   ' Define as largest possible menu text.
   hMenu = GetMenu(frmX.hwnd)   ' retrieve menu handle.
   BuffStr = Space(80)
   With mnuItemInfo   ' Initialize the UDT.
          .cbSize = Len(mnuItemInfo)   ' 44
          .dwTypeData = BuffStr & Chr(0)
          .fType = MF_STRING
          .cch = Len(mnuItemInfo.dwTypeData)   ' 80
          .fState = MFS_DEFAULT
          .fMask = MIIM_ID Or MIIM_DATA Or MIIM_TYPE Or MIIM_SUBMENU
    End With
' Use item break point position for the '3' below (zero-based list).
   hSubMenu = GetSubMenu(hMenu, MenuNumber)
   If GetMenuItemInfo(hSubMenu, ItemPerCol, True, mnuItemInfo) = 0 Then
      'MsgBox "GetMenuItemInfo failed. Error: " & Err.LastDllError, , "Error"
    Else
      mnuItemInfo.fType = mnuItemInfo.fType Or MF_MENUBARBREAK
      If SetMenuItemInfo(hSubMenu, ItemPerCol, True, mnuItemInfo) = 0 Then
         'MsgBox "SetMenuItemInfo failed. Error: " & Err.LastDllError, , "Error"
      End If
   End If
   DrawMenuBar (frmX.hwnd)   ' Repaint top level Menu.

End Sub

Public Function DoEditOperation(TextBox As TextBox, ByVal Operation As EditOperations)

SendMessageLong TextBox.hwnd, Operation, 0, 0

End Function
Public Function xClearUndo(TextCtl As Control) As Boolean
    
xClearUndo = CBool(SendMessage(TextCtl.hwnd, EM_EMPTYUNDOBUFFER, ByVal CLng(0), ByVal CLng(0)))

End Function
Public Sub Clear(hwndEditCtl As Long)
  SendMessageLong hwndEditCtl, WM_CLEAR, 0, 0
End Sub
Public Sub Cut(hwndEditCtl As Long)
  SendMessageLong hwndEditCtl, WM_CUT, 0, 0
End Sub
Public Sub Copy(hwndEditCtl As Long)
  
  SendMessageLong hwndEditCtl, WM_COPY, 0, 0

End Sub
Public Function GetLineCount(TextBox As TextBox) As Long

GetLineCount = SendMessageLong(TextBox.hwnd, EM_GETLINECOUNT, 0, 0)

End Function
Public Function RTF_Can_Paste_Text(ByVal hWndRTF As Long) As Boolean

RTF_Can_Paste_Text = SendMessageLong(hWndRTF, EM_CANPASTE, CF_TEXT, 0) <> 0

End Function

Public Sub Paste(hwndCtl As Long)
'paste anything (text,graphics,...)

SendMessageLong hwndCtl, WM_PASTE, 0, 0

End Sub

Public Sub TextBoxSelectAll(txtBox As TextBox)
'Select all text in a text box

    Call SendMessageLong(txtBox.hwnd, EM_SETSEL, 0, -1)

End Sub
Public Sub SetMenuIconIcon(hwnd As Long, MenuIndex As Long, SubIndex As Long, pic1 As Picture, pic2 As Picture)
Dim hMenu As Long, hSubMenu As Long, hID As Long

'Get the menuhandle of the form
hMenu = GetMenu(hwnd)

'Get the handle of the first submenu
hSubMenu = GetSubMenu(hMenu, MenuIndex)

'Get the menuId of the first entry
hID = GetMenuItemID(hSubMenu, SubIndex)

'Add the bitmap
SetMenuItemBitmaps hMenu, hID, MF_BITMAP, pic1, pic2

End Sub

Sub ExecWait(ByVal JobToDo As String)

         Dim hProcess As Long
         Dim RetVal As Long
         'The next line launches JobToDo as icon,

         'captures process ID
         hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(JobToDo, 1))

         Do

             'Get the status of the process
             GetExitCodeProcess hProcess, RetVal

             'Sleep command recommended as well as DoEvents
             DoEvents: Sleep 100

         'Loop while the process is active
         Loop While RetVal = STILL_ACTIVE


End Sub

Public Sub SetMenuIcon(hwnd As Long, MenuIndex As Long, SubIndex As Long, pic As Picture)
Dim hMenu As Long, hSubMenu As Long, hID As Long

'Get the menuhandle of the form
hMenu = GetMenu(hwnd)

'Get the handle of the first submenu
hSubMenu = GetSubMenu(hMenu, MenuIndex)

'Get the menuId of the first entry
hID = GetMenuItemID(hSubMenu, SubIndex)

'Add the bitmap
SetMenuItemBitmaps hMenu, hID, MF_BITMAP, pic, pic

End Sub
Public Function CanUndo(TextCtl As Control) As Boolean

Dim iRetVal As Long

iRetVal = SendMessageLong(TextCtl.hwnd, EM_CANUNDO, 0&, 0&)
CanUndo = (iRetVal <> 0)

'CanUndo = CBool(SendMessage(TextCtl.hwnd, EM_CANUNDO, 0&, 0&))

End Function
Public Sub Undo(ByVal hTextCtl As Long)

    SendMessageLong hTextCtl, EM_UNDO, 0&, 0&

End Sub
Public Sub AddBorderToAllTextBoxes(frmX As Form)

Dim X As Control

On Error Resume Next
For Each X In frmX.Controls
        If TypeOf X Is TextBox Then
                AddOfficeBorder X
        End If
Next

End Sub


Public Sub AddOfficeBorder(ctlX As Control)
    
    Dim lngRetVal As Long
    
    'Retrieve the current border style
    lngRetVal = GetWindowLong(ctlX.hwnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
    'Apply the changes
    SetWindowLong ctlX.hwnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos ctlX.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Sub

Public Sub HHelp_Show(ByVal ChmFileName As String, HtmFileName As String)

Call HtmlHelp(0, ChmFileName, HH_DISPLAY_TOPIC, HtmFileName)

End Sub


Public Sub HHelp_Close()

Call HtmlHelp(0, "", HH_CLOSE_ALL, "")

End Sub



Public Sub CButtons(frmX As Form, Optional Identifier As String)
' Button.Style must be GRAPHICAL

Dim ctl As Control

For Each ctl In frmX      'loop trough all the controls on the form
    
    '3 Methods of doing it
    'If LCase(Left(Control.Name, Len(Identifier))) = LCase(Identifier) Then
    'If TypeName(Control) = "CommandButton" Then
    If TypeOf ctl Is CommandButton Then
                SendMessage ctl.hwnd, &HF4&, &H0&, 0&
    End If

Next ctl

End Sub

Public Sub SetTopMost(ByVal lhWnd As Long, ByVal bTopMost As Boolean)
'
' Set the hwnd of the window topmost or not topmost
'
    Dim lUseVal  As Long
    Dim lRet As Long
    
    lUseVal = IIF(bTopMost, HWND_TOPMOST, HWND_NOTOPMOST)
    
    lRet = SetWindowPos(lhWnd, lUseVal, 0, 0, 0, 0, flags)
    
    If lRet < 0 Then
'
' Couldn't do operation - handle error here
'
'        DisplayWinAPIError lRet
    End If

End Sub


' Comments  : Allow only numbers in a textbox
' Returns   : The Style of the textbox before the change.
Public Function NumbersOnly(TextBox As TextBox)
Dim DefaultStyle As Long

DefaultStyle = GetWindowLong(TextBox.hwnd, GWL_STYLE)
NumbersOnly = SetWindowLong(TextBox.hwnd, GWL_STYLE, DefaultStyle Or ES_NUMBER)

End Function
Public Function UpperCaseOnly(tBox As TextBox)

Dim DefaultStyle As Long
DefaultStyle = GetWindowLong(tBox.hwnd, GWL_STYLE)
UpperCaseOnly = SetWindowLong(tBox.hwnd, GWL_STYLE, DefaultStyle Or ES_UPPERCASE)

End Function

' Comments  : Allow only lowercase letters in a textbox
' Returns   : The Style of the textbox before the change.
Public Function LowerCaseOnly(tBox As TextBox)
Dim DefaultStyle As Long
DefaultStyle = GetWindowLong(tBox.hwnd, GWL_STYLE)
LowerCaseOnly = SetWindowLong(tBox.hwnd, GWL_STYLE, DefaultStyle Or ES_LOWERCASE)
End Function


' Comments  : Sets the style of a textbox.
' Returns   : The new style.
Public Function SetStyle(tBox As TextBox, NewStyle As Long)
SetStyle = SetWindowLong(tBox.hwnd, GWL_STYLE, NewStyle)
End Function


' Comments  : Gets the current style of a textbox.
' Returns   : The Style of the textbox.
Public Function GetStyle(tBox As TextBox)
GetStyle = GetWindowLong(tBox.hwnd, GWL_STYLE)
End Function

Public Function StyleNumberToText(tBox As TextBox)
Dim StyleNum  As Long
Dim StyleText As String

StyleNum = GetStyle(tBox)

Select Case StyleNum
    Case 1409360064: StyleText = "Number"
    Case 1409351880: StyleText = "Uppercase"
    Case 1409351888: StyleText = "Lowercase"
    Case Else: StyleText = "Other"
End Select

StyleNumberToText = StyleText
End Function

Public Sub SetWordWrap(RichTextBox As RichTextBox, WordWrap As Boolean)

If WordWrap Then
    'Enable word wrap:
    SendMessageLong RichTextBox.hwnd, EM_SETTARGETDEVICE, 0, 0
Else
    'Disable word wrap:
    SendMessageLong RichTextBox.hwnd, EM_SETTARGETDEVICE, 0, 1
End If

End Sub



Public Sub Rtf_SelChange(rtf As RichTextBox, Row As Long, Col As Long)
    
    Row = rtf.GetLineFromChar(rtf.SelStart) + 1
    Col = rtf.SelStart - SendMessage(rtf.hwnd, EM_LINEINDEX, -1, 0&) + 1

End Sub

Sub UnFlatAllBtns(frmX As Form)

Dim btnX As Control
For Each btnX In frmX.Controls
    If Left(btnX.Name, 3) = "cmd" Then
            UnbtnFlat btnX
    End If
    
Next btnX

End Sub
Public Sub BtnFlat(cmdFlat As CommandButton)
    SetWindowLong cmdFlat.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    cmdFlat.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Sub

Public Function UnbtnFlat(cmdFlat As CommandButton)
    SetWindowLong cmdFlat.hwnd, GWL_STYLE, WS_CHILD
    cmdFlat.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function

Sub FlatAllBtns(frmX As Form)

Dim btnX As Control
For Each btnX In frmX.Controls
    If TypeOf btnX Is CommandButton Then
            BtnFlat btnX
    End If
    
Next btnX

End Sub



Sub ToolFlat(ControlName As Control, flat As Boolean)
    Dim style As Long
    Dim hToolbar As Long
    Dim r As Long
       
'Now Make it Flat
    'First get the hWnd
    hToolbar = FindWindowEx(ControlName.hwnd, 0&, "ToolbarWindow32", vbNullString)
    'get Style
    style = SendMessageLong(hToolbar, TB_GETSTYLE, 0&, 0&)
    'Change style
    If (style And TBSTYLE_FLAT) And Not flat Then
        style = style Xor TBSTYLE_FLAT
    ElseIf flat Then
        style = style Or TBSTYLE_FLAT
    End If
    'Set the Style
    r = SendMessageLong(hToolbar, TB_SETSTYLE, 0, style)
    'Now show what we've done, this isn't neccesary if used in form_load
    ControlName.Refresh
End Sub

