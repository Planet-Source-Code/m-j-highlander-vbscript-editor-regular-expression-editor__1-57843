Attribute VB_Name = "OpenFileDlg_ViewMode"
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.

Option Explicit

'var for exposed property
Private m_lvInitialView As Long

'windows version constants
Private Const VER_PLATFORM_WIN32_NT As Long = 2
Private Const OSV_LENGTH As Long = 76
Private Const OSVEX_LENGTH As Long = 88
Private OSV_VERSION_LENGTH As Long  'our const to hold appropriate OSV length

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

'windows messages & notifications etc
Private Const WM_COMMAND = &H111
Public Const WM_NOTIFY As Long = &H4E&
Public Const WM_INITDIALOG As Long = &H110
Public Const CDN_FIRST As Long = -601
Public Const CDN_INITDONE As Long = (CDN_FIRST - &H0&)
Public Const MAX_PATH As Long = 260

'openfilename constants
Private Const OFN_ENABLEHOOK As Long = &H20
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_ENABLESIZING As Long = &H800000
Private Const OFN_EX_NOPLACESBAR As Long = &H1
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_PATHMUSTEXIST = &H800


'this is the version 5+ definition of
'the OPENFILENAME structure containing
'three additional members providing
'additional options on Windows 2000
'or later. The SetOSVersion routine
'will assign either OSV_LENGTH (76)
'or OSVEX_LENGTH (88) to the OSV_VERSION_LENGTH
'variable declared above. This variable, rather
'than Len(OFN) is used to assign the required
'value to the OPENFILENAME structure's nStructSize
'member which tells the OS if extended features
'- primarily the Places Bar - are supported.
Private Type OPENFILENAME
  nStructSize       As Long
  hWndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  sFile             As String
  nMaxFile          As Long
  sFileTitle        As String
  nMaxTitle         As Long
  sInitialDir       As String
  sDialogTitle      As String
  flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  sDefFileExt       As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
  pvReserved        As Long
  dwReserved        As Long
  flagsEx           As Long
End Type

Private OFN As OPENFILENAME

'defined As Any to support either the
'OSVERSIONINFO or OSVERSIONINFOEX structure
Private Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
  (lpVersionInformation As Any) As Long
  
Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
   Alias "GetOpenFileNameA" _
  (pOpenfilename As OPENFILENAME) As Long

Private Declare Function FindWindowEx Lib "user32" _
   Alias "FindWindowExA" _
  (ByVal hWndParent As Long, _
   ByVal hWndChildAfter As Long, _
   ByVal lpClassName As String, _
   ByVal lpWindowName As String) As Long

Private Declare Function GetParent Lib "user32" _
  (ByVal hwnd As Long) As Long

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Enum ViewType
'constants for listview view state provided by Brad Martinez
    SHVIEW_ICON = &H7029
    SHVIEW_LIST = &H702B
    SHVIEW_REPORT = &H702C
    SHVIEW_THUMBNAIL = &H702D
    SHVIEW_TILE = &H702E
End Enum

Private Function TrimNullChar(ByVal Text As String) As String
Dim lPos As Long

Text = LeftTo(Text, vbNullChar)

TrimNullChar = Trim$(Text)

End Function
Private Function LeftTo(ByVal Text As String, ByVal ToWhat As String) As String
Dim lPos As Long

lPos = InStr(1, Text, ToWhat, vbTextCompare)
If lPos = 0 Then
    LeftTo = Text
Else
    LeftTo = Left(Text, lPos - 1)
End If


End Function
Public Function OpenFileDlg(ByVal hWndOwner As Long, Optional ByVal FileName As String = "", Optional ByVal Filter As String = "", Optional ByVal InitialDir As String = "", Optional ByVal ViewMode As ViewType = SHVIEW_LIST, Optional ByVal ShowPlacesBar As Boolean = True) As String
Dim sFilters As String, sFileName As String
Dim OFN As OPENFILENAME

If OSV_VERSION_LENGTH = 0 Then SetOSVersion

'filters for the dialog

If Filter = "" Then
    sFilters = "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
Else
    sFilters = Replace(Filter, "|", vbNullChar) & vbNullChar & vbNullChar
End If

If FileName = "" Then
    sFileName = vbNullChar & Space$(MAX_PATH) & vbNullChar & vbNullChar
Else
    sFileName = FileName & Space$(MAX_PATH) & vbNullChar & vbNullChar
End If

'populate the structure
With OFN
   .nStructSize = OSV_VERSION_LENGTH
   .hWndOwner = hWndOwner
   .sFilter = sFilters
   .nFilterIndex = 0
   .sFile = sFileName
   .nMaxFile = Len(.sFile)
   .sDefFileExt = "*.*" & vbNullChar & vbNullChar
   .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
   .nMaxTitle = Len(.sFileTitle)
   .sInitialDir = InitialDir & vbNullChar & vbNullChar
   .sDialogTitle = "Open"
   .flags = OFN_EXPLORER Or OFN_ENABLEHOOK Or OFN_ENABLESIZING Or OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
   
   If ShowPlacesBar = False Then
        'hide places bar
        .flagsEx = OFN_EX_NOPLACESBAR
    End If

   .fnHook = FARPROC(AddressOf OFNHookProc)

End With

OFN_SetInitialView = ViewMode

If GetOpenFileName(OFN) Then
    OpenFileDlg = TrimNullChar(OFN.sFile)
Else
    OpenFileDlg = ""
End If

End Function
Public Function FARPROC(pfn As Long) As Long
  
  'A dummy procedure that receives and returns
  'the return value of the AddressOf operator.
 
  'Obtain and set the address of the callback
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)

  FARPROC = pfn

End Function
Public Function IsWin2000Plus() As Boolean

  'returns True if running Windows 2000 or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWin2000Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                      (osv.dwVerMajor = 5 And osv.dwVerMinor >= 0)
  
   End If

End Function
Public Property Let OFN_SetInitialView(ByVal initview As Long)

   m_lvInitialView = initview
   
End Property
Public Function OFNHookProc(ByVal hwnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

   Dim hWndParent As Long
   Dim hwndLv As Long
   Static bLvSetupDone As Boolean
   
   Select Case uMsg
      Case WM_INITDIALOG
        'Initdialog is set when the dialog
        'has been created and is ready to
        'be displayed, so set our flag
        'to prevent re-executing the code
        'in the wm_notify message. This is
        'required as the dialog receives a
        'number of WM_NOTIFY messages throughout
        'the life of the dialog. If this is not
        'done, and the user chooses a different
        'view, on the next WM_NOTIFY message
        'the listview would be reset to the
        'initial view, probably ticking off
        'the user. The variable is declared
        'static to preserve values between
        'calls; it will be automatically reset
        'on subsequent showing of the dialog.
         bLvSetupDone = False
         
        'other WM_INITDIALOG code here, such
        'as caption or button changing, or
        'centering the dialog.
      
      Case WM_NOTIFY
               
            If bLvSetupDone = False Then
               
             'hwnd is the handle to the dialog
             'hwndParent is the handle to the common control
             'hwndLv is the handle to the listview itself
               hWndParent = GetParent(hwnd)
               hwndLv = FindWindowEx(hWndParent, 0, "SHELLDLL_DefView", vbNullChar)
               
               If hwndLv > 0 Then
                  Call SendMessage(hwndLv, WM_COMMAND, ByVal m_lvInitialView, ByVal 0&)
                 
                 'since we found the lv hwnd, assume the
                 'command was received and set the flag
                 'to prevent recalling this routine
                  bLvSetupDone = True
               End If  'hwndLv

            End If  'bLvSetupDone

         Case Else
         
   End Select

End Function
Public Sub SetOSVersion()
  
   Select Case IsWin2000Plus()
      Case True
         OSV_VERSION_LENGTH = OSVEX_LENGTH '5.0+ structure size
      
      Case Else
         OSV_VERSION_LENGTH = OSV_LENGTH 'pre-5.0 structure size
   End Select

End Sub
