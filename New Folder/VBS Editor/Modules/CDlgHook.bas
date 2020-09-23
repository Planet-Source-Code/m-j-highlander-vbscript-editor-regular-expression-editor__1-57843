Attribute VB_Name = "CDlgHook"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

' Hook and notification support:
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    Code As Long
End Type

Private Type OFNOTIFYshort
    hdr As NMHDR
    lpOFN As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2

Private Const WM_INITDIALOG = &H110
Private Const WM_USER = &H400
Private Const WM_DESTROY = &H2
Private Const WM_NOTIFY = &H4E

Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
Private Const CDM_HIDECONTROL = (CDM_FIRST + &H5)
Private Const CDM_GETSPEC = (CDM_FIRST + &H0)
Private Const CDM_GETFILEPATH = (CDM_FIRST + &H1)
Private Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)

Private Const H_MAX As Long = &HFFFF + 1
Private Const CDN_FIRST = (H_MAX - 601)
Private Const CDN_SELCHANGE = (CDN_FIRST - &H1)

Private Const ID_OPEN = &H1  'Open or Save button
Private Const ID_CANCEL = &H2 'Cancel Button
Private Const ID_HELP = &H40E 'Help Button
'Private Const ID_READONLY = &H410 'Read-only check box
'Private Const ID_FILETYPELABEL = &H441 'FileType label
'Private Const ID_FILELABEL = &H442 'FileName label
'Private Const ID_FOLDERLABEL = &H443 'Folder label
'Private Const ID_LIST = &H461 'Parent of file list
'Private Const ID_FORMAT = &H470 'FileType combo box
'Private Const ID_FOLDER = &H471 'Folder combo box
Private Const ID_FILETEXT = &H480 'FileName text box

Private Const SW_NORMAL = 1
Private Const WM_PASTE = &H302

Public cdlgPic As New CdlgEx

Public rtb As Control
Public pix As Control

Public TwipsInHimetric As Single
Public DlgHwnd As Long
Dim sPath As String


Public Function lHookAddress(lPtr As Long) As Long
  lHookAddress = lPtr
End Function

Public Function DialogHookFunction(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Select Case wMsg
          Case WM_INITDIALOG
               MoveDlg hDlg
               ModifyDlg hDlg
          Case WM_NOTIFY
               Dim tOFNs As OFNOTIFYshort
               Dim tNMH As NMHDR
               CopyMemory tNMH, ByVal lParam, Len(tNMH)
               If tNMH.Code = CDN_SELCHANGE Then
                 ' Changed selected file:
                  Dim ret As Integer, sTmp As String
                  sTmp = String(259, 0)
                  ret = SendMessage(GetParent(hDlg), CDM_GETFILEPATH, 260, ByVal sTmp)
                  If ret > 1 Then
                     sPath = Left$(sTmp, ret - 1)
                     ViewFile sPath
                  End If
               End If
          Case WM_DESTROY
               ' Here you can add user's notification
               ' before exiting
               Set rtb = Nothing
          Case Else
   End Select
End Function

Private Sub MoveDlg(hDlg)
  Dim rct As RECT
  If cdlgPic.Left = 0 And cdlgPic.Top = 0 Then Exit Sub
  DlgHwnd = GetParent(hDlg)
  GetWindowRect DlgHwnd, rct
  MoveWindow DlgHwnd, cdlgPic.Left, cdlgPic.Top, rct.Right - rct.Left, rct.Bottom - rct.Top, 1
End Sub

Private Sub ModifyDlg(hDlg)
  Dim sClass As String
  Dim h As Long, i As Long, k As Long
  Dim rc As RECT, pt As POINTAPI, bDone As Boolean
  Dim tEdge As Long, rEdge As Long, rct As RECT
  DlgHwnd = GetParent(hDlg)
  If cdlgPic.OKText <> "" Then SendMessage DlgHwnd, CDM_SETCONTROLTEXT, ID_OPEN, ByVal cdlgPic.OKText
  If cdlgPic.CancelText <> "" Then SendMessage DlgHwnd, CDM_SETCONTROLTEXT, ID_CANCEL, ByVal cdlgPic.CancelText
  If cdlgPic.HelpText <> "" Then SendMessage DlgHwnd, CDM_SETCONTROLTEXT, ID_HELP, ByVal cdlgPic.HelpText
  h = GetWindow(DlgHwnd, GW_CHILD)
  Do
    sClass = Space$(128)
    k = GetClassName(h, ByVal sClass, 128)
    sClass = Left$(sClass, k)
    If sClass = "ComboBox" And Not bDone Then
       bDone = True
       GetWindowRect h, rc
       rEdge = rc.Right - 1
    End If
    If bDone Then
       If sClass = "ListBox" Then
          GetWindowRect h, rc
          pt.x = rc.Left
          pt.Y = rc.Top
          tEdge = rc.Top
          ScreenToClient DlgHwnd, pt
          MoveWindow h, pt.x, pt.Y, rEdge - rc.Left, rc.Bottom - rc.Top, 1
       End If
    End If
    h = GetWindow(h, GW_HWNDNEXT)
  Loop While h <> 0
  
'MsgBox SetParent(rtb.hWnd, DlgHwnd)
SetParent rtb.hwnd, DlgHwnd

  rtb.Visible = True
  pt.x = rEdge + 3
  pt.Y = tEdge - 1
  ScreenToClient DlgHwnd, pt
  GetWindowRect DlgHwnd, rct
  MoveWindow rtb.hwnd, pt.x, pt.Y, rct.Right - rEdge - 8, rc.Bottom - rc.Top + 2, 1
'MsgBox MoveWindow(rtb.hWnd, pt.X, pt.Y, rct.Right - rEdge - 8, rc.Bottom - rc.Top + 2, 1)
End Sub

Private Sub ViewFile(sFile As String)
  
   If Dir$(sFile, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Sub
   If (GetAttr(sFile) And vbDirectory) = vbDirectory Then Exit Sub
'   On Error GoTo ErrHandler
   Select Case Right$(sFile, 4)
          Case ".bmp", ".jpg", ".gif"
               Set rtb.Picture = LoadPicture(sPath)
          Case Else
                'do nothing
   End Select
  ' rtb.Refresh
   
   Exit Sub
ErrHandler:
    MsgBox Err.Description
   On Error GoTo 0

End Sub
