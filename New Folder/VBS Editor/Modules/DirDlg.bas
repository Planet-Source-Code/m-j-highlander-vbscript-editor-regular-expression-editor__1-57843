Attribute VB_Name = "BrowseDirectorysOnly"
Option Explicit
'=====================================================================================
' Usage:
'
'    Dim folder As String
'    folder = BrowseForFolder(Me, "Select A Directory", "C:\startdir\anywhere")
'    If Len(folder) = 0 Then Exit Sub  'User Selected Cancel
'
'=====================================================================================
Private Const BIF_STATUSTEXT = &H4&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
  hwndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type

Private m_CurrentDirectory As String   'The current directory
Public Function BrowseForFolder(owner As Form, Title As String, StartDir As String) As String


  Dim lpIDList As Long
  Dim szTitle As String
  Dim sBuffer As String
  Dim tBrowseInfo As BrowseInfo
  m_CurrentDirectory = StartDir & vbNullChar

  szTitle = Title
  With tBrowseInfo
    .hwndOwner = owner.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
    .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
  End With

  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    BrowseForFolder = sBuffer
  Else
    BrowseForFolder = ""
  End If
  
End Function
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
  
  Dim lpIDList As Long
  Dim ret As Long
  Dim sBuffer As String
  
  On Error Resume Next  'Sugested by MS to prevent an error from
                        'propagating back into the calling process.
     
  Select Case uMsg
  
    Case BFFM_INITIALIZED
      Call SendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
      
    Case BFFM_SELCHANGED
      sBuffer = Space(MAX_PATH)
      
      ret = SHGetPathFromIDList(lp, sBuffer)
      If ret = 1 Then
        Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
      End If
      
  End Select
  
  BrowseCallbackProc = 0
  
End Function

' This function allows you to assign a function pointer to a variable.
Private Function GetAddressofFunction(Add As Long) As Long
  GetAddressofFunction = Add
End Function
