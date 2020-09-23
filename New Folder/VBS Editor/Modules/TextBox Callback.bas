Attribute VB_Name = "TextBox_Callback"
Option Explicit

Public OldTextBoxProc As Long
Public OldAboutTextBoxProc As Long


Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4)
Public Const WM_CONTEXTMENU = &H7B


Public Function NewAboutTextBoxProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg <> WM_CONTEXTMENU Then _
        NewAboutTextBoxProc = CallWindowProc( _
            OldAboutTextBoxProc, hWnd, Msg, wParam, _
            lParam)
'Add in form_load():
'    OldAboutTextBoxProc  = SetWindowLong( _
'       Text1.hWnd, GWL_WNDPROC, _
'        AddressOf NewAboutTextBoxProc)
End Function

' *********************************************
' Pass along all messages except the one that
' makes the context menu appear.
' *********************************************
Public Function NewTextBoxProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg <> WM_CONTEXTMENU Then _
        NewTextBoxProc = CallWindowProc( _
            OldTextBoxProc, hWnd, Msg, wParam, _
            lParam)
'Add in form_load():
'    OldTextBoxProc  = SetWindowLong( _
'       Text1.hWnd, GWL_WNDPROC, _
'        AddressOf NewTextBoxProc)
End Function

