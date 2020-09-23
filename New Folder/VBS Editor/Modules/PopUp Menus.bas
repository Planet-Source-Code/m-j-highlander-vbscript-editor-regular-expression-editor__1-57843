Attribute VB_Name = "PopUpMenu"
Option Explicit

Private Type POINT
    X As Long
    y As Long
End Type

Private Const MF_ENABLED = &H0&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const TPM_RIGHTBUTTON = &H2&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_NONOTIFY = &H80&
Private Const TPM_RETURNCMD = &H100&
Private Const MF_BITMAP = &H4&

Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal sCaption As String) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, nIgnored As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Sub SetPopUpMenuIcon(hwndMenu As Long, SubIndex As Long, pic As Picture)
Dim hID As Long

'Get the menuId of the specified menu item
hID = GetMenuItemID(hwndMenu, SubIndex)

'Add the bitmap
SetMenuItemBitmaps hwndMenu, hID, MF_BITMAP, pic, pic
 
End Sub

Public Function PopUpEx(ByVal PixForm As Form, ByVal FunctionName As String, ParamArray param()) As Long
Dim iMenu As Long, hMenu As Long, nMenus As Long
Dim p As POINT

'Get the current cursor pos in screen coordinates
GetCursorPos p

'Create an empty popup menu
hMenu = CreatePopupMenu()

'Determine number of strings in paramarray
nMenus = 1 + UBound(param)

'Put each string in the menu
For iMenu = 1 To nMenus
' the AppendMenu function has been superseeded by the InsertMenuItem
' function, but it is a bit easier to use.
    If Trim$(CStr(param(iMenu - 1))) = "-" Then
        'if the parameter is a single dash, a separator is drawn
        AppendMenu hMenu, MF_SEPARATOR, iMenu, ""
    Else
        AppendMenu hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(param(iMenu - 1))
    End If
Next iMenu


If FunctionName <> "" Then
    CallByName PixForm, FunctionName, VbMethod, hMenu
End If

' Show the menu at the current cursor location;
' the flags make the menu aligned to the right (!); enable the right button to select
' an item; prohibit the menu from sending messages and make it return the index of
' the selected item.
' the TrackPopupMenu function returns when the user selected a menu item or cancelled
' the window handle used here may be any window handle from your application
' the return value is the (1-based) index of the menu item or 0 in case of cancelling
iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.y, 0, GetForegroundWindow(), 0)

' Release and destroy the menu
DestroyMenu hMenu

' Return the selected menu item's index
PopUpEx = iMenu

End Function
Public Function PopUp(ParamArray param()) As Long
Dim iMenu As Long, hMenu As Long, nMenus As Long
Dim p As POINT

'Get the current cursor pos in screen coordinates
GetCursorPos p

'Create an empty popup menu
hMenu = CreatePopupMenu()

'Determine number of strings in paramarray
nMenus = 1 + UBound(param)

'Put each string in the menu
For iMenu = 1 To nMenus
' the AppendMenu function has been superseeded by the InsertMenuItem
' function, but it is a bit easier to use.
    If Trim$(CStr(param(iMenu - 1))) = "-" Then
        'if the parameter is a single dash, a separator is drawn
        AppendMenu hMenu, MF_SEPARATOR, iMenu, ""
    Else
        AppendMenu hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(param(iMenu - 1))
    End If
Next iMenu




' Show the menu at the current cursor location;
' the flags make the menu aligned to the right (!); enable the right button to select
' an item; prohibit the menu from sending messages and make it return the index of
' the selected item.
' the TrackPopupMenu function returns when the user selected a menu item or cancelled
' the window handle used here may be any window handle from your application
' the return value is the (1-based) index of the menu item or 0 in case of cancelling
iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.y, 0, GetForegroundWindow(), 0)

' Release and destroy the menu
DestroyMenu hMenu

' Return the selected menu item's index
PopUp = iMenu

End Function
Public Function PopUpImg(ByVal ImageList As ImageList, ParamArray param()) As Long
Dim iMenu As Long, hMenu As Long, nMenus As Long
Dim p As POINT

'Get the current cursor pos in screen coordinates
GetCursorPos p

'Create an empty popup menu
hMenu = CreatePopupMenu()

'Determine number of strings in paramarray
nMenus = 1 + UBound(param)

'Put each string in the menu
For iMenu = 1 To nMenus
' the AppendMenu function has been superseeded by the InsertMenuItem
' function, but it is a bit easier to use.
    If Trim$(CStr(param(iMenu - 1))) = "-" Then
        'if the parameter is a single dash, a separator is drawn
        AppendMenu hMenu, MF_SEPARATOR, iMenu, ""
    Else
        AppendMenu hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(param(iMenu - 1))
        SetPopUpMenuIcon hMenu, iMenu - 1, ImageList.ListImages(iMenu).Picture
    End If
Next iMenu




' Show the menu at the current cursor location;
' the flags make the menu aligned to the right (!); enable the right button to select
' an item; prohibit the menu from sending messages and make it return the index of
' the selected item.
' the TrackPopupMenu function returns when the user selected a menu item or cancelled
' the window handle used here may be any window handle from your application
' the return value is the (1-based) index of the menu item or 0 in case of cancelling
iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.y, 0, GetForegroundWindow(), 0)

' Release and destroy the menu
DestroyMenu hMenu

' Return the selected menu item's index
PopUpImg = iMenu

End Function
