Attribute VB_Name = "File_PropertiesDlg"
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

Private Const MAX_PATH = 260


Private Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hwnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    lpIDList      As Long     'Optional
    lpClass       As String   'Optional
    hkeyClass     As Long     'Optional
    dwHotKey      As Long     'Optional
    hIcon         As Long     'Optional
    hProcess      As Long     'Optional
End Type

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Declare Function ShellExecuteEx Lib "shell32" _
   Alias "ShellExecuteExA" (SEI As SHELLEXECUTEINFO) As Long

Public Function TrimNull(ByVal sItem As String) As String
   
  'Return a string without the chr$(0) terminator.
   Dim pos As Integer

   pos = InStr(sItem, Chr$(0))
   
   If pos Then
        TrimNull = Left$(sItem, pos - 1)
   Else
        TrimNull = sItem
   End If

End Function
Public Sub ShowFileProperties(ByVal FileName As String)

Dim SEI As SHELLEXECUTEINFO
 
If FileName = "" Then Exit Sub
 
With SEI
   .cbSize = Len(SEI)
   .fMask = SEE_MASK_NOCLOSEPROCESS Or _
            SEE_MASK_INVOKEIDLIST Or _
            SEE_MASK_FLAG_NO_UI
   .hwnd = 0  'parent form hwnd?
   .lpVerb = "properties"
   .lpFile = FileName
   .lpParameters = vbNullChar
   .lpDirectory = vbNullChar
   .nShow = 0
   .hInstApp = 0
   .lpIDList = 0
End With
 
Call ShellExecuteEx(SEI)
    
End Sub

