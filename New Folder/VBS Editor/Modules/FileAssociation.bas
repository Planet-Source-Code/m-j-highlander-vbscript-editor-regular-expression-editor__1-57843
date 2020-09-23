Attribute VB_Name = "FileAssociation"
Option Explicit

' Modifications by M. Highlander <mdsy@ny.com>,<m.highlander@gmail.com>
' By Kim Pedersen, vcoders@get2net.dk
' Inspiration from VBnet at http://www.mvps.org/vbnet
'
' Usage Example:
' CreateAssociation("txt","C:\Windows\Npotepad", "Notepad" , "Notepad Text Editor", "C:\Windows\Notepad.exe,0")

' Notice that it creates entries in 2 places in the Registry (on WinXP that is, other Windows i dunno!)
' HKEY_CLASSES_ROOT
' HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0

Public Sub CreateAssociation(ByVal Extension As String, Optional ByVal AppPath As String, Optional ByVal AppName As String, Optional ByVal AppDesc As String, Optional ByVal IconLib As String)

'Full Application Path without the ".exe" Extension
If AppPath = "" Then AppPath = IIF(Right(App.Path, 1) = "\", App.Path & App.EXEName, App.Path & "\" & App.EXEName)

'Since we will add ".exe" later, remove it if it exists
If Right(LCase(AppPath), 4) = ".exe" Then AppPath = Left(AppPath, Len(AppPath) - 4)

If IconLib = "" Then IconLib = AppPath & ".exe,0"

AppPath = AppPath & ".exe ""%1"""

If AppName = "" Then AppName = App.EXEName

If AppDesc = "" Then AppDesc = AppName & " Application"

'just in case
Extension = Replace(Extension, ".", "")

CreateNewKey "." & Extension, HKEY_CLASSES_ROOT
SetKeyValue "." & Extension, "", AppName, REG_SZ

CreateNewKey AppName & "\shell\open\command", HKEY_CLASSES_ROOT
SetKeyValue AppName & "\shell\open\command", "", AppPath, REG_SZ

SetKeyValue AppName, "", AppDesc, REG_SZ

CreateNewKey AppName & "\DefaultIcon", HKEY_CLASSES_ROOT
SetKeyValue AppName & "\DefaultIcon", "", IconLib, REG_SZ

End Sub
Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)

'handle to the new key
Dim hKey As Long
'result of the RegCreateKeyEx function
Dim r As Long
r = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, r)
Call RegCloseKey(hKey)

End Sub
Private Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)

'result of the SetValueEx function
Dim r As Long
'handle of opened key
Dim hKey As Long
'open the specified key
r = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, KEY_ALL_ACCESS, hKey)
r = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
Call RegCloseKey(hKey)

End Sub
Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
Dim nValue As Long
Dim sValue As String

Select Case lType

    Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
    Case REG_DWORD
            nValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, nValue, 4)

End Select

End Function
