VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CINIFileAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
'Same function, but Vars as ANY:
Private Declare Function WritePrivateProfileStringAny Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private ms_FileName As String
Private ms_Section As String
Private ms_Key As String
Private ms_DefaultValue As String

Public Function EnumSections() As String()
Dim v As String, idx As Integer

v = GetValue(FileName, vbNullString, vbNullString, "")

'Remove Trailing Null Char
If Right$(v, 1) = Chr$(0) Then v = Left$(v, Len(v) - 1)

EnumSections = Split(v, Chr$(0))


End Function

Public Function RemoveSection() As Boolean

     RemoveSection = WritePrivateProfileStringAny(Section, 0&, 0&, FileName)

End Function

Public Property Get DefaultValue() As String
       DefaultValue = ms_DefaultValue
End Property

Public Property Let DefaultValue(ByVal sNewValue As String)
       ms_DefaultValue = sNewValue
End Property

Public Function EnumKeys() As String()
Dim v As String, idx As Integer

v = GetValue(FileName, Section, vbNullString, "")

'Remove Trailing Null Char
If Right$(v, 1) = Chr$(0) Then v = Left$(v, Len(v) - 1)

EnumKeys = Split(v, Chr$(0))


End Function
Public Function RemoveKey() As Boolean

     RemoveKey = WritePrivateProfileStringAny(Section, Key, 0&, FileName)

End Function

Public Property Get Value() As String
       
       Value = GetValue(FileName, Section, Key, DefaultValue)
       
End Property

Public Property Let Value(ByVal sNewValue As String)
       
       SetValue FileName, Section, Key, sNewValue
       
End Property

Public Property Get Key() As String
       Key = ms_Key
End Property

Public Property Let Key(ByVal sNewValue As String)
       ms_Key = sNewValue
End Property

Public Property Get Section() As String
       Section = ms_Section
End Property

Public Property Let Section(ByVal sNewValue As String)
       ms_Section = sNewValue
End Property

Public Property Get FileName() As String
       FileName = ms_FileName
End Property

Public Property Let FileName(ByVal sNewValue As String)
       ms_FileName = sNewValue
End Property

Public Function SetValue(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean

SetValue = CBool(WritePrivateProfileString(Section, Key, Value, FileName))

End Function

Public Function GetValue(ByVal FileName As String, ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As String) As String
Dim Buffer As String '* 4096
Dim Length As Long

Buffer = Space$(4096)

Length = GetPrivateProfileString( _
        Section, Key, DefaultValue, Buffer, Len(Buffer), FileName)

GetValue = Left$(Buffer, Length)

End Function

