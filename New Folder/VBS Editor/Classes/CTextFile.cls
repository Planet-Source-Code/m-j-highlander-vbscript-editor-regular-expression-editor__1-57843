VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Enum FileOpenMode

        OpenForInput = 0
        OpenForOutput = 1
        OpenForAppend = 2
        
End Enum

Private m_iFileHandle As Integer
Private m_sFileName As String


Public Function CreateTempFile() As String
'Create a temp file and open for output
'Return_Value: Created file name (can be used in deleting it)


Dim sBaseName As String
Dim sTempDir As String
Dim sTempFile As String
Dim iRand As Integer

Randomize
iRand = CInt(Rnd * 9999)
sBaseName = "~ctxtmp@" & Format(iRand, "0000")
sTempDir = Environ("Temp")

If Right(sTempDir, 1) <> "\" Then
        sTempDir = sTempDir & "\"
End If
sTempFile = sTempDir & sBaseName & ".tmp"
sTempFile = LCase(Trim(sTempFile))

'Open file for output
FileOpen sTempFile, OpenForOutput

'Return()
CreateTempFile = sTempFile

End Function

Public Property Get FileName() As String
        FileName = LCase$(m_sFileName)
End Property

Private Property Let FileName(ByVal sNewValue As String)
'This is a Private method, making "FileName" Read-Only outside Class

        m_sFileName = LCase$(sNewValue)
        
End Property

Public Function FileOpen(sFileName As String, OpenMode As FileOpenMode) As Boolean
Dim iFileNum As Integer

On Error GoTo Error_FileOpen

iFileNum = FreeFile

Select Case OpenMode
        Case OpenForInput
                Open sFileName For Input As iFileNum
        Case OpenForOutput
                Open sFileName For Output As iFileNum
        Case OpenForAppend
                Open sFileName For Append As iFileNum
        Case Else
        'dunno!
End Select

'Assign to properties
FileHandle = iFileNum
FileName = sFileName
FileOpen = True

Exit Function
Error_FileOpen:
    Err.Clear
    FileOpen = False
End Function
Public Function FileClose()
        
        Close FileHandle
        ' Invalidate File Handle and File Name
        FileHandle = vbEmpty
        FileName = vbNullString
        
End Function

Public Sub Flush()
'Force writing to file without closing it

Dim lOpenMode As Long
Dim sTmpName As String

'Get current open mode
lOpenMode = FileAttr(FileHandle, 1)

'Save name
sTmpName = FileName

'Close, thus writing contents
FileClose

'ReOpen:
'NOTE that this might change FILE HANDLE
'this is no problem since it should never be saved in a var

Select Case lOpenMode
        Case 2, 8 '2=output,8=append
                ' BUT we open for_append anyway
                ' coz we want to resume at end of file
                FileOpen sTmpName, OpenForAppend
        '('''''''''Case 32=binary --> CBinFile!!!)
        Case Else
        'flush is invalid for Open_For_Input mode
End Select

        
End Sub

Public Function ReadAll() As String
Dim iTemp As Integer
Dim bArray() As Byte

On Error GoTo ReadAll_Error
If LOF(FileHandle) = 0 Then
    ReadAll = ""
    Exit Function
End If

ReDim bArray(0 To LOF(FileHandle) - 1) As Byte

Close FileHandle  ' CLOSE FOR NOW

iTemp = FreeFile  ' OPEN IN BINARY MODE
Open FileName For Binary Access Read As iTemp
bArray = Input$(LOF(iTemp), iTemp)
ReadAll = bArray
Close #iTemp
FileOpen FileName, OpenForInput  'REOPEN IN TEXT MODE (and SEEK TO START)

Exit Function
ReadAll_Error:
    Err.Clear
    ReadAll = ""

End Function
Public Function ReadLnArray(Optional ByVal IgnoreEmptyLines As Boolean = False) As String()
Dim sTmp As String
Dim idx As Long
Dim sTmpArray() As String
ReDim sTmpArray(0 To 999) As String

idx = 0
Do While Not VBA.EOF(FileHandle)
        Line Input #FileHandle, sTmp
        If IgnoreEmptyLines And sTmp = "" Then
            'do nothing, don't add empty lines
        Else ' not an empty line or IgnoreEmptyLines=False
            sTmpArray(idx) = sTmp
            idx = idx + 1
        End If
        
        If idx > UBound(sTmpArray) Then
                'redim at 1000's for max speed
                ReDim Preserve sTmpArray(0 To UBound(sTmpArray) + 1000)
        End If
Loop
' (-1) last idx=idx+1 is not used
ReDim Preserve sTmpArray(0 To idx - 1)
ReadLnArray = sTmpArray

End Function
Public Property Get EOF() As Boolean
' To avoide ambeguity and/or clash we use VBA.EOF()

If VBA.EOF(FileHandle) Then
        EOF = True
Else
        EOF = False
End If

End Property
Public Function ReadLn() As String
Dim sTmp As String

Line Input #FileHandle, sTmp
ReadLn = sTmp

End Function




Public Sub WriteLnStr(Optional ByVal sStrToWrite As String)

        Print #FileHandle, sStrToWrite

End Sub

Public Sub WriteStr(ByVal sStrToWrite As String)

        Print #FileHandle, sStrToWrite;

End Sub


Public Property Get FileHandle() As Integer
        FileHandle = m_iFileHandle
End Property

Private Property Let FileHandle(ByVal iNewValue As Integer)
'This is a Private method, making "FileHandle" Read-Only outside Class

        m_iFileHandle = iNewValue
        
End Property
