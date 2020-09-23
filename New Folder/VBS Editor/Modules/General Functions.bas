Attribute VB_Name = "General_Functions"
Option Explicit

Private Const ATTR_READONLY = 1    'Read-Only file
Private Const ATTR_VOLUME = 8      'Volume label
Private Const ATTR_ARCHIVE = 32    'File has changed since last back-up
Private Const ATTR_NORMAL = 0      'Normal files
Private Const ATTR_HIDDEN = 2      'Hidden files
Private Const ATTR_SYSTEM = 4      'System files
Private Const ATTR_DIRECTORY = 16  'Directory

Private Const ATTR_DIR_ALL = ATTR_DIRECTORY + ATTR_READONLY + ATTR_ARCHIVE + ATTR_HIDDEN + ATTR_SYSTEM
Private Const ATTR_ALL_FILES = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM Or ATTR_READONLY Or ATTR_ARCHIVE

Public Function RunScriptEx(ByVal ScriptCode As String, ByVal InputText As String) As String
Dim ScriptCtl As MSScriptControl.ScriptControl

Dim idx As Integer, pos As Integer, bFound As Boolean
Dim sResult As String

'On Error GoTo RunScriptEx_Error

Set ScriptCtl = New MSScriptControl.ScriptControl
ScriptCtl.Language = "VBScript"

If Trim$(ScriptCode) = "" Then
    RunScriptEx = ""
    Exit Function
End If

If InStr(LCase$(ScriptCode), "function") = 0 Then
    MsgBox "Cannot find any Functions, immediate mode is not supported", vbCritical, "Script Error"
    RunScriptEx = ""
    Exit Function
End If


ScriptCtl.Reset

'ADD OBJECTS AND CLASSES
'-----------------------
'Add object only, properties not public:
'scr1.AddObject "Clipboard", Clipboard, False

'Add functions as public!
'this MUST be added first. to OVER-RIDE some functions from CAxiomFunction
ScriptCtl.AddObject "CVBScriptEx", New CVBScriptEx, True                     'Extra Functions
ScriptCtl.AddObject "CAxiom", New CAxiomFunction, True                       'Macro Functions
'ScriptCtl.AddObject "CHTML_Tag_Functions", New CHTML_Tag_Functions, True     'HTML Tag Functions

ScriptCode = Process_VBS_Xtensions(ScriptCode)

ScriptCtl.AddCode ScriptCode & vbCrLf & gs_XVBScript

bFound = False
pos = 0
For idx = 1 To ScriptCtl.Procedures.Count
    If LCase$(ScriptCtl.Procedures(idx).Name) = "main" Then
        pos = idx
        bFound = True
        Exit For
    End If
Next

If bFound Then
    If ScriptCtl.Procedures(pos).NumArgs = 0 Then
        sResult = ScriptCtl.Run("Main")
    Else
        sResult = ScriptCtl.Run("Main", InputText)
    End If

ElseIf ScriptCtl.Procedures.Count > 0 Then
    If ScriptCtl.Procedures(1).NumArgs = 0 Then
        sResult = ScriptCtl.Run(ScriptCtl.Procedures(1).Name)
    Else
        sResult = ScriptCtl.Run(ScriptCtl.Procedures(1).Name, InputText)
    End If
    
Else
    'no Procs: DO NOTHING! immediate mode not allowed.
End If

RunScriptEx = sResult

Exit Function
RunScriptEx_Error:
    MsgBox "VBScript Error " & ScriptCtl.Error.Number & vbCrLf & ScriptCtl.Error.Description, vbCritical, "Oops"
    ScriptCtl.Error.Clear

End Function
Public Function ForceRename(ByVal sSrcFile As String, ByVal sTgtFile As String) As Boolean

On Error Resume Next

If FileExists(sTgtFile) Then Kill sTgtFile
Name sSrcFile As sTgtFile
ForceRename = True

If Err Then
    Err.Clear
    ForceRename = False
End If

End Function
Public Function AddStrIfNotExist(ByVal Text As String, ByVal StrToAdd As String) As String

If Right(Text, Len(StrToAdd)) <> StrToAdd Then

    AddStrIfNotExist = Text & StrToAdd
Else

    AddStrIfNotExist = Text
End If


End Function
Public Function ConvertToNum(ByVal StrNum As String) As Long

'Allow commas in numbers
If StrNum = "" Then StrNum = "0"
ConvertToNum = CLng(Replace(StrNum, ",", ""))

End Function
Public Function CreateHTMLIndex(ByRef InFiles() As String, Optional ByVal Target As String = "") As String
Dim idx As Long, sTarget As String

If Target = "" Then
    sTarget = ""
Else
    sTarget = "target=""" & Target & """"

End If


For idx = LBound(InFiles) To UBound(InFiles)
    InFiles(idx) = "<a href=""" & Replace(InFiles(idx), "\", "/") & """ " & sTarget & ">" & ExtractFileName(InFiles(idx)) & "</a><br>"
Next

CreateHTMLIndex = Join(InFiles, vbCrLf)

End Function
Public Function EnQuote(ByVal Text As String) As String

    EnQuote = Chr(34) & Text & Chr(34)

End Function
Public Function GetDirName(ByVal FileName As String) As String
'Extract the Directory name from a full file name
Dim sTemp As String
Dim iPos As Integer


iPos = InStrRev(FileName, "\")
If iPos = 0 Then
        
        GetDirName = ""

ElseIf DirExists(FileName) Then
        'it is already a Dir name!
        GetDirName = FileName
        If Right(sTemp, 1) <> "\" Then sTemp = sTemp & "\"

Else
        
        sTemp = Left(FileName, iPos)
        If Right(sTemp, 1) <> "\" Then sTemp = sTemp & "\"
        GetDirName = sTemp

End If

End Function
Public Sub CenterFormUp(ByVal frmX As Form)

    frmX.Left = (Screen.Width - frmX.Width) / 2
    frmX.Top = (Screen.Height - frmX.Height) / 3

End Sub
Function GetScriptFileContent(ByVal FileName As String) As String
Dim sTemp As String
Dim File As New CTextFile


If File.FileOpen(FileName, OpenForInput) Then
        sTemp = File.ReadAll
        File.FileClose

'not found, try looking in Scripts subfolder:
ElseIf File.FileOpen(RemoveSlash(App.Path) & "\Scripts\" & FileName, OpenForInput) Then
        sTemp = File.ReadAll
        File.FileClose
Else
        'Couldn't load file
        sTemp = ""
End If

GetScriptFileContent = sTemp
Set File = Nothing

End Function
Public Function EscapeSnippetChars(ByVal sText As String) As String

sText = Replace$(sText, "~", "~~")
sText = Replace$(sText, vbTab, "~t")
sText = Replace$(sText, vbCrLf, "~n")

EscapeSnippetChars = sText

End Function
Public Function Process_VBS_Xtensions(ByVal ScriptText As String) As String

'Line-continuation char, works for strings too!
ScriptText = Replace(ScriptText, "~" & vbCrLf, "")

'handle escape chars
If RX_Test(ScriptText, "^[ \t]*?option[ \t]*?useescapes[ \t]*?\r\n", True) Then
    ScriptText = RX_GenericReplace(ScriptText, "^[ \t]*?option[ \t]*?useescapes[ \t]*?\r\n", vbCrLf, True)
    ScriptText = RX_EscapeQuotedContents(ScriptText)
Else
    ' Option "UseEsapes" is not present, default behaviour
    ' ScriptText = ScriptText
End If

'Error-Handling Statements:
ScriptText = RX_GenericReplace(ScriptText, "^ *?err\.supress *?\r\n", "On Error Resume Next" & vbCrLf, True)
ScriptText = RX_GenericReplace(ScriptText, "^ *?err\.allow *?\r\n", "On Error Goto 0" & vbCrLf, True)


'convert every call to Write() to a call to Writes(), because Write is reserved in VB6 but not in VBScript
ScriptText = RX_ReplaceWriteFunction(ScriptText)

'convert "FOR index IN array" syntax to "FOR index=0 TO LBOUND(array)"
ScriptText = RX_ReplaceForInArray(ScriptText)

' Handle Multi-Line Comments ( C-Style /* .... */ )
''''''''''''''''''ScriptText = RX_GenericReplace(ScriptText, "/\*[^\v]*?\*/", "", False)

' Add Included file contents ( C-Style #INCLUDE )
ScriptText = RX_ReplaceIncludes(ScriptText)

' Handle array creation syntax  ( a= [...] )
ScriptText = RX_ExpandArrays(ScriptText)

Process_VBS_Xtensions = ScriptText

End Function
Public Function RemoveFromPath(ByVal sPath As String, ByVal iLevels As Integer) As String
Dim iPos As Integer, idx As Integer, sTemp As String

iPos = 0

For idx = 1 To iLevels
    iPos = InStr(iPos + 1, sPath, "\")
    If iPos = 0 Then Exit For   'no need to go on , besides it will roll back to the beginning
Next

'in case we asked for more levels than there's available, now we take as much as exists:
If iPos = 0 Then iPos = InStrRev(sPath, "\")

If iPos > 0 Then

    RemoveFromPath = Right(sPath, Len(sPath) - iPos)

Else
    
    RemoveFromPath = sPath

End If


End Function
Public Function UnEscapeSnippetChars(ByVal sText As String) As String

sText = Replace$(sText, "~~", Chr$(7))

sText = Replace$(sText, "~t", vbTab)
sText = Replace$(sText, "~n", vbCrLf)

sText = Replace$(sText, Chr$(7), "~")

UnEscapeSnippetChars = sText

End Function
Public Function Min(ByVal ValA As Variant, ByVal ValB As Variant) As Variant

Min = IIF(ValA < ValB, ValA, ValB)

End Function

Public Function RunScript(ByVal ScriptCode As String, ByVal InputText As String) As String

Dim sTemp As String
Dim frmX As Form

sTemp = InputText
                                 ' make a "copy"
Set frmX = New frmVBScript       ' in order not to alter frmVBScript

'Load frmX      'implied, no need to do it explicitly

frmX.DontColorize = True
frmX.Tag = sTemp
frmX.Visible = False
frmX.txtScript.Text = ScriptCode
frmX.btnRun.Value = True  ' fire the Click() event
sTemp = frmX.Tag

Unload frmX
Set frmX = Nothing

If sTemp = "CANCEL" Then sTemp = ""

RunScript = sTemp

End Function
Public Function Max(ByVal ValA As Variant, ByVal ValB As Variant) As Variant

Max = IIF(ValA > ValB, ValA, ValB)

End Function
Public Function AddToFileName(ByVal FileName As String, ByVal StrAdd As String) As String
Dim sExt As String
Dim sBaseName As String

sExt = ExtractFileExtension(FileName)
sBaseName = Left$(FileName, Len(FileName) - Len(sExt))
If Right$(sBaseName, 1) = "." Then sBaseName$ = Left$(sBaseName, Len(sBaseName) - 1)

If sExt <> "" Then
    AddToFileName = sBaseName & StrAdd & "." & sExt
Else
    AddToFileName = sBaseName & StrAdd
End If

End Function
Public Function ChangeFileExtension(ByVal FileName As String, ByVal NewExtension As String) As String
Dim sOldExt As String, sBaseName As String

NewExtension = Replace$(NewExtension, ".", "")

sOldExt = ExtractFileExtension(FileName)

sBaseName = Left$(FileName, Len(FileName) - Len(sOldExt))
If Right$(sBaseName, 1) <> "." Then sBaseName = sBaseName & "."

If (NewExtension = "" And Right$(sBaseName, 1) = ".") Then
    sBaseName = Left$(sBaseName, Len(sBaseName) - 1)
End If

ChangeFileExtension = sBaseName & NewExtension

End Function
Public Function ExtractFileExtension(ByVal FileName As String) As String

Dim pos As Integer


pos = InStrRev(FileName, ".")

If pos = 0 Then
    ExtractFileExtension = ""

Else

    ExtractFileExtension = Right$(FileName, Len(FileName) - pos)

End If

End Function
Public Function CreatePath(ByVal Path As String) As Boolean
On Error Resume Next
    
Dim v As Variant
Dim idx As Integer
Dim sFolder As String
Dim lower As Integer, upper As Integer

If Right$(Path, 1) = "\" Then Path = Left$(Path, Len(Path) - 1)
v = Split(Path, "\")
If IsArray(v) Then
    lower = LBound(v)
    upper = UBound(v)
    sFolder = v(lower) & "\" & v(lower + 1) ' drive + first folder

    MkDir sFolder
    For idx = lower + 2 To upper
        sFolder = sFolder & "\" & v(idx)
        MkDir sFolder
    Next
End If

If DirExists(sFolder) Then
    CreatePath = True
Else
    CreatePath = False
End If

End Function

Public Function GetLongFileName(ByVal ShortFileName As String) As String

Dim intPos As Integer
Dim strLongFileName As String
Dim strDirName As String

'Format the filename for later processing
ShortFileName = ShortFileName & "\"

'Grab the position of the first real slash
intPos = InStr(4, ShortFileName, "\")

'Loop round all the directories and files
'in ShortFileName, grabbing the full names
'of everything within it.

While intPos

    strDirName = Dir(Left(ShortFileName, intPos - 1), _
        vbNormal + vbHidden + vbSystem + vbDirectory)
    
    If strDirName = "" Then
        GetLongFileName = ""
        Exit Function
    End If
    
    strLongFileName = strLongFileName & "\" & strDirName
    intPos = InStr(intPos + 1, ShortFileName, "\")
    
Wend

'Return the completed long file name
GetLongFileName = Left(ShortFileName, 2) & strLongFileName
  
End Function

Function GetArgs(ByVal Func As String) As Variant
' Func is of the form   FuncName(Arg1,Arg2,...)

If InStr(Func, "(") = 0 Then
    GetArgs = ""
    Exit Function
Else

    Func = Trim(Mid(Func, InStr(Func, "(") + 1, InStr(Func, ")") - InStr(Func, "(") - 1))
    GetArgs = Split(Func, ",")
End If

End Function
Public Function CharsToDec(ByVal Chars As String) As String
Dim idx As Long
Dim bChars() As Byte
Dim sTemp() As String

If Chars = "" Then Exit Function

ReDim bChars(0 To Len(Chars) - 1)
ReDim sTemp(0 To Len(Chars) - 1)

bChars = StrConv(Chars, vbFromUnicode) ' VB Strings are Double-Byte Unicode


For idx = 0 To Len(Chars) - 1
            sTemp(idx) = Format$(bChars(idx)) ', "000")
Next idx

CharsToDec = Join(sTemp, " ")


End Function

Public Function CharsToHex(ByVal Chars As String) As String
Dim idx As Long
Dim bChars() As Byte
Dim sTemp() As String

If Chars = "" Then Exit Function

ReDim bChars(0 To Len(Chars) - 1)
ReDim sTemp(0 To Len(Chars) - 1)

bChars = StrConv(Chars, vbFromUnicode) ' VB Strings are Double-Byte Unicode


For idx = 0 To Len(Chars) - 1
            sTemp(idx) = Right$("0" & Hex$(bChars(idx)), 2)
Next idx

CharsToHex = Join(sTemp, " ")


End Function
Public Function ExtractFileName(ByVal FilePath As String) As String

' Extract the File name from a full file path

Dim iLastSlash As Integer

iLastSlash = InStrRev(FilePath, "\")

If iLastSlash = 0 Then
        ExtractFileName = FilePath
Else
    ExtractFileName = Right$(FilePath, Len(FilePath) - iLastSlash)
End If


End Function

Function RemoveSlash(ByVal sPath As String) As String

sPath = Trim$(sPath)

If Right$(sPath, 1) = "\" Then
    RemoveSlash = Left(sPath, Len(sPath) - 1)
Else
    RemoveSlash = sPath
End If


End Function

Function DirExists(ByVal DirName As String) As Boolean
Dim tmp As String
Dim iResult As Integer

If Trim$(DirName) = "" Then
            DirExists = False
            Exit Function
End If

iResult = 0
If Dir$(DirName, ATTR_DIR_ALL) <> "" Then
    iResult = GetAttr(DirName) And ATTR_DIRECTORY
End If

If iResult = 0 Then   'Directory not found, or the passed argument is a filename.
    DirExists = False
Else
    DirExists = True
End If

End Function

Function ExtractDirName(ByVal FileName As String) As String
'THIS FUNCTION IS FOR KILL!!! REPLACED BY THE SUPIRIOR FUNCTION GetDirName()

    Dim tmp$
    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
        ExtractDirName = ""
        Exit Function
    End If
    
    Do While pos <> 0
        PrevPos = pos
        pos = InStr(pos + 1, FileName, "\")
    Loop

    tmp = Left(FileName, PrevPos)
    If Right(tmp, 1) = "\" Then tmp = Left(tmp, Len(tmp) - 1)
    ExtractDirName = tmp
    
End Function
Function FileExists(ByVal FileName As String) As Boolean
On Error GoTo DirError

If Trim$(FileName) = "" Then
            FileExists = False
            Exit Function
End If

If Dir$(FileName, ATTR_ALL_FILES) = "" Then
    FileExists = False
Else
    FileExists = True
End If

Exit Function
DirError:
    Err.Clear
    FileExists = False
End Function
