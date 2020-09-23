VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBScriptEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Functions added to VBScript

Option Explicit

Public OutStr As String

Public FSO As Scripting.FileSystemObject

Public GUI As Object

Public Function CreateFile(ByVal FileName As String, Optional ByVal FileSize As Long = 0) As Boolean
Dim ff As Integer
Dim bArray() As Byte
'On Error GoTo CreateFile_Error

ff = FreeFile
Open FileName For Binary Access Write As ff

If FileSize > 0 Then
    ReDim bArray(1 To FileSize)
    Put #ff, , bArray
End If

Close #ff

CreateFile = True

Exit Function
CreateFile_Error:
    Err.Clear
    CreateFile = False

End Function
Public Function SplitEx(ByVal Expression As String, _
                        Optional ByVal Delimiter As String = " ", _
                        Optional ByVal Limit As Long = -1, _
                        Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant

    SplitEx = Text_Functions.SplitEx(Expression, Delimiter, Limit, Compare)

End Function
Public Function TrimWhiteSpace(ByVal Text As String) As String

    TrimWhiteSpace = RegExp_Functions.TrimWhiteSpace(Text)

End Function

Public Function RegExpExtract0(ByVal Text As String, ByVal Pattern As String) As String
    
    Dim sArray() As String

    sArray = RegExp_Functions.RX_GenericExtractToArray(Text, Pattern)
    RegExpExtract0 = Join(sArray, Chr(0))


End Function
Public Function Writes(Optional ByVal sText As String = "") As String

OutStr = OutStr & sText

End Function

Public Function vbCounter(Optional ByVal Start As Variant, Optional ByVal Step As Variant) As Long
Static cntr As Long
Static IncStep As Long

If IsMissing(Step) And IsMissing(Start) Then
    'do nothing
ElseIf IsMissing(Step) Then
    IncStep = 1
Else
    IncStep = Step
End If

If IncStep = 0 Then IncStep = 1

If IsMissing(Start) Then
    cntr = cntr + IncStep
Else
    cntr = Start - IncStep
End If


vbCounter = cntr

End Function
Public Function Chrs(ParamArray vValues() As Variant) As String
'sample call: Chrs(65,66,67) --> "ABC"

Dim idx As Long

If UBound(vValues) > -1 Then

    For idx = LBound(vValues) To UBound(vValues)
            vValues(idx) = Chr(vValues(idx))
    Next
    
    Chrs = Join(vValues, "")

Else

    Chrs = ""

End If

End Function
Public Function stringf(ByVal Text As String, ParamArray vValues() As Variant) As String
'sample call:  stringf("coco is %s but %s. %s","good,naughty,wow!")

Dim idx As Long
Const SPECIAL_CHAR = "%"

If UBound(vValues) > -1 Then

    'Hide %%, only needed to insert %s (without it being replaced)
    Text = Replace(Text, SPECIAL_CHAR & SPECIAL_CHAR, Chr(7))

    For idx = LBound(vValues) To UBound(vValues)
            Text = Replace(Text, SPECIAL_CHAR & "s", vValues(idx), 1, 1, vbTextCompare)
    Next

    'Remove redundant %s
    Text = Replace(Text, SPECIAL_CHAR & "s", "")

    'Unhide %% and turn to %
    Text = Replace(Text, Chr(7), SPECIAL_CHAR)
    stringf = Text

Else

    stringf = Text

End If

End Function
Public Function Sort(ByRef Lines As Variant, Optional ByVal Ascending As Boolean = True, Optional ByVal IgnoreCase As Boolean = True) As Variant
Dim sTempArray As Variant
Dim sLines() As String
Dim idx As Long, lower As Long, upper As Long


' Although 'Lines' is passed ByRef, it is not changed.

' If Not IsArray(Lines) Then: execution will stop and an error will be raised automatically!

lower = LBound(Lines)
upper = UBound(Lines)

ReDim sLines(lower To upper)
ReDim sTempArray(lower To upper)

For idx = lower To upper
    sLines(idx) = Lines(idx)
Next

Sort_String.StrSort sLines, Ascending, IgnoreCase


For idx = lower To upper
    sTempArray(idx) = sLines(idx)
Next

Sort = sTempArray

End Function
Public Function StrJoin(ParamArray vStrings() As Variant) As String
Dim sDelimiter As String
Dim sTemp As String


If UBound(vStrings) > -1 Then
    
    If UBound(vStrings) = 0 Then
        StrJoin = vStrings(0)
    Else
        sDelimiter = vStrings(UBound(vStrings))
        vStrings(UBound(vStrings)) = ""
        sTemp = Join(vStrings, sDelimiter)
        StrJoin = Left(sTemp, Len(sTemp) - Len(sDelimiter))
    End If

Else

    StrJoin = ""

End If

End Function
Public Function IIF(ByVal Condition As Boolean, ByVal TruePart As Variant, ByVal FalsePart As Variant) As Variant

'If Condition = True Then
'    IIF = TruePart
'Else
'    IIF = FalsePart
'End If

IIF = VBA.IIF(Condition, TruePart, FalsePart)

End Function
Public Function LeftTo(ByVal MainStr As String, ByVal SubStr As String)

LeftTo = Text_Functions.LeftTo(MainStr, SubStr)

End Function
Public Function MultiInstr(ByVal Start As Long, ByVal Text As String, ByVal LookFor As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Long

MultiInstr = Text_Functions.MultiInstr(Start, Text, LookFor, Compare)

End Function
Function MultiInstrRev(ByVal Start As Long, ByVal Text As String, ByVal LookFor As String, ByVal Compare As VbCompareMethod) As Long

MultiInstrRev = Text_Functions.MultiInstrRev(Start, Text, LookFor, Compare)

End Function
Public Function RightTo(ByVal MainStr As String, ByVal SubStr As String)

RightTo = Text_Functions.RightTo(MainStr, SubStr)

End Function

Public Function StrCat(ParamArray vStrings() As Variant) As String

StrCat = Join(vStrings, "")

End Function
Public Function StrFill(ByVal Count As Long, ByVal StrToFill As String) As String

StrFill = Text_Functions.StrFill(Count, StrToFill)

End Function
Public Function TextComp(ByVal Text1 As String, ByVal Text2 As String, Optional ByVal MatchCase As Boolean = False) As Boolean
Dim iCompareMethod As Integer, iResult As Integer

If MatchCase Then
    iCompareMethod = vbBinaryCompare
Else
    iCompareMethod = vbTextCompare
End If

iResult = StrComp(Text1, Text2, iCompareMethod)

If iResult = 0 Then
    TextComp = True
Else
    TextComp = False
End If

End Function
Public Function Copy(ByVal Text As String) As String

If Text <> "" Then

        Copy = Text
        Clipboard.Clear
        Clipboard.SetText Text, vbCFText
End If

End Function
Public Property Get CurrentDir() As String
    
    CurrentDir = AxiomMain.CurrentDir

' DIDN`T WORK!
'Public Property Let CurrentDir(ByVal sNewValue As String)
'    AxiomMain.CurrentDir = RemoveSlash(ExtractDirName(sNewValue))
'End Property

End Property
Public Function LoadFile(ByVal FileName As String) As String
'OVER-RIDE

Dim F As CTextFile


Set F = New CTextFile

F.FileOpen FileName, OpenForInput
LoadFile = F.ReadAll

F.FileClose

Set F = Nothing

End Function

Public Function MidStr(ByVal sString As String, ByVal sNewValue As String, ByVal lStart As Long, Optional ByVal lLength As Long = -1) As String
'Could be used in Macros, but mot implemented there!

Dim sTemp

sTemp = sString

If lLength = -1 Then
    Mid$(sTemp, lStart) = sNewValue
Else
    Mid$(sTemp, lStart, lLength) = sNewValue
End If

MidStr = sTemp

End Function
Public Function Format(ByVal Str As String, ByVal Formatting As String) As String
   
   Format = VBA.Format$(Str, Formatting)

End Function
Public Function AxiomPath() As String

AxiomPath = RemoveSlash(App.Path) & "\"

End Function

Public Function Paste() As String

If Clipboard.GetFormat(vbCFText) Then

        Paste = Clipboard.GetText(vbCFText)

End If

End Function
Public Function vbQuote() As String

    vbQuote = """"

End Function
Public Function Writeln(Optional ByVal sText As String = "") As String

OutStr = OutStr & sText & vbCrLf

End Function

Private Sub Class_Initialize()

Set FSO = New Scripting.FileSystemObject

On Error Resume Next
   Set GUI = CreateObject("Axiom.Forms")
If Err Then
   MsgBox "Cannot create object"
   Err.Clear
End If

End Sub
