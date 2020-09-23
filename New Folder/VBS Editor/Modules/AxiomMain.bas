Attribute VB_Name = "AxiomMain"
''''''''''''''[ Program Execution Starts Here ]''''''''''''''
Option Explicit

Public Const APP_TITLE = "Axiom HTML"

Public Type TagInfo
    Name As String
    IsSingle As Boolean
    Description As String
End Type

Public gs_XVBScript As String

' Check the "Property" Procs to see how Properties
' are superior to Public Vars
Private ms_CurrentDir As String
Private ms_CurrentFile As String

Private mb_InputIsDirty As Boolean
Private mb_OutputIsDirty As Boolean
Private mb_PadIsDirty As Boolean


Public Enum WhichTextbox
    All_TextBoxes = 0
    Input_TextBox = 1
    Output_TextBox = 2
    Pad_TextBox = 3
    
End Enum

Public Enum OutputFormat
    Format_Other = 0
    Format_Html = 1
    Format_Text = 2
End Enum

Public gs_FindWhat As String
Public gl_Options As Long
Public gl_Pos As Long

Public go_CustomCommands As CCustom
'Public AxiomSettings As CAxiomSettings
'Public go_HTMLTags As CHTMLTags
'Public go_MRU As CMRUList
'Public go_PlugIns As CPlugIns
Private mb_OpenedAsReadOnly As Boolean

Public gEditorSpacesPerTab As Integer

Public gOutputFormat As OutputFormat

Public Type ColorVals
    ColorName As String
    ColorLong As Long
    ColorHex As String
End Type

Public ga_ColorVals(0 To 15) As ColorVals

Public Sub InitColors()

'ga_ColorVals().ColorName=
'ga_ColorVals().ColorHex=
'ga_ColorVals().ColorLong=

ga_ColorVals(1).ColorName = "Black"
ga_ColorVals(1).ColorHex = "#000000"
ga_ColorVals(1).ColorLong = &H0&

ga_ColorVals(12).ColorName = "Silver"
ga_ColorVals(12).ColorHex = "#C0C0C0"
ga_ColorVals(12).ColorLong = &HC0C0C0

ga_ColorVals(4).ColorName = "Gray"
ga_ColorVals(4).ColorHex = "#808080"
ga_ColorVals(4).ColorLong = &H808080

ga_ColorVals(14).ColorName = "White"
ga_ColorVals(14).ColorHex = "#FFFFFF"
ga_ColorVals(14).ColorLong = &HFFFFFF

ga_ColorVals(7).ColorName = "Maroon"
ga_ColorVals(7).ColorHex = "#800000"
ga_ColorVals(7).ColorLong = &H80&

ga_ColorVals(11).ColorName = "Red"
ga_ColorVals(11).ColorHex = "#FF0000"
ga_ColorVals(11).ColorLong = &HFF&

ga_ColorVals(10).ColorName = "Purple"
ga_ColorVals(10).ColorHex = "#800080"
ga_ColorVals(10).ColorLong = &H800080

ga_ColorVals(3).ColorName = "Fuchsia"
ga_ColorVals(3).ColorHex = "#FF00FF "
ga_ColorVals(3).ColorLong = &HFF00FF

ga_ColorVals(5).ColorName = "Green"
ga_ColorVals(5).ColorHex = "#008000"
ga_ColorVals(5).ColorLong = &H8000&

ga_ColorVals(6).ColorName = "Lime"
ga_ColorVals(6).ColorHex = "#00FF00"
ga_ColorVals(6).ColorLong = &HFF00&

ga_ColorVals(9).ColorName = "Olive"
ga_ColorVals(9).ColorHex = "#808000"
ga_ColorVals(9).ColorLong = &H8080&

ga_ColorVals(15).ColorName = "Yellow"
ga_ColorVals(15).ColorHex = "#FFFF00"
ga_ColorVals(15).ColorLong = &HFFFF&

ga_ColorVals(8).ColorName = "Navy"
ga_ColorVals(8).ColorHex = "#000080"
ga_ColorVals(8).ColorLong = &H800000

ga_ColorVals(2).ColorName = "Blue"
ga_ColorVals(2).ColorHex = "#0000FF"
ga_ColorVals(2).ColorLong = &HFF0000

ga_ColorVals(13).ColorName = "Teal"
ga_ColorVals(13).ColorHex = "#008080"
ga_ColorVals(13).ColorLong = &H808000

ga_ColorVals(0).ColorName = "Aqua"
ga_ColorVals(0).ColorHex = "#00FFFF"
ga_ColorVals(0).ColorLong = &HFFFF00


End Sub
Public Function LoadScriptTemplate()
'if file not found, no error will be raised, an empty str will be returned
Dim sTemp As String
Dim F As CTextFile

Set F = New CTextFile
F.FileOpen RemoveSlash(App.Path) & "\Scripts\Template.vbs", OpenForInput
sTemp = F.ReadAll
F.FileClose
Set F = Nothing

'sTemp = Process_VBS_Xtensions(sTemp)

LoadScriptTemplate = sTemp

End Function
Public Function LoadXVBScript()
'if file not found, no error will be raised, an empty str will be returned
Dim sTemp As String
Dim F As CTextFile

Set F = New CTextFile
F.FileOpen RemoveSlash(App.Path) & "\Scripts\Autoload.vbs", OpenForInput
sTemp = F.ReadAll
F.FileClose
Set F = Nothing

sTemp = Process_VBS_Xtensions(sTemp)

LoadXVBScript = sTemp

End Function
Public Property Get OpenedAsReadOnly() As Boolean
       OpenedAsReadOnly = mb_OpenedAsReadOnly
End Property

Public Property Let OpenedAsReadOnly(ByVal bNewValue As Boolean)
       mb_OpenedAsReadOnly = bNewValue
End Property

Public Property Get CurrentFile() As String
       CurrentFile = ms_CurrentFile
End Property

Public Property Let CurrentFile(ByVal sNewValue As String)
       ms_CurrentFile = sNewValue
End Property

Public Sub Main() '<-----------------[Program Execution Starts Here]
Dim CmdLineArg  As String
Dim Result As VbMsgBoxResult


InitCommonControls  'WinXP specific

InitEntityInfo

gs_XVBScript = LoadXVBScript()

Set go_CustomCommands = New CCustom
'Set go_HTMLTags = New CHTMLTags
'Set go_MRU = New CMRUList
'Set AxiomSettings = New CAxiomSettings
'Set go_PlugIns = New CPlugIns

go_CustomCommands.Load
'go_PlugIns.PlugInsFolder = (RemoveSlash(App.Path) & "\PlugIns")
'go_HTMLTags.LoadFromFile (RemoveSlash(App.Path) & "\html_tags.ini")

gEditorSpacesPerTab = 8

'AxiomSettings.LoadSettings

'Load frmAxiomHTMLMain
'Load frmHTMLTags 'must be loaded to be hidden later?
'Load frmHidden

'// Handle Command Line
CmdLineArg = UnQuote(Command)   'Win Me specific?

If FileExists(CmdLineArg) Then
        CmdLineArg = GetLongFileName(CmdLineArg) 'Win 98 /95?
        
        If InStr(1, CmdLineArg, "?") > 0 Then
            MsgBox "Cannot Open  " & CmdLineArg & vbCrLf & "Filename and/or pathname contain invalid or unsupported chars", vbCritical, "Error"
        Else

'                    frmAxiomHTMLMain.MainText.OpenFile CmdLineArg
'                    frmAxiomHTMLMain.MainText.EnableCRLF = True
            
'                    frmAxiomHTMLMain.cdlg.InitDir = ExtractDirName(CmdLineArg)
                    CurrentDir = CmdLineArg
                    CurrentFile = CmdLineArg
'                    frmAxiomHTMLMain.Caption = ExtractFileName(CurrentFile) & " - Axiom"
'                    go_MRU.Add CmdLineArg
'                    frmAxiomHTMLMain.UpdateMRU
                    ''''''''''''''''''''''''''FIX INVALID NEW-LINE CHARS
            '        If AxiomSettings.CheckInvalidNewLine = True Then
            '            If RX_Check_Invalid_NewLine(frmAxiomHTMLMain.MainText.Text) Then
            '                Result = MsgBox("File has invalid New-Line chars, Fix them now?" & vbCrLf _
            '                                & "If this is a Binary File choose NO, else choose YES" _
            '                              , vbYesNo + vbQuestion, "Message")
            '                If Result = vbYes Then
            '                    frmAxiomHTMLMain.MainText.Text = FixNewLineChars(frmAxiomHTMLMain.MainText.Text)
            '                    MainText.Modified = True
            '                End If
            '            End If
            '        End If
                '''''''''''''''''''''
                If (GetAttr(CmdLineArg) And vbReadOnly) Then
                    OpenedAsReadOnly() = True
'                    frmAxiomHTMLMain.Status.Panels.Item("state").Text = "Read-Only"
                Else
                    OpenedAsReadOnly() = False
'                    frmAxiomHTMLMain.Status.Panels.Item("state").Text = ""
                End If
        End If

End If

'frmAxiomHTMLMain.Show

'frmVBScript.AutoLoad  'should we reset frmVBScript.IsDirty here???

End Sub
Public Property Get CurrentDir() As String
    
    CurrentDir = ms_CurrentDir
    
End Property

Public Property Let CurrentDir(ByVal sNewValue As String)
    
    ms_CurrentDir = RemoveSlash(ExtractDirName(sNewValue))
    
End Property
