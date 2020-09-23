VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVBScript 
   Caption         =   "Edit/Execute Script"
   ClientHeight    =   7575
   ClientLeft      =   1080
   ClientTop       =   1890
   ClientWidth     =   10935
   Icon            =   "vbScriptControl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   729
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txtScript 
      Height          =   3735
      Left            =   720
      TabIndex        =   9
      Top             =   1380
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   6588
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      TextRTF         =   $"vbScriptControl.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton btnRun 
      Caption         =   "BTN_RUN"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   5460
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ImageList imglstScriptEditor 
      Left            =   4980
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":0619
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":083B
            Key             =   "view"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":0A5D
            Key             =   "hilite"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":0C7F
            Key             =   "unhilite"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":0EA1
            Key             =   "vbs"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":10C3
            Key             =   "save"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":12E5
            Key             =   "run"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":1507
            Key             =   "quick"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":1C15
            Key             =   "new"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":23D3
            Key             =   "comment"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":25F5
            Key             =   "vbshelp"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":2D03
            Key             =   "uncomment"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":2F25
            Key             =   "special"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":3147
            Key             =   "escape"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":3369
            Key             =   "saveas"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":4075
            Key             =   "indent"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":4297
            Key             =   "outdent"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":44B9
            Key             =   "regexp"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":46DB
            Key             =   "help"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbScriptControl.frx":4DE9
            Key             =   "axshelp"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrScriptEditor 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   4260
      Top             =   5460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picPlaceHolder 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   570
      Width           =   10935
      Begin AxiomHTML.ShadowedSeperator SGLine1 
         Height          =   30
         Left            =   0
         Top             =   0
         Width           =   22500
         _ExtentX        =   39688
         _ExtentY        =   53
      End
      Begin AxiomHTML.CoolButton btnEscapeText 
         Height          =   375
         Left            =   9840
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   90
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "Text"
         ShowFocusRect   =   0   'False
      End
      Begin AxiomHTML.CoolButton btnArray 
         Height          =   375
         Left            =   9420
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Convert Block of text to array"
         Top             =   90
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "A( )"
         ShowFocusRect   =   0   'False
      End
      Begin VB.ComboBox cboExternal 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5580
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   2940
      End
      Begin VB.ComboBox cboPresets 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   2940
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "External Functions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   4
         Top             =   180
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   180
         Width           =   510
      End
   End
   Begin MSScriptControlCtl.ScriptControl scr1 
      Left            =   3420
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   2700
      Picture         =   "vbScriptControl.frx":54F7
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuEscapes 
      Caption         =   "mnuEscapes"
      Visible         =   0   'False
      Begin VB.Menu mnuBackslash 
         Caption         =   "Backslash"
      End
      Begin VB.Menu mnuCrLf 
         Caption         =   "Newline (CRLF)"
      End
      Begin VB.Menu mnuTab 
         Caption         =   "Tab"
      End
      Begin VB.Menu mnuQuote 
         Caption         =   "Quote"
      End
      Begin VB.Menu hyph 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDisable 
         Caption         =   "Disable Special Symbols"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmVBScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Property:
Public DontColorize As Boolean



Private xFlag As Boolean
Private IsDirty As Boolean
Private bInMargin As Boolean
Private msFindWhat As String
Private mlSearchPos As Long

Public Sub AutoLoad()
    
    Dim sFileName As String
    Dim oFile As New CTextFile

    sFileName = RemoveSlash(App.Path) & "\LatestScript.txt"
    
    oFile.FileOpen sFileName, OpenForInput
    frmVBScript.txtScript.Text = oFile.ReadAll
    oFile.FileClose

    Set oFile = Nothing
    
End Sub
Public Sub AutoSave()
    
    Dim sFileName As String
    Dim oFile As New CTextFile

    sFileName = RemoveSlash(App.Path) & "\LatestScript.txt"
    
    oFile.FileOpen sFileName, OpenForOutput
    oFile.WriteStr frmVBScript.txtScript.Text
    oFile.FileClose

    Set oFile = Nothing
    
End Sub
Private Function CheckIncludedFilename(ByVal FileName As String) As String
On Error GoTo CheckIncludedFilename_Error
Dim sFileName As String
Dim sTemp As String

sFileName = Trim$(FileName)

If Dir(sFileName) <> "" Then
        sTemp = RemoveSlash(CurDir) & "\" & sFileName
'not found or no path ? try looking in Scripts subfolder:
ElseIf Dir(RemoveSlash(App.Path) & "\Scripts\" & sFileName) <> "" Then
        sTemp = RemoveSlash(App.Path) & "\Scripts\" & sFileName
Else
        sTemp = ""
End If

If sFileName = "" Then sTemp = ""

CheckIncludedFilename = sTemp


Exit Function
CheckIncludedFilename_Error:
    MsgBox Err.Description, vbCritical, "Error: [CheckIncludedFilename_Error]"
    Err.Clear

End Function
Private Sub CreateToolBar()

With tbrScriptEditor

    Set .ImageList = imglstScriptEditor

    .Buttons.Add , "sep_1", , tbrSeparator

    .Buttons.Add , "new", , tbrDefault, "new"
    .Buttons("new").ToolTipText = "New"

    .Buttons.Add , "open", , tbrDefault, "open"
    .Buttons("open").ToolTipText = "Open"

    .Buttons.Add , "quick", , tbrDefault, "quick"
    .Buttons("quick").ToolTipText = "Quick Open"
    
    .Buttons.Add , "save", , tbrDefault, "save"
    .Buttons("save").ToolTipText = "Save"

    .Buttons.Add , "saveas", , tbrDefault, "saveas"
    .Buttons("saveas").ToolTipText = "Save As"

    .Buttons.Add , "sep_2", , tbrSeparator

    .Buttons.Add , "comment", , tbrDefault, "comment"
    .Buttons("comment").ToolTipText = "Comment"

    .Buttons.Add , "uncomment", , tbrDefault, "uncomment"
    .Buttons("uncomment").ToolTipText = "Uncomment"

    .Buttons.Add , "indent", , tbrDefault, "indent"
    .Buttons("indent").ToolTipText = "Indent"

    .Buttons.Add , "outdent", , tbrDefault, "outdent"
    .Buttons("outdent").ToolTipText = "Outdent"

    .Buttons.Add , "sep_3", , tbrSeparator

    .Buttons.Add , "hilite", , tbrDefault, "hilite"
    .Buttons("hilite").ToolTipText = "Highlight"

    .Buttons.Add , "unhilite", , tbrDefault, "unhilite"
    .Buttons("unhilite").ToolTipText = "Unhilite"

    .Buttons.Add , "sep_35", , tbrSeparator
    
    .Buttons.Add , "run", , tbrDefault, "run"
    .Buttons("run").ToolTipText = "Run [F5]"

    .Buttons.Add , "sep_4", , tbrSeparator

    .Buttons.Add , "vbs", , tbrDefault, "vbs"
    .Buttons("vbs").ToolTipText = "Show Standard VBScript Code"

    .Buttons.Add , "view", , tbrDefault, "view"
    .Buttons("view").ToolTipText = "View/Edit Included File"

'    .Buttons.Add , "array", , tbrDefault, "array"
'    .Buttons("array").ToolTipText = "Array"

'    .Buttons.Add , "sep_5", , tbrSeparator
    .Buttons.Add , "regexp", , tbrDefault, "regexp"
    .Buttons("regexp").ToolTipText = "Edit Pattern"

    .Buttons.Add , "escape", , tbrDefault, "escape"
    .Buttons("escape").ToolTipText = "Escape RegExp Special Chars:    \^$*+{}?.:=!|[]-(),"

    .Buttons.Add , "special", , tbrDropdown, "special"
    .Buttons("special").ToolTipText = "Insert Special Char"
    .Buttons("special").ButtonMenus.Add , "backslash", "Backslash" & vbTab & "\\"
    .Buttons("special").ButtonMenus.Add , "newline", "Newline (CR+LF)" & vbTab & "\n"
    .Buttons("special").ButtonMenus.Add , "tab", "Tab" & vbTab & "\t"
    .Buttons("special").ButtonMenus.Add , "quote", "Quote" & vbTab & "\q"


    .Buttons.Add , "sep_6", , tbrSeparator

    .Buttons.Add , "vbshelp", , tbrDefault, "vbshelp"
    .Buttons("vbshelp").ToolTipText = "VBScript Help [F1]"

    .Buttons.Add , "axshelp", , tbrDefault, "axshelp"
    .Buttons("axshelp").ToolTipText = "Axiom VBScript Help"

End With

End Sub
Private Function GetIncludeFilename() As String
Dim sCurFileName As String
Dim sCurFileTitle As String

cdlg1.InitDir = RemoveSlash(App.Path) & "\Scripts"
cdlg1.Filter = "VBScript Files|*.vbs|All Files|*.*"
cdlg1.DialogTitle = "Select Include File"
cdlg1.FilterIndex = 2
'save original filename
sCurFileName = cdlg1.FileName
cdlg1.FileName = ""

cdlg1.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly

On Error Resume Next
    cdlg1.ShowOpen
If Err Then Err.Clear
On Error GoTo 0

If ExtractDirName(cdlg1.FileName) = RemoveSlash(App.Path) & "\Scripts" Then
    GetIncludeFilename = ExtractFileName(cdlg1.FileName)
Else
    GetIncludeFilename = cdlg1.FileName
End If

'restore original filename
cdlg1.FileName = sCurFileName

End Function
Private Sub btnClose_Click()
    Me.Tag = "CANCEL"
    Me.Hide

End Sub

Private Sub btnArray_Click()
Dim sTemp As String, sArray As String
Dim vArray As Variant
Dim idx As Long

sTemp = txtScript.SelText

If sTemp = "" Then Exit Sub

'Escape Quotes
sTemp = Replace(sTemp, Chr(34), Chr(34) & Chr(34))

sArray = InputBox("Type a name for the Array", "Data Required", "vArray")
    
If sArray = "" Then Exit Sub

vArray = Split(sTemp, vbCrLf)
For idx = LBound(vArray) To UBound(vArray)
    vArray(idx) = sArray & "(" & CStr(idx) & ")=" & Chr(34) & vArray(idx) & Chr(34)
Next

sTemp = Join(vArray, vbCrLf)

sTemp = "Dim " & sArray & "(" & CStr(UBound(vArray)) & ")" & vbCrLf & sTemp
txtScript.SelText = sTemp

End Sub
Private Sub btnComment_Click()
Dim sTemp As String
Dim vTemp As Variant
Dim idx As Long

sTemp = txtScript.SelText
If sTemp <> "" Then
    vTemp = Split(sTemp, vbCrLf)
        For idx = LBound(vTemp) To UBound(vTemp)
            vTemp(idx) = "'" & vTemp(idx)
        Next idx
    sTemp = Join(vTemp, vbCrLf)
    txtScript.SelText = sTemp
End If

SyntaxColorize txtScript ', QBColor(2), RGB(165, 105, 195), vbBlue

txtScript.SetFocus

End Sub
Private Sub btnEscape_Click()

If txtScript.SelLength > 0 Then
    txtScript.SelText = EscapeRegExpChars(txtScript.SelText)
End If

txtScript.SetFocus

End Sub

Private Sub btnEscapeText_Click()
Dim sTemp As String, sConst As String
Dim vArray As Variant
Dim idx As Long

sTemp = txtScript.SelText

If sTemp = "" Then Exit Sub

'Escape Quotes
sTemp = Replace(sTemp, Chr(34), Chr(34) & Chr(34))

sConst = InputBox("Type a name for the string constant", "Data Required", "sTemp")
    
If sConst = "" Then Exit Sub

vArray = Split(sTemp, vbCrLf)
For idx = LBound(vArray) To UBound(vArray) - 1
    vArray(idx) = Chr(34) & vArray(idx) & Chr(34) & " & vbCrLf & _"
Next

sTemp = Join(vArray, vbCrLf) & vArray(UBound(vArray)) & Chr(34) & vbCrLf

sTemp = "Dim " & sConst & vbCrLf & sConst & " = " & sTemp

txtScript.SelText = sTemp

End Sub

Private Sub btnHelp_Click()

HHelp_Show RemoveSlash(App.Path) & "\Vbscript.chm", "/html/VBSTOC.htm"

End Sub

Private Sub btnHelp2_Click()

HHelp_Show RemoveSlash(App.Path) & "\Axiom VBScript.chm", "Axiom_VBS.htm"

End Sub
Private Sub btnIndent_Click()
Dim sTemp As String
Dim vTemp As Variant
Dim idx As Long

sTemp = txtScript.SelText
If sTemp <> "" Then
    vTemp = Split(sTemp, vbCrLf)
        For idx = LBound(vTemp) To UBound(vTemp)
            vTemp(idx) = Space$(gEditorSpacesPerTab) & vTemp(idx)
        Next idx
    sTemp = Join(vTemp, vbCrLf)
    txtScript.SelText = sTemp
End If
SyntaxColorize txtScript ', QBColor(2), RGB(165, 105, 195), vbBlue
txtScript.SetFocus

End Sub

Private Sub btnMenu_MouseDown()

'PopUpMenu mnuEscapes, 2, btnMenu.Left, btnMenu.Top + btnMenu.Height
'txtScript.SetFocus

End Sub
Private Sub btnNew_Click()
Dim Result As VbMsgBoxResult

If IsDirty Then
    Result = MsgBox("Contents have changed, Continue?", vbOKCancel + vbDefaultButton2 + vbQuestion)
    If Result = vbCancel Then Exit Sub
End If

Me.Caption = "Edit/Execute Script"
cdlg1.FileName = ""

SyntaxColorizeAll txtScript, LoadScriptTemplate() ', QBColor(2), RGB(165, 105, 195), vbBlue
txtScript.SelColor = vbBlack

IsDirty = False

txtScript.SetFocus

End Sub
Private Sub btnOpen_Click()
Dim TextFile As New CTextFile
Dim Result As VbMsgBoxResult
Dim sCurrentFilename As String
Dim sFileContents As String

If IsDirty Then
    Result = MsgBox("Contents have changed, Continue?", vbOKCancel + vbDefaultButton2 + vbQuestion)
    If Result = vbCancel Then Exit Sub
End If

sCurrentFilename = cdlg1.FileName

cdlg1.InitDir = RemoveSlash(App.Path) & "\Scripts"
cdlg1.Filter = "VBScript Files|*.vbs|All Files|*.*"
cdlg1.DialogTitle = "Open"
cdlg1.FileName = ""
cdlg1.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly

On Error Resume Next
    cdlg1.ShowOpen
    If Err Then          'user clicked "Cancel"
        Err.Clear
        cdlg1.FileName = sCurrentFilename
        txtScript.SetFocus
        Exit Sub
    End If
On Error GoTo 0

TextFile.FileOpen cdlg1.FileName, OpenForInput
sFileContents = TextFile.ReadAll
TextFile.FileClose

'dont change to .FileTitle, i know what i'm doing!
Me.Caption = "Edit/Execute Script - [" & ExtractFileName(cdlg1.FileName) & "]"
IsDirty = False

Set TextFile = Nothing

'txtScript.Text = sFileContents
SyntaxColorizeAll txtScript, sFileContents ', QBColor(2), RGB(165, 105, 195), vbBlue

IsDirty = False

txtScript.SetFocus

End Sub
Private Sub btnOutdent_Click()
Dim sTemp As String
Dim vTemp As Variant
Dim idx As Long

sTemp = txtScript.SelText
If sTemp <> "" Then
    vTemp = Split(sTemp, vbCrLf)
        For idx = LBound(vTemp) To UBound(vTemp)
            If Left(vTemp(idx), gEditorSpacesPerTab) = Space$(gEditorSpacesPerTab) Then
                vTemp(idx) = Right(vTemp(idx), Len(vTemp(idx)) - gEditorSpacesPerTab)
            ElseIf Left(vTemp(idx), 1) = Space$(1) Then
                vTemp(idx) = LTrim$(vTemp(idx))
            End If
        Next idx
    sTemp = Join(vTemp, vbCrLf)
    txtScript.SelText = sTemp
End If

SyntaxColorize txtScript ', QBColor(2), RGB(165, 105, 195), vbBlue
txtScript.SetFocus

End Sub

Private Sub btnRegExpEdit_Click()
Dim frmX As frmRegExpGenerator
Dim sResult  As String


Set frmX = New frmRegExpGenerator

frmX.Value = txtScript.SelText
frmX.Show vbModal

sResult = frmX.Value
If sResult <> "" And frmX.Canceled = False Then
    txtScript.SelText = sResult
End If

Unload frmX
Set frmX = Nothing

txtScript.SetFocus

End Sub
Private Sub btnRun_Click()
Dim idx As Integer, pos As Integer, bFound As Boolean
Dim sResult As String, sTextArg As String
Dim ScriptText As String

On Error GoTo ScriptError

ScriptText = txtScript.Text
If Trim$(ScriptText) = "" Then
    Beep
    txtScript.SetFocus
    Exit Sub
End If

If InStr(LCase$(ScriptText), "function") = 0 Then
    Beep
    MsgBox "Cannot find any Functions, immediate mode is not supported", vbCritical, "Script Error"
    txtScript.SetFocus
    Exit Sub
End If


sTextArg = Me.Tag
scr1.Reset

'ADD OBJECTS AND CLASSES
'-----------------------
'Add object only, properties not public:
'scr1.AddObject "Clipboard", Clipboard, False

'Add functions as public!
'this MUST be added first. to OVER-RIDE some functions from CAxiomFunction
scr1.AddObject "CVBScriptEx", New CVBScriptEx, True                     'Extra Functions
scr1.AddObject "CAxiom", New CAxiomFunction, True                       'Macro Functions
'scr1.AddObject "CForm", New frmScriptObject, True                       'InputForm
'scr1.AddObject "CInputForm", New frmInputText, True                     'InputText
'scr1.AddObject "COptionForm", New frmAxOptionForm, True                 'OptionForm
'scr1.AddObject "CHTML_Tag_Functions", New CHTML_Tag_Functions, True     'HTML Tag Functions

ScriptText = Process_VBS_Xtensions(ScriptText)

scr1.AddCode ScriptText & vbCrLf & gs_XVBScript

bFound = False
pos = 0
For idx = 1 To scr1.Procedures.Count
    If LCase$(scr1.Procedures(idx).Name) = "main" Then
        pos = idx
        bFound = True
        Exit For
    End If
Next

If bFound Then
    If scr1.Procedures(pos).NumArgs = 0 Then
        sResult = scr1.Run("Main")
    Else
        sResult = scr1.Run("Main", sTextArg)
    End If

ElseIf scr1.Procedures.Count > 0 Then
    If scr1.Procedures(1).NumArgs = 0 Then
        sResult = scr1.Run(scr1.Procedures(1).Name)
    Else
        sResult = scr1.Run(scr1.Procedures(1).Name, sTextArg)
    End If
    
Else
    'no Procs: DO NOTHING! immediate mode not allowed.
End If

Me.Tag = sResult
MsgBox sResult
'If Me.Visible Then Me.Hide

Exit Sub
ScriptError:
    MsgBox "VBScript Error: " & scr1.Error.Number & vbCrLf & "Description: " & scr1.Error.Description, vbCritical, "Oops"
    SelText txtScript, scr1.Error.Line, scr1.Error.Line
    scr1.Error.Clear
    Err.Clear
    If frmVBScript.Visible Then txtScript.SetFocus
    
End Sub
Private Sub btnSaveAs_Click()
Dim TextFile As New CTextFile

cdlg1.Filter = "VBScript Files|*.vbs|All Files (*.*)|*.*"
cdlg1.InitDir = RemoveSlash(App.Path) & "\Scripts"
cdlg1.DialogTitle = "Save As"
cdlg1.flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly

On Error Resume Next
    cdlg1.ShowSave
If Err Then Exit Sub
On Error GoTo 0

TextFile.FileOpen cdlg1.FileName, OpenForOutput
TextFile.WriteStr txtScript.Text
TextFile.FileClose

Me.Caption = "Edit/Execute Script - [" & ExtractFileName(cdlg1.FileName) & "]"
IsDirty = False

Set TextFile = Nothing

txtScript.SetFocus

End Sub
Private Sub btnSave_Click()
Dim TextFile As New CTextFile

If cdlg1.FileName = "" Then
    cdlg1.Filter = "VBScript Files|*.vbs|All Files (*.*)|*.*"
    cdlg1.InitDir = RemoveSlash(App.Path) & "\Scripts"
    cdlg1.DialogTitle = "Save"
    cdlg1.flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly

    On Error Resume Next
        cdlg1.ShowSave
    If Err Then Exit Sub
    On Error GoTo 0
End If


TextFile.FileOpen cdlg1.FileName, OpenForOutput
TextFile.WriteStr txtScript.Text
TextFile.FileClose

Me.Caption = "Edit/Execute Script - [" & ExtractFileName(cdlg1.FileName) & "]"
IsDirty = False

Set TextFile = Nothing

txtScript.SetFocus

End Sub
Private Sub btnUnComment_Click()

Dim sTemp As String
Dim vTemp As Variant
Dim idx As Long

sTemp = txtScript.SelText
If sTemp <> "" Then
    vTemp = Split(sTemp, vbCrLf)
        For idx = LBound(vTemp) To UBound(vTemp)
            If Left(vTemp(idx), 1) = "'" Then
                vTemp(idx) = Right(vTemp(idx), Len(vTemp(idx)) - 1)
            End If
        Next idx
    sTemp = Join(vTemp, vbCrLf)
    txtScript.SelText = sTemp
End If

SyntaxColorize txtScript ', QBColor(2), RGB(165, 105, 195), vbBlue
txtScript.SetFocus

End Sub
Private Sub btnViewIncluded_Click()
Dim sFileName As String
Dim sTemp As String, iStart As Integer, iEnd As Integer
Dim iSelStart As Integer

sTemp = txtScript.Text

If sTemp = "" Then txtScript.SetFocus: Exit Sub

If txtScript.SelText = "" Then

        iSelStart = txtScript.SelStart
        If iSelStart = 0 Then iSelStart = 1
        
        iEnd = InStr(iSelStart, sTemp, vbCrLf)
        If iEnd = 0 Then iEnd = 1
        iStart = InStrRev(sTemp, vbCrLf, iSelStart)
        If iStart = 0 Then iStart = 1
        
        sTemp = Mid(sTemp, iStart, iEnd - iStart)
        sFileName = CrLfTabTrim(Replace(sTemp, "#include", "", 1, -1, vbTextCompare))
Else

        sFileName = txtScript.SelText
End If
'MsgBox "*" & sFileName & "*"
sFileName = CheckIncludedFilename(sFileName)

If sFileName <> "" Then
    Shell "notepad.exe " & sFileName, vbNormalFocus
End If


End Sub
Private Sub btnViewVBScript_Click()
Dim sTemp As String
Dim ScriptText  As String

ScriptText = txtScript.Text
sTemp = Process_VBS_Xtensions(ScriptText)

frmViewText.View sTemp, "View Standard VBScript Code", True

End Sub

Private Sub btnUnderline_Click()

End Sub

Private Sub cboExternal_Click()

On Error Resume Next
    'if opened then closed without picking an item: cboExternal.Text=""
    txtScript.SelText = cboExternal.Text
    txtScript.SetFocus


End Sub

Private Sub cboPresets_Click()
Dim sTemp As String

Select Case cboPresets.Text
    Case "Function"
        sTemp = "Function " & vbCrLf & vbCrLf & "End Function" & vbCrLf & vbCrLf
    Case "Sub"
        sTemp = "Sub " & vbCrLf & vbCrLf & "End Sub" & vbCrLf & vbCrLf
    Case "If...Then...Else"
        sTemp = "If   Then" & vbCrLf & vbCrLf & "Else " & vbCrLf & vbCrLf & "End If" & vbCrLf & vbCrLf
    Case "For...Next"
        sTemp = "For  =  To " & vbCrLf & vbCrLf & "Next" & vbCrLf & vbCrLf
    Case "Do...Loop"
        sTemp = "Do " & vbCrLf & vbCrLf & "Loop" & vbCrLf & vbCrLf
    Case "InputBox"
        sTemp = " =InputBox("" "")" & vbCrLf
    Case "InputForm"
        sTemp = " =InputForm("" "","" "")" & vbCrLf
    Case "MsgBox"
        sTemp = "MsgBox " & vbCrLf
    Case "#SPECIALSYMBOLS OFF"
        sTemp = "#SPECIALSYMBOLS OFF" & vbCrLf
    Case "#INCLUDE"
        sTemp = "#INCLUDE " & GetIncludeFilename() & vbCrLf
    
    Case Else

End Select

txtScript.SelText = sTemp

txtScript.SetFocus

End Sub
Private Sub AddXtraFunctions()
Dim idx As Integer

scr1.Reset
scr1.AddCode gs_XVBScript

For idx = 1 To scr1.Procedures.Count
    If LCase(scr1.Procedures(idx).Name) <> "axiomautoloadmain" Then
            cboExternal.AddItem scr1.Procedures(idx).Name
    End If
Next

scr1.Reset

End Sub

Private Sub Command1_Click()

End Sub

Private Sub CoolButton1_Click()

End Sub
Private Sub Form_Activate()
On Error Resume Next

    txtScript.SetFocus

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyA Then
    If Shift = vbCtrlMask Then
        txtScript.SelStart = 0
        txtScript.SelLength = Len(txtScript.Text)
        xFlag = True 'used in KeyPress to prevent the Beep
    End If

ElseIf KeyCode = vbKeyF5 Then
    btnRun_Click

ElseIf KeyCode = vbKeyF1 Then
    btnHelp_Click
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If xFlag = True Then 'KeyDown() Detected a CTRL+A
    KeyAscii = 0     'This is done to prevent the Beep
    xFlag = False
End If

End Sub


Private Sub Form_Load()
'AutoLoad 'load last edited script

CreateToolBar

IsDirty = False

mnuBackslash.Caption = "Backslash" & vbTab & "\\"
mnuCrLf.Caption = "Newline (CR+LF)" & vbTab & "\n"
mnuTab.Caption = "Tab" & vbTab & "\t"
mnuQuote.Caption = "Quote" & vbTab & "\q"

cboPresets.AddItem "Function"
cboPresets.AddItem "Sub"
cboPresets.AddItem "If...Then...Else"
cboPresets.AddItem "For...Next"
cboPresets.AddItem "Do...Loop"
cboPresets.AddItem "InputBox"
cboPresets.AddItem "MsgBox"
cboPresets.AddItem "_________________________________________________"
cboPresets.AddItem "#INCLUDE"
cboPresets.AddItem "#SPECIALSYMBOLS OFF"

AddXtraFunctions

'@  LeftMargin(txtScript) = 6
SetComboHeight Me, cboPresets, 15
SetComboHeight Me, cboExternal, 15

cboPresets.Top = (picPlaceHolder.ScaleHeight - cboPresets.Height) / 2
cboExternal.Top = (picPlaceHolder.ScaleHeight - cboExternal.Height) / 2

txtScript.SelIndent = 10       'Left Margin
SetWordWrap txtScript, False   'Disablw Word-wrap

If DontColorize = False Then
    'txtScript.Text = LoadScriptTemplate()
    SyntaxColorizeAll txtScript, LoadScriptTemplate() ', QBColor(2), RGB(165, 105, 195), vbBlue
End If

IsDirty = False

CenterFormUp Me

'GetScriptKeywords   'for colorize

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Result As VbMsgBoxResult

If IsDirty Then
    Result = MsgBox("Contents have changed, Continue?", vbOKCancel + vbDefaultButton2 + vbQuestion)
    If Result = vbCancel Then Cancel = True
End If



End Sub
Private Sub Form_Resize()

On Error Resume Next

txtScript.Top = tbrScriptEditor.Height + picPlaceHolder.Height
txtScript.Left = 0
txtScript.Width = Me.ScaleWidth
txtScript.Height = Me.ScaleHeight - tbrScriptEditor.Height - picPlaceHolder.Height

SGLine1.Left = 0
SGLine1.Width = Me.ScaleWidth


End Sub
Private Sub mnuBackslash_Click()
txtScript.SelText = "\\"
End Sub

Private Sub mnuCrLf_Click()
txtScript.SelText = "\n"
End Sub


Private Sub mnuDisable_Click()

mnuDisable.Checked = Not mnuDisable.Checked
If mnuDisable.Checked = True Then
    mnuCrLf.Enabled = False
    mnuTab.Enabled = False
    mnuQuote.Enabled = False
    mnuBackslash.Enabled = False
Else
    mnuCrLf.Enabled = True
    mnuTab.Enabled = True
    mnuQuote.Enabled = True
    mnuBackslash.Enabled = True
End If

End Sub

Private Sub mnuEscapes_Click()
If mnuDisable.Checked = True Then
    mnuCrLf.Enabled = False
    mnuTab.Enabled = False
    mnuQuote.Enabled = False
    mnuBackslash.Enabled = False
Else
    mnuCrLf.Enabled = True
    mnuTab.Enabled = True
    mnuQuote.Enabled = True
    mnuBackslash.Enabled = True
End If

End Sub


Private Sub mnuQuote_Click()

txtScript.SelText = "\q"

End Sub


Private Sub mnuTab_Click()
txtScript.SelText = "\t"
End Sub


Private Sub pic1_Click()

txtScript.SetFocus

End Sub


Private Sub tbrScriptEditor_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim bTemp As Boolean

Select Case LCase(Button.Key)
    
    Case "new": btnNew_Click

    Case "open": btnOpen_Click
    
    'Case "quick": QuickOpen

    Case "save": btnSave_Click

    Case "saveas": btnSaveAs_Click

    Case "comment": btnComment_Click

    Case "uncomment": btnUnComment_Click

    Case "indent": btnIndent_Click

    Case "outdent": btnOutdent_Click

    Case "hilite":
        bTemp = IsDirty 'save "Modified" state
        SelFontBackColor txtScript, RGB(255, 255, 210)
        IsDirty = bTemp 'restore "Modified" state
        txtScript.SelLength = 0
    
    Case "unhilite":
        bTemp = IsDirty 'save "Modified" state
        RemoveFontBackColor txtScript ', RemoveAll:=True
        IsDirty = bTemp 'restore "Modified" state
        txtScript.SelLength = 0

    Case "run": btnRun_Click
    
    Case "vbs": btnViewVBScript_Click
    
    Case "view": btnViewIncluded_Click

    Case "regexp": btnRegExpEdit_Click

    Case "escape": btnEscape_Click

    Case "special":
        'PopUpMenu mnuEscapes, 2, Button.Left, Button.Top + Button.Height
        'txtScript.SetFocus

    Case "vbshelp": HHelp_Show RemoveSlash(App.Path) & "\Vbscript.chm", "/html/VBSTOC.htm"

    Case "axshelp": HHelp_Show RemoveSlash(App.Path) & "\Axiom VBScript.chm", "Axiom_VBS.htm"


    Case Else

End Select

End Sub
Private Sub tbrScriptEditor_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

Select Case ButtonMenu.Key

    Case "backslash": txtScript.SelText = "\\"
    
    Case "newline": txtScript.SelText = "\n"
    
    Case "tab": txtScript.SelText = "\t"
    
    Case "quote": txtScript.SelText = "\q"
    
End Select


End Sub
Private Sub txtScript_Change()

IsDirty = True

End Sub

Private Sub txtScript_KeyDown(KeyCode As Integer, Shift As Integer)

If (Shift = vbCtrlMask) And (KeyCode = vbKeyV) Then
    'Trap [CTRL]+[V] = PASTE
    'To prevent pasting of images and formated text
    KeyCode = 0
    'If Clipboard.GetFormat(vbCFText) Then
    RTF_Paste_Text txtScript.hwnd
    'End If

End If

End Sub
Private Sub txtScript_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then
    KeyAscii = 0
    If txtScript.SelLength = 0 Then
        txtScript.SelText = Space$(gEditorSpacesPerTab)
    Else
        btnIndent_Click
    End If

End If

End Sub
Private Sub txtScript_KeyUp(KeyCode As Integer, Shift As Integer)
Dim bTemp As Boolean
    
Select Case KeyCode
    Case vbKeyReturn, vbKeySpace, vbKeyUp, vbKeyDown
        bTemp = IsDirty 'preserve modified state
        SyntaxColorize txtScript ', QBColor(2), RGB(165, 105, 195), vbBlue
        IsDirty = bTemp
End Select

End Sub
Private Sub txtScript_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If x < Screen.TwipsPerPixelX * 8 Then
    txtScript.MousePointer = 99
    txtScript.MouseIcon = frmVBScript.imgCursor.Picture
    bInMargin = True
Else
    txtScript.MousePointer = 0
    bInMargin = False
End If

End Sub
Private Sub txtScript_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lResult As Long, lCurPos As Long

If Button = vbLeftButton Then
    If txtScript.SelLength > 0 Then Exit Sub
    
    If bInMargin = True Then
        SelText txtScript, CurrentLine(txtScript), CurrentLine(txtScript)
    End If

ElseIf Button = vbRightButton Then

    lResult = PopUp("Undo", "-", "Cut", "Copy", "Paste", "Delete", "-", "Select All", "-", "Find...", "Find Next")
    '                 1      2     3       4       5        6       7       8          9     10      11
    
    Select Case lResult
        
        Case 1: 'Undo
            Undo txtScript.hwnd
        
        Case 3: 'Cut
            Cut txtScript.hwnd
    
        Case 4: 'Copy
            Copy txtScript.hwnd
            
        Case 5: 'Paste
            'If Clipboard.GetFormat(vbCFText) Then
             RTF_Paste_Text txtScript.hwnd
            'End If
        
        Case 6: 'Delete
            txtScript.SelText = ""
            
        Case 8: 'Select All
            SelectAll txtScript.hwnd
        
        Case 10: 'Find Next
            msFindWhat = InputBox("Find what?", "Find", "")
            'If ml_OutTextSearchPos = 0 Then ml_OutTextSearchPos = OutputText.SelStart
            mlSearchPos = txtScript.Find(msFindWhat, txtScript.SelStart, Len(txtScript.Text)) + 1

        Case 11: 'Find Next
            If mlSearchPos = 0 Then
                mlSearchPos = txtScript.SelStart
                msFindWhat = InputBox("Find what?", "Find", "")
            End If
            mlSearchPos = txtScript.Find(msFindWhat, mlSearchPos, Len(txtScript.Text)) + 1
        
        Case Else
            'we never get here!
    End Select

End If

End Sub
