VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegExpGenerator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regular Expression Pattern Editor"
   ClientHeight    =   5310
   ClientLeft      =   2700
   ClientTop       =   2730
   ClientWidth     =   7350
   Icon            =   "RegExpGen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Un-Escape dbl-quotes"
      Height          =   195
      Left            =   180
      TabIndex        =   26
      Top             =   4680
      Width           =   2115
   End
   Begin VB.CheckBox chkVBScriptCompatible 
      Caption         =   "Return VBScript Compatible Pattern"
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      ToolTipText     =   "Propperly escapes the quote char for VBScript use"
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox txtRegExpPattern 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   4920
      Width           =   7200
   End
   Begin VB.CommandButton btnAutoEscape 
      Caption         =   "Escape selected text"
      Height          =   315
      Left            =   60
      TabIndex        =   23
      ToolTipText     =   "Escape RegExp Special Chars:  \^$*+{}?.:=!|[]-(),"
      Top             =   4020
      Width           =   1755
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   3540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegExpGen.frx":000C
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegExpGen.frx":022E
            Key             =   "saveas"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegExpGen.frx":0F3A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegExpGen.frx":115C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegExpGen.frx":137E
            Key             =   "selectall"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegExpGen.frx":15A0
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegExpGen.frx":17C2
            Key             =   "new"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegExpGen.frx":19E4
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegExpGen.frx":1C06
            Key             =   "copy"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   370
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4050
      Width           =   1230
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   370
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4050
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   370
      Left            =   4642
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4050
      Width           =   1230
   End
   Begin VB.CheckBox chkEscapeAll 
      Caption         =   "Escape non-printable also"
      Height          =   300
      Left            =   300
      TabIndex        =   19
      ToolTipText     =   "Escape Carriage Return, Line Feed and Tab"
      Top             =   4350
      Width           =   2985
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   490
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   570
      Width           =   7350
      Begin VB.CheckBox chkNonGreedy 
         Caption         =   "Non-Greedy Matching"
         Height          =   255
         Left            =   5280
         TabIndex        =   20
         Top             =   480
         Width           =   2085
      End
      Begin AxiomHTML.ShadowedSeperator ShadowedSeperator1 
         Height          =   30
         Left            =   -120
         Top             =   30
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   53
      End
      Begin AxiomHTML.CoolButton btnChar 
         Height          =   300
         Index           =   0
         Left            =   1035
         TabIndex        =   8
         Top             =   450
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "RegExpGen.frx":1E28
      End
      Begin AxiomHTML.CoolButton btnRepetition 
         Height          =   270
         Index           =   0
         Left            =   5460
         TabIndex        =   6
         Top             =   120
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   476
         Caption         =   ""
         MaskColor       =   0
         ShowFocusRect   =   0   'False
      End
      Begin AxiomHTML.CoolButton btnLiterals 
         Height          =   270
         Index           =   0
         Left            =   780
         TabIndex        =   3
         Top             =   135
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   476
         Caption         =   "*"
         MaskColor       =   0
         ShowFocusRect   =   0   'False
      End
      Begin AxiomHTML.CoolButton btnChar 
         Height          =   300
         Index           =   1
         Left            =   1350
         TabIndex        =   9
         Top             =   450
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "RegExpGen.frx":1FAA
      End
      Begin AxiomHTML.CoolButton btnChar 
         Height          =   300
         Index           =   2
         Left            =   1665
         TabIndex        =   10
         Top             =   450
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "RegExpGen.frx":212C
      End
      Begin AxiomHTML.CoolButton btnChar 
         Height          =   300
         Index           =   3
         Left            =   1995
         TabIndex        =   11
         Top             =   450
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "RegExpGen.frx":22AE
      End
      Begin AxiomHTML.CoolButton btnChar 
         Height          =   300
         Index           =   4
         Left            =   2280
         TabIndex        =   12
         Top             =   450
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "RegExpGen.frx":2430
      End
      Begin AxiomHTML.CoolButton btnChar 
         Height          =   300
         Index           =   5
         Left            =   2580
         TabIndex        =   13
         Top             =   450
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         MaskColor       =   14737632
         ShowFocusRect   =   0   'False
         BackPicture     =   "RegExpGen.frx":25B2
      End
      Begin AxiomHTML.CoolButton btnChar 
         Height          =   300
         Index           =   6
         Left            =   2880
         TabIndex        =   14
         Top             =   450
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         MaskColor       =   14737632
         ShowFocusRect   =   0   'False
         BackPicture     =   "RegExpGen.frx":2734
      End
      Begin AxiomHTML.CoolButton btnChar 
         Height          =   300
         Index           =   7
         Left            =   3180
         TabIndex        =   15
         Top             =   450
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "RegExpGen.frx":28B6
      End
      Begin AxiomHTML.CoolButton btnChar 
         Height          =   300
         Index           =   8
         Left            =   3495
         TabIndex        =   16
         Top             =   450
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "RegExpGen.frx":2A38
      End
      Begin AxiomHTML.CoolButton btnChar 
         Height          =   300
         Index           =   9
         Left            =   3825
         TabIndex        =   18
         ToolTipText     =   "ASCII Char"
         Top             =   450
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Caption         =   ""
         MaskColor       =   12632256
         ShowFocusRect   =   0   'False
         BackPicture     =   "RegExpGen.frx":2BBA
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charachters"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repetition"
         Height          =   195
         Left            =   4680
         TabIndex        =   5
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Literals"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   165
         Width           =   495
      End
   End
   Begin VB.TextBox txtRegExp 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1440
      Width           =   7200
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   2520
      Top             =   3660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuCharClasses 
      Caption         =   "Char Classes"
      Begin VB.Menu mnuCharClassesAnyOf 
         Caption         =   "Any of"
      End
      Begin VB.Menu mnuCharClassesAllExcept 
         Caption         =   "All except"
      End
      Begin VB.Menu mnuCharClassesAnyChar 
         Caption         =   "Any char (except NewLine)"
      End
      Begin VB.Menu mnuCharClassesWordChar 
         Caption         =   "Word char"
      End
      Begin VB.Menu mnuCharClassesNonWordChar 
         Caption         =   "Non-word char"
      End
      Begin VB.Menu mnuCharClassesDigit 
         Caption         =   "Digit"
      End
      Begin VB.Menu mnuCharClassesNonDigit 
         Caption         =   "Non-digit"
      End
      Begin VB.Menu mnuCharClassesWhiteSpace 
         Caption         =   "white space"
      End
      Begin VB.Menu mnuCharClassesNonWhiteSpace 
         Caption         =   "Non-white space"
      End
   End
   Begin VB.Menu mnuLiterals 
      Caption         =   "Literals"
      Begin VB.Menu mnuLiteralsQuest 
         Caption         =   "?"
      End
      Begin VB.Menu mnuLiteralsAst 
         Caption         =   "*"
      End
      Begin VB.Menu mnuLiteralsPlus 
         Caption         =   "+"
      End
      Begin VB.Menu mnuLiteralsDot 
         Caption         =   "."
      End
      Begin VB.Menu mnuLiteralsPipe 
         Caption         =   "|"
      End
      Begin VB.Menu mnuLiteralsCurlyOpen 
         Caption         =   "{"
      End
      Begin VB.Menu mnuLiteralsCurlyClose 
         Caption         =   "}"
      End
      Begin VB.Menu mnuLiteralsBackSlash 
         Caption         =   "\"
      End
      Begin VB.Menu mnuLiteralsBracketsOpen 
         Caption         =   "["
      End
      Begin VB.Menu mnuLiteralsBracketsClose 
         Caption         =   "]"
      End
      Begin VB.Menu mnuLiteralsParaOpen 
         Caption         =   "("
      End
      Begin VB.Menu mnuLiteralsParaClose 
         Caption         =   ")"
      End
      Begin VB.Menu mnuLiteralsCaret 
         Caption         =   "^"
      End
      Begin VB.Menu mnuLiteralsDollar 
         Caption         =   "$"
      End
      Begin VB.Menu mnuLiterals_seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiteralsCr 
         Caption         =   "Carriage Return"
      End
      Begin VB.Menu mnuLiteralsLineFeed 
         Caption         =   "Line Feed"
      End
      Begin VB.Menu mnuLiteralsTab 
         Caption         =   "Tab"
      End
      Begin VB.Menu mnuLiterals_seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiteralsHex 
         Caption         =   "ASCII Hex Value"
      End
      Begin VB.Menu mnuLiteralsUnicode 
         Caption         =   "Unicode Hex Value"
      End
   End
   Begin VB.Menu mnuRepetition 
      Caption         =   "Repetition"
      Begin VB.Menu mnuRepetitionItem 
         Caption         =   "0 or more times"
         Index           =   0
      End
      Begin VB.Menu mnuRepetitionItem 
         Caption         =   "1 or more times"
         Index           =   1
      End
      Begin VB.Menu mnuRepetitionItem 
         Caption         =   "0 or 1 time"
         Index           =   2
      End
      Begin VB.Menu mnuRepetitionItem 
         Caption         =   "n times"
         Index           =   3
      End
      Begin VB.Menu mnuRepetitionItem 
         Caption         =   "At least n times"
         Index           =   4
      End
      Begin VB.Menu mnuRepetitionItem 
         Caption         =   "At least n times at most m times"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPosition 
      Caption         =   "Position"
      Begin VB.Menu mnuPositionStart 
         Caption         =   "Start of line"
      End
      Begin VB.Menu mnuPositionEnd 
         Caption         =   "End of line"
      End
      Begin VB.Menu mnuPositionWordBoundry 
         Caption         =   "Word boundry"
      End
      Begin VB.Menu mnuPositionNonWordBoundry 
         Caption         =   "Non-word boundry"
      End
   End
   Begin VB.Menu mnuAdvanced 
      Caption         =   "Advanced"
      Begin VB.Menu mnuAdvancedAlternate 
         Caption         =   "Alternation"
      End
      Begin VB.Menu mnuAdvancedGroup 
         Caption         =   "Grouping (Capturing submatch)"
      End
      Begin VB.Menu mnuAdvancedNonCapture 
         Caption         =   "Non-capturing match"
      End
      Begin VB.Menu mnuAdvancedHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdvancedPositive 
         Caption         =   "Positive Lookahead"
      End
      Begin VB.Menu mnuAdvancedNegative 
         Caption         =   "Negative Lookahead"
      End
      Begin VB.Menu mnuAdvancedHyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdvancedBackRef 
         Caption         =   "Backreference"
      End
   End
   Begin VB.Menu mnuCustom 
      Caption         =   "Custom"
      Begin VB.Menu mnuCustomItem 
         Caption         =   "(none found)"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmRegExpGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ms_Value As String

Private IsDirty As Boolean

'Property to indicate that [Cancel] was pressed:
Public Canceled As Boolean

Private Sub New_Click()
Dim Result As VbMsgBoxResult

If IsDirty Then
    Result = MsgBox("Contents have changed, Continue?", vbOKCancel + vbDefaultButton2 + vbQuestion)
    If Result = vbCancel Then Exit Sub
End If

Me.Caption = "RegExp Pattern Editor"
cdlg1.FileName = ""

txtRegExp.Text = ""

IsDirty = False

txtRegExp.SetFocus

End Sub
Private Sub SaveAs_Click()
Dim TextFile As New CTextFile

cdlg1.Filter = "RegExp Files|*.rx|All Files (*.*)|*.*"
cdlg1.InitDir = RemoveSlash(App.Path) & "\RegExps"
cdlg1.DialogTitle = "Save As"
cdlg1.flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly

On Error Resume Next
    cdlg1.ShowSave
If Err Then Exit Sub
On Error GoTo 0

TextFile.FileOpen cdlg1.FileName, OpenForOutput
TextFile.WriteStr txtRegExp.Text
TextFile.FileClose

Me.Caption = "RegExp Pattern Editor - [" & ExtractFileName(cdlg1.FileName) & "]"
IsDirty = False

Set TextFile = Nothing

txtRegExp.SetFocus

End Sub
Private Sub CreateToolBar()

With Toolbar1

    .ImageList = ImageList1
    
    .Buttons.Add , "sep_1", , tbrSeparator
    
    .Buttons.Add , "new", , tbrDefault, "new"
    .Buttons("new").ToolTipText = "New"

    .Buttons.Add , "open", , tbrDefault, "open"
    .Buttons("open").ToolTipText = "Open"
    
    .Buttons.Add , "save", , tbrDefault, "save"
    .Buttons("save").ToolTipText = "Save"
    
    .Buttons.Add , "saveas", , tbrDefault, "saveas"
    .Buttons("saveas").ToolTipText = "Save As"

    .Buttons.Add , "sep_1a", , tbrSeparator

    .Buttons.Add , "selectall", , tbrDefault, "selectall"
    .Buttons("selectall").ToolTipText = "Select All"

    .Buttons.Add , "sep_2", , tbrSeparator

    .Buttons.Add , "cut", , tbrDefault, "cut"
    .Buttons("cut").ToolTipText = "Cut"

    .Buttons.Add , "copy", , tbrDefault, "copy"
    .Buttons("copy").ToolTipText = "Copy"

    .Buttons.Add , "paste", , tbrDefault, "paste"
    .Buttons("paste").ToolTipText = "Paste"

    .Buttons.Add , "undo", , tbrDefault, "undo"
    .Buttons("undo").ToolTipText = "Undo"

End With

End Sub
Private Sub LoadCustomRegExp()
Dim idx As Long
Dim sTempArray() As String
Dim objIniFile As CINIFileAccess

Set objIniFile = New CINIFileAccess
objIniFile.FileName = RemoveSlash(App.Path) & "\Axiom.ini"
objIniFile.Section = "Custom RegExp"

sTempArray = objIniFile.EnumKeys

For idx = LBound(sTempArray) To UBound(sTempArray)
    
    If idx > 0 Then Load mnuCustomItem(idx)
    mnuCustomItem(idx).Caption = sTempArray(idx)
    mnuCustomItem(idx).Visible = True
    objIniFile.Key = sTempArray(idx)
    mnuCustomItem(idx).Tag = objIniFile.Value

Next

If UBound(sTempArray) = -1 Then
    mnuCustomItem(0).Caption = "(none found)"
    mnuCustomItem(0).Enabled = False
End If

Set objIniFile = Nothing

End Sub
Private Function TranslateMetaChar(ByVal MetaChar As String) As String

Dim S As String

Select Case MetaChar

Case "\d"
    S = "Any Digit Char [0-9]"
Case "\D"
    S = "Anyt Non-Digit Char[^0-9]"

Case Else

    S = MetaChar
End Select


TranslateMetaChar = S

End Function
Public Property Get Value() As String

    Value = ms_Value

End Property
Public Property Let Value(ByVal sNewValue As String)

       ms_Value = sNewValue
       txtRegExp.Text = sNewValue

End Property
Private Function CharsToHex(ByVal sCharList As String) As String
Dim idx As Long
Dim ch As Integer
Dim sTemp As String

For idx = 1 To Len(sCharList)
    ch = Asc(Mid$(sCharList, idx, 1))
    sTemp = sTemp & "\x" & Format$(Hex$(ch), "00")
Next

CharsToHex = sTemp

End Function

Private Sub btnAutoEscape_Click()
Dim sTemp As String


If txtRegExp.SelLength > 0 Then
    sTemp = txtRegExp.SelText
Else
    sTemp = txtRegExp.Text
End If

sTemp = EscapeRegExpChars(sTemp)

If chkEscapeAll.Value = vbChecked Then
    sTemp = EscapeNonPrintableChars(sTemp)
End If

If txtRegExp.SelLength > 0 Then
    'Only selection
    txtRegExp.SelText = sTemp
Else
    'No selection, so do NOTHING
    'txtRegExp.Text = sTemp
End If
    
txtRegExp.SetFocus

End Sub
Private Sub btnChar_Click(Index As Integer)
'Dim sTemp As String

If Index = 9 Then
    'sTemp = InputBox("Enter Hex ASCII Value", "ASCII", "")
    'If sTemp <> "" Then txtRegExp.SelText = "\x" & UCase$(sTemp)
    PopUpMenu frmHidden.mnuASCIITable, vbBoth
Else
    txtRegExp.SelText = btnChar(Index).Tag
End If


txtRegExp.SetFocus

End Sub
Private Sub btnLiterals_Click(Index As Integer)

txtRegExp.SelText = "\" & btnLiterals(Index).Caption
txtRegExp.SetFocus

End Sub

Private Sub btnRepetition_Click(Index As Integer)
Dim sResult As String

Select Case Index

Case 0
    sResult = "*"

Case 1
    sResult = "+"
Case 2
    sResult = "?"
Case 3
    sResult = InputBox("How many times?", "Repeat", "")
    If sResult <> "" Then
        sResult = "{" & sResult & "}"
    End If
Case 4
    sResult = InputBox("How many times?", "Repeat", "")
    If sResult <> "" Then
        sResult = "{" & sResult & ",}"
    End If
Case 5
    sResult = InputBox("Enter n,m ", "Repeat", "")
    If sResult <> "" Then
        sResult = "{" & sResult & "}"
    End If
End Select

If sResult <> "" Then
        If chkNonGreedy.Value = vbChecked Then
            sResult = sResult & "?"
        End If
        
        txtRegExp.SelText = sResult
End If

txtRegExp.SetFocus

End Sub

Private Sub chkEscapeAll_Click()

txtRegExp.SetFocus

End Sub
Private Sub chkNonGreedy_Click()

txtRegExp.SetFocus

End Sub

Private Sub cmdCancel_Click()

Canceled = True
'Value = ""
Me.Hide

End Sub
Private Sub cmdHelp_Click()

HHelp_Show RemoveSlash(App.Path) & "\RegExp Help.chm", "RegExpHelp.htm"

'no good, steals focus from help file!
'txtRegExp.SetFocus

End Sub
Private Sub cmdOk_Click()
Dim sTemp As String

sTemp = txtRegExpPattern.Text

If chkVBScriptCompatible.Value = vbChecked Then
    sTemp = Replace$(sTemp, Chr$(34), Chr$(34) & Chr$(34))
End If

Value = Replace(sTemp, Chr$(172), " ")

Canceled = False
Me.Hide

End Sub
Private Sub Form_Activate()

On Error Resume Next
txtRegExp.SetFocus

End Sub

Private Sub Save_Click()
Dim TextFile As New CTextFile

If cdlg1.FileName = "" Then
    cdlg1.Filter = "RegExp Files|*.rx|All Files (*.*)|*.*"
    cdlg1.InitDir = RemoveSlash(App.Path) & "\RegExps"
    cdlg1.DialogTitle = "Save"
    cdlg1.flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly

    On Error Resume Next
        cdlg1.ShowSave
    If Err Then Exit Sub
    On Error GoTo 0
End If

TextFile.FileOpen cdlg1.FileName, OpenForOutput
TextFile.WriteStr txtRegExp.Text
TextFile.FileClose

Me.Caption = "RegExp Pattern Editor - [" & ExtractFileName(cdlg1.FileName) & "]"
IsDirty = False

Set TextFile = Nothing

txtRegExp.SetFocus

End Sub

Private Sub Open_Click()
Dim TextFile As New CTextFile
Dim Result As VbMsgBoxResult
Dim sCurrentFilename As String
Dim sFileContents As String

If IsDirty Then
    Result = MsgBox("Contents have changed, Continue?", vbOKCancel + vbDefaultButton2 + vbQuestion)
    If Result = vbCancel Then Exit Sub
End If

sCurrentFilename = cdlg1.FileName

cdlg1.InitDir = RemoveSlash(App.Path) & "\RegExps"
cdlg1.Filter = "RegExp Files|*.rx|All Files|*.*"
cdlg1.DialogTitle = "Open"
cdlg1.FileName = ""
cdlg1.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly

On Error Resume Next
    cdlg1.ShowOpen
    If Err Then          'user clicked "Cancel"
        Err.Clear
        cdlg1.FileName = sCurrentFilename
        txtRegExp.SetFocus
        Exit Sub
    End If
On Error GoTo 0

TextFile.FileOpen cdlg1.FileName, OpenForInput
sFileContents = TextFile.ReadAll
TextFile.FileClose

'dont change to .FileTitle, i know what i'm doing!
Me.Caption = "RegExp Pattern Editor - [" & ExtractFileName(cdlg1.FileName) & "]"
IsDirty = False

Set TextFile = Nothing

txtRegExp.Text = sFileContents

IsDirty = False

txtRegExp.SetFocus

End Sub
Private Sub Form_Load()
Dim idx As Integer

CreateToolBar
CButtons Me

SetLeftMargin txtRegExp, 8

LoadCustomRegExp

For idx = 1 To 13
    Load btnLiterals(idx)
    btnLiterals(idx).Visible = True
    btnLiterals(idx).Left = btnLiterals(idx - 1).Width + btnLiterals(idx - 1).Left
Next
btnLiterals(0).Caption = "?"
btnLiterals(1).Caption = "*"
btnLiterals(2).Caption = "+"
btnLiterals(3).Caption = "."
btnLiterals(4).Caption = "|"
btnLiterals(5).Caption = "{"
btnLiterals(6).Caption = "}"
btnLiterals(7).Caption = "\"
btnLiterals(8).Caption = "["
btnLiterals(9).Caption = "]"
btnLiterals(10).Caption = "("
btnLiterals(11).Caption = ")"
btnLiterals(12).Caption = "^"
btnLiterals(13).Caption = "$"

For idx = 1 To 5
    Load btnRepetition(idx)
    btnRepetition(idx).Visible = True
    btnRepetition(idx).Left = btnRepetition(idx - 1).Width + btnRepetition(idx - 1).Left
Next
btnRepetition(0).Caption = "0+"
btnRepetition(0).ToolTipText = "0 Or More Times" & vbTab & "*"
btnRepetition(1).Caption = "1+"
btnRepetition(1).ToolTipText = "1 Or More Times"
btnRepetition(2).Caption = "0/1"
btnRepetition(2).ToolTipText = "0 Or 1 Time"
btnRepetition(3).Caption = "n"
btnRepetition(3).ToolTipText = "n Times"
btnRepetition(4).Caption = ">n"
btnRepetition(4).ToolTipText = "At Least n Times"
btnRepetition(5).Caption = "n,m"
btnRepetition(5).ToolTipText = "At Least n At Most m Times"

For idx = 1 To 9
    btnChar(idx).Visible = True
    btnChar(idx).Top = btnChar(0).Top
    btnChar(idx).Left = btnChar(idx - 1).Width + btnChar(idx - 1).Left
Next
btnChar(0).Tag = "."
btnChar(0).ToolTipText = "Any Charachter"
btnChar(1).Tag = "\d"
btnChar(1).ToolTipText = "Digit"
btnChar(2).Tag = "\D"
btnChar(2).ToolTipText = "Non-Digit"
btnChar(3).Tag = "\w"
btnChar(3).ToolTipText = "Alphanumeric"
btnChar(4).Tag = "\W"
btnChar(4).ToolTipText = "Non-Alphanumeric"
btnChar(5).Tag = "\s"
btnChar(5).ToolTipText = "Whitespace"
btnChar(6).Tag = "\S"
btnChar(6).ToolTipText = "Non-Whitespace"
btnChar(7).Tag = "\r\n"
btnChar(7).ToolTipText = "Newline (CR+LF)"
btnChar(8).Tag = "\t"
btnChar(8).ToolTipText = "Tab"
btnChar(9).Tag = "\x"
btnChar(9).ToolTipText = "ASCII Char"

'mnuAnyChar.Caption = "Any Charachter" & vbTab & "."
'mnuDigit.Caption = "Digit" & vbTab & "\d"
'mnuNonDigit.Caption = "Non-Digit" & vbTab & "\D"
'mnuTab.Caption = "Tab" & vbTab & "\t"
'mnuAlphaNum.Caption = "Alphanumeric" & vbTab & "\w"
'mnuNonAlphaNum.Caption = "Non-Alphanumeric" & vbTab & "\W"
'mnuWhiteSpace.Caption = "Whitespace" & vbTab & "\s"
'mnuNonWhiteSpace.Caption = "Non-Whitespace" & vbTab & "\S"
'mnuCR.Caption = "Newline (CR+LF)" & vbTab & "\r\n"
'mnuASCII.Caption = "ASCII Code"
'mnuZeroOrMore.Caption = "0 Or More Times" & vbTab & "*"
'mnuOneOrMore.Caption = "1 Or More Times" & vbTab & "+"
'mnuZeroOrOne.Caption = "0 Or 1 Time" & vbTab & "?"
'mnuNTimes.Caption = "n Times" & vbTab & "{n}"
'mnuAtLeast.Caption = "At Least n Times" & vbTab & "{n,}"
'mnuAtLeastAtMost.Caption = "At Least n At Most m Times" & vbTab & "{n,m}"

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then 'clicked the [x] button

    Cancel = True
    Value = ""
    Me.Hide

End If


End Sub

Private Sub mnuAdvancedAlternate_Click()

txtRegExp.SelText = "|"
txtRegExp.SetFocus

End Sub

Private Sub mnuAdvancedBackRef_Click()
Dim sTemp As String

sTemp = InputBox("Enter previous submatch number", "Backreference", "")
If sTemp <> "" Then
    txtRegExp.SelText = "(" & txtRegExp.SelText & ")\" & sTemp
End If

txtRegExp.SetFocus

End Sub
Private Sub mnuAdvancedGroup_Click()

txtRegExp.SelText = "(" & txtRegExp.SelText & ")"
txtRegExp.SetFocus

End Sub

Private Sub mnuAdvancedNegative_Click()

txtRegExp.SelText = "(?!" & txtRegExp.SelText & ")"
txtRegExp.SetFocus

End Sub

Private Sub mnuAdvancedNonCapture_Click()


txtRegExp.SelText = "(?:" & txtRegExp.SelText & ")"
txtRegExp.SetFocus

End Sub

Private Sub mnuAdvancedPositive_Click()

txtRegExp.SelText = "(?=" & txtRegExp.SelText & ")"
txtRegExp.SetFocus

End Sub

Private Sub mnuCharClassesAllExcept_Click()

Dim sResult As String

sResult = InputBox("Type chars or char range (denoted by a hyphen)", "Enter Chars", "")
If sResult <> "" Then
    txtRegExp.SelText = "[^" & sResult & "]"
End If

txtRegExp.SetFocus

End Sub

Private Sub mnuCharClassesAnyChar_Click()

txtRegExp.SelText = "."
txtRegExp.SetFocus

End Sub

Private Sub mnuCharClassesAnyOf_Click()
Dim sResult As String

sResult = InputBox("Type chars or char range (denoted by a hyphen)", "Enter Chars", "")
If sResult <> "" Then
    txtRegExp.SelText = "[" & sResult & "]"
End If

txtRegExp.SetFocus

End Sub
Private Sub mnuCharClassesDigit_Click()

txtRegExp.SelText = "\d"
txtRegExp.SetFocus

End Sub

Private Sub mnuCharClassesNonDigit_Click()

txtRegExp.SelText = "\D"
txtRegExp.SetFocus

End Sub

Private Sub mnuCharClassesNonWhiteSpace_Click()

txtRegExp.SelText = "\S"
txtRegExp.SetFocus

End Sub

Private Sub mnuCharClassesNonWordChar_Click()

txtRegExp.SelText = "\W"
txtRegExp.SetFocus

End Sub

Private Sub mnuCharClassesWhiteSpace_Click()

txtRegExp.SelText = "\s"
txtRegExp.SetFocus

End Sub


Private Sub mnuCharClassesWordChar_Click()

txtRegExp.SelText = "\w"
txtRegExp.SetFocus

End Sub

Private Sub mnuCustomItem_Click(Index As Integer)

txtRegExp.SelText = mnuCustomItem(Index).Tag
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsAst_Click()

txtRegExp.SelText = "\*"
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsBackSlash_Click()

txtRegExp.SelText = "\\"
txtRegExp.SetFocus


End Sub

Private Sub mnuLiteralsBracketsClose_Click()

txtRegExp.SelText = "\]"
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsBracketsOpen_Click()

txtRegExp.SelText = "\["
txtRegExp.SetFocus

End Sub


Private Sub mnuLiteralsCaret_Click()

txtRegExp.SelText = "\^"
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsCr_Click()

txtRegExp.SelText = "\r"
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsCurlyClose_Click()

txtRegExp.SelText = "\}"
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsCurlyOpen_Click()

txtRegExp.SelText = "\{"
txtRegExp.SetFocus


End Sub

Private Sub mnuLiteralsDollar_Click()

txtRegExp.SelText = "\$"
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsDot_Click()

txtRegExp.SelText = "\."
txtRegExp.SetFocus


End Sub

Private Sub mnuLiteralsHex_Click()
Dim sTemp As String

sTemp = InputBox("Enter Hex ASCII Value", "ASCII", "")
If sTemp <> "" Then txtRegExp.SelText = "\x" & UCase$(sTemp)
txtRegExp.SetFocus

End Sub
Private Sub mnuLiteralsLineFeed_Click()

txtRegExp.SelText = "\n"
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsParaClose_Click()

txtRegExp.SelText = "\)"
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsParaOpen_Click()

txtRegExp.SelText = "\("
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsPipe_Click()

txtRegExp.SelText = "\|"
txtRegExp.SetFocus

End Sub


Private Sub mnuLiteralsPlus_Click()

txtRegExp.SelText = "\+"
txtRegExp.SetFocus


End Sub


Private Sub mnuLiteralsQuest_Click()

txtRegExp.SelText = "\?"
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsTab_Click()

txtRegExp.SelText = "\t"
txtRegExp.SetFocus

End Sub

Private Sub mnuLiteralsUnicode_Click()
Dim sTemp As String

sTemp = InputBox("Enter Hex Unicode Value", "ASCII", "")
If sTemp <> "" Then txtRegExp.SelText = "\u" & UCase$(sTemp)
txtRegExp.SetFocus

End Sub

Private Sub mnuPositionEnd_Click()

txtRegExp.SelText = "$"
txtRegExp.SetFocus

End Sub

Private Sub mnuPositionNonWordBoundry_Click()

txtRegExp.SelText = "\B"
txtRegExp.SetFocus

End Sub
Private Sub mnuPositionStart_Click()

txtRegExp.SelText = "^"
txtRegExp.SetFocus

End Sub

Private Sub mnuPositionWordBoundry_Click()

txtRegExp.SelText = "\b"
txtRegExp.SetFocus

End Sub

Private Sub mnuRepetitionItem_Click(Index As Integer)
Dim sResult As String

Select Case Index

Case 0
    sResult = "*"

Case 1
    sResult = "+"
Case 2
    sResult = "?"
Case 3
    sResult = InputBox("How many times?", "Repeat", "")
    If sResult <> "" Then
        sResult = "{" & sResult & "}"
    End If
Case 4
    sResult = InputBox("How many times?", "Repeat", "")
    If sResult <> "" Then
        sResult = "{" & sResult & ",}"
    End If
Case 5
    sResult = InputBox("Enter n,m ", "Repeat", "")
    If sResult <> "" Then
        sResult = "{" & sResult & "}"
    End If
End Select

If sResult <> "" Then
        If chkNonGreedy.Value = vbChecked Then
            sResult = sResult & "?"
        End If
        
        txtRegExp.SelText = sResult
End If

txtRegExp.SetFocus

End Sub
Private Sub Picture1_Click()

txtRegExp.SetFocus

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "selectall"
        SelText txtRegExp, 0, GetLineCount(txtRegExp)
    
    Case "new": New_Click
        'To allow UNDO
        'SelText txtRegExp, 0, GetLineCount(txtRegExp)
        'DoEditOperation txtRegExp, Edit_Clear
    
    Case "open": Open_Click
    
    Case "save": Save_Click

    Case "saveas": SaveAs_Click
    
    Case "undo"
        DoEditOperation txtRegExp, Edit_Uundo

    Case "cut"
        DoEditOperation txtRegExp, Edit_Cut
    
    Case "copy"
        DoEditOperation txtRegExp, Edit_Copy

    Case "paste"
        DoEditOperation txtRegExp, Edit_Paste

End Select

txtRegExp.SetFocus


End Sub
Private Sub txtRegExp_Change()
Dim sTemp As String

sTemp = RX_GenericReplace(txtRegExp.Text, "(\t.*?$)|(\r\n)", "")
'txtRegExpPattern.Text = Replace(sTemp, " ", Chr$(176))
txtRegExpPattern.Text = Replace(sTemp, " ", Chr$(172))

IsDirty = True

End Sub
Private Sub txtRegExp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'txtRegExp.ToolTipText = TranslateMetaChar(txtRegExp.SelText)

End Sub

