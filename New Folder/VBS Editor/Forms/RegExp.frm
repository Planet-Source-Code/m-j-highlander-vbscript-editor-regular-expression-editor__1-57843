VERSION 5.00
Begin VB.Form frmRegExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regular Expression  Extract / Replace"
   ClientHeight    =   2085
   ClientLeft      =   2220
   ClientTop       =   3480
   ClientWidth     =   8190
   Icon            =   "RegExp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSpecialReplace 
      Height          =   345
      Left            =   7260
      Picture         =   "RegExp.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Insert Non-printable Char"
      Top             =   600
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   20010
      Width           =   375
   End
   Begin VB.CommandButton btnRegExpEdit 
      Height          =   345
      Left            =   7710
      Picture         =   "RegExp.frx":0172
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Edit Pattern"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton btnEscape 
      Height          =   345
      Left            =   7260
      Picture         =   "RegExp.frx":02D8
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Escape RegExp Special Chars:  \^$*+{}?.:=!|[]-(),"
      Top             =   120
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   20000
      Width           =   375
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Match case"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   1170
      Width           =   1455
   End
   Begin VB.CheckBox chkWord 
      Caption         =   "Whole word only"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   1485
      Width           =   1635
   End
   Begin VB.TextBox txtReplace 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   5535
   End
   Begin VB.CheckBox chkReplace 
      Caption         =   "&Replace With"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      TabIndex        =   2
      Top             =   600
      Width           =   1470
   End
   Begin VB.TextBox txtPattern 
      Height          =   315
      Left            =   1455
      TabIndex        =   1
      Top             =   120
      Width           =   5760
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1890
      Picture         =   "RegExp.frx":043E
      Top             =   990
      Width           =   225
   End
   Begin VB.Label lblTip 
      Caption         =   $"RegExp.frx":04DC
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   2160
      TabIndex        =   6
      Top             =   1020
      Width           =   5715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Search Pattern"
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
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
   Begin VB.Menu mnuSpecialChars 
      Caption         =   "mnuSpecialChars"
      Visible         =   0   'False
      Begin VB.Menu mnuPara 
         Caption         =   "Newline (CR+LF)"
      End
      Begin VB.Menu mnuTab 
         Caption         =   "Tab"
      End
      Begin VB.Menu mnuBackSlash 
         Caption         =   "Tilde"
      End
   End
End
Attribute VB_Name = "frmRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ms_Pattern As String
Private ms_ReplaceWith As String

Private mb_MatchCase As Boolean
Private mb_WholeWords As Boolean

Public Property Get WholeWords() As Boolean
       WholeWords = mb_WholeWords
End Property

Public Property Let WholeWords(ByVal bNewValue As Boolean)
       mb_WholeWords = bNewValue
End Property

Public Property Get MatchCase() As Boolean
       MatchCase = mb_MatchCase
End Property

Public Property Let MatchCase(ByVal bNewValue As Boolean)
       mb_MatchCase = bNewValue
End Property

Public Property Get ReplaceWith() As String
       ReplaceWith = ms_ReplaceWith
End Property

Public Property Let ReplaceWith(ByVal sNewValue As String)
       ms_ReplaceWith = sNewValue
End Property

Public Property Get Pattern() As String
       Pattern = ms_Pattern
End Property

Public Property Let Pattern(ByVal sNewValue As String)
       ms_Pattern = sNewValue
End Property

Private Sub btnEscape_Click()

If txtPattern.SelLength > 0 Then
    'Only selection
    txtPattern.SelText = EscapeRegExpChars(txtPattern.SelText)
Else
    'No selection, so do the Entire text
    txtPattern.Text = EscapeRegExpChars(txtPattern.Text)
End If
    
txtPattern.SetFocus

End Sub

Private Sub btnHelp_Click()


End Sub

Private Sub btnEscape_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim sText As String

If Button = vbRightButton Then

    sText = "Escape RegExp Special Chars:" & vbTab & "\^$*+{}?.:=!|[]-()," & vbCrLf & _
            "Some chars, (like \ and ^) have special meanings in a Regular Expression Pattern" & vbCrLf & _
            "To denote the actual literals, the chars must be escaped, that is, preceeded by \"

    HH_ShowPopUp sText

End If

End Sub
Private Sub btnRegExpEdit_Click()
Dim frmX As frmRegExpGenerator
Dim sResult  As String


Set frmX = New frmRegExpGenerator

frmX.Value = txtPattern.Text
frmX.Show vbModal

sResult = frmX.Value
If sResult <> "" And frmX.Canceled = False Then
    txtPattern.Text = sResult
End If

Unload frmX
Set frmX = Nothing

End Sub
Private Sub chkReplace_Click()

If chkReplace.Value = vbChecked Then
    txtReplace.SetFocus
End If

End Sub

Private Sub cmdCancel_Click()
Pattern = ""
ReplaceWith = ""

Me.Hide

End Sub

Private Sub cmdHelp_Click()
'HHelp_Show RemoveSlash(App.Path) & "\regexp.chm", "RegExp Pattern.html"

HHelp_Show RemoveSlash(App.Path) & "\RegExp Help.chm", "RegExpHelp.htm"

End Sub

Private Sub cmdOk_Click()

Pattern = txtPattern.Text
If chkReplace.Value = vbChecked Then
    ReplaceWith = EscapeChars_ForFindReplace(txtReplace.Text)
Else
    ReplaceWith = vbNullChar
End If

If chkCase.Value = vbChecked Then
    MatchCase = True
Else
    MatchCase = False
End If

If chkWord.Value = vbChecked Then
    Pattern = "\b" & Pattern & "\b"
End If

Me.Hide

End Sub
Private Sub cmdSpecialReplace_Click()

PopUpMenu mnuSpecialChars, 2, cmdSpecialReplace.Left, cmdSpecialReplace.Top + cmdSpecialReplace.Height
txtReplace.SetFocus

End Sub

Private Sub FakeButton1_Click()


End Sub

Private Sub FakeButton2_Click()


End Sub

Private Sub Form_Activate()

'App.HelpFile = App.Path & "\popups.chm::/popups.txt"

txtPattern.SetFocus
TextBoxSelectAll txtPattern

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF1 Then
    btnHelp_Click
End If

End Sub
Private Sub Form_Load()

CButtons Me
'AddBorderToAllTextBoxes Me

'If AxiomSettings.XPThemesSupported Then
'    MakeXPButton btnEscape
'    MakeXPButton btnRegExpEdit
'    MakeXPButton cmdSpecialReplace
'Else
    btnEscape.Caption = "esc"
    btnRegExpEdit.Caption = "RX"
    cmdSpecialReplace.Caption = "..."
'End If

SetLeftMarginsAll Me, 8

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then 'clicked the X button

    Pattern = ""        'Property
    ReplaceWith = ""    'Property
    Cancel = True
    Me.Hide

End If

End Sub
Private Sub mnuBackslash_Click()

txtReplace.SelText = "~~"
txtReplace.SetFocus

End Sub

Private Sub mnuPara_Click()

txtReplace.SelText = "~n"
txtReplace.SetFocus

End Sub

Private Sub mnuTab_Click()

txtReplace.SelText = "~t"
txtReplace.SetFocus

End Sub

Private Sub txtPattern_DblClick()

TextBoxSelectAll txtPattern

End Sub

Private Sub txtReplace_DblClick()

TextBoxSelectAll txtReplace

End Sub

