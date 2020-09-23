VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHidden 
   Caption         =   "Hidden Form!!!!!!!!!!!!!!"
   ClientHeight    =   4935
   ClientLeft      =   2280
   ClientTop       =   3015
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   6540
   Begin VB.PictureBox picToLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5100
      Picture         =   "Hidden.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   1380
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picToRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5100
      Picture         =   "Hidden.frx":007E
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   1695
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   2400
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   2130
      Width           =   675
   End
   Begin VB.PictureBox picJustify 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1980
      Picture         =   "Hidden.frx":00FC
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   450
      Width           =   195
   End
   Begin VB.PictureBox picCenter 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1680
      Picture         =   "Hidden.frx":017A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   390
      Width           =   195
   End
   Begin VB.PictureBox picRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1260
      Picture         =   "Hidden.frx":01F8
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   390
      Width           =   195
   End
   Begin VB.PictureBox picLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   990
      Picture         =   "Hidden.frx":0276
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   390
      Width           =   195
   End
   Begin MSComDlg.CommonDialog cdlgColor 
      Left            =   810
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   3
   End
   Begin VB.Menu mnuASCIITable 
      Caption         =   "mnuASCIITable"
      Begin VB.Menu mnuASCIIItem 
         Caption         =   "mnuASCIIItem"
         Index           =   0
      End
   End
   Begin VB.Menu mnuAsciiHtml 
      Caption         =   "mnuAsciiHtml"
      Begin VB.Menu mnuAsciiHtmlItem 
         Caption         =   "mnuAsciiHtmlItem"
         Index           =   0
      End
   End
   Begin VB.Menu mnuPara 
      Caption         =   "mnuPara"
      Begin VB.Menu mnuDefault 
         Caption         =   "Default"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCenter 
         Caption         =   "Center"
      End
      Begin VB.Menu mnuJustify 
         Caption         =   "Justify"
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "Left"
      End
      Begin VB.Menu mnuRight 
         Caption         =   "Right"
      End
   End
   Begin VB.Menu mnuColors 
      Caption         =   "mnuColors"
      Begin VB.Menu mnuItem 
         Caption         =   "mnuItem"
         Index           =   0
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColorCustom 
         Caption         =   "Custom..."
      End
   End
   Begin VB.Menu mnuHeader 
      Caption         =   "mnuHeader"
      Begin VB.Menu mnuHeaderN 
         Caption         =   "Header 1"
         Index           =   0
      End
      Begin VB.Menu mnuHeaderN 
         Caption         =   "Header 2"
         Index           =   1
      End
      Begin VB.Menu mnuHeaderN 
         Caption         =   "Header 3"
         Index           =   2
      End
      Begin VB.Menu mnuHeaderN 
         Caption         =   "Header 4"
         Index           =   3
      End
      Begin VB.Menu mnuHeaderN 
         Caption         =   "Header 5"
         Index           =   4
      End
      Begin VB.Menu mnuHeaderN 
         Caption         =   "Header 6"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPads 
      Caption         =   "mnuPads"
      Begin VB.Menu mnuOut_to_In 
         Caption         =   "&Move Output to Input"
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearOutput 
         Caption         =   "&Clear Output"
      End
   End
End
Attribute VB_Name = "frmHidden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CreateColorsMenu()
Dim idx As Long

For idx = 1 To 15
    Load mnuItem(idx)
    Load picColor(idx)
Next

InitColors

For idx = 0 To 15
    mnuItem(idx).Caption = " " & ga_ColorVals(idx).ColorName
    picColor(idx).Line (0, 0)-(picColor(idx).Width, picColor(idx).Height), ga_ColorVals(idx).ColorLong, BF
    picColor(idx).Picture = picColor(idx).Image
    SetMenuIcon Me.hwnd, 3, idx, picColor(idx).Picture
Next

End Sub
Private Sub Form_Load()
Dim idx As Integer

SetMenuIcon Me.hwnd, 2, 2, picCenter
SetMenuIcon Me.hwnd, 2, 3, picJustify
SetMenuIcon Me.hwnd, 2, 4, picLeft
SetMenuIcon Me.hwnd, 2, 5, picRight
SetMenuIcon Me.hwnd, 5, 0, picToLeft

mnuLeft.Caption = "  Left"
mnuRight.Caption = "  Right"
mnuJustify.Caption = "  Justify"
mnuCenter.Caption = "  Center"

mnuOut_to_In.Caption = "&Move Output to Input" & vbTab & "Alt+Left"


CreateColorsMenu


'-------------------RegExp
For idx = 128 To 255
    Load frmHidden.mnuASCIIItem(idx)
    frmHidden.mnuASCIIItem(idx).Caption = Chr$(idx)
Next
frmHidden.mnuASCIIItem(0).Visible = False

For idx = 20 To 120 Step 20
    MenuCols frmHidden, 0, idx
Next

'-------------------HTML
Load mnuAsciiHtmlItem(32): mnuAsciiHtmlItem(32).Caption = "space": mnuAsciiHtmlItem(32).Tag = Chr$(32)
Load mnuAsciiHtmlItem(34): mnuAsciiHtmlItem(34).Caption = Chr$(34): mnuAsciiHtmlItem(34).Tag = Chr$(34)
Load mnuAsciiHtmlItem(38): mnuAsciiHtmlItem(38).Caption = "&&": mnuAsciiHtmlItem(38).Tag = Chr$(38)
Load mnuAsciiHtmlItem(60): mnuAsciiHtmlItem(60).Caption = Chr$(60): mnuAsciiHtmlItem(60).Tag = Chr$(60)
Load mnuAsciiHtmlItem(62): mnuAsciiHtmlItem(62).Caption = Chr$(62): mnuAsciiHtmlItem(62).Tag = Chr$(62)
'seperator only
Load mnuAsciiHtmlItem(63): mnuAsciiHtmlItem(63).Caption = "-": mnuAsciiHtmlItem(63).Tag = Chr$(0)

For idx = 128 To 255
    Select Case idx
        Case 129, 141, 143, 144, 157
            'dont add
        Case Else
            Load mnuAsciiHtmlItem(idx)
            mnuAsciiHtmlItem(idx).Caption = Chr$(idx)
            mnuAsciiHtmlItem(idx).Tag = Chr$(idx)
    End Select
Next
frmHidden.mnuAsciiHtmlItem(0).Visible = False

For idx = 20 To 120 Step 20
    MenuCols frmHidden, 1, idx
Next

End Sub
Private Sub mnuAsciiHtmlItem_Click(Index As Integer)
Dim sName As String

sName = EntityInfo((Asc(mnuAsciiHtmlItem(Index).Tag))).Name

If sName = "" Then    'use number instead
    sName = "&#" & EntityInfo((Asc(mnuAsciiHtmlItem(Index).Tag))).Code & ";"
End If

'frmAxiomHTMLMain.MainText.SelText = sName

End Sub
Private Sub mnuASCIIItem_Click(Index As Integer)

 frmRegExpGenerator.txtRegExp.SelText = "\x" & Hex$(Asc(mnuASCIIItem(Index).Caption))
 
End Sub

Private Sub mnuCenter_Click()
'EncloseTag frmAxiomHTMLMain.MainText, "<p align=""center"">", "</p>"
End Sub

Private Sub mnuClearOutput_Click()

'frmAxiomHTMLMain.OutputText.Text = ""

End Sub
Private Sub mnuColorCustom_Click()
Dim sColor As String

On Error Resume Next

cdlgColor.ShowColor

If Err Then
    Err.Clear
    Exit Sub
End If

sColor = ColorToHex(cdlgColor.Color)
'EncloseTag frmAxiomHTMLMain.MainText, "", sColor


End Sub

Private Sub mnuDefault_Click()

'EncloseTag frmAxiomHTMLMain.MainText, "<p>", "</p>"

End Sub

Private Sub mnuHeaderN_Click(Index As Integer)
Dim sHeader As String

sHeader = Format(Index + 1)

'EncloseTag frmAxiomHTMLMain.MainText, "<h" & sHeader & ">", "</h" & sHeader & ">"

End Sub
Private Sub mnuItem_Click(Index As Integer)

'EncloseTag frmAxiomHTMLMain.MainText, "", Quote & ga_ColorVals(Index).ColorName & Quote
'EncloseTag frmAxiomHTMLMain.MainText, "", ga_ColorVals(Index).ColorName

End Sub
Private Sub mnuJustify_Click()
'EncloseTag frmAxiomHTMLMain.MainText, "<p align=""justify"">", "</p>"
End Sub


Private Sub mnuLeft_Click()

'EncloseTag frmAxiomHTMLMain.MainText, "<p align=""left"">", "</p>"

End Sub


Private Sub mnuOut_to_In_Click()

'frmAxiomHTMLMain.mnuOut_to_In_Click

End Sub
Private Sub mnuRight_Click()
'EncloseTag frmAxiomHTMLMain.MainText, "<p align=""right"">", "</p>"
End Sub


