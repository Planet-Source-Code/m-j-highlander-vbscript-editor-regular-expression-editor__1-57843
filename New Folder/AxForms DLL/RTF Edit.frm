VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSmall_HTML_Editor 
   Caption         =   "Edit HTML"
   ClientHeight    =   5340
   ClientLeft      =   3225
   ClientTop       =   2835
   ClientWidth     =   6585
   Icon            =   "RTF Edit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   6585
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5340
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBottom 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2940
      ScaleHeight     =   495
      ScaleWidth      =   2745
      TabIndex        =   2
      Top             =   4620
      Width           =   2745
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5340
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":0A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":0B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":0CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":0E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":0F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":10F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":124E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":15EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":1752
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":18BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":1A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":24EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RTF Edit.frx":3342
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5953
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      TextRTF         =   $"RTF Edit.frx":34A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Center"
            Object.ToolTipText     =   "Align Center"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Object.ToolTipText     =   "Font"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Text Color"
            Object.ToolTipText     =   "Text Color"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bullet"
            Object.ToolTipText     =   "Bullet"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Date/Time"
            Object.ToolTipText     =   "Date/Time"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Insert File"
            Object.ToolTipText     =   "Insert File"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSmall_HTML_Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bCanceled As Boolean

Public Function Exec(Optional ByVal DefaultText As String, Optional ByVal DefaultText_IsHTML As Boolean = False) As String
On Error Resume Next

If DefaultText_IsHTML Then
    Me.rtbText.TextRTF = HTMLtoRTF(DefaultText)
Else
    Me.rtbText.Text = DefaultText
End If

Me.rtbText.SelStart = 0
Me.rtbText.SelLength = Len(Me.rtbText.Text)
Me.rtbText.SetFocus

Me.Show vbModal

If bCanceled Then
    Exec = vbNullString
Else
    Exec = rtf2html3(Me.rtbText.TextRTF, "+CR")
End If

If Err Then
    Err.Clear
End If

End Function
Private Sub cmdCancel_Click()

bCanceled = True
Me.Hide

End Sub

Private Sub cmdOk_Click()

'Clipboard.Clear
'Clipboard.SetText rtf2html3(rtbText.TextRTF, "+CR")

bCanceled = False
Me.Hide

End Sub
Private Sub Form_Load()
Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next

rtbText.Left = 0
rtbText.Top = Toolbar.Height
rtbText.Width = ScaleWidth
rtbText.Height = ScaleHeight - Toolbar.Height - picBottom.Height
picBottom.Left = ScaleWidth - picBottom.Width
picBottom.Top = Toolbar.Height + rtbText.Height

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

Case "Bold"
If rtbText.SelBold = True Then
    rtbText.SelBold = False
Else
    rtbText.SelBold = True
End If

Case "Italic"
If rtbText.SelItalic = True Then
    rtbText.SelItalic = False
Else
    rtbText.SelItalic = True
End If

Case "Underline"
If rtbText.SelUnderline = True Then
    rtbText.SelUnderline = False
Else
    rtbText.SelUnderline = True
End If

Case "Align Left"
rtbText.SelAlignment = rtfLeft

Case "Align Right"
rtbText.SelAlignment = rtfRight

Case "Align Center"
rtbText.SelAlignment = rtfCenter

Case "Font"
CD1.Flags = cdlCFBoth Or cdlCFEffects
CD1.ShowFont
With rtbText
.SelFontName = CD1.FontName
.SelFontSize = CD1.FontSize
.SelStrikeThru = CD1.FontStrikethru
.SelUnderline = CD1.FontUnderline
.SelBold = CD1.FontBold
.SelItalic = CD1.FontItalic
End With

Case "Text Color"
CD1.ShowColor
With rtbText
.SelColor = CD1.Color
End With

Case "Find"
    MsgBox "To Do..."

Case "Bullet"
rtbText.SelBullet = True

Case "Date/Time"
rtbText.SelText = Date & " " & Time

Case "Insert File"
'    CD1.DialogTitle = "Choose file to insert..."
'    CD1.Filter = "Text File(*.txt)|*.txt|All Files (*.*)|*.*|"
'    CD1.InitDir = App.Path
'    CD1.ShowOpen
'    If Len(CD1.Filename) > 0 Then
'    End If
    MsgBox "To Do..."

End Select

End Sub
