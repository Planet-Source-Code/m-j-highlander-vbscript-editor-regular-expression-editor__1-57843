VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewText 
   ClientHeight    =   7260
   ClientLeft      =   2550
   ClientTop       =   1845
   ClientWidth     =   8085
   Icon            =   "ViewText.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   484
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   539
   Begin RichTextLib.RichTextBox txtView 
      Height          =   5955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   10504
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"ViewText.frx":014A
   End
   Begin VB.Image imgScript 
      Height          =   240
      Left            =   1140
      Picture         =   "ViewText.frx":01DB
      Top             =   6420
      Width           =   240
   End
   Begin VB.Image imgPage 
      Height          =   240
      Left            =   420
      Picture         =   "ViewText.frx":0325
      Top             =   6360
      Width           =   240
   End
End
Attribute VB_Name = "frmViewText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function View(ByVal sText As String, Optional ByVal sCaption As String = "", Optional ByVal ColorizeVB As Boolean = False)

txtView.SelIndent = 10

If ColorizeVB Then
    SyntaxColorizeAll txtView, sText ', QBColor(2), RGB(165, 105, 195), vbBlue
    Icon = imgScript.Picture
Else
    txtView.Text = sText
    Icon = imgPage.Picture
End If

txtView.SelStart = 0

Me.Caption = sCaption

Me.Show vbModal

End Function
Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then Unload Me


End Sub

Private Sub Form_Load()

SetWordWrap txtView, False
'LeftMargin(txtView) = 6

End Sub
Private Sub Form_Resize()

On Error Resume Next

txtView.Top = 0
txtView.Left = 0
txtView.Width = Me.ScaleWidth
txtView.Height = Me.ScaleHeight

End Sub

Private Sub txtView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim lResult As Long, lCurPos As Long

If Button <> vbRightButton Then
    Exit Sub
End If


lResult = PopUp("Copy Selection", "-", "Copy All")
'                 1      2               3

Select Case lResult
    
    Case 1: 'Copy
        Copy txtView.hWnd
        
    Case 3: 'Select All
        SelectAll txtView.hWnd
        Copy txtView.hWnd
    
    Case Else
        'overkill, we never get here!
End Select


End Sub
