VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInputHTML 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter HTML"
   ClientHeight    =   4860
   ClientLeft      =   2685
   ClientTop       =   2145
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6585
   Begin RichTextLib.RichTextBox rchHTML 
      Height          =   3735
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6588
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      TextRTF         =   $"InputHTML.frx":0000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   5235
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1230
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   405
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1230
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "lblCaption"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   690
   End
End
Attribute VB_Name = "frmInputHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bCanceled As Boolean
Public Function Exec(ByVal Caption As String, Optional ByVal DefaultText As String) As String
On Error Resume Next

Me.lblCaption.Caption = Caption
Me.rchHTML.Text = DefaultText
Me.rchHTML.SelStart = 0
Me.rchHTML.SelLength = Len(Me.rchHTML.Text)
Me.rchHTML.SetFocus

Me.Show vbModal

If bCanceled Then
    Exec = vbNullString
Else

    Exec = Me.rchHTML.Text
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

bCanceled = False
Me.Hide

End Sub
Private Sub Form_Load()

CButtons Me
'SetLeftMargin rchHTML, 8

End Sub
Private Sub RichTextBox1_Change()
    
    Bubble_Change_HTML rchHTML

End Sub


Private Sub rchHTML_Change()

    Bubble_Change_HTML rchHTML

End Sub

Private Sub rchHTML_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Bubble_KeyDown KeyCode, Shift, rchHTML

End Sub
