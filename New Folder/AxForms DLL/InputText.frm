VERSION 5.00
Begin VB.Form frmInputText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Text"
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
   Begin VB.TextBox txtText 
      Height          =   3750
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   540
      Width           =   6450
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   5235
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1230
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   405
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   2
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
Attribute VB_Name = "frmInputText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bCanceled As Boolean
Public Function Exec(ByVal Caption As String, Optional ByVal DefaultText As String) As String
On Error Resume Next

Me.lblCaption.Caption = Caption
Me.txtText.Text = DefaultText
Me.txtText.SelStart = 0
Me.txtText.SelLength = Len(Me.txtText.Text)
Me.txtText.SetFocus

Me.Show vbModal

If bCanceled Then
    Exec = vbNullString
Else

    Exec = Me.txtText.Text
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
SetLeftMargin txtText, 8

End Sub
