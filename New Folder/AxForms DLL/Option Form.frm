VERSION 5.00
Begin VB.Form frmAxOptionForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select an Option"
   ClientHeight    =   2850
   ClientLeft      =   2685
   ClientTop       =   2145
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7155
   Begin Axiom.ShadowedSeperator ShadowedSeperator1 
      Height          =   30
      Left            =   60
      Top             =   2280
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   53
   End
   Begin VB.TextBox txtOther 
      Height          =   310
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   4875
   End
   Begin VB.OptionButton optOptionButton 
      Caption         =   "Option1"
      Height          =   255
      Index           =   4
      Left            =   300
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   6795
   End
   Begin VB.OptionButton optOptionButton 
      Caption         =   "Option1"
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   6795
   End
   Begin VB.OptionButton optOptionButton 
      Caption         =   "Option1"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   6795
   End
   Begin VB.OptionButton optOptionButton 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   6795
   End
   Begin VB.OptionButton optOptionButton 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   6795
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   5835
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1230
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   405
      Left            =   4425
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1230
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Label1"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   6975
   End
End
Attribute VB_Name = "frmAxOptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private bCanceled As Boolean
Public Function Exec(ParamArray Args() As Variant) As String
Dim sLeft As String, sRight As String
Dim iReturn As Integer

On Error Resume Next

Dim idx As Integer

ReDim sTemp(0 To UBound(Args)) As Variant

lblPrompt.Caption = Args(0)
For idx = 0 To UBound(Args) - 1
    'picArgs(idx).Visible = True
    'SplitAt Args(idx), "|", sLeft, sRight
    'lblArgs(idx).Caption = sLeft  'Args(idx)
    'txtArgs(idx).Text = sRight
    optOptionButton(idx).Caption = Replace(Args(idx + 1), "&", "&&")
    If optOptionButton(idx).Caption <> "" Then optOptionButton(idx).Visible = True
    If LCase(optOptionButton(idx).Caption) = "other" Then
        txtOther.Visible = True
        txtOther.Top = optOptionButton(idx).Top
        txtOther.Left = optOptionButton(idx).Left + TextWidth("OTHER") + optOptionButton(idx).Height
    End If
Next

Me.Show vbModal

If bCanceled Then
    Exec = vbNullString
Else
    For idx = 0 To optOptionButton.Count - 1
        If optOptionButton(idx).Value = True Then
            iReturn = idx
            Exit For
        End If
        'sTemp(idx) = txtArgs(idx).Text
    Next
    Exec = Replace(optOptionButton(idx).Caption, "&&", "&")   'sTemp 'Join(sTemp, vbCrLf)
    If LCase(Exec) = "other" Then
        Exec = txtOther.Text
    End If
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
SetLeftMarginsAll Me, 8

End Sub

Private Sub optOptionButton_Click(Index As Integer)

If LCase(optOptionButton(Index).Caption) = "other" Then
    
    txtOther.SetFocus

End If

End Sub
