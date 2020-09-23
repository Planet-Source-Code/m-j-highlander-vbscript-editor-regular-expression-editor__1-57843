VERSION 5.00
Begin VB.Form frmInputBoxes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Values"
   ClientHeight    =   4860
   ClientLeft      =   2205
   ClientTop       =   3180
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6585
   Begin Axiom.ShadowedSeperator ShadowedSeperator1 
      Height          =   30
      Left            =   120
      Top             =   4140
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   53
   End
   Begin VB.PictureBox picArgs 
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   4
      Left            =   90
      ScaleHeight     =   765
      ScaleWidth      =   6390
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   6390
      Begin VB.TextBox txtArgs 
         Height          =   330
         Index           =   4
         Left            =   30
         TabIndex        =   15
         Top             =   300
         Width           =   6210
      End
      Begin VB.Label lblArgs 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Index           =   4
         Left            =   45
         TabIndex        =   16
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.PictureBox picArgs 
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   3
      Left            =   90
      ScaleHeight     =   765
      ScaleWidth      =   6390
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2550
      Visible         =   0   'False
      Width           =   6390
      Begin VB.TextBox txtArgs 
         Height          =   330
         Index           =   3
         Left            =   30
         TabIndex        =   12
         Top             =   300
         Width           =   6210
      End
      Begin VB.Label lblArgs 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Index           =   3
         Left            =   45
         TabIndex        =   13
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.PictureBox picArgs 
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   2
      Left            =   90
      ScaleHeight     =   765
      ScaleWidth      =   6390
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1740
      Visible         =   0   'False
      Width           =   6390
      Begin VB.TextBox txtArgs 
         Height          =   330
         Index           =   2
         Left            =   30
         TabIndex        =   9
         Top             =   300
         Width           =   6210
      End
      Begin VB.Label lblArgs 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   10
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.PictureBox picArgs 
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   1
      Left            =   90
      ScaleHeight     =   765
      ScaleWidth      =   6390
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   6390
      Begin VB.TextBox txtArgs 
         Height          =   330
         Index           =   1
         Left            =   30
         TabIndex        =   6
         Top             =   300
         Width           =   6210
      End
      Begin VB.Label lblArgs 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   7
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.PictureBox picArgs 
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   0
      Left            =   90
      ScaleHeight     =   765
      ScaleWidth      =   6390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   6390
      Begin VB.TextBox txtArgs 
         Height          =   330
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   300
         Width           =   6210
      End
      Begin VB.Label lblArgs 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   3
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   5235
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1230
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   405
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1230
   End
End
Attribute VB_Name = "frmInputBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private bCanceled As Boolean
Public Function Exec(ParamArray Args() As Variant) As Variant
Dim sLeft As String, sRight As String

On Error Resume Next

Dim idx As Integer

ReDim sTemp(0 To UBound(Args)) As Variant

For idx = 0 To UBound(Args)
    If Args(idx) <> "" Then
        picArgs(idx).Visible = True
        SplitAt Args(idx), "|", sLeft, sRight
        lblArgs(idx).Caption = sLeft  'Args(idx)
        txtArgs(idx).Text = sRight
    End If
Next

Me.Show vbModal

If bCanceled Then
    Exec = vbNullString
Else
    For idx = 0 To UBound(Args)
        sTemp(idx) = txtArgs(idx).Text
    Next
    Exec = sTemp 'Join(sTemp, vbCrLf)
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
Private Sub txtArgs_GotFocus(Index As Integer)

TextBoxSelectAll txtArgs(Index)

End Sub
