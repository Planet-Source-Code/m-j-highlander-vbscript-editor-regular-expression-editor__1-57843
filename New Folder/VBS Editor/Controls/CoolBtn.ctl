VERSION 5.00
Begin VB.UserControl CoolButton 
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   39
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ToolboxBitmap   =   "CoolBtn.ctx":0000
End
Attribute VB_Name = "CoolButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========= Local ==========================
Public Enum CoolBtnBorderStyle
    cbtn_RAISED_OUTER = &H1
    cbtn_RAISED_INNER = &H4
    cbtn_RAISED = &H5
    cbtn_EDGE_ETCHED = &H6
    cbtn_EDGE_BUMP = &H9
End Enum

Public Event Click()
Public Event MouseDown()
Public Event MouseUp()
Public Event MouseMove()

Private m_BorderStyle As CoolBtnBorderStyle
Private ms_Caption As String
Private m_Picture As Picture
Private mb_ShowFocusRect As Boolean
Private m_MaskColor As OLE_COLOR


'========== API Declerations ==============
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNFACE = 15
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_GRAYTEXT = 17
Private Const PS_SOLID = 0

Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
' CUSTOMIZED EDGE styles (TO COMBINE THE CONSTANTS)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

Private Const BF_FLAT = &H4000
Private Const BF_LEFT = &H1
Private Const BF_MONO = &H8000
Private Const BF_MIDDLE = &H800
Private Const BF_RIGHT = &H4
Private Const BF_SOFT = &H1000
Private Const BF_TOP = &H2
Private Const BF_ADJUST = &H2000
Private Const BF_BOTTOM = &H8
Private Const BF_DIAGONAL = &H10
Private Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Sub TransBlt(OutDstDC, DstDC, SrcDC, SrcRect As RECT, DstX, DstY, TransColor As Long)
  'DstDC=Device context into which image must be drawn transparently
  'OutDstDC=Device context into image is actually drawn, even though it is made transparent in terms of DstDC
  'Src=Device context of source to be made transparent in color TransColor
  'SrcRect=rectangular region within SrcDC to be made transparent in terms of DstDC, and drawn to OutDstDC
  'DstX, DstY =coordinates in OutDstDC (and DstDC) where tranparent bitmap must go
  
  Rem In most cases, OutDstDC and DstDC will be the same
  
  Dim nRet As Long, W As Integer, H As Integer
  Dim MonoMaskDC As Long, hMonoMask As Long
  Dim MonoInvDC As Long, hMonoInv As Long
  Dim ResultDstDC As Long, hResultDst As Long
  Dim ResultSrcDC As Long, hResultSrc As Long
  Dim hPrevMask As Long, hPrevInv As Long, hPrevSrc As Long, hPrevDst As Long
  W = SrcRect.Right - SrcRect.Left + 1
  H = SrcRect.Bottom - SrcRect.Top + 1
  
  'create monochrome mask and inverse masks
  MonoMaskDC = CreateCompatibleDC(DstDC)
  MonoInvDC = CreateCompatibleDC(DstDC)
  hMonoMask = CreateBitmap(W, H, 1, 1, ByVal 0&)
  hMonoInv = CreateBitmap(W, H, 1, 1, ByVal 0&)
  hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
  hPrevInv = SelectObject(MonoInvDC, hMonoInv)
  
  'create keeper DCs and bitmaps
  ResultDstDC = CreateCompatibleDC(DstDC)
  ResultSrcDC = CreateCompatibleDC(DstDC)
  hResultDst = CreateCompatibleBitmap(DstDC, W, H)
  hResultSrc = CreateCompatibleBitmap(DstDC, W, H)
  hPrevDst = SelectObject(ResultDstDC, hResultDst)
  hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
  
  'copy src to monochrome mask
  Dim OldBC As Long
  OldBC = SetBkColor(SrcDC, TransColor)
  nRet = BitBlt(MonoMaskDC, 0, 0, W, H, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
  TransColor = SetBkColor(SrcDC, OldBC)
  
  'create inverse of mask
  nRet = BitBlt(MonoInvDC, 0, 0, W, H, MonoMaskDC, 0, 0, vbNotSrcCopy)
  
  'get background
  nRet = BitBlt(ResultDstDC, 0, 0, W, H, DstDC, DstX, DstY, vbSrcCopy)
  'AND with Monochrome mask
  nRet = BitBlt(ResultDstDC, 0, 0, W, H, MonoMaskDC, 0, 0, vbSrcAnd)
  'get overlapper
  nRet = BitBlt(ResultSrcDC, 0, 0, W, H, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
  'AND with inverse monochrome mask
  nRet = BitBlt(ResultSrcDC, 0, 0, W, H, MonoInvDC, 0, 0, vbSrcAnd)
  'XOR these two
  nRet = BitBlt(ResultDstDC, 0, 0, W, H, ResultSrcDC, 0, 0, vbSrcInvert)
  
  'output results
  nRet = BitBlt(OutDstDC, DstX, DstY, W, H, ResultDstDC, 0, 0, vbSrcCopy)
  
  'clean up
  hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
  DeleteObject hMonoMask
  hMonoInv = SelectObject(MonoInvDC, hPrevInv)
  DeleteObject hMonoInv
  hResultDst = SelectObject(ResultDstDC, hPrevDst)
  DeleteObject hResultDst
  hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
  DeleteObject hResultSrc
  DeleteDC MonoMaskDC
  DeleteDC MonoInvDC
  DeleteDC ResultDstDC
  DeleteDC ResultSrcDC

End Sub
Public Property Get BorderStyle() As CoolBtnBorderStyle
       BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As CoolBtnBorderStyle)
       m_BorderStyle = NewValue
       PropertyChanged "BorderStyle"
       UserControl_Paint

End Property
Public Property Get MaskColor() As OLE_COLOR
       MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal NewValue As OLE_COLOR)
       m_MaskColor = NewValue
       PropertyChanged "MaskColor"
       UserControl_Paint
End Property
Public Property Let Value(ByVal bNewValue As Boolean)
'WRITE-ONLY property!
       
       If bNewValue Then
            RaiseEvent Click
       End If

End Property
Public Sub About()
Attribute About.VB_UserMemId = -552
    MsgBox "CoolButton by highlander <mdsy@ny.com", vbInformation, "About"
End Sub
Public Property Get ShowFocusRect() As Boolean
       ShowFocusRect = mb_ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal bNewValue As Boolean)
       mb_ShowFocusRect = bNewValue
End Property

Private Sub DrawPic(bDown As Boolean)
    Dim ButtonTop As Long
    Dim BkColor As Long
    Dim TLng1 As Double, TLng2 As Double
    Dim cx As Long, cy As Long
    Dim W As Long, H As Long
    
    Dim hMemDC As Long
    Dim hOldBmp As Long
    
Dim r As RECT

r.Left = 0
r.Top = 0
r.Right = ScaleWidth
r.Bottom = ScaleHeight

    If (m_Picture Is Nothing) Then
        'no pic =  m_Picture.Width = 0
    Else
            'Check picture dimensions
        cx = UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels)
        cy = UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels)
        W = 0 '(UserControl.ScaleWidth - cx) / 2
        H = 0 '(UserControl.ScaleHeight - cy) / 2
        If bDown = True Then
            W = W + 1
            H = H + 1
        End If
        hMemDC = CreateCompatibleDC(UserControl.hdc)
        hOldBmp = SelectObject(hMemDC, m_Picture.Handle)
                
        'BitBlt UserControl.hdc, W, H, CX, CY, hMemDC, 0&, 0&, vbSrcCopy
        TransBlt UserControl.hdc, UserControl.hdc, hMemDC, r, 0&, 0&, MaskColor
        'StretchBlt UserControl.hdc, 0&, ButtonTop, 12&, 24&, hMemDC, 0&, 0&, cx, cy, vbSrcCopy
        
        SelectObject hMemDC, hOldBmp
        DeleteDC hMemDC
    End If

End Sub

Public Property Get BackPicture() As Picture
    Set BackPicture = m_Picture
End Property

Public Property Set BackPicture(ByVal picNewValue As Picture)
   
  Set m_Picture = picNewValue
       PropertyChanged "BackPicture"
       UserControl_Paint

End Property

Public Property Get Font() As IFontDisp
    
    Set Font = UserControl.Font
    
    Font.Bold = UserControl.Font.Bold
    Font.Charset = UserControl.Font.Charset
    Font.Italic = UserControl.Font.Italic
    Font.Name = UserControl.Font.Name
    Font.SIZE = UserControl.Font.SIZE
    Font.Strikethrough = UserControl.Font.Strikethrough
    Font.Underline = UserControl.Font.Underline
    Font.Weight = UserControl.Font.Weight

End Property

Public Property Set Font(ByVal fontNewValue As IFontDisp)
    'chk.Font = fontNewValue
    
    UserControl.Font.Bold = fontNewValue.Bold
    UserControl.Font.Charset = fontNewValue.Charset
    UserControl.Font.Italic = fontNewValue.Italic
    UserControl.Font.Name = fontNewValue.Name
    UserControl.Font.SIZE = fontNewValue.SIZE
    UserControl.Font.Strikethrough = fontNewValue.Strikethrough
    UserControl.Font.Underline = fontNewValue.Underline
    UserControl.Font.Weight = fontNewValue.Weight
    PropertyChanged "Font"
    UserControl_Paint

End Property


Public Property Get ForeColor() As OLE_COLOR
       ForeColor = UserControl.ForeColor
        
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
       UserControl.ForeColor = NewValue
       PropertyChanged "ForeColor"
       UserControl_Paint

End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501

       BackColor = UserControl.BackColor

End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)

       UserControl.BackColor = NewValue
       PropertyChanged "BackColor"
       UserControl_Paint

End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
       Caption = ms_Caption
End Property

Public Property Let Caption(ByVal sNewValue As String)
       ms_Caption = sNewValue
       PropertyChanged "Caption"
       UserControl_Paint
End Property

Private Sub UserControl_Click()

RaiseEvent Click

End Sub
Private Sub UserControl_GotFocus()

Dim r As RECT

r.Left = 3
r.Top = 3
r.Right = ScaleWidth - 3
r.Bottom = ScaleHeight - 3

If ShowFocusRect Then DrawFocusRect hdc, r
'Refresh



End Sub

Private Sub UserControl_InitProperties()

Caption = Extender.Name
BorderStyle = BDR_RAISEDINNER
MaskColor = vbWhite

'Set Picture = Nothing

End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)


If KeyAscii = vbKeySpace Then
    RaiseEvent Click
End If


End Sub
Private Sub UserControl_LostFocus()
UserControl_Paint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> vbLeftButton Then Exit Sub

Dim r As RECT
Cls

DrawPic True

r.Left = 0
r.Top = 0
r.Right = ScaleWidth
r.Bottom = ScaleHeight
DrawEdge hdc, r, BDR_SUNKENOUTER, BF_RECT

r.Left = 2
r.Top = 2
r.Right = ScaleWidth
r.Bottom = ScaleHeight
DrawText hdc, Caption, Len(Caption), r, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER

RaiseEvent MouseDown

End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

RaiseEvent MouseMove

End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> vbLeftButton Then Exit Sub

Dim r As RECT
Cls

DrawPic False

r.Left = 0
r.Top = 0
r.Right = ScaleWidth
r.Bottom = ScaleHeight
DrawEdge hdc, r, BorderStyle, BF_RECT

DrawText hdc, Caption, Len(Caption), r, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER


RaiseEvent MouseUp

UserControl_GotFocus

End Sub
Private Sub UserControl_Paint()
Dim r As RECT

Cls
DrawPic False

r.Left = 0
r.Top = 0
r.Right = ScaleWidth
r.Bottom = ScaleHeight

DrawEdge hdc, r, BorderStyle, BF_RECT

DrawText hdc, Caption, Len(Caption), r, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Caption = PropBag.ReadProperty("Caption", Extender.Name)

BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
MaskColor = PropBag.ReadProperty("MaskColor", vbWhite)

ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", True)

BorderStyle = PropBag.ReadProperty("BorderStyle", BDR_RAISEDINNER)

Font.Bold = PropBag.ReadProperty("Font.Bold", Ambient.Font.Bold)
Font.Charset = PropBag.ReadProperty("Font.Charset", Ambient.Font.Charset)
Font.Italic = PropBag.ReadProperty("Font.Italic", Ambient.Font.Italic)
Font.Name = PropBag.ReadProperty("Font.Name", Ambient.Font.Name)
Font.SIZE = PropBag.ReadProperty("Font.Size", Ambient.Font.SIZE)
Font.Strikethrough = PropBag.ReadProperty("Font.Strikethrough", Ambient.Font.Strikethrough)
Font.Underline = PropBag.ReadProperty("Font.Underline", Ambient.Font.Underline)
Font.Weight = PropBag.ReadProperty("Font.Weight", Ambient.Font.Weight)

Set BackPicture = PropBag.ReadProperty("BackPicture", Nothing)
'Set m_picture = PropBag.ReadProperty("Picture", Nothing)

UserControl_Paint

End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

PropBag.WriteProperty "Caption", Caption, Extender.Name

PropBag.WriteProperty "BackColor", BackColor, Ambient.BackColor
PropBag.WriteProperty "ForeColor", ForeColor, Ambient.ForeColor
PropBag.WriteProperty "MaskColor", MaskColor, vbWhite

PropBag.WriteProperty "ShowFocusRect", ShowFocusRect, True
PropBag.WriteProperty "BorderStyle", BorderStyle, BDR_RAISEDINNER

PropBag.WriteProperty "Font.bold", Font.Bold, Ambient.Font.Bold
PropBag.WriteProperty "Font.Charset", Font.Charset, Ambient.Font.Charset
PropBag.WriteProperty "Font.Italic", Font.Italic, Ambient.Font.Italic
PropBag.WriteProperty "Font.Name", Font.Name, Ambient.Font.Name
PropBag.WriteProperty "Font.size", Font.SIZE, Ambient.Font.SIZE
PropBag.WriteProperty "Font.Strikethrough", Font.Strikethrough, Ambient.Font.Strikethrough
PropBag.WriteProperty "Font.Underline", Font.Underline, Ambient.Font.Underline
PropBag.WriteProperty "Font.Weight", Font.Weight, Ambient.Font.Weight

PropBag.WriteProperty "BackPicture", BackPicture, Nothing


End Sub
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    
    Enabled = UserControl.Enabled

End Property
Public Property Let Enabled(ByVal bNewValue As Boolean)
    
    UserControl.Enabled = bNewValue

End Property
