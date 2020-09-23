VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextSnippet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type TextSnippet
    Name As String
    Text As String
End Type

Private SnippetsArray() As TextSnippet

Private mi_UpperIndex As Integer
Public Property Let Text(ByVal Index As Integer, ByVal sNewVal As String)
    
    SnippetsArray(Index).Text = sNewVal
    
End Property


Public Property Let Name(ByVal Index As Integer, ByVal sNewVal As String)
    
    SnippetsArray(Index).Name = sNewVal
    
End Property

Private Property Let Count(ByVal iNewValue As Integer)
    
    mi_UpperIndex = iNewValue
    
End Property
Public Sub RemoveItem(Index As Long)
Const Err_Subscript_Out_Of_Range = 9

Dim idx As Long

'If UBound(m_TheArray) = 0 Then
'    m_TheArray(0) = ""
'    mb_ListIsEmpty = True
'Else
    Select Case Index
    Case Is > Count - 1, Is < 0
        Err.Raise Err_Subscript_Out_Of_Range
    Case Else
        
        For idx = Index To UBound(SnippetsArray) - 1
            SnippetsArray(idx) = SnippetsArray(idx + 1)
        Next idx
    
        If UBound(SnippetsArray) > 0 Then
            ReDim Preserve SnippetsArray(LBound(SnippetsArray) To UBound(SnippetsArray) - 1)
            Count = UBound(SnippetsArray) + 1
        Else
            'Erase SnippetsArray
            Count = 0
        End If

    End Select
'End If

End Sub
Public Sub Add(ByVal Name As String, ByVal Text As String)
Dim idx As Integer


If Count > 0 Then
    ReDim Preserve SnippetsArray(0 To UBound(SnippetsArray) + 1)
End If


idx = UBound(SnippetsArray)
SnippetsArray(idx).Name = Name
SnippetsArray(idx).Text = Text


Count = Count + 1


Sort SnippetsArray

End Sub
Public Property Get Count() As Integer
    
    Count = mi_UpperIndex
    
End Property
Public Property Get Text(ByVal Index As Integer) As String
    
    Text = SnippetsArray(Index).Text
    
End Property
Public Property Get Name(ByVal Index As Integer) As String
    
    Name = SnippetsArray(Index).Name
    
End Property

Public Sub Save()

Dim idx As Long



Dim objIniFile As CINIFileAccess
Set objIniFile = New CINIFileAccess
objIniFile.Filename = RemoveSlash(App.Path) & "\Snippets.ini"
objIniFile.Section = "Text Snippets"
objIniFile.RemoveSection

If Count > 0 Then
    For idx = LBound(SnippetsArray) To UBound(SnippetsArray)
        objIniFile.Key = SnippetsArray(idx).Name
        objIniFile.Value = EscapeSnippetChars(SnippetsArray(idx).Text)
    Next
End If

Set objIniFile = Nothing

End Sub
Private Sub Swap(ByRef var1 As TextSnippet, ByRef var2 As TextSnippet)
    Dim X As TextSnippet
    X = var1
    var1 = var2
    var2 = X
End Sub

Private Sub Sort(SortArray() As TextSnippet)

Dim sVal1 As String, sVal2 As String

Dim Row As Long
Dim MaxRow As Long
Dim MinRow As Long
Dim Swtch As Long
Dim Limit As Long
Dim Offset As Long

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0

      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit

            sVal1 = LCase(SortArray(Row).Name)
            sVal2 = LCase(SortArray(Row + Offset).Name)
            If sVal1 > sVal2 Then
               Swap SortArray(Row), SortArray(Row + Offset)

               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub
Public Sub Load()
Dim idx As Long, vTemp As Variant
Dim sTempArray() As String


Dim objIniFile As CINIFileAccess
Set objIniFile = New CINIFileAccess
objIniFile.Filename = RemoveSlash(App.Path) & "\Snippets.ini"
objIniFile.Section = "Text Snippets"

sTempArray = objIniFile.EnumKeys

If UBound(sTempArray) < 0 Then
    Count = 0
    Exit Sub
End If

ReDim SnippetsArray(LBound(sTempArray) To UBound(sTempArray))

For idx = LBound(sTempArray) To UBound(sTempArray)
    SnippetsArray(idx).Name = sTempArray(idx)
    objIniFile.Key = sTempArray(idx)
    SnippetsArray(idx).Text = UnEscapeSnippetChars(objIniFile.Value)
Next

Count = UBound(sTempArray) + 1

Sort SnippetsArray

Set objIniFile = Nothing

End Sub
Private Sub Class_Initialize()

ReDim SnippetsArray(0 To 0)
Count = 0

End Sub

Private Sub Class_Terminate()
'MsgBox "U"
End Sub


