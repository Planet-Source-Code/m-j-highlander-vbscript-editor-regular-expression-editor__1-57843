VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type CustomCommands
    Caption As String
    Path As String
    Contents As String
End Type

Private CustomCmdsArray() As CustomCommands

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
        
        For idx = Index To UBound(CustomCmdsArray) - 1
            CustomCmdsArray(idx) = CustomCmdsArray(idx + 1)
        Next idx
    
        ReDim Preserve CustomCmdsArray(LBound(CustomCmdsArray) To UBound(CustomCmdsArray) - 1)
        
    End Select
'End If

End Sub

Public Sub Add(ByVal Caption As String, ByVal Path As String)
Dim idx As Integer
Dim objTextFile As CTextFile

ReDim Preserve CustomCmdsArray(0 To UBound(CustomCmdsArray) + 1)

idx = UBound(CustomCmdsArray)
CustomCmdsArray(idx).Caption = Caption
CustomCmdsArray(idx).Path = Path

Set objTextFile = New CTextFile
objTextFile.FileOpen RemoveSlash(App.Path) & "\" & Path, OpenForInput
CustomCmdsArray(idx).Contents = objTextFile.ReadAll
objTextFile.FileClose
Set objTextFile = Nothing

Sort CustomCmdsArray


End Sub

Public Property Get Count() As Integer
    
    Count = UBound(CustomCmdsArray) + 1
    
End Property
Public Property Get Path(ByVal Index As Integer) As String
    
    Path = CustomCmdsArray(Index).Path
    
End Property
Public Property Get Contents(ByVal Index As Integer) As String
    
    Contents = CustomCmdsArray(Index).Contents
    
End Property


Public Property Get Caption(ByVal Index As Integer) As String
    
    Caption = CustomCmdsArray(Index).Caption
    
End Property
Public Function Save()

Dim idx As Long

Dim objIniFile As CINIFileAccess
Set objIniFile = New CINIFileAccess
objIniFile.Filename = RemoveSlash(App.Path) & "\Axiom.ini"
objIniFile.Section = "Custom Functions"
objIniFile.RemoveSection
For idx = LBound(CustomCmdsArray) To UBound(CustomCmdsArray)
    objIniFile.Key = CustomCmdsArray(idx).Caption
    objIniFile.Value = CustomCmdsArray(idx).Path
Next

Set objIniFile = Nothing

End Function
Private Sub Swap(ByRef var1 As CustomCommands, ByRef var2 As CustomCommands)
    Dim x As CustomCommands
    x = var1
    var1 = var2
    var2 = x
End Sub

Private Sub Sort(SortArray() As CustomCommands)

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

            sVal1 = LCase(SortArray(Row).Caption)
            sVal2 = LCase(SortArray(Row + Offset).Caption)
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
Dim objTextFile As New CTextFile


Dim objIniFile As CINIFileAccess
Set objIniFile = New CINIFileAccess
objIniFile.Filename = RemoveSlash(App.Path) & "\Axiom.ini"
objIniFile.Section = "Custom Functions"

sTempArray = objIniFile.EnumKeys

ReDim CustomCmdsArray(LBound(sTempArray) To UBound(sTempArray))

For idx = LBound(sTempArray) To UBound(sTempArray)
    CustomCmdsArray(idx).Caption = sTempArray(idx)
    objIniFile.Key = sTempArray(idx)
    CustomCmdsArray(idx).Path = objIniFile.Value
    'MsgBox CustomCmdsArray(idx).Path
    objTextFile.FileOpen RemoveSlash(App.Path) & "\" & CustomCmdsArray(idx).Path, OpenForInput
    CustomCmdsArray(idx).Contents = objTextFile.ReadAll
    objTextFile.FileClose
Next

Sort CustomCmdsArray

Set objTextFile = Nothing
Set objIniFile = Nothing

End Sub

