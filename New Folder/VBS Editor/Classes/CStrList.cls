VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private m_TheArray() As String
Private mb_ListIsEmpty As Boolean

Public Sub ConcatLines(StrList As CStrList, Optional ByVal Delimiter As String = vbCrLf)
Dim MinCount As Long, idx As Long

MinCount = Min(Me.ItemCount, StrList.ItemCount)

For idx = 0 To MinCount - 1
    Item(idx) = Item(idx) & Delimiter & StrList.Item(idx)
Next

End Sub

Public Sub ExchangeItems(ByVal Index1 As Long, ByVal Index2 As Long)

      Swap m_TheArray(Index1), m_TheArray(Index2)

End Sub

Public Sub AssignTo(StrList As CStrList)
'Copy items, NOT set as same object

'Set StrList = Me

 StrList.Text = Text

End Sub

Public Sub AppendStringList(StrList As CStrList)
Dim idx As Integer

For idx = 0 To StrList.ItemCount - 1
    AddItem StrList.Item(idx)
Next

End Sub

Public Sub AssignStringList(StrList As CStrList)
'Copy items, NOT set as same object

Text = StrList.Text

End Sub

Public Sub SaveFile(ByVal FileName As String)
Dim iFF As Integer


iFF = FreeFile
Open FileName For Binary Access Write As #iFF
    Put #iFF, , Text
Close iFF

End Sub
Public Sub Clear()
       
       ReDim m_TheArray(0 To 0)
       mb_ListIsEmpty = True
    
End Sub

Public Sub LoadFile(ByVal FileName As String)
Dim iFF As Integer
Dim sTemp As String


iFF = FreeFile
Open FileName For Binary Access Read As #iFF
    sTemp = Input$(LOF(iFF), iFF)
Close iFF


Text = sTemp

End Sub

Public Sub RemoveItem(Index As Long)
Const Err_Subscript_Out_Of_Range = 9

Dim idx As Long

If UBound(m_TheArray) = 0 Then
    m_TheArray(0) = ""
    mb_ListIsEmpty = True
Else
    Select Case Index
    Case Is > ItemCount - 1, Is < 0
        Err.Raise Err_Subscript_Out_Of_Range
    Case Else
        
        For idx = Index To UBound(m_TheArray) - 1
            m_TheArray(idx) = m_TheArray(idx + 1)
        Next idx
    
        ReDim Preserve m_TheArray(LBound(m_TheArray) To UBound(m_TheArray) - 1)
        
    End Select
End If

End Sub

Private Sub Text2Array(ByVal sText As String, ByRef sArray() As String)
'Dim X As Long
    ' sText should not be Empty:
    ' Check for it in the calling routine.
    
    Dim vTmpArray As Variant
    Dim idx As Integer
    
    vTmpArray = Split(sText, vbCrLf)
    
'If vTmpArray(UBound(vTmpArray)) = "" Then x = 1 Else x = 0
    
    ReDim sArray(LBound(vTmpArray) To UBound(vTmpArray))
    
    For idx = LBound(vTmpArray) To UBound(vTmpArray)
    
        sArray(idx) = vTmpArray(idx)
'        MsgBox sArray(idx)
    Next idx

End Sub

Private Function Array2Text(sArray() As String) As String

Array2Text = Join(sArray, vbCrLf)

End Function
Public Sub InsertItem(TextLine As String, Index As Long)
Dim idx As Long

If Index > ItemCount Then 'Add to end of list
    AddItem TextLine
Else
    'Add one cell to end of Array
    ReDim Preserve m_TheArray(LBound(m_TheArray) To UBound(m_TheArray) + 1)
    'Shift values. up to down , else values will be lost
    For idx = UBound(m_TheArray) To Index + 1 Step -1
        m_TheArray(idx) = m_TheArray(idx - 1)
    Next idx
    'Add value in position
    m_TheArray(Index) = TextLine
    
End If


End Sub

Public Sub AddItem(TextLine As String)

If mb_ListIsEmpty = False Then
    ReDim Preserve m_TheArray(LBound(m_TheArray) To UBound(m_TheArray) + 1)
End If

m_TheArray(UBound(m_TheArray)) = TextLine

mb_ListIsEmpty = False

End Sub


Private Sub ShellSortAsc(SortArray() As String, AllLowerCase As Boolean)
'The fastets sort algorithm!
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
            If AllLowerCase Then
                sVal1 = LCase(SortArray(Row))
                sVal2 = LCase(SortArray(Row + Offset))
            Else
                sVal1 = SortArray(Row)
                sVal2 = SortArray(Row + Offset)
            End If
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


Private Sub ShellSortDesc(SortArray() As String, AllLowerCase As Boolean)
'The fastets sort algorithm!
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
            If AllLowerCase Then
                sVal1 = LCase(SortArray(Row))
                sVal2 = LCase(SortArray(Row + Offset))
            Else
                sVal1 = SortArray(Row)
                sVal2 = SortArray(Row + Offset)
            End If
            If sVal1 < sVal2 Then
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





Private Sub Swap(ByRef var1 As String, ByRef var2 As String)
    Dim x As String
    x = var1
    var1 = var2
    var2 = x
End Sub

Public Sub Sort(Optional Ascending As Boolean = True, Optional AllLowerCase As Boolean = True)

If Ascending Then
    ShellSortAsc m_TheArray, AllLowerCase
Else
    ShellSortDesc m_TheArray, AllLowerCase
End If

End Sub

Public Function Find(StrToFind As String, Optional Compare As VbCompareMethod) As Long

Dim idx As Long
Dim bFound As Boolean

bFound = False
For idx = 0 To ItemCount - 1
    If InStr(1, Item(idx), StrToFind, Compare) <> 0 Then
        bFound = True
        Exit For
    End If
Next idx

If bFound Then
    Find = idx
Else
    Find = -1
End If

End Function

Public Property Get Text() As String

       Text = Array2Text(m_TheArray)

End Property

Public Property Let Text(ByVal sNewValue As String)
    
    Text2Array sNewValue, m_TheArray
    mb_ListIsEmpty = False

End Property

Public Function IndexOf(StrToFind As String) As Long
Dim idx As Long
Dim bFound As Boolean

bFound = False
For idx = 0 To ItemCount - 1
    If Item(idx) = StrToFind Then
        bFound = True
        Exit For
    End If
Next idx

If bFound Then
    IndexOf = idx
Else
    IndexOf = -1
End If

End Function

Public Property Get ItemCount() As Long

If mb_ListIsEmpty Then
    ItemCount = 0
Else
    ItemCount = UBound(m_TheArray) + 1
End If

End Property
Private Property Let ItemCount(ByVal lNewValue As Long)
' this method is PRIVATE = READ-ONLY outside class
End Property

Public Property Get Item(ByVal idx As Long) As String
Attribute Item.VB_UserMemId = 0
    
    Item = m_TheArray(idx)
    
End Property

Public Property Let Item(ByVal idx As Long, ByVal sNewValue As String)
    
    m_TheArray(idx) = sNewValue
    
End Property


Private Sub Class_Initialize()
   
   ReDim m_TheArray(0 To 0)
   mb_ListIsEmpty = True

End Sub

