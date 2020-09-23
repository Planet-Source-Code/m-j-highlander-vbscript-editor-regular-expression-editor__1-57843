Attribute VB_Name = "Sort_String"
Option Explicit

' By: Brian Cidern
' http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.24342/lngWId.1/qx/vb/scripts/ShowCode.htm

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                               lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
Public Sub QuickSort(varArray() As String, Optional l_First As Long = -1, Optional l_Last As Long = -1)
'This is faster than ShellSort, but is a Recursive function
Dim l_Low As Long, l_Middle As Long, l_High As Long
Dim v_Test As Variant

If l_First = -1 Then
    l_First = LBound(varArray)
End If

If l_Last = -1 Then
    l_Last = UBound(varArray)
End If

If l_First < l_Last Then
    
    l_Middle = (l_First + l_Last) / 2
    v_Test = varArray(l_Middle)
    l_Low = l_First
    l_High = l_Last
    Do
        While varArray(l_Low) < v_Test
            l_Low = l_Low + 1
        Wend


        While varArray(l_High) > v_Test
            l_High = l_High - 1
        Wend
        If (l_Low <= l_High) Then
            SwapStrings varArray(l_Low), varArray(l_High)
            l_Low = l_Low + 1
            l_High = l_High - 1
        End If
    Loop While (l_Low <= l_High)
    
    If l_First < l_High Then
        QuickSort varArray, l_First, l_High
    End If

    If l_Low < l_Last Then
        QuickSort varArray, l_Low, l_Last
    End If

End If

End Sub
Private Sub SwapStrings(pbString1 As String, pbString2 As String)
    Dim l_Hold As Long
    CopyMemory l_Hold, ByVal VarPtr(pbString1), 4
    CopyMemory ByVal VarPtr(pbString1), ByVal VarPtr(pbString2), 4
    CopyMemory ByVal VarPtr(pbString2), l_Hold, 4
End Sub
Public Sub ShellSortAsc(SortArray() As String, ByVal AllLowerCase As Boolean)
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


Public Sub ShellSortDesc(SortArray() As String, ByVal AllLowerCase As Boolean)

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




Public Sub StrSort(Lines() As String, ByVal Ascending As Boolean, ByVal AllLowerCase As Boolean)

If Ascending Then
    ShellSortAsc Lines(), AllLowerCase
Else
    ShellSortDesc Lines(), AllLowerCase
End If


End Sub
Private Sub Swap(ByRef var1 As String, ByRef var2 As String)
    Dim X As String
    X = var1
    var1 = var2
    var2 = X
End Sub
