'String Functions
'====================
Public Function SplitAt(ByVal Text , ByVal LookFor , ByRef LeftSplit , ByRef RightSplit)
        Dim pos
        pos = InStr(Text, LookFor)
        If pos <> 0 Then
                LeftSplit = Left(Text, pos - 1)
                RightSplit = Right(Text, Len(Text) - pos - Len(LookFor) + 1)
        Else
                LeftSplit = ""
                RightSplit = ""
        End If
        
        SplitAt = ""  'VOID
End Function
'----------------------------------------------------------------------------------------
Public Function DelStr(ByVal Text , ByVal Start , ByVal Length )
        'If "Start" is larger than the length of "Text", no characters are deleted.
        'If "Length" specifies more characters than remaining starting at "Start",
        'DelStr removes the rest of the string.
        
        Dim sLeft , sRight , sTemp , lLenRight
        
        If (Start > Len(Text) Or Start = 0 Or Length = 0) Then
                sTemp = Text
        Else
                sLeft = Left(Text, Start - 1)
                lLenRight = Len(Text) - Start - Length + 1
                If lLenRight > 0 Then
                        sRight = Right(Text, lLenRight)
                Else
                        sRight = ""
                End If
                sTemp = sLeft & sRight
        End If
        
        DelStr = sTemp

End Function
'----------------------------------------------------------------------------------------
Public Function DelLeft(ByVal Text, ByVal Count )

        ' Deletes "Count" chars from the left of "Text".
        ' If Count <=0 delete nothing
        ' If Count >= length of "Text" delte all
        
        If (Count >= Len(Text)) Then  ' delete all
                        DelLeft = ""
                        
        ElseIf Count <= 0 Then        'delete nothing
                        DelLeft = Text
        
        Else
                        DelLeft = Right(Text, Len(Text) - Count)
        End If

End Function
'-----------------------------------------------------------
Public Function DelLeftTo(ByVal Text , ByVal SubStr , ByVal Inclusive , ByVal MatchCase)

        Dim lPos
        
        If CBool(MatchCase) = False Then
                lPos = InStr(1, Text, SubStr, vbTextCompare)   'dont match case
        Else
                lPos = InStr(1, Text, SubStr, vbBinaryCompare) 'case sensitive
        End If
        
        If (lPos = 0) Then      ' search string not found, return whole string
                    DelLeftTo = Text
        Else
                    If CBool(Inclusive) = True Then
                                DelLeftTo = Right(Text, Len(Text) - lPos - Len(SubStr) + 1)
                    Else
                                DelLeftTo = Right(Text, Len(Text) - lPos + 1)
                    End If
        End If
        
End Function
'-----------------------------------------------------------

Public Function DelRightTo(ByVal Text , ByVal SubStr , ByVal Inclusive , ByVal MatchCase )

        Dim lPos
        
        If CBool(MatchCase) = False Then
                lPos = InStrRev(Text, SubStr, -1, vbTextCompare)   'dont match case
        Else
                lPos = InStrRev(Text, SubStr, -1, vbBinaryCompare) 'case sensitive
        End If
        
        If (lPos = 0) Then
                    DelRightTo = Text
        Else
                    If CBool(Inclusive) = True Then
                                DelRightTo = Left(Text, lPos - 1)
                    Else
                                DelRightTo = Left(Text, lPos + Len(SubStr) - 1)
                    End If
        End If
        
End Function

'-----------------------------------------------------------
Public Function DelRight(ByVal Text , ByVal Count )

        ' Deletes "Count" chars from the right of "Text".
        ' If Count <=0 delete nothing
        ' If Count >= length of "Text" delte all
        
        If (Count >= Len(Text)) Then  ' delete all
                        DelRight = ""
                        
        ElseIf Count <= 0 Then        'delete nothing
                        DelRight = Text
        
        Else
                        DelRight = Left(Text, Len(Text) - Count)
        End If

End Function

'-----------------------------------------------------------
