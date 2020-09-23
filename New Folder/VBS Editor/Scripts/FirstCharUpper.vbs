' This function changes the first letter to Upper-Case
' and the rest to Lower-Case.
' Axiom's built-in function converts only first letter

Function FirstCharUpper (Text)

sArray=Split(Text,vbcrlf)

For idx = LBound(sArray) To UBound(sArray)

    If sArray(idx) <> "" Then
        sArray(idx) = UCase(left(sArray(idx), 1)) & LCase(Right(sArray(idx),Len(sArray(idx))-1))
    End If
Next

FirstCharUpper = Join(sArray, vbCrLf)

End Function
