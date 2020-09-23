Public Function EncDec (Secret)
Dim L,X,PassWord,iChar,Result

' Calling the Function on a string Encrypts it
' Calling the Function on an encrypted  string Decrypts it

    PassWord =InputBox("Enter Password ")

    L = Len(PassWord)
    if L=0 Then Exit Function

    For X = 1 To Len(Secret)
        iChar = Asc(Mid(PassWord, (X Mod L) - L * ((X Mod L) = 0), 1))
        Result = Result & Chr(Asc(Mid(Secret, X, 1)) Xor iChar)
    Next

EncDec = Result

End Function 
