'-----------------------------------------------------------

Function ExtractFileExt(FileName)
Dim pos

pos = InStrRev(FileName, ".")

If pos = 0 Then
    ExtractFileExt = ""
Else
    ExtractFileExt = Right(FileName, Len(FileName) - pos)
End If

End Function

'-----------------------------------------------------------

Function ExtractFileName(FilePath)
Dim iLastSlash

    iLastSlash = InStrRev(FilePath, "\")
    
    If iLastSlash = 0 Then
            ExtractFileName = FilePath
    Else
        ExtractFileName = Right(FilePath, Len(FilePath) - iLastSlash)
    End If
End Function

'-----------------------------------------------------------

Function ChangeFileExt(FileName,NewExtension)
Dim OldExt , BaseName

NewExtension = Replace(NewExtension, ".", "")

OldExt = ExtractFileExt(FileName)

BaseName = Left(FileName, Len(FileName) - Len(OldExt))
If Right(BaseName, 1) <> "." Then BaseName = BaseName & "."

If (NewExtension = "" And Right(BaseName, 1) = ".") Then
    BaseName = Left(BaseName, Len(BaseName) - 1)
End If

ChangeFileExt = BaseName & NewExtension
End Function

'-----------------------------------------------------------

Function ExtractDirName(FileName )

    Dim tmp
    Dim pos 
    Dim PrevPos 

    pos = InStr(FileName, "\")
    If pos = 0 Then
        ExtractDirName = ""
        Exit Function
    End If
    
    Do While pos <> 0
        PrevPos = pos
        pos = InStr(pos + 1, FileName, "\")
    Loop

    tmp = Left(FileName, PrevPos)
    If Right(tmp, 1) = "\" Then tmp = Left(tmp, Len(tmp) - 1)
    ExtractDirName = tmp
    
End Function
