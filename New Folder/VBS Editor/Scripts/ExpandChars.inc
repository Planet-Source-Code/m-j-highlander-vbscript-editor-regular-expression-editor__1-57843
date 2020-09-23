
Public Function ExpandChars(Text)
' Replaces \nnn with the equiv char
' the number is in DECIMAL, leftmost zeroes are ignored
' a single backslash "\" followed by non-numerical chars is omitted
' a double backslash "\\" is converted to single
' if number is larger than 255 the modulous of 256 will be used.
' SAMPLE CALL:
ExpandChars("no \34is\034 good\013\10cool")

Dim sTemp

sTemp = Replace(Text, "\\", "%DOUBLE_SLASHES%")
Do
        sTemp = ExpandFirstChar(sTemp)
        If InStr(sTemp, "\") = 0 Then Exit Do
Loop

ExpandChars = Replace(sTemp, "%DOUBLE_SLASHES%", "\")

End Function

Private Function ExpandFirstChar(Text)
' Used by ExpandChars()

Dim backslash_pos
Dim idx
Dim sTemp
Dim ch
Dim sReplaceWith
Dim vTemp

sTemp = ""
backslash_pos = InStr(Text, "\")
idx = backslash_pos + 1
Do While idx <= Len(Text)
        ch = Mid(Text, idx, 1)
        If IsNumeric(ch) Then
                sTemp = sTemp & ch
                idx = idx + 1
         Else
                Exit Do
        End If
Loop

If sTemp <> "" Then
        vTemp = sTemp
        vTemp = vTemp Mod 256     ' keep in range 0..255
        sReplaceWith = Chr(vTemp)
Else
        sReplaceWith = ""
End If
ExpandFirstChar = Replace(Text, "\" & sTemp, sReplaceWith,1 , 1) 'replace only once

End Function

