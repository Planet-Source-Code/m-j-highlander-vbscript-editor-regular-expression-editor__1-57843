Function  ASCII_To_Entity()

For i = 1 To 255
    D = i & "&nbsp;&nbsp;&nbsp; " & "&#" & i & ";" & "<BR>" & vbCrLf
    All = All & D
Next

ASCII_To_Entity = All

End Function

