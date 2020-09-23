
Public Function Main(Text)
Dim C
Set C = New RegExp

'\w Matches any word character. Equivalent to [A-Za-z0-9_]
'C.Pattern ="%(\w{2})"

C.Pattern ="%([0-9ABCDEF]{2})"   'no need for error handling.
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("R"))

End Function



Function R(Match,SubMatch1 , Index , FullText)
'        On Error Resume Next  ' in case of % not followed by hex value.
        R=Chr("&h" & SubMatch1)
'        If Err then R=Match
End Function

