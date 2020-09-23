
Function Main(Text)

Set C = New RegExp

C.Pattern ="(<[^>]*>)"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("R"))

End Function


Function R(Match ,SubMatch1, Index , FullText)

R = lcase(Match)

End Function
