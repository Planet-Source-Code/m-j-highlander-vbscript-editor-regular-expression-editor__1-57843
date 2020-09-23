
Function Main(Text)

Set C = New RegExp

C.Pattern ="http://www.online-literature.com/dumas/man_in_the_iron_mask/(\d\d)/"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("R"))

End Function


Function R(Match ,SubMatch1, Index , FullText)

R = "files/Chapter " & SubMatch1 & ".htm"


End Function
