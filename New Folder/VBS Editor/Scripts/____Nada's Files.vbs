
Function Main(Text)

Set C = New RegExp

C.Pattern ="<\!\-\- TOP AD CODE BEGIN \-\->[^\v]+?<\!\-\- END FEATURED CODE \-\->"
C.Global = True
C.IgnoreCase = True

Text = C.Replace(Text ,"")

C.Pattern ="<\!\-\- AMAZON \-\->[^\v]+?</BODY>"

Main = C.Replace(Text ,"</body>")


End Function


Function R(Match ,SubMatch1, Index , FullText)

R = "files/Chapter " & SubMatch1 & ".htm"


End Function
