
'******************** Change HTML Tags to Lower Case ******************

Function Main(Text)
Dim C

Set C = New RegExp

C.Pattern ="<[^\v]*?>"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("Replacer"))

End Function



Private Function Replacer( Match , Index , FullText )

        Replacer = LCase(Match)

End Function
