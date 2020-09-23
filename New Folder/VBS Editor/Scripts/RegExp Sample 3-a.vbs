
Function Main(Text)

Pattern ="<BODY[^\v]*?bgcolor=([^\v]+?(\s|""))[^\v]*?>"

Main = RegExpReplaceFunc(Text ,Pattern,"RR")

End Function

Function RR(Match,SubMatch1 , SubMatch2 ,Index , FullText)
        RR = Replace(Match,SubMatch1,"""WHITE""")
End Function

