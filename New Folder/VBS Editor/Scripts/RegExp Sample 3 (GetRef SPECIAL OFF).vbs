
Function Main(Text)


Set C = New RegExp

C.Pattern ="<BODY[^\v]*?bgcolor=(.+?(\s|"")).*?>"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("RR"))


'Main = C.Replace(Text ,GetRef("R"))
'same result ... less bells and whistles !
'Main = C.Replace(Text ,"<BODY BGCOLOR=$1>")

End Function


Function R(Match,SubMatch1 , Index , FullText)
        R="<BODY BGCOLOR=" & SubMatch1 & ">"
End Function

Function RR(Match,SubMatch1 , SubMatch2 ,Index , FullText)
        'msgBox SubMatch1
        RR = Replace(Match,SubMatch1,"""WHITE""")
End Function

