
Function Main(Text)


Set C = New RegExp

C.Pattern ="<meta[^\v]*?>"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("RR"))

End Function



Private Function RR( Match , Index , FullText )

        If InStr(LCase(Match),"http-equiv")>0 Then
                RR = Match
        Else
                RR = ""
        End If

End Function

