
' public vars:
Dim sLookFor

Function Main(Text)
Dim C

sLookFor = InputBox("Remove <A>...</A> Tags if HREF contains:")
If sLookFor = "" Then Exit Function

Set C = New RegExp
C.Global = True
C.IgnoreCase = True

C.Pattern ="<a[^\v]*?href=""([^\v]*?)""[^\v]*?</A>"

Main = C.Replace(Text ,GetRef("R"))

End Function

'*********************************
Private Function R(Match ,SubMatch, Index , FullText)

        If Instr(LCase(SubMatch),LCase(sLookFor))>0 then
                R = ""
        Else
                R = Match
        End If

End Function
