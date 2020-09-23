
Function Main(Text)


Set C = New RegExp
C.Global =false
C.IgnoreCase = True


C.Pattern ="<[^\v]*?<BODY.*?>"
sTemp = C.Replace(Text ,"")

'intentionaly greedy:
C.Pattern ="</body[^\v]*>"
Main = C.Replace(sTemp ,"")


End Function

