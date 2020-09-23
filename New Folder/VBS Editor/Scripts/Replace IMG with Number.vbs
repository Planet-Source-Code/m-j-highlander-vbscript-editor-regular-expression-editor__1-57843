
Dim gCntr , gString
gString = "File #"

Function Main(Text)

Set C = New RegExp

C.Pattern ="<img[^\v]*?src=[^\v]*?>"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("R"))

End Function


Function R(Match , Index , FullText)
Dim sTemp
        gCntr = gCntr +1
        R = gString & Format(gCntr,"000")
End Function

'// Helper Functions:
