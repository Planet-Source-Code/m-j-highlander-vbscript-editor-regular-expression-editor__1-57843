Option UseEscapes

Public Function Main(Text)
Dim v,f,idx

v=Split(Text,vbCrLf)

For idx=0 to ubound(v)
        s = LoadFile ( v(idx) )
        s = replace(RegExpExtractSubMatch ( s , "<title>([^\\x00]*?)</title>"),vbcrlf,"")
        if s = "" then s = ExtractFileName(v(idx))
        v(idx) = "<a href=\q" & Replace(v(idx),"\\","/") & "\q>" & s & "</a><br>"
Next


Main =  Join(v,vbcrlf)
Main = "<font face=\qVerdana\q size=\q2\q>\n" & Main & "\n</font>"
Main = "<html>\n<head>\n<title></title>\n</head>\n<body>\n" & Main  & "\n</body>\n</html>"


End Function

#INCLUDE FileName Functions.inc

