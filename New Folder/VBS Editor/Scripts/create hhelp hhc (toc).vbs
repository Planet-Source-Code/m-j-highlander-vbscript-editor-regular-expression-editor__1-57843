Option UseEscapes

Public Function Main ( ByVal Text )

V = Split(Text,vbCrLf)
S = "\t\t<LI> <OBJECT type=\qtext/sitemap\q>\n\t\t\t<param name=\qName\q value=\q%s\q>\n\t\t\t<param name=\qLocal\q value=\q%s\q></OBJECT>\n"

for idx=0 to ubound(V)
        vv=split(V(idx),"*")
        V(idx) = stringf(s,vv(1),vv(0))
next

Main = Join(V,vbcrlf)

End Function

