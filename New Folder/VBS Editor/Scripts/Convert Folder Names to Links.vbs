Option UseEscapes

Function Main(Text)
Dim v,f,idx

v=Split(Text,vbCrLf)

For idx=lbound(v) to ubound(v)
        v(idx) = "<A href=\q" & v(idx) & "\\" & v(idx) & ".htm" & "\q>" & v(idx) & "</a><BR>"
Next

Main = Join(v,vbcrlf)

End Function
