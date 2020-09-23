Option UseEscapes

Public Function Main ( ByVal Text )

v=split(text,vbcrlf)

for idx in v
        v(idx)="N(" & idx & ")=\q" & v(idx) & "\q"
next

Main = join(v,vbcrlf)

End Function

