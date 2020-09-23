
Function ReverseLinesOrder (Text)

v=Split(Text,vbCrLf)

R = v

for idx=lbound(v) to ubound(v)
        R(idx)=V(ubound(v)-idx)
next

ReverseLinesOrder = Join(R,"\n")

End Function