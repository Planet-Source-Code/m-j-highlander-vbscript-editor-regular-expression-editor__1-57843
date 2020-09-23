
Public Function Main ( ByVal Text )

s=split(text,vbcrlf)

for idx in s
        s(idx)="MOVE " & s(idx) & ".zip " & s(idx)
next

Main = join(s,vbcrlf)

End Function

