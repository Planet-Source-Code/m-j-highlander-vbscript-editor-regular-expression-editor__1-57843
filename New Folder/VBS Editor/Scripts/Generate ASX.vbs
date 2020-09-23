Option UseEscapes

Public Function Main ( ByVal Text )

a = ["together_3.mpg", "4th.mpg", "amanda.mpg", "asst.mpg", "bums.mpg", "cherry.mpg", "mira_1.mpg", "mira_2.mpg", "together_1.mpg", "together_2.mpg"]

for idx=lbound(a) to ubound(a)
        a(idx)=stringf("<Entry><ref href=\q%s\q /></Entry>",a(idx))
next

Main = StrJoin("<ASX version=\q3\q>\n" , join(a,vbCrLf) , "\n</ASX>\n", vbcrlf)

End Function

