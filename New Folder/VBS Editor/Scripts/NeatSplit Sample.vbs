
Public Function Main ( ByVal Text )
Dim V

V = NeatSplit (Text , vbCrLf)

for idx=0 to ubound(v)
        msgbox v(idx)
next

Main = Join(V,vbCrLf)

End Function
