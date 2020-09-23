
Function Main(Text)
Dim C

Set C = New RegExp

C.Pattern ="<A[^\v]*?>\s*</A>"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,"")


End Function
