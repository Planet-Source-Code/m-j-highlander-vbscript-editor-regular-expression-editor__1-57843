
Function Main(Text)


Set C = New RegExp

C.Pattern = ".*?a\r\n"
C.Global = True
C.IgnoreCase = True


ReDim V(2000)
idx=0

For Each m In C.Execute(Text)
                V(idx)= m.Value
                idx = idx + 1
Next

ReDim Preserve V(idx)

Main = Join(V,"")


End Function

