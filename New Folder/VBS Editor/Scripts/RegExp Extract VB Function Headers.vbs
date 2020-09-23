Function Main(Text)


Set C = New RegExp

C.Pattern = "^[ \t]*(?:public|private)*[ \t]*function(.*?)\r\n"
C.Global = True
C.IgnoreCase = True
C.MultiLine=True

ReDim V(2000)
idx=0

For Each m In C.Execute(Text)
                V(idx)= m.submatches(0)
                If Instr(V(idx),"(") Then V(idx)=Left(V(idx), Instr(V(idx),"(") -1)
                idx = idx + 1
Next

ReDim Preserve V(idx)

Main = Join(V,vbcrlf)


End Function

