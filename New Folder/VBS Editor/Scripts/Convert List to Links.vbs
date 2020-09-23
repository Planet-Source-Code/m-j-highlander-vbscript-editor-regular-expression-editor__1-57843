Option UseEscapes
#INCLUDE FileName Functions.inc

Function Main(Text)
Dim v,f,idx

v=Split(Text,vbCrLf)

For idx=lbound(v) to ubound(v)
         f = ChangeFileExt(ExtractFileName(v(idx)),"")
'        f = v(idx)
        If v(idx) <>"" Then v(idx)="<A HREF=\qFiles/" & v(idx) & "\q>" & f & "</A><BR>"
Next

Main = Join(v,vbcrlf)

End Function
