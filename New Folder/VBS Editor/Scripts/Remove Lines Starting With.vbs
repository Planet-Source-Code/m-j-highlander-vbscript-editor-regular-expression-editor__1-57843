Function Main(Text)
dim v,word

word =InputBox("Enter String","")

If Word="" or Text="" Then Exit Function

v=Split(Text,vbCrLf)

for idx=lbound(v) to ubound(v)
        If lcase(left(v(idx),len(word)))=lcase(word) then
                v(idx)="~!@#"
        End If
Next

v=Filter(v,"~!@#",False,vbTextCompare)

Main=Join(v,vbCrLf)

End Function

