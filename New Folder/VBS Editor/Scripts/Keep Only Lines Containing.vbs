Function Main(Text)
dim v,word

word =InputBox("Enter String","")

If Word="" or Text="" Then Exit Function

v=Filter(Split(Text,vbCrLf),Word,True,vbTextCompare)

'for idx=lbound(v) to ubound(v)
'    v(idx)=v(idx)
'Next

Main=Join(v,vbCrLf)

End Function

