

Function RX_ExtractJpegLinks(Html)

Tag = "A"
sOpenTag = "<" & Tag
sCloseTag = "</" & Tag & ">"

Set objRegExp = New RegExp
Set objRegExp2 = New RegExp

objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp2.IgnoreCase = True
objRegExp2.Global = True

objRegExp.Pattern = sOpenTag & "[^\v]*?" & sCloseTag
objRegExp2.Pattern = "href[^\v]*?jpg"".?>"

ReDim V(1000)
idx=0

For Each m In objRegExp.Execute(Html)

'        If Instr(lcase(m.Value),".jpg")>0 then
        If objRegExp2.Test(m.Value) Then
                V(idx)= Replace(m.Value,vbCrLf,"")
                idx = idx + 1
        End If

Next

ReDim Preserve V(idx)

RX_ExtractJpegLinks = Join(V,vbCrLf)

End Function
