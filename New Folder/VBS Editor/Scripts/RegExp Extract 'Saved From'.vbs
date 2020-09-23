
Function RX_ExtractSavedFromUrl (Text)

Set objRegExp = New RegExp
objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = "<\!\-\- saved from url\=\(.*?\)(.+?) \-\->"


ReDim V(2000)
idx=0

For Each m In objRegExp.Execute(Text)
                V(idx)= m.SubMatches(0)
                idx = idx + 1
Next

ReDim Preserve V(idx)

RX_ExtractSavedFromUrl = Join(V,"")

End Function
