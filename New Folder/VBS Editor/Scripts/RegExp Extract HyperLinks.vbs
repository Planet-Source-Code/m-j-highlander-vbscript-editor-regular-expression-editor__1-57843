Function RX_ExtractHyperLinks (Text)
'recommended when coding RegExp... to avoid escaping the regexp escapes !


Set objRegExp = New RegExp
objRegExp.IgnoreCase = True
objRegExp.Global = True

objRegExp.Pattern = "(\w+):\/\/([^/:]+)(:\d*)?([^# ""]*)"


ReDim V(2000)
idx=0

For Each m In objRegExp.Execute(Text)
                V(idx)= m.Value
                idx = idx + 1
Next

ReDim Preserve V(idx)

RX_ExtractHyperLinks = Join(V,vbCrLf)

End Function
