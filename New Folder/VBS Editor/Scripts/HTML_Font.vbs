Function HTML_Font (Text)

s=inputbox("Font Size")
n=inputbox("Font Name")

HTML_Font = "<FONT FACE=""" & n & """ SIZE=""" & s & """>" & vbcrlf & Text & vbcrlf & "</FONT>"

End Function
