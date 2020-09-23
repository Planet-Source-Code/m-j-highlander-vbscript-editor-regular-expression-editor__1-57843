
Public Function Main ( ByVal Text )

' Listing 2. Code to Replace Substrings That Match Regular Expressions

Set regexp = CreateObject("VBScript.RegExp")
regexp.Global = True
regexp.IgnoreCase = True
buf = "If you comply with regular expression, go at www.regexp.com or www.re.com."
regexp.Pattern = "www.\w+\.\w+"

Set matches = regexp.Execute(buf)
For Each m In matches
' BEGIN CALLOUT A
	temp = Left(buf, m.FirstIndex-1)
' END CALLOUT A
	buf = temp & Replace(buf, m.value, _
		"http://" & m.Value, m.FirstIndex, 1)
	msgbox buf
Next
MsgBox buf


Main = ""

End Function

