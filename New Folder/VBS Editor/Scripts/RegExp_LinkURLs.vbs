
Function Main (Text)

Main = LinkURLs(Text)

End Function


Function LinkURLs(byVal strIn)
	dim re, sOut

	set re = New RegExp
	re.global = true
	re.ignorecase = true
	re.multiline = true

	' pattern that parses and finds any Internet URLs:
                re.pattern = "((mailto\:|(news|(ht|f)tp(s?))\://){1}\S+)"

	' replace method of RegExp object uses a remembered
	' pattern denoted as $1 to link the found URL(s)
	sOut = re.replace( strIn, "<A HREF=""$1"" TARGET=""_new"">$1</A>")

	set re = Nothing
	LinkURLs = sOut

End Function
