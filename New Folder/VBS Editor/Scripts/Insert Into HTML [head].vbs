
Public Function InsertIntoHead ( ByVal Text )
Dim sHead
sHead = gui.InputMultiLine("Enter Text to insert into <head>...</head> tag:","<base target=""_blank"">")

Text = RegExpReplace ( Text , "(<head[^\v]*?)</head>" , "$1" & vbcrlf & sHead & vbcrlf & "</head>" )

InsertIntoHead = Text

End Function
