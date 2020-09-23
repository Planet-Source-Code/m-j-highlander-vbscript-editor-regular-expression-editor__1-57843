
Function Main(Text)

' the double-slash is required since Axiom-VBScript has special
' symbols (\\ \n \t \q)
'Since we're not turning special symbols off


Set C = New RegExp

C.Pattern = "<BODY[^\\v]*?bgcolor=(.*?)\\s.*?>"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text , "<BODY BGCOLOR=\q$1\q>")

End Function

