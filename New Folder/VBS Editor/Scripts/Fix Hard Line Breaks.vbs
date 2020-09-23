Option UseEscapes


Public Function Main ( ByVal Text )

'Replace dbl-newlines with smthng quite strange!
Text = Replace(Text , "\n\n" , "!^#@%$&*")

Text = RemoveLineBreaks (Text)

'Restore
Text = Replace(Text , "!^#@%$&*","\n\n")

Main = CompactSpaces(Text)

End Function