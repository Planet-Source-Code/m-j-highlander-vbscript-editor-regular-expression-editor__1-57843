Option UseEscapes

Public Function Main ( ByVal Text )

Text = RemoveLineBreaks (Text)
Text = CompactSpaces (Text)
Text = Replace(Text , ">" ,">\n")
Text = LinesTrim(Text , True , True ,True )

Main = Text

End Function

