Function Main(Text)

Text=HtmlToText(Text)
Text=TrimSpaces(Text,1 ,1  ,1 )
Text=CompactBlankLines(Text)

Main=Text

End Function
