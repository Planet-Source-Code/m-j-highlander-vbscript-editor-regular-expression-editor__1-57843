Function Main(Text)

Text=Replace(Text,",","</TD><TD>")
Text=Replace(Text,"\n","</TD></TR>\n<TR><TD>")

Text="<TABLE border=\q1\q>\n<TR><TD>" & Text

Text=Left(Text,len(Text)-8) & "\n</TABLE>"


Main=Text

End Function
