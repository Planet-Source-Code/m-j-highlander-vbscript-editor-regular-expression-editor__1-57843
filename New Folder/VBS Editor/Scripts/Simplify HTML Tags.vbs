
'Doesn't work well

Public Function Main ( ByVal Text )

Text = RX_RemoveTagCRLF(Text)

Text = RemoveTagAndContents ( Text , "script",false)

Text = RegExpReplace ( Text , "<(\w+)\s*?.*?>" , "<$1>")

Main = Text

End Function

#INCLUDE Remove CRLFs from HTML Tags.vbs
