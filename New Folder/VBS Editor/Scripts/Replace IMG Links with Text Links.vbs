
Public Function Main ( ByVal Text )

Text = ReplaceTagAndContents(Text , "img" , true , "Download File...")
Text = ReplaceAll(Text , "</a>" , "</a><br>" , False)

Main = Text

End Function

