
Public Function Main ( ByVal Text )

Text = RemoveLineBreaks (Text)

Text = Replace ( Text , ">" , ">" & vbCrLf )

Main = Text

End Function

