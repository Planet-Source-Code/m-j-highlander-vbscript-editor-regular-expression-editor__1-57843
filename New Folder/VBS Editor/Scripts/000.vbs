
Public Function Main ( ByVal Text )

Text = RemoveTagandContents (Text, "script" , false)
Text = RemoveTagPath( Text , "src" , true )
Text = RemoveTagPath( Text , "href" , true )
Text = PutCSSinHTML (Text)
Text = ValidateImageTags (Text)


Main = Text

End Function

