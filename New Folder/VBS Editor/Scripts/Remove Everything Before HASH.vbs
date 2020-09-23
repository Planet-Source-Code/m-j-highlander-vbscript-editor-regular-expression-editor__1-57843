
Public Function Main ( ByVal Text )


Main = RegExpReplace(Text, "href\="".*?#" , "href=""#")

End Function

