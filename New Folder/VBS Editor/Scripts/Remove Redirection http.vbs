
Public Function Main ( ByVal Text )

Main = RegExpReplace(Text , "http\://.+?http" , "http")

End Function
