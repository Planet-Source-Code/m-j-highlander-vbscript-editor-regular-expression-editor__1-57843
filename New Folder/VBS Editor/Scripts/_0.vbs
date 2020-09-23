
Function Main(Text)

Set C = New RegExp

C.Pattern ="bgColor\=(.+?)(>|\s)"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("Replacer_Function"))

End Function



Function Replacer_Function(Match,SubMatch1 , SubMatch2 ,Index , FullText)

        Replacer_Function = Replace(Match,SubMatch1,"""white""")

End Function
