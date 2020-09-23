
Public Function Main ( ByVal Text )

p="(\w+)(\s*=\s*)(\d+)"
i=0

Main = RegExpExtractSubMatch(Text,p,i)

End Function
