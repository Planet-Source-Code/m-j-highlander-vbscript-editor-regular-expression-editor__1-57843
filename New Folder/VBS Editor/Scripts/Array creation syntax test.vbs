Option UseEscapes

Public Function Main ( ByVal Text )


mar=["my","new","place, well","not","new"]

'Equivalent to:
'mar=Array("my","new","place, well","not","new")


Main = join(mar,"*")

End Function

