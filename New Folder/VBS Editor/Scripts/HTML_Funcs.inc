Option UseEscapes
'-----------------------------------------------------------

Private Function EnTag(Text,Tag)

EnTag = "<" & Tag & ">" & Text & "</" & Tag & ">"

End Function

'-----------------------------------------------------------
Private Function Link(Text,Href)

Link = "<a href=\q" & Href & "\q>" & Text & "</a>"

End Function

'-----------------------------------------------------------
