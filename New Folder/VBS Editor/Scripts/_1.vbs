Public Function Main ( ByVal Text )
Dim V , idx , LinkHref() , LinkText

Text = ExtractTagWithContents (Text , "A")
Text = Replace (Text , vbCrLf , " ")
Text = Replace (Text , "</a>" , "</a>" & vbCrLf,1,-1,vbTextCompare)
Text = CompactSpaces ( Text )
'V = Split ( Text , vbCrLf)

'ReDim LinkHref(Ubound(V)) , LinkText(Ubound(V))

'For idx in V
 '       LinkHref(idx)=RegExpExtractSubMatch ( V(idx) , "href=""(.*?)""")
 '       LinkText(idx)=RegExpExtractSubMatch ( V(idx) , "<a.*?>(.*?)</a>")
'Next

V = RegExpExtract0 ( Text , "href=""(.*?)""")

v=split(v,chr(0))

Main = Join(v,vbCrLf)

End Function
