Option UseEscapes

Public Function Main ( ByVal Text )
Dim img,idx

for idx=310 to 340
        img = Replace("http://www.pwgalleries.com/*/image/_02.jpg","*",idx)
        img = "<img src=\q" & img & "\q>"
        img = "<A href=\q" & Replace("http://www.pwgalleries.com/*","*",idx) & "\q>" & img & "</a>"
        WriteLn img
Next


Main = OutStr

End Function

