
' Public Variables
Dim gsNewValue

Public Function Main(Text)
Dim sTag , sAttr , V

V = InputForm("Tag","Attribute Name","New Attribute Value")

If IsArray(v) Then
        sTag = CStr(V(0))
        sAttr = CStr(V(1))
        gsNewValue = CStr(V(2))


        Pattern ="<" & sTag & "[^\v>]*?" & sAttr & "=([^\v]+?(\s|""))[^\v]*?>"
        ' "<Tag" followed by any char except ">" 
        
        Main = RegExpReplaceFunc(Text ,Pattern,"ReplacerFunc")
        
End If


End Function



'------------------- Helper Functions --------------------------

Function ReplacerFunc(Match,SubMatch1 , SubMatch2 ,Index , FullText)
        ReplacerFunc = Replace(Match,SubMatch1, vbQuote & gsNewValue & vbQuote )
End Function

