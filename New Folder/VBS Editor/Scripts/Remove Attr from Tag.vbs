
Public Function Main(Text)
Dim sTag , sAttr , V

V = GUI.InputForm("Tag","Attribute To Remove")

If IsArray(v) Then
        sTag = CStr(V(0))
        sAttr = CStr(V(1))

        Pattern ="<" & sTag & "[^\v>]*?" & "(" & sAttr & "=[^\v]+?(\s|""))[^\v]*?>"
        ' "<Tag" followed by any char except ">" 
        
        Main = RegExpReplaceFunc(Text ,Pattern,"ReplacerFunc")
        
End If


End Function



'------------------- Helper Functions --------------------------

Function ReplacerFunc(Match,SubMatch1 , SubMatch2 ,Index , FullText)
        ReplacerFunc = Replace(Match,SubMatch1, "" )
End Function

