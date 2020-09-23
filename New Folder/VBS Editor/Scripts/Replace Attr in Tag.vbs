
' Public Vars
Dim gsReplaceWith


Function Main(Text)
Dim V , sTag , sAttr , sPattern

V = GUI.InputForm("Find Tag|IMG","Find Attribute|ALT","Replace Atrribute With, Leave empty to Remove")

If IsArray(V) Then
        sTag = V(0)
        sAttr = V(1)
        gsReplaceWith = V(2)
        sPattern ="<" & sTag & "\s+[^\v]*?(" & sAttr & "\s*\=\s*""*?[^\v]+?"")\s*[^\v]*?>"
        Main = RegExpReplaceFunc ( Text , sPattern , "RepFunc")
End If

End Function


Function RepFunc(Match,SubMatch1 , Index , FullText)

        RepFunc = Replace ( Match , SubMatch1 , gsReplaceWith)

End Function
