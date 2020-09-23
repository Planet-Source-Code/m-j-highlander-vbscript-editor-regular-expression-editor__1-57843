
Public Function InsertIntoBody ( ByVal Text )

Dim V , sUp , sBottom

V = InputForm("Insert into <body> tag at top:","Insert into <body> tag at bottom:")

If IsArray(V) Then
        sUp = V(lbound(V))
        sBottom = V(1)
        Text = RegExpReplace ( Text , "<body([^\v]*?)>([^\v]*?)</body>" ,  "<body$1>" & sUp & vbCrLf & "$2" & vbCrLF & sBottom & "</body>")
        InsertIntoBody = Text
End If


End Function
