
Public Function Main ( ByVal Text )

S = LoadFile ( AxiomPath & "Scripts\mplayerocx.txt" )

If S<>"" Then

        V = GUI.InputForm ("Media File Name","Width|150","Height|50")
        If Not IsArray (V) Then Exit Function
        
        S = Replace ( S , "%filename%" , V(0) )
        S = Replace ( S , "%width%" , V(1) )
        S = Replace ( S , "%height%" , V(2) )

        Main = S

End If

End Function
