' Description:        Generate WinXP Manifest File


Public Function Main ( ByVal Text )

S = LoadFile ( AxiomPath & "Scripts\manifest_template.txt" )

If S<>"" Then

        V = GUI.InputForm ( "Company Name|CompanyName","Product Name|ProductName","App Name|AppName","Description|Description")
        If Not IsArray ( V) Then Exit Function
        
        S = Replace ( S , "%CompanyName%" , V(0) )
        S = Replace ( S , "%ProductName%" , V(1) )
        S = Replace ( S , "%AppName%" , V(2) )
        S = Replace ( S , "%Description%" , V(3) )

        Main = S

End If

End Function
