
Public Function RX_RemoveTagCRLF(Text)

        Set C = New RegExp
        
        C.Pattern ="<[^\v]*?>"
        C.Global = True
        C.IgnoreCase = True
        
        RX_RemoveTagCRLF = C.Replace(Text ,GetRef("R"))
        
End Function


Private Function R(Match , Index , FullText)

        Dim sTemp
        
        'first remove CRLFs
        sTemp = Replace (Match , vbCrlf , " " )
        
        'now compact spaces
        R = RegExpReplace ( sTemp , " +" , " ")
        
End Function
