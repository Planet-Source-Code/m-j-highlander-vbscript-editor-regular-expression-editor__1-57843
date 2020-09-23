
Public Function RemoveTrailingGarbage ( ByVal Text )

        sVar =InputBox("Garbage starts with... (Cancel for no garbage)", "Input Box" , "&amp;")
        
        If sVar ="" Then
                RemoveTrailingGarbage = Text
        Else
                RemoveTrailingGarbage = RegExpReplace (Text , sVar & ".*?""" , vbQuote)
        End If
        
End Function


