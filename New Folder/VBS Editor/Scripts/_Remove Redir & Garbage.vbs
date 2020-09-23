
Public Function Main ( ByVal Text )

Text = RegExpReplace(Text ,"http\://.+?http" , "http")

Text = RemoveTrailingGarbage (Text)

Main = Text

End Function



'=================================================


Private Function RemoveTrailingGarbage ( ByVal Text )

        'sVar =InputBox("Garbage starts with... (Cancel for no garbage)", "Input Box" , "&amp;")
        sVar = GUI.OptionForm ("Garbage starts with... (Cancel for no garbage)","&amp;","%3F","Other")
        
        If sVar ="" Then
                RemoveTrailingGarbage = Text
        Else
                sVar = RegExpEscape(sVar)
                RemoveTrailingGarbage = RegExpReplace (Text , sVar & ".*?""" , vbQuote)
        End If
        
End Function
