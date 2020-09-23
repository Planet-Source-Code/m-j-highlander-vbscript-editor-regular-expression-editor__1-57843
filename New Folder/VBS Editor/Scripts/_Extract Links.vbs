' Extract And tidy-up Links
'(makes extensive use of Macro functions)
'**********************************


Public Function Main ( ByVal Text )

        Dim V,idx

        ' Extract "saved from" url, without the "http://www."
        sUrl = RegExpExtractSubMatch ( Text , "<\!\-\- saved from url\=\(.*?\)http\://www\.(.+?) \-\->")

        'Remove <Script> tags since they could confuse:
        Text = RemoveTagAndContents (Text , "SCRIPT",false)
        
        'Extract links only:
        Text = ExtractTagWithContents (Text, "A")
        
        'Remove Redirection from HREF's:
        Text = RegExpReplace(Text, "http\://.+?http" , "http")
        
        'Format link-per-line:
        Text = RemoveLineBreaks (Text)
        Text = ReplaceAll(Text, "</a>" ,"</a><br>" & vbCrLf , False)
        Text = CompactSpaces (Text)
        
        'Sort alphabetical,ignorung case:
        Text = LinesSort(Text, true,true)
        

        Text = RemoveInTagScript (Text)
        
        'Add HTML page structure
        Text = "<HTML><HEAD>" &vbCrlf & "<TITLE>" & vbCrLf & _
        sUrl & "</TITLE>" & vbCrlf & "<BASE TARGET=""_blank""></HEAD>" & _
        vbCrLf & "<BODY>" & vbCrLf & Text
        
        Text = Text &  vbCrLf & "</BODY></HTML>"
        
        Text = RemoveTrailingGarbage (Text)
        
        Main = Text


End Function


'================== Helper Functions ====================

Private Function RemoveInTagScript ( Byval Text)

        Text = RegExpReplace(Text , "OnMouseOver=""[^\v]*?""" ,"" )
        Text = RegExpReplace(Text , "OnMouseUp=""[^\v]*?""" , "")
        Text = RegExpReplace(Text , "OnMouseOut=""[^\v]*?""" , "")
        Text = RegExpReplace(Text , "OnClick=""[^\v]*?""" , "")
        Text = RegExpReplace(Text , "OnLoad=""[^\v]*?""" , "")
        Text = RegExpReplace(Text , "OnExit=""[^\v]*?""" , "")
        Text = RegExpReplace(Text , "OnFocus=""[^\v]*?""" , "")
        
        RemoveInTagScript = Text
        
End Function

'--------------------------------------------------------------------------------

Private Function RemoveTrailingGarbage ( ByVal Text )

        sVar =InputBox("Garbage starts with... (Cancel for no garbage)", "Input Box" , "&amp;")
        
        If sVar ="" Then
                RemoveTrailingGarbage = Text
        Else
                RemoveTrailingGarbage = RegExpReplace (Text , sVar & ".*?""" , vbQuote)
        End If
        
End Function

