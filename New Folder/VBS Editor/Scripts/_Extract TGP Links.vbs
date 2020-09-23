' Extract Thumbnail Links from TGP's       

Public Function Main ( ByVal Text )

        Dim V,idx

        'sUrl = RegExpExtractSubMatch ( Text , "<\!\-\- saved from url\=\(.*?\)(.+?) \-\->")
        ' extracts without "http://www." :
        sUrl = RegExpExtractSubMatch ( Text , "<\!\-\- saved from url\=\(.*?\)http\://www\.(.+?) \-\->")

        'Remove <Script> tags since they could confuse:
        Text = RemoveTagAndContents (Text , "SCRIPT",false)
        
        'Extract links only:
        Text = ExtractTagWithContents (Text, "A")
        
        'Remove Redirection from HREF's:
        Text = RegExpReplace(Text, "http\://.+?http" , "http")
        
        'Format link-per-line:
        Text = RemoveLineBreaks (Text)
        Text = ReplaceAll(Text, "</a>" ,"</a>" & vbCrLf , False)
        Text = CompactSpaces (Text)
        
        'Sort alphabetical,ignorung case:
        Text = LinesSort(Text, true,true)
        
        'Remove text-only links (no <img> tag):
        V = Split (Text,vbCrlf)
        for idx=lbound(v) to ubound(v)
                If InStr(1,v(idx),"<img",vbTextCompare)=0 then
                        v(idx)=""
                End If
        Next
        Text = Join(V,vbCrLf)
        Text = LinesRemoveBlank (Text)
        
        'Remove <img> tag attrib's:
        'Text = RegExpReplace(Text, "height=\d+" , "")
        'Text = RegExpReplace(Text, "width=\d+" , "")

        'Set Image Border:
        Text = RegExpReplace(Text, "border=\d+" , "border=3")
        
        Text = RemoveInTagScript (Text)
        
        'Add HTML page structure
        Text = "<html><head>" &vbCrlf & "<title>" & _
        sUrl & "</title>" & vbCrlf & "</head>" & _
        vbCrLf & "<body link=#e4e4e4 vlink=red>" & vbCrLf & Text
        
        Text = Text &  vbCrLf & "</body></html>"
        
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
' no longer used!
Private Function RX_ExtractSavedFromUrl (ByVal Text)

        Set objRegExp = New RegExp
        objRegExp.IgnoreCase = True
        objRegExp.Global = True

        objRegExp.Pattern = "<\!\-\- saved from url\=\(.*?\)(.+?) \-\->"
        
        
        ReDim V(2000)
        idx=0
        
        For Each m In objRegExp.Execute(Text)
                        V(idx)= m.SubMatches(0)
                        idx = idx + 1
        Next
        
        ReDim Preserve V(idx)
        
        RX_ExtractSavedFromUrl = Join(V,"")
        
End Function

'--------------------------------------------------------------------------------

Private Function RemoveTrailingGarbage ( ByVal Text )

'        sVar = InputBox("Garbage starts with... (Cancel for no garbage)", "Input Box" , "&amp;")
        sVar = GUI.OptionForm ("Garbage starts with... (Cancel for no garbage)","&amp;","%3F","Other")
        sVar = RegExpEscape(sVar)
        
        If sVar ="" Then
                RemoveTrailingGarbage = Text
        Else
                RemoveTrailingGarbage = RegExpReplace (Text , sVar & ".*?""" , vbQuote)
        End If
        
End Function


