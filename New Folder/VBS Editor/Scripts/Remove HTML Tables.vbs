
Public Function Main ( ByVal Text )

Text = ReplaceTagKeepContent(Text , "table" , "<P>" , "</P>")
Text = ReplaceTagKeepContent(Text , "tbody" , "" , "")
Text = ReplaceTagKeepContent(Text , "tr" , "" , "<BR>")
Text = ReplaceTagKeepContent(Text , "td" , "&nbsp;" , "&nbsp;")

If MsgBox("Remove <DIV> ?",vbYesNo)=vbYes Then
        Text = ReplaceTagKeepContent(Text , "div" , "<P>" , "</P>")
End If

Main = Text

End Function
