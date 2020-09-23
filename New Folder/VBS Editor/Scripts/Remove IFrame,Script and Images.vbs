
Public Function Main ( ByVal Text )

Text = RemoveTagAndContents(Text , "iframe" ,true )
Text = RemoveTagAndContents(Text , "script" ,false )
'Text = RemoveTagAndContents(Text , "marquee" ,false )

result = MsgBox("Remove <IMG> ?",vbYesNo)
If result = vbYes Then
        Text = RemoveTagAndContents(Text , "img" , true )
End If

Main = Text

End Function

