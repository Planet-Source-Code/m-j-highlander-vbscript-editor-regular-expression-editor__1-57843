
Public Function RemoveInTagScript ( Byval Text)

'RegExpReplace(""javascript[^\v]*?"" ,""#"" )

        Text = RegExpReplace(Text , "OnMouseOver=""[^\v]*?""" ,"" )
        Text = RegExpReplace(Text , "OnMouseUp=""[^\v]*?""" , "")
        Text = RegExpReplace(Text , "OnMouseOut=""[^\v]*?""" , "")
        Text = RegExpReplace(Text , "OnClick=""[^\v]*?""" , "")
        Text = RegExpReplace(Text , "OnLoad=""[^\v]*?""" , "")
        Text = RegExpReplace(Text , "OnExit=""[^\v]*?""" , "")
        Text = RegExpReplace(Text , "OnFocus=""[^\v]*?""" , "")
        
        RemoveInTagScript = Text
        
End Function
