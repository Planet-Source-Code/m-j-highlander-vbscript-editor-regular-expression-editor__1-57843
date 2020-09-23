
Public Function RemoveHardLineBreaks ( ByVal Text )

Text = Replace ( Text , vbcrlf & vbcrlf , chr(7) )
Text = Replace ( Text , vbcrlf , "" )
Text = Replace ( Text , chr(7) , vbcrlf & vbcrlf )

RemoveHardLineBreaks = Text

End Function

