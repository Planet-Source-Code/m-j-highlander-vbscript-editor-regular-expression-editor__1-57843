Option UseEscapes

Public Function Main ( ByVal Text )

'With Toolbar
'    .Buttons.Add , "new", , tbrDefault, "new"
'   .Buttons("new").ToolTipText = "New"


Buttons = ["new","open","save","saveas","help"]
ToolBar = "Toolbar1"

writeln "With " & Toolbar & "\n"

writeln "\t.Buttons.Add , \qsep_1\q, , tbrSeparator\n"

For idx In Buttons
        WriteLn stringf("\t.Buttons.Add , \q%s\q, , tbrDefault, \q%s\q",Buttons(idx),Buttons(idx))
        WriteLn stringf("\t.Buttons(\q%s\q).ToolTipText = \q%s\q\n",Buttons(idx),SentenceCase(Buttons(idx)))
Next

writeln "End With"

writeln vbCrLf & vbCrLf

writeln "Select Case Button.Key\n"

For idx In Buttons
        writeln "\tCase \q" & Buttons(idx) & "\q"
        writeln
Next

writeln "End Select\n"

Main = OutStr

End Function
