Function HTML_FirstChar(Text)

sColor=inputbox("Font Color")
sSize=inputbox("Font Size")

a = Split(Text, " ")

For idx = LBound(a) To UBound(a)
        Ch = Left(a(idx), 1)
        NewCh = "<FONT COLOR=""%COLOR%"" SIZE=""%SIZE%"">" & Ch & "</FONT>"
        NewCh = Replace(NewCh, "%COLOR%", sColor)
        NewCh = Replace(NewCh, "%SIZE%", sSize)
        a(idx) = NewCh & Right(a(idx),len(a(idx))-1)
Next 


Text = Join(a, " ")

' Handle first char of each line

a=""
a = Split(Text, vbCrLf)
For idx = LBound(a) + 1 To UBound(a)
        Ch = Left(a(idx), 1)
        NewCh = "<FONT COLOR=""%COLOR%"" SIZE=""%SIZE%"">" & Ch & "</FONT>"
        NewCh = Replace(NewCh, "%COLOR%", sColor)
        NewCh = Replace(NewCh, "%SIZE%", sSize)
       if len(a(idx))>0 then
            a(idx) = NewCh & Right ( a(idx) , len(a(idx))-1 )
       end if
Next 
Text = Join(a, vbCrLf)


HTML_FirstChar = Text
End Function
