
Public Function Main ( ByVal Text )

dim b()
a = split(Text,vbCrlf)

redim b(ubound(a))

For i In a

        b(i)=TableRow(a(i))
next

Main = "<table border=""1"">" & Join(b) & "</table>"

End Function



'-----------------------------------------------------

private Function TableRow ( sLine )

v = split ( sLine, "," )

For idx In v
        v(idx) = "<td>" & v(idx) & "</td>"
next

s = join (v,vbCrLf)
s = "<tr>" & s & "</tr>"
TableRow = s


End function

