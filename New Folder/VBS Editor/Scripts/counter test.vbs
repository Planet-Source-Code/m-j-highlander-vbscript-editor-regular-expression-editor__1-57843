Option UseEscapes

Public Function Main ( ByVal Text )
dim linx(25)

' initialize auto-increment counter: start=0, step=1
vbCounter 0,1

for i = vbKeyA to vbKeyZ

        j= vbCounter        'don't use it more than once if you want the same value, it will auto-increment !

        linx(j)=stringf("http://www.caratuleo.com/lista.php?letra=%s&pag=1",lcase(chr(i)))
        linx(j)="<A HREF=\q" & linx(j) & "\q>" & chr(i) & "</a> | "

next

Main = join(linx,"\n")

End Function

