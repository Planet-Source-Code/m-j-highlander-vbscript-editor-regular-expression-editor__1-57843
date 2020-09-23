
Public Function Main ( ByVal Text )
Dim vHead(22),sHead
vHead(0)="<script language=""vbscript"">"
vHead(1)="dim vImgs (1000)"
vHead(2)="dim vFlagImgsVisible"
vHead(3)="Function Main()"
vHead(4)="vFlagImgsVisible = true"
vHead(5)="for i=0 to document.images.length-1"
vHead(6)="    vImgs(i)=document.images(i).src"
vHead(7)="next"
vHead(8)="End Function"
vHead(9)="Function PopImgs()"
vHead(10)="if vFlagImgsVisible=true then"
vHead(11)="   for i=0 to document.images.length-1"
vHead(12)="           document.images(i).src="""""
vHead(13)="           document.images(i).alt="""""
vHead(14)="   next"
vHead(15)="else"
vHead(16)="   for i=0 to document.images.length-1"
vHead(17)="           document.images(i).src=vImgs(i)"
vHead(18)="   next"
vHead(19)="end if"
vHead(20)="vFlagImgsVisible = not vFlagImgsVisible"
vHead(21)="End function"
vHead(22)="</script>"
sHead = Join (vHead,vbcrlf)

Text = RegExpReplace ( Text , "(<head[^\v]*?)</head>" , "$1" & vbcrlf & sHead & vbcrlf & "</head>" )

Text = RegExpReplace ( Text , "(<body[^\v]*?)>" , "$1 OnLoad=""Main"">" & vbCrlf & "<span OnClick=""PopImgs"">POP</span><BR>")



Main = Text

End Function
