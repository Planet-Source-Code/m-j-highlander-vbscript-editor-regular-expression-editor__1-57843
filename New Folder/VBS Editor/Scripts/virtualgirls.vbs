
Option UseEscapes

Function Main()

T = "http://banners.toteme.com/virtuagirl/galleries/*_1/index.php"

S=AxiomPath() &  ".txt"
S=loadFile(S)
V=Split(S,vbCrLf)

S=""
For idx=lbound(V) to ubound(V)
        A = Replace(T,"*", V(idx))
        S = S & A & vbcrlf
Next

S = List2Links(S)

sHeader = "<html>\n<head>\n<title></title>\n<base target=\q_blank\q>\n</head>\n<body>\n<center>\n\n"

S = AddFirst(S,sHeader)

Main = S

End Function


'************** Helper Functions **************

Private Function List2Links(Text)
Dim V,idx,f

V=Split(Text,vbCrLf)

For idx=lbound(v) to ubound(v)
       'f = GetFileName (v(idx)) 
       f = v(idx)
       If v(idx) <> "" Then v(idx)="<a href=\q" & v(idx) & "\q>" & f & "</a><br>"
Next

List2Links = Join(v,vbcrlf)

End Function

Private Function GetFileName(FilePath)
Dim iLastSlash

    iLastSlash = InStrRev(FilePath, "/")
    
    If iLastSlash = 0 Then
            GetFileName = FilePath
    Else
            GetFileName = Right(FilePath, Len(FilePath) - iLastSlash)
    End If

End Function
