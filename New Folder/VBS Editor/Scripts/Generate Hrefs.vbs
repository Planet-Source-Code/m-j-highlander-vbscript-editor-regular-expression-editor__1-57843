Option UseEscapes

Function Main()

dim sHeader,sDigits,T,S,A

' change the following string
' but notice the "*" , it's the numeric value placeholder


T = "http://www.sietname.net/folder/page*.html"
S = ""
sDigits = "000"

For idx=100 to 1 step -1
        A = Replace(T,"*", Format(idx,sDigits))
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
