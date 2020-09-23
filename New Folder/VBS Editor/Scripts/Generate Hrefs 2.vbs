Option UseEscapes

Function Main()

' Notice the "*" , it's the numeric value placeholder:
' http://www.sietname.net/folder/page*.html

Dim S,V,iStart,iStop,iStep,iNumDigits,sSiteName,sHeader,sDigits,T,A

V = InputForm("URL (replace * for number):|http://www.sietname.net/folder/page*.html","Start|1","End|100","Number of Digits|3")

If IsArray(v) Then
        S = ""
        iStart = CLng(V(1))
        iStop = CLng(V(2))
        iNumDigits = CLng(V(3))
        sSiteName = V(0)

        sDigits = String(iNumDigits,"0")
        If iStart > iStop Then iStep = -1 Else iStep =1
        For idx=iStart to IStop Step iStep
                A = Replace(sSiteName,"*", Format(idx,sDigits))
                S = S & A & vbcrlf
        Next
        S = List2Links(S)
        sHeader = "<html>\n<head>\n<title></title>\n<base target=\q_blank\q>\n</head>\n<body>\n<center>\n\n"
        S = AddFirst(S,sHeader)
Else
        S=""
End If
        
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
