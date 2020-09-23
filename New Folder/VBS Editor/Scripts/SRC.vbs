
Function Main(Text)


Set C = New RegExp

C.Pattern ="src\="".*?/(.*?)"""
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("RR"))

End Function



Private Function RR( Match , SubMatch1 , Index , FullText )
        RR = "SRC=" & vbQuote & "imgs/"  & SubMatch1 & vbQuote
End Function

Private Function RemoveChars (ByVal Text , ByVal CharList)
dim i,ch

For i=1 to Len(CharList)
        ch=Mid(CharList,i,1)
        Text = Replace(Text,ch,"")
Next

RemoveChars =Text

End Function
