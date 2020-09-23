
Function Main(Text)


Set C = New RegExp

C.Pattern ="<A([^\v]*?)>([^\v]*?)</A>"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("RR"))

End Function



Private Function RR( Match , SubMatch1 , SubMatch2 , Index , FullText )
        RR = "<A " & RemoveChars(SubMatch1," " & vbCrLf & vbTab) & ">" & _
                 RemoveChars(SubMatch2," " & vbCrLf & vbTab)  & "</A>"
End Function

Private Function RemoveChars (ByVal Text , ByVal CharList)
dim i,ch

For i=1 to Len(CharList)
        ch=Mid(CharList,i,1)
        Text = Replace(Text,ch,"")
Next

RemoveChars =Text

End Function
