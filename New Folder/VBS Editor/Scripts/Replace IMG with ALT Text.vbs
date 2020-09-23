
Function Main(Text)

Set C = New RegExp

C.Pattern ="<img[^\v]*?alt=([^\v]*?)>"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("R"))

End Function


Function R(Match ,SubMatch1, Index , FullText)
Dim sTemp,iPos

If Left(SubMatch1,1)=vbQuote then
        iPos=InStr(2,SubMatch1,vbQuote)
        sTemp=Mid(SubMatch1,2,iPos-2)
Else
        iPos=MultiInStr(1,SubMatch1,vbQuote & " " & vbTab & vbCrLf)
        sTemp=Mid(SubMatch1,1,iPos-1)
End If

        R=sTemp
End Function

''''''''Helper Function
Private Function MultiInstr(iStart , sText , sLookFor )
Const MAXINT = 32767
Dim iLen,chLookFor,idx,iPos,iFirstPos

iFirstPos = MAXINT
iLen = Len(sLookFor)
ReDim LookForChars(iLen)
For idx = 1 To iLen
    LookForChars(idx) = Mid(sLookFor, idx, 1)
    iPos = InStr(iStart, sText, LookForChars(idx))
    If (iPos <> 0 And iPos < iFirstPos) Then iFirstPos = iPos
Next 

MultiInstr = iFirstPos

End Function
