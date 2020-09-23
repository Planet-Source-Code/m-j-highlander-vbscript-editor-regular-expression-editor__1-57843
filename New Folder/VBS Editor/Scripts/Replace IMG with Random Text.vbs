
Function Main(Text)

Set C = New RegExp

C.Pattern ="<img[^\v]*?src=[^\v]*?>"
C.Global = True
C.IgnoreCase = True

Main = C.Replace(Text ,GetRef("R"))

End Function


Function R(Match , Index , FullText)
Dim sVar,iVar
        iVar = RndInt(6,12)
        sVar = RndStr(iVar)
        R=sVar
End Function

'// Helper Functions:

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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function RndStr (StrLen)
' generate a random string containing small case letters only

For idx = 1 To StrLen
	ch = Chr(RndInt(97, 122))
	tmp = tmp & ch
Next
RndStr = tmp

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function RndInt (Lower, Upper)

Randomize Timer
RndInt = Int(Rnd * (Upper - Lower + 1)) + Lower

End Function
