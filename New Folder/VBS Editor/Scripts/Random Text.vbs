Function RandomText()

'  Generate random text.
'  Useful in creating a piece of text for testing purposes

NumWords=InputBox("Number of Words?")
Randomize

For i = 1 To NumWords
    lngth = Rnd * 8
    tmp = tmp & " " & RndStr(lngth)
Next

RandomText = Trim(tmp)

End Function


Function RndStr (StrLen)
' generate a random string containing small case letters only

For idx = 1 To StrLen
	ch = Chr(RndInt(97, 122))
	tmp = tmp & ch
Next

RndStr = tmp

End Function

Function RndInt (Lower, Upper)

Randomize Timer
RndInt = Int(Rnd * (Upper - Lower + 1)) + Lower

End Function
