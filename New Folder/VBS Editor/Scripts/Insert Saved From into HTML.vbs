
Public Function Main ( ByVal Text )

sLines = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbCrLf & _
"<!-- saved from url=(*)$URL$ -->" & vbCrLf


sURL = INputBox ( "Enter URL:","Input","http://")

if sURL = "" Then Exit Function

if Right(sURL,1) <> "/" Then sURL = sURL & "/"

sLines = Replace ( sLines , "$URL$", sURL)
sLines = Replace ( sLines , "*", Format(Len(sURL),"0000"))

Main = sLines & Text

End Function

