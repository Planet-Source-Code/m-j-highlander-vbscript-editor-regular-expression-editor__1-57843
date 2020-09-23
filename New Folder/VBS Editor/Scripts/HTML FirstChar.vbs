Public Function HTML_FirstChar(ByVal sText)
Dim idx , a , Ch , NewCh , V , sColor , sSize



V = InputForm("Font Color","Font Size")

If Not IsArray(V) Then Exit Function

sColor = V(0)
sSize  = V(1)

a = Split(sText, " ")
For idx = LBound(a) To UBound(a)
        Ch = Left(a(idx), 1)
        Select Case Ch
        Case vbCr,vbLF
        Case Else
                NewCh = "<FONT COLOR=""%COLOR%"" SIZE=""%SIZE%"">" & Ch & "</FONT>"
                NewCh = Replace(NewCh, "%COLOR%", sColor)
                NewCh = Replace(NewCh, "%SIZE%", sSize)
                a(idx) = Replace(a(idx), Ch, NewCh,1 , 1)
        End Select
Next
sText = Join(a, " ")
''''''''''''''''''''''''''''''
a = Split(sText, vbCrLf)
For idx = LBound(a) + 1 To UBound(a)
                Ch = Left(a(idx), 1)
                Select Case Ch
                Case vbCr,vbLF
                Case Else
                        NewCh = "<FONT COLOR=""%COLOR%"" SIZE=""%SIZE%"">" & Ch & "</FONT>"
                        NewCh = Replace(NewCh, "%COLOR%", sColor)
                        NewCh = Replace(NewCh, "%SIZE%", sSize)
                        a(idx) = Replace(a(idx), Ch, NewCh,1 , 1)
                End Select
Next

'use this for multi-line text
'sText = Join(a, "<br>" & vbCrLf)

sText = Join(a,vbCrLf)

HTML_FirstChar = sText
End Function