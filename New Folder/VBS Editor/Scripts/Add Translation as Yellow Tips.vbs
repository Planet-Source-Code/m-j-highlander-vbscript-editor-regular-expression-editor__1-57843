Option UseEscapes

Public Function Main ( ByVal Text )

If Text = "" Then
        Exit Function
End If

sTips =GUI.InputMultiLine("Type or paste line-by-line translation:")

vTips = Split (sTips,vbCrLf)
vText = Split (Text,vbCrLf)

For idx = 0 To UBound(vTips)
        If vText(idx) <> "" Then
                vText(idx) = "<span title=\q" & vTips(idx) & "\q>" & vText(idx) & "</span>"
        End If
Next

Main = Join (vText,"<br>" & vbCrLf)

End Function

