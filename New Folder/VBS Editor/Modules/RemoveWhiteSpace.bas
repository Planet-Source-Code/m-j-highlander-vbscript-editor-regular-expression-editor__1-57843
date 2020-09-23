Function RemoveWhiteSpace(strText)
 Dim RegEx
 Set RegEx = New RegExp
 RegEx.Pattern = "\s+"
 RegEx.Multiline = True
 RegEx.Global = True
 strText = RegEx.Replace(strText, " ")
 RemoveWhiteSpace = strText
End Function

