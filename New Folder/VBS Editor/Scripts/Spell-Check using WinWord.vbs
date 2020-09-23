Option UseEscapes

Public Function Main ( ByVal Text )


Main = ChkSpelling(Text)

End Function

Function ChkSpelling(TextValue)
' Comments to alexangelopoulos@hotmail.com

     Dim objWord, objDocument, strReturnValue

     Set objWord = CreateObject("word.Application")
     objWord.WindowState = 2
     objWord.Visible = False

     'Create a new instance of Document
     Set objDocument = objWord.Documents.Add( , , 1, True)
     objDocument.Content=TextValue
     objDocument.CheckSpelling

     'Return checked text and quit Word
     strReturnValue = objDocument.Content
     objDocument.Close False
     objWord.Application.Quit True

     ChkSpelling = strReturnValue

End function
