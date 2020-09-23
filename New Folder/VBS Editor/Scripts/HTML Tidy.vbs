
Public Function Main ( ByVal Text )
Dim TidyObj

Err.Supress
Set TidyObj = CreateObject("TidyCOM.TidyObject")
If Err Then
        MsgBox "Cannot Create Object of Type: ""TidyCOM.TidyObject"" " & vbCrLF & """TidyCOM.dll"" is not Registered",vbCritical,"Oops"
        Exit Function
End IF
Err.Allow

With TidyObj.Options
     
    ' generator meta tag
    .TidyMark = False
    
    'DocType:   'valid values: "auto" , "omit" , "strict" , "loose" , "any text you like..."
    .Doctype = "omit"
  
    'Default ALT text for IMG tags
    .AltText = "Image..."
    
    'Break before <BR> tag
    .BreakBeforeBr = True
    
    'Charachter encoding:   ascii=1 , iso2022=4 , latin1=2 , raw=0 , utf8=3
    .CharEncoding = ascii

    'Replace presentational tags and attrs by style rules
    .Clean = True
    
    'Discard empty paragraphs
    .DropEmptyParas = True

    'Discard <font> and <center> tags
    .DropFontTags = False

    'Enclose text in blocks whitin <P>'s
    .EncloseBlockText = False
    
    'Enclose text in BODY whitin <P>'s
    .EncloseText = True

    'Replace '\' in URLs by '/'
    .FixBackslash = True

    ' Fix bad comments
    .FixBadComments = True

    'Suppress optional end tags
    .HideEndtags = False
    
    'Indentation,   values: AutoIndent=2 , IndentBlocks=1 , NoIndent=0
    .Indent = AutoIndent
    
    ' Indent attributes
    .IndentAttributes = false

    'Number of spaces for indentation
    .IndentSpaces = 4

    'Preserve whitespace characters within attributes
    .LiteralAttributes = False

    'Output numeric character entities
    .NumericEntities = True

    'Uppercase tags
    .UppercaseTags = False

    'Right margin for line wrapping
    .Wrap = 0 'disable
    
    'Wrap attribute values
    .WrapAttributes = False


End With


Main = TidyObj.TidyMemToMem(Text)

End Function

