'DESCRIPTION:        Convert VB Code to Colored HTML using GeWare VB2HTML Library ver 1.01 (gwVB2HTMLlib.dll)
'                               "gwVB2HTMLlib.dll" is an ActiveX DLL and must be registered first.
'
'Class Name (AppName.ObjectType) : "gwVB2HTMLLib.gwVB2HTML"
'
'Methods:
'  ParseFile(FileName As String) As String

'  ParseString(Source As String) As String

'
'Properties (w/ Default Values):
'  CommentColor As Long = 32768
'  KeywordColor As Long = 8388608
'  NormalColor As Long = 0
'  FontName As String = "Courier New"



Public Function Main ( ByVal Text )

dim objDLL

Set objDLL = CreateObject("gwVB2HTMLLib.gwVB2HTML")

Text = objDLL.ParseString ( Text )


Set objDLL = Nothing

Main = Text

End Function

