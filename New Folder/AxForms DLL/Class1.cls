VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Forms"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function BrowseForFile(Optional ByVal Filename As String, _
                              Optional ByVal Filter As String, _
                              Optional ByVal Title As String, _
                              Optional ByVal FileMustExist As Boolean = True, _
                              Optional ByVal MultiSelect As Boolean = False _
                              ) As String

On Error GoTo BrowseForFiles_ErrorHandler

With frmHidden.FileDialog
    .CancelError = True
    .Filename = Filename
    .Filter = Filter
    .DialogTitle = Title
    .Flags = cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNExplorer
    If FileMustExist Then .Flags = .Flags Or cdlOFNFileMustExist
    If MultiSelect Then
        .Flags = .Flags Or cdlOFNAllowMultiselect
        .MaxFileSize = 32767    '= maximum possible value 32K
    End If
    .ShowOpen
    BrowseForFile = .Filename
    
End With

Exit Function
BrowseForFiles_ErrorHandler:
        
        If Err.Number = cdlCancel Then 'Cancel was selected
            BrowseForFile = ""
            Err.Clear
        Else       'Other Error
            BrowseForFile = ""
            Err.Raise Err.Number    'Raise the Error to the Calling Function
        End If

End Function
Public Function InputHTMLSource(Optional ByVal DefaultText As String, Optional ByVal Caption As String = "Type HTML Source:") As String

   InputHTMLSource = frmInputHTML.Exec(Caption, DefaultText)

   Unload frmInputHTML

End Function
Public Function InputHTML(Optional ByVal DefaultText As String, Optional ByVal DefaultText_IsHTML As Boolean = False) As String

   InputHTML = frmSmall_HTML_Editor.Exec(DefaultText, DefaultText_IsHTML)

   Unload frmSmall_HTML_Editor

End Function
Public Function BrowseForFolder(Optional ByVal InitDir As String = "", Optional ByVal Title As String = "Select a Folder:") As String

BrowseForFolder = BrowseDir.BrowseForFolder(Title, InitDir)

End Function
Public Function InputForm(Optional ByVal Text1 As String, _
                          Optional ByVal Text2 As String, _
                          Optional ByVal Text3 As String, _
                          Optional ByVal Text4 As String, _
                          Optional ByVal Text5 As String) As Variant

   InputForm = frmInputBoxes.Exec(Text1, Text2, Text3, Text4, Text5)
   
   Unload frmInputBoxes

End Function
Public Function InputMultiline(ByVal Caption As String, Optional ByVal DefaultText As String) As String

   InputMultiline = frmInputText.Exec(Caption, DefaultText)
   
   Unload frmInputText


End Function
Public Function OptionForm(ByVal Prompt As String, _
                               Optional ByVal Option1 As String, _
                               Optional ByVal Option2 As String, _
                               Optional ByVal Option3 As String, _
                               Optional ByVal Option4 As String, _
                               Optional ByVal Option5 As String) As String

   OptionForm = frmAxOptionForm.Exec(Prompt, Option1, Option2, Option3, Option4, Option5)
   
   Unload frmAxOptionForm

End Function
