' VBS Spell Checker
' This script calls Word to spell check, you need to have Microsoft Word Installed.
' Adopted from a program by George - Interactive PsyberTechnology Developers Group (IPDG3)
' www.ipdg3.com - info@ipdg3.com - Aurora, IL. USA

Public Function WinWord_SpellCheck( ByVal Text )

Dim oSpellChecker , stText , stNew_Text , iPosition


Set oSpellChecker = CreateObject("Word.Basic")

oSpellChecker.FileNew
oSpellChecker.Insert Text
oSpellChecker.ToolsSpelling
oSpellChecker.EditSelectAll
stText = oSpellChecker.Selection()
oSpellChecker.FileExit 2        ' 2 = Exit without saving

If Right(stText, 1) = vbCr Then stText = Left(stText, Len(stText) - 1)
stNew_Text = ""
iPosition = InStr(stText, vbCr)
Do While iPosition > 0
        stNew_Text = stNew_Text & Left(stText, iPosition - 1) & vbCrLf
        stText = Right(stText, Len(stText) - iPosition)
        iPosition = InStr(stText, vbCr)
Loop

stNew_Text = stNew_Text & stText

WinWord_SpellCheck = stNew_Text

'MsgBox "Spell Check Complete"


End Function

