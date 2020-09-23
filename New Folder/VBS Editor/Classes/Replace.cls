VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRXEntityDecodeANSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Replacer(Match As String, Index As Long, FullText As String) As String
Attribute Replacer.VB_UserMemId = 0
Dim idx As Long, NumEnt As Long

Replacer = Match ' in case not found

If InStr(Match, "#") > 0 Then
    'Numeric Entity
    NumEnt = Val(Mid$(Match, 3, Len(Match) - 3))
    If NumEnt < 256 Then ' index = char code
                If EntityInfo(NumEnt).Asc127 = "REMOVE" Then
                    Replacer = ""
                ElseIf EntityInfo(NumEnt).Char = "" Then 'Unlisted Entry
                    Replacer = Chr$(NumEnt)
                Else
                    Replacer = EntityInfo(NumEnt).Char
                End If
    Else
                For idx = 256 To 300
                    If EntityInfo(idx).Code = NumEnt Then
                        Replacer = EntityInfo(idx).Char
                        Exit For
                    End If
                Next
                
    End If
Else
    'Named Entity
    For idx = 0 To 300
        If EntityInfo(idx).Name = Match Then
            Replacer = EntityInfo(idx).Char
            Exit For
        End If
    Next

End If


End Function
Private Sub unused()
'            Select Case Match
'
'            'ASCII_Entities:
'                Case "&nbsp;": Replacer = " "
'                Case "&gt;": Replacer = ">"
'                Case "&lt;": Replacer = "<"
'                Case "&quot;": Replacer = """"
'                Case "&amp;": Replacer = "&"
'
'            'Symbol_Entities:
'                Case "&iexcl;": Replacer = Chr$(161)
'                Case "&curren;": Replacer = Chr$(164)
'                Case "&cent;": Replacer = Chr$(162)
'                Case "&pound;": Replacer = Chr$(163)
'                Case "&yen;": Replacer = Chr$(165)
'                Case "&brvbar;": Replacer = Chr$(166)
'                Case "&sect;": Replacer = Chr$(167)
'                Case "&uml;": Replacer = Chr$(168)
'                Case "&copy;": Replacer = Chr$(169)
'                Case "&ordf;": Replacer = Chr$(170)
'                Case "&laquo;": Replacer = Chr$(171)
'                Case "&not;": Replacer = Chr$(172)
'                Case "&shy;": Replacer = Chr$(173)
'                Case "&reg;": Replacer = Chr$(174)
'                Case "&macr;": Replacer = Chr$(175)
'                Case "&deg;": Replacer = Chr$(176)
'                Case "&plusmn;": Replacer = Chr$(177)
'                Case "&sup2;": Replacer = Chr$(178)
'                Case "&sup3;": Replacer = Chr$(179)
'                Case "&acute;": Replacer = Chr$(180)
'                Case "&micro;": Replacer = Chr$(181)
'                Case "&para;": Replacer = Chr$(182)
'                Case "&middot;": Replacer = Chr$(183)
'                Case "&cedil;": Replacer = Chr$(184)
'                Case "&sup1;": Replacer = Chr$(185)
'                Case "&ordm;": Replacer = Chr$(186)
'                Case "&raquo;": Replacer = Chr$(187)
'                Case "&frac14;": Replacer = Chr$(188)
'                Case "&frac12;": Replacer = Chr$(189)
'                Case "&frac34;": Replacer = Chr$(190)
'                Case "&iquest;": Replacer = Chr$(191)
'                Case "&times;": Replacer = Chr$(215)
'                Case "&divide;": Replacer = Chr$(247)
'
'            'Character_Entities:
'                Case "&Agrave;": Replacer = Chr$(192)
'                Case "&Aacute;": Replacer = Chr$(193)
'                Case "&Acirc;": Replacer = Chr$(194)
'                Case "&Atilde;": Replacer = Chr$(195)
'                Case "&Auml;": Replacer = Chr$(196)
'                Case "&Aring;": Replacer = Chr$(197)
'                Case "&AElig;": Replacer = Chr$(198)
'                Case "&Ccedil;": Replacer = Chr$(199)
'                Case "&Egrave;": Replacer = Chr$(200)
'                Case "&Eacute;": Replacer = Chr$(201)
'                Case "&Ecirc;": Replacer = Chr$(202)
'                Case "&Euml;": Replacer = Chr$(203)
'                Case "&Igrave;": Replacer = Chr$(204)
'                Case "&Iacute;": Replacer = Chr$(205)
'                Case "&Icirc;": Replacer = Chr$(206)
'                Case "&Iuml;": Replacer = Chr$(207)
'                Case "&ETH;": Replacer = Chr$(208)
'                Case "&Ntilde;": Replacer = Chr$(209)
'                Case "&Ograve;": Replacer = Chr$(210)
'                Case "&Oacute;": Replacer = Chr$(211)
'                Case "&Ocirc;": Replacer = Chr$(212)
'                Case "&Otilde;": Replacer = Chr$(213)
'                Case "&Ouml;": Replacer = Chr$(214)
'                Case "&Oslash;": Replacer = Chr$(216)
'                Case "&Ugrave;": Replacer = Chr$(217)
'                Case "&Uacute;": Replacer = Chr$(218)
'                Case "&Ucirc;": Replacer = Chr$(219)
'                Case "&Uuml;": Replacer = Chr$(220)
'                Case "&Yacute;": Replacer = Chr$(221)
'                Case "&THORN;": Replacer = Chr$(222)
'                Case "&szlig;": Replacer = Chr$(223)
'                Case "&agrave;": Replacer = Chr$(224)
'                Case "&aacute;": Replacer = Chr$(225)
'                Case "&acirc;": Replacer = Chr$(226)
'                Case "&atilde;": Replacer = Chr$(227)
'                Case "&auml;": Replacer = Chr$(228)
'                Case "&aring;": Replacer = Chr$(229)
'                Case "&aelig;": Replacer = Chr$(230)
'                Case "&ccedil;": Replacer = Chr$(231)
'                Case "&egrave;": Replacer = Chr$(232)
'                Case "&eacute;": Replacer = Chr$(233)
'                Case "&ecirc;": Replacer = Chr$(234)
'                Case "&euml;": Replacer = Chr$(235)
'                Case "&igrave;": Replacer = Chr$(236)
'                Case "&iacute;": Replacer = Chr$(237)
'                Case "&icirc;": Replacer = Chr$(238)
'                Case "&iuml;": Replacer = Chr$(239)
'                Case "&eth;": Replacer = Chr$(240)
'                Case "&ntilde;": Replacer = Chr$(241)
'                Case "&ograve;": Replacer = Chr$(242)
'                Case "&oacute;": Replacer = Chr$(243)
'                Case "&ocirc;": Replacer = Chr$(244)
'                Case "&otilde;": Replacer = Chr$(245)
'                Case "&ouml;": Replacer = Chr$(246)
'                Case "&oslash;": Replacer = Chr$(248)
'                Case "&ugrave;": Replacer = Chr$(249)
'                Case "&uacute;": Replacer = Chr$(250)
'                Case "&ucirc;": Replacer = Chr$(251)
'                Case "&uuml;": Replacer = Chr$(252)
'                Case "&yacute;": Replacer = Chr$(253)
'                Case "&thorn;": Replacer = Chr$(254)
'                Case "&yuml;": Replacer = Chr$(255)
'
'                Case Else: Replacer = Match
'            End Select

End Sub

