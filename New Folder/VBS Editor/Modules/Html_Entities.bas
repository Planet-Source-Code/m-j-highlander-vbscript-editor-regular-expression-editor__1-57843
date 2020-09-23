Attribute VB_Name = "HTML_Entities"
Option Explicit

Public Type EntityInfo
    Name As String
    Code As Long
    Char As String
    Asc127 As String
End Type

Public EntityInfo(0 To 300) As EntityInfo

Public Sub InitEntityInfo()

If EntityInfo(131).Code = 131 <> 0 Then Exit Sub  ' already initialized

Call InitEntityInfo1
Call InitEntityInfo2
Call InitEntityInfo3

End Sub

Private Sub InitEntityInfo3()
'''' All Codes here are duplicates

EntityInfo(256).Code = 8482    ' same as 153
EntityInfo(256).Char = Chr$(153)
EntityInfo(256).Name = "&trade;"
EntityInfo(256).Asc127 = "(tm)"

EntityInfo(257).Code = 338     ' same as 140
EntityInfo(257).Char = Chr$(140)
EntityInfo(257).Name = "&OElig;"
EntityInfo(257).Asc127 = "OE"

EntityInfo(258).Code = 339     ' 156
EntityInfo(258).Char = Chr$(156)
EntityInfo(258).Name = "&oelig;"
EntityInfo(258).Asc127 = "oe"

EntityInfo(259).Code = 352     ' 138
EntityInfo(259).Char = Chr$(138)
EntityInfo(259).Name = "&Scaron;"
EntityInfo(259).Asc127 = "S"

EntityInfo(260).Code = 353     ' 154
EntityInfo(260).Char = Chr$(154)
EntityInfo(260).Name = "&scaron;"
EntityInfo(260).Asc127 = "s"

EntityInfo(261).Code = 376    ' 159
EntityInfo(261).Char = Chr$(159)
EntityInfo(261).Name = "&Yuml;"
EntityInfo(261).Asc127 = "Y"

EntityInfo(262).Code = 710    ' 136
EntityInfo(262).Char = Chr$(136)
EntityInfo(262).Name = "&circ;"
EntityInfo(262).Asc127 = "^"

EntityInfo(263).Code = 732     ' 152
EntityInfo(263).Char = Chr$(152)
EntityInfo(263).Name = "&tilde;"
EntityInfo(263).Asc127 = "~"

EntityInfo(264).Code = 8194
EntityInfo(264).Char = ""
EntityInfo(264).Name = "&ensp;"
EntityInfo(264).Asc127 = " "

EntityInfo(265).Code = 8195
EntityInfo(265).Char = ""
EntityInfo(265).Name = "&emsp;"
EntityInfo(265).Asc127 = " "

EntityInfo(266).Code = 8201
EntityInfo(266).Char = ""
EntityInfo(266).Name = "&thinsp;"
EntityInfo(266).Asc127 = " "

EntityInfo(267).Code = 8204
EntityInfo(267).Char = ""
EntityInfo(267).Name = "&zwnj;"
EntityInfo(267).Asc127 = "|"

EntityInfo(268).Code = 8205
EntityInfo(268).Char = ""
EntityInfo(268).Name = "&zwj;"
EntityInfo(268).Asc127 = "|"

EntityInfo(269).Code = 8206
EntityInfo(269).Char = ""
EntityInfo(269).Name = "&lrm;"
EntityInfo(269).Asc127 = ""

EntityInfo(270).Code = 8207
EntityInfo(270).Char = ""
EntityInfo(270).Name = "&rlm;"
EntityInfo(270).Asc127 = ""

EntityInfo(271).Code = 8211  '150
EntityInfo(271).Char = Chr$(150)
EntityInfo(271).Name = "&ndash;"
EntityInfo(271).Asc127 = "-"

EntityInfo(272).Code = 8212  '151
EntityInfo(272).Char = Chr$(151)
EntityInfo(272).Name = "&mdash;"
EntityInfo(272).Asc127 = "-"

EntityInfo(273).Code = 8216   '145
EntityInfo(273).Char = Chr$(145)
EntityInfo(273).Name = "&lsquo;"
EntityInfo(273).Asc127 = "`"  ' OR "'"

EntityInfo(274).Code = 8217  '146
EntityInfo(274).Char = Chr$(146)
EntityInfo(274).Name = "&rsquo;"
EntityInfo(274).Asc127 = "'"

EntityInfo(275).Code = 8218     '130
EntityInfo(275).Char = Chr$(130)
EntityInfo(275).Name = "&sbquo;"
EntityInfo(275).Asc127 = "'"

EntityInfo(276).Code = 8220   '147
EntityInfo(276).Char = Chr$(147)
EntityInfo(276).Name = "&ldquo;"
EntityInfo(276).Asc127 = """"

EntityInfo(277).Code = 8221   '148
EntityInfo(277).Char = Chr$(148)
EntityInfo(277).Name = "&rdquo;"
EntityInfo(277).Asc127 = """"

EntityInfo(278).Code = 8222  '132
EntityInfo(278).Char = Chr$(132)
EntityInfo(278).Name = "&bdquo;"
EntityInfo(278).Asc127 = """"

EntityInfo(279).Code = 8224   '134
EntityInfo(279).Char = Chr$(134)
EntityInfo(279).Name = "&dagger;"
EntityInfo(279).Asc127 = ""

EntityInfo(280).Code = 8225   '135
EntityInfo(280).Char = Chr$(135)
EntityInfo(280).Name = "&Dagger;"
EntityInfo(280).Asc127 = ""

EntityInfo(281).Code = 8240     '137
EntityInfo(281).Char = Chr$(137)
EntityInfo(281).Name = "&permil;"
EntityInfo(281).Asc127 = "0/00"  'per 1000

EntityInfo(282).Code = 8249      '139
EntityInfo(282).Char = Chr$(139)
EntityInfo(282).Name = "&lsaquo;"
EntityInfo(282).Asc127 = "<"

EntityInfo(283).Code = 8250   '155
EntityInfo(283).Char = Chr$(155)
EntityInfo(283).Name = "&rsaquo;"
EntityInfo(283).Asc127 = ">"

EntityInfo(284).Code = 8364   ' 128
EntityInfo(284).Char = Chr$(128)
EntityInfo(284).Name = "&euro;"
EntityInfo(284).Asc127 = "euro"

EntityInfo(285).Code = 402    'also 131
EntityInfo(285).Char = Chr$(131)
EntityInfo(285).Name = "&fnof;"
EntityInfo(285).Asc127 = "f"

EntityInfo(286).Code = 8230   'also 133
EntityInfo(286).Char = Chr$(133)
EntityInfo(286).Name = "&hellip;"
EntityInfo(286).Asc127 = "..."

EntityInfo(287).Code = 8226   'also 149
EntityInfo(287).Char = Chr$(149)
EntityInfo(287).Name = "&bull;"
EntityInfo(287).Asc127 = "*"



End Sub
Private Sub InitEntityInfo1()
Dim idx As Integer
' ASCII 127 (0-126)
''''''''''''''''''''
' 0 to 31 - Control Chars (Considered Binary not Text chars)
' Not used in HTML, Except: 9=Tab, 10=LineFeed, 13=Carraige Return
' (Multiple) Tab, CrLf, and Space are used directly and they all
' evaluate to a single space.

' 32 to 126 (inclusive) are used directly (no entities), EXCEPT:
' Quote, Space, Ampersand, Less-Than, Greater-Than

'Although 9,10 and 13 are used in html, using them as Entities makes
' no sense, so we remove all range 0 --> 31
For idx = 0 To 31
    EntityInfo(idx).Code = idx
    'EntityInfo(idx).Char = ""
    'EntityInfo(idx).Name = ""
    EntityInfo(idx).Asc127 = "REMOVE"  'special value, used in calling Subs
Next idx

EntityInfo(32).Code = 32  '<----------- ?
EntityInfo(32).Char = " "
EntityInfo(32).Name = "&nbsp;"
EntityInfo(32).Asc127 = " "

EntityInfo(34).Code = 34
EntityInfo(34).Char = Chr$(34)
EntityInfo(34).Name = "&quot;"
EntityInfo(34).Asc127 = Chr$(34)

EntityInfo(38).Code = 38
EntityInfo(38).Char = "&"
EntityInfo(38).Name = "&amp;"
EntityInfo(38).Asc127 = "&"

EntityInfo(60).Code = 60
EntityInfo(60).Char = "<"
EntityInfo(60).Name = "&lt;"
EntityInfo(60).Asc127 = "<"

EntityInfo(62).Code = 62
EntityInfo(62).Char = ">"
EntityInfo(62).Name = "&gt;"
EntityInfo(62).Asc127 = ">"

'''''''''''''''''''''''''''''
' ASCII 255 (127-255)
'''''''''''''''''''''''''''''
EntityInfo(127).Code = 127  'UNUSED
EntityInfo(127).Char = ""
EntityInfo(127).Name = ""
EntityInfo(127).Asc127 = ""

EntityInfo(128).Code = 128
EntityInfo(128).Char = Chr$(128)
EntityInfo(128).Name = "&euro;"
EntityInfo(128).Asc127 = "euro"

EntityInfo(129).Code = 129  'UNUSED
EntityInfo(129).Char = ""
EntityInfo(129).Name = ""
EntityInfo(129).Asc127 = ""

EntityInfo(130).Code = 130
EntityInfo(130).Char = Chr$(130)
EntityInfo(130).Name = "&sbquo;"
EntityInfo(130).Asc127 = "'"

EntityInfo(131).Code = 131   'also 402
EntityInfo(131).Char = Chr$(131)
EntityInfo(131).Name = "&fnof;"
EntityInfo(131).Asc127 = "f"

EntityInfo(132).Code = 132
EntityInfo(132).Char = Chr$(132)
EntityInfo(132).Name = "&bdquo;"
EntityInfo(132).Asc127 = """"

EntityInfo(133).Code = 133 ''''''''''''''''''''''ALSO 8230
EntityInfo(133).Char = Chr$(133)
EntityInfo(133).Name = "&hellip;"
EntityInfo(133).Asc127 = "..."

EntityInfo(134).Code = 134
EntityInfo(134).Char = Chr$(134)
EntityInfo(134).Name = "&dagger;"
EntityInfo(134).Asc127 = ""

EntityInfo(135).Code = 135
EntityInfo(135).Char = Chr$(135)
EntityInfo(135).Name = "&Dagger;"
EntityInfo(135).Asc127 = ""

EntityInfo(136).Code = 136
EntityInfo(136).Char = Chr$(136)
EntityInfo(136).Name = "&circ;"
EntityInfo(136).Asc127 = "^"

EntityInfo(137).Code = 137
EntityInfo(137).Char = Chr$(137)
EntityInfo(137).Name = "&permil;"
EntityInfo(137).Asc127 = "0/00"  'per 1000

EntityInfo(138).Code = 138
EntityInfo(138).Char = Chr$(138)
EntityInfo(138).Name = "&Scaron;"
EntityInfo(138).Asc127 = "S"

EntityInfo(139).Code = 139
EntityInfo(139).Char = Chr$(139)
EntityInfo(139).Name = "&lsaquo;"
EntityInfo(139).Asc127 = "<"

EntityInfo(140).Code = 140
EntityInfo(140).Char = Chr$(140)
EntityInfo(140).Name = "&OElig;"
EntityInfo(140).Asc127 = "OE"

EntityInfo(141).Code = 141   '''UNUSED
EntityInfo(141).Char = ""
EntityInfo(141).Name = ""
EntityInfo(141).Asc127 = ""

EntityInfo(142).Code = 142 '''''''''''''NO NAME ?''''''''''''''
EntityInfo(142).Char = Chr$(142)
EntityInfo(142).Name = ""
EntityInfo(142).Asc127 = "Z"

EntityInfo(143).Code = 143   '''UNUSED
EntityInfo(143).Char = ""
EntityInfo(143).Name = ""
EntityInfo(143).Asc127 = ""

EntityInfo(144).Code = 144   '''UNUSED
EntityInfo(144).Char = ""
EntityInfo(144).Name = ""
EntityInfo(144).Asc127 = ""


EntityInfo(145).Code = 145
EntityInfo(145).Char = Chr$(145)
EntityInfo(145).Name = "&lsquo;"
EntityInfo(145).Asc127 = "`"  ' OR "'"

EntityInfo(146).Code = 146
EntityInfo(146).Char = Chr$(146)
EntityInfo(146).Name = "&rsquo;"
EntityInfo(146).Asc127 = "'"

EntityInfo(147).Code = 147
EntityInfo(147).Char = Chr$(147)
EntityInfo(147).Name = "&ldquo;"
EntityInfo(147).Asc127 = """"

EntityInfo(148).Code = 148
EntityInfo(148).Char = Chr$(148)
EntityInfo(148).Name = "&rdquo;"
EntityInfo(148).Asc127 = """"

EntityInfo(149).Code = 149      ''''''ALSO 8226
EntityInfo(149).Char = Chr$(149)
EntityInfo(149).Name = "&bull;"
EntityInfo(149).Asc127 = "*"

EntityInfo(150).Code = 150
EntityInfo(150).Char = Chr$(150)
EntityInfo(150).Name = "&ndash;"
EntityInfo(150).Asc127 = "-"

EntityInfo(151).Code = 151
EntityInfo(151).Char = Chr$(151)
EntityInfo(151).Name = "&mdash;"
EntityInfo(151).Asc127 = "-"

EntityInfo(152).Code = 152
EntityInfo(152).Char = Chr$(152)
EntityInfo(152).Name = "&tilde;"
EntityInfo(152).Asc127 = "~"

EntityInfo(153).Code = 153
EntityInfo(153).Char = Chr$(153)
EntityInfo(153).Name = "&trade;"
EntityInfo(153).Asc127 = "(tm)"

EntityInfo(154).Code = 154
EntityInfo(154).Char = Chr$(154)
EntityInfo(154).Name = "&scaron;"
EntityInfo(154).Asc127 = "s"

EntityInfo(155).Code = 155
EntityInfo(155).Char = Chr$(155)
EntityInfo(155).Name = "&rsaquo;"
EntityInfo(155).Asc127 = ">"

EntityInfo(156).Code = 156
EntityInfo(156).Char = Chr$(156)
EntityInfo(156).Name = "&oelig;"
EntityInfo(156).Asc127 = "oe"

EntityInfo(157).Code = 157   '''UNUSED
EntityInfo(157).Char = ""
EntityInfo(157).Name = ""
EntityInfo(157).Asc127 = ""

EntityInfo(158).Code = 158   '''NO NAME ?'''''''''''''''''''
EntityInfo(158).Char = Chr$(158)
EntityInfo(158).Name = ""
EntityInfo(158).Asc127 = "z"

EntityInfo(159).Code = 159
EntityInfo(159).Char = Chr$(159)
EntityInfo(159).Name = "&Yuml;"
EntityInfo(159).Asc127 = "Y"

EntityInfo(160).Code = 160
EntityInfo(160).Char = " "
EntityInfo(160).Name = "&nbsp;"
EntityInfo(160).Asc127 = " "

EntityInfo(161).Code = 161
EntityInfo(161).Char = Chr$(161)
EntityInfo(161).Name = "&iexcl;"
EntityInfo(161).Asc127 = "!"

EntityInfo(162).Code = 162
EntityInfo(162).Char = Chr$(162)
EntityInfo(162).Name = "&cent;"
EntityInfo(162).Asc127 = "cent"

EntityInfo(163).Code = 163
EntityInfo(163).Char = Chr$(163)
EntityInfo(163).Name = "&pound;"
EntityInfo(163).Asc127 = "pound"

EntityInfo(164).Code = 164
EntityInfo(164).Char = Chr$(164)
EntityInfo(164).Name = "&curren;"
EntityInfo(164).Asc127 = ""

EntityInfo(165).Code = 165
EntityInfo(165).Char = Chr$(165)
EntityInfo(165).Name = "&yen;"
EntityInfo(165).Asc127 = "yen"

EntityInfo(166).Code = 166
EntityInfo(166).Char = Chr$(166)
EntityInfo(166).Name = "&brvbar;"
EntityInfo(166).Asc127 = "|"

EntityInfo(167).Code = 167
EntityInfo(167).Char = Chr$(167)
EntityInfo(167).Name = "&sect;"
EntityInfo(167).Asc127 = ""

EntityInfo(168).Code = 168
EntityInfo(168).Char = Chr$(168)
EntityInfo(168).Name = "&uml;"
EntityInfo(168).Asc127 = ""

EntityInfo(169).Code = 169
EntityInfo(169).Char = Chr$(169)
EntityInfo(169).Name = "&copy;"
EntityInfo(169).Asc127 = "(C)"

EntityInfo(170).Code = 170
EntityInfo(170).Char = Chr$(170)
EntityInfo(170).Name = "&ordf;"
EntityInfo(170).Asc127 = "a"

EntityInfo(171).Code = 171
EntityInfo(171).Char = Chr$(171)
EntityInfo(171).Name = "&laquo;"
EntityInfo(171).Asc127 = """"

EntityInfo(172).Code = 172
EntityInfo(172).Char = Chr$(172)
EntityInfo(172).Name = "&not;"
EntityInfo(172).Asc127 = "-"

EntityInfo(173).Code = 173
EntityInfo(173).Char = Chr$(173)
EntityInfo(173).Name = "&shy;"
EntityInfo(173).Asc127 = ""

EntityInfo(174).Code = 174
EntityInfo(174).Char = Chr$(174)
EntityInfo(174).Name = "&reg;"
EntityInfo(174).Asc127 = "(R)"

EntityInfo(175).Code = 175
EntityInfo(175).Char = Chr$(175)
EntityInfo(175).Name = "&macr;"
EntityInfo(175).Asc127 = " "

EntityInfo(176).Code = 176
EntityInfo(176).Char = Chr$(176)
EntityInfo(176).Name = "&deg;"
EntityInfo(176).Asc127 = "deg"

EntityInfo(177).Code = 177
EntityInfo(177).Char = Chr$(177)
EntityInfo(177).Name = "&plusmn;"
EntityInfo(177).Asc127 = "+/-"

EntityInfo(178).Code = 178
EntityInfo(178).Char = Chr$(178)
EntityInfo(178).Name = "&sup2;"
EntityInfo(178).Asc127 = "^2"

EntityInfo(179).Code = 179
EntityInfo(179).Char = Chr$(179)
EntityInfo(179).Name = "&sup3;"
EntityInfo(179).Asc127 = "^3"

EntityInfo(180).Code = 180
EntityInfo(180).Char = Chr$(180)
EntityInfo(180).Name = "&acute;"
EntityInfo(180).Asc127 = "'"

EntityInfo(181).Code = 181
EntityInfo(181).Char = Chr$(181)
EntityInfo(181).Name = "&micro;"
EntityInfo(181).Asc127 = "micro"

EntityInfo(182).Code = 182
EntityInfo(182).Char = Chr$(182)
EntityInfo(182).Name = "&para;"
EntityInfo(182).Asc127 = ""

EntityInfo(183).Code = 183
EntityInfo(183).Char = Chr$(183)
EntityInfo(183).Name = "&middot;"
EntityInfo(183).Asc127 = "."

EntityInfo(184).Code = 184
EntityInfo(184).Char = Chr$(184)
EntityInfo(184).Name = "&cedil;"
EntityInfo(184).Asc127 = " "

EntityInfo(185).Code = 185
EntityInfo(185).Char = Chr$(185)
EntityInfo(185).Name = "&sup1;"
EntityInfo(185).Asc127 = "^1"

EntityInfo(186).Code = 186
EntityInfo(186).Char = Chr$(186)
EntityInfo(186).Name = "&ordm;"
EntityInfo(186).Asc127 = "o"

EntityInfo(187).Code = 187
EntityInfo(187).Char = Chr$(187)
EntityInfo(187).Name = "&raquo;"
EntityInfo(187).Asc127 = """"

EntityInfo(188).Code = 188
EntityInfo(188).Char = Chr$(188)
EntityInfo(188).Name = "&frac14;"
EntityInfo(188).Asc127 = "1/4"

EntityInfo(189).Code = 189
EntityInfo(189).Char = Chr$(189)
EntityInfo(189).Name = "&frac12;"
EntityInfo(189).Asc127 = "1/2"

EntityInfo(190).Code = 190
EntityInfo(190).Char = Chr$(190)
EntityInfo(190).Name = "&frac34;"
EntityInfo(190).Asc127 = "3/4"

EntityInfo(191).Code = 191
EntityInfo(191).Char = Chr$(191)
EntityInfo(191).Name = "&iquest;"
EntityInfo(191).Asc127 = "?"

End Sub
Private Sub InitEntityInfo2()
' Codes 192 to 255 (Chars with Accents)

EntityInfo(192).Code = 192
EntityInfo(192).Char = Chr$(192)
EntityInfo(192).Name = "&Agrave;"
EntityInfo(192).Asc127 = "A"

EntityInfo(193).Code = 193
EntityInfo(193).Char = Chr$(193)
EntityInfo(193).Name = "&Aacute;"
EntityInfo(193).Asc127 = "A"

EntityInfo(194).Code = 194
EntityInfo(194).Char = Chr$(194)
EntityInfo(194).Name = "&Acirc;"
EntityInfo(194).Asc127 = "A"

EntityInfo(195).Code = 195
EntityInfo(195).Char = Chr$(195)
EntityInfo(195).Name = "&Atilde;"
EntityInfo(195).Asc127 = "A"

EntityInfo(196).Code = 196
EntityInfo(196).Char = Chr$(196)
EntityInfo(196).Name = "&Auml;"
EntityInfo(196).Asc127 = "A"

EntityInfo(197).Code = 197
EntityInfo(197).Char = Chr$(197)
EntityInfo(197).Name = "&Aring;"
EntityInfo(197).Asc127 = "A"

EntityInfo(198).Code = 198
EntityInfo(198).Char = Chr$(198)
EntityInfo(198).Name = "&AElig;"
EntityInfo(198).Asc127 = "AE"

EntityInfo(199).Code = 199
EntityInfo(199).Char = Chr$(199)
EntityInfo(199).Name = "&Ccedil;"
EntityInfo(199).Asc127 = "C"

EntityInfo(200).Code = 200
EntityInfo(200).Char = Chr$(200)
EntityInfo(200).Name = "&Egrave;"
EntityInfo(200).Asc127 = "E"

EntityInfo(201).Code = 201
EntityInfo(201).Char = Chr$(201)
EntityInfo(201).Name = "&Eacute;"
EntityInfo(201).Asc127 = "E"

EntityInfo(202).Code = 202
EntityInfo(202).Char = Chr$(202)
EntityInfo(202).Name = "&Ecirc;"
EntityInfo(202).Asc127 = "E"

EntityInfo(203).Code = 203
EntityInfo(203).Char = Chr$(203)
EntityInfo(203).Name = "&Euml;"
EntityInfo(203).Asc127 = "E"

EntityInfo(204).Code = 204
EntityInfo(204).Char = Chr$(204)
EntityInfo(204).Name = "&Igrave;"
EntityInfo(204).Asc127 = "I"

EntityInfo(205).Code = 205
EntityInfo(205).Char = Chr$(205)
EntityInfo(205).Name = "&Iacute;"
EntityInfo(205).Asc127 = "I"

EntityInfo(206).Code = 206
EntityInfo(206).Char = Chr$(206)
EntityInfo(206).Name = "&Icirc;"
EntityInfo(206).Asc127 = "I"

EntityInfo(207).Code = 207
EntityInfo(207).Char = Chr$(207)
EntityInfo(207).Name = "&Iuml;"
EntityInfo(207).Asc127 = "I"

EntityInfo(208).Code = 208
EntityInfo(208).Char = Chr$(208)
EntityInfo(208).Name = "&ETH;"
EntityInfo(208).Asc127 = ""

EntityInfo(209).Code = 209
EntityInfo(209).Char = Chr$(209)
EntityInfo(209).Name = "&Ntilde;"
EntityInfo(209).Asc127 = "N"

EntityInfo(210).Code = 210
EntityInfo(210).Char = Chr$(210)
EntityInfo(210).Name = "&Ograve;"
EntityInfo(210).Asc127 = "O"

EntityInfo(211).Code = 211
EntityInfo(211).Char = Chr$(211)
EntityInfo(211).Name = "&Oacute;"
EntityInfo(211).Asc127 = "O"

EntityInfo(212).Code = 212
EntityInfo(212).Char = Chr$(212)
EntityInfo(212).Name = "&Ocirc;"
EntityInfo(212).Asc127 = "O"

EntityInfo(213).Code = 213
EntityInfo(213).Char = Chr$(213)
EntityInfo(213).Name = "&Otilde;"
EntityInfo(213).Asc127 = "O"

EntityInfo(214).Code = 214
EntityInfo(214).Char = Chr$(214)
EntityInfo(214).Name = "&Ouml;"
EntityInfo(214).Asc127 = "O"

EntityInfo(215).Code = 215
EntityInfo(215).Char = Chr$(215)
EntityInfo(215).Name = "&times;"
EntityInfo(215).Asc127 = "*"

EntityInfo(216).Code = 216
EntityInfo(216).Char = Chr$(216)
EntityInfo(216).Name = "&Oslash;"
EntityInfo(216).Asc127 = "O"

EntityInfo(217).Code = 217
EntityInfo(217).Char = Chr$(217)
EntityInfo(217).Name = "&Ugrave;"
EntityInfo(217).Asc127 = "U"

EntityInfo(218).Code = 218
EntityInfo(218).Char = Chr$(218)
EntityInfo(218).Name = "&Uacute;"
EntityInfo(218).Asc127 = "U"

EntityInfo(219).Code = 219
EntityInfo(219).Char = Chr$(219)
EntityInfo(219).Name = "&Ucirc;"
EntityInfo(219).Asc127 = "U"

EntityInfo(220).Code = 220
EntityInfo(220).Char = Chr$(220)
EntityInfo(220).Name = "&Uuml;"
EntityInfo(220).Asc127 = "U"

EntityInfo(221).Code = 221
EntityInfo(221).Char = Chr$(221)
EntityInfo(221).Name = "&Yacute;"
EntityInfo(221).Asc127 = "Y"

EntityInfo(222).Code = 222
EntityInfo(222).Char = Chr$(222)
EntityInfo(222).Name = "&THORN;"
EntityInfo(222).Asc127 = ""

EntityInfo(223).Code = 223
EntityInfo(223).Char = Chr$(223)
EntityInfo(223).Name = "&szlig;"
EntityInfo(223).Asc127 = ""

EntityInfo(224).Code = 224
EntityInfo(224).Char = Chr$(224)
EntityInfo(224).Name = "&agrave;"
EntityInfo(224).Asc127 = "a"

EntityInfo(225).Code = 225
EntityInfo(225).Char = Chr$(225)
EntityInfo(225).Name = "&aacute;"
EntityInfo(225).Asc127 = "a"

EntityInfo(226).Code = 226
EntityInfo(226).Char = Chr$(226)
EntityInfo(226).Name = "&acirc;"
EntityInfo(226).Asc127 = "a"

EntityInfo(227).Code = 227
EntityInfo(227).Char = Chr$(227)
EntityInfo(227).Name = "&atilde;"
EntityInfo(227).Asc127 = "a"

EntityInfo(228).Code = 228
EntityInfo(228).Char = Chr$(228)
EntityInfo(228).Name = "&auml;"
EntityInfo(228).Asc127 = "a"

EntityInfo(229).Code = 229
EntityInfo(229).Char = Chr$(229)
EntityInfo(229).Name = "&aring;"
EntityInfo(229).Asc127 = "a"

EntityInfo(230).Code = 230
EntityInfo(230).Char = Chr$(230)
EntityInfo(230).Name = "&aelig;"
EntityInfo(230).Asc127 = "a"

EntityInfo(231).Code = 231
EntityInfo(231).Char = Chr$(231)
EntityInfo(231).Name = "&ccedil;"
EntityInfo(231).Asc127 = "c"

EntityInfo(232).Code = 232
EntityInfo(232).Char = Chr$(232)
EntityInfo(232).Name = "&egrave;"
EntityInfo(232).Asc127 = "e"

EntityInfo(233).Code = 233
EntityInfo(233).Char = Chr$(233)
EntityInfo(233).Name = "&eacute;"
EntityInfo(233).Asc127 = "e"

EntityInfo(234).Code = 234
EntityInfo(234).Char = Chr$(234)
EntityInfo(234).Name = "&ecirc;"
EntityInfo(234).Asc127 = "e"

EntityInfo(235).Code = 235
EntityInfo(235).Char = Chr$(235)
EntityInfo(235).Name = "&euml;"
EntityInfo(235).Asc127 = "e"

EntityInfo(236).Code = 236
EntityInfo(236).Char = Chr$(236)
EntityInfo(236).Name = "&igrave;"
EntityInfo(236).Asc127 = "i"

EntityInfo(237).Code = 237
EntityInfo(237).Char = Chr$(237)
EntityInfo(237).Name = "&iacute;"
EntityInfo(237).Asc127 = "i"

EntityInfo(238).Code = 238
EntityInfo(238).Char = Chr$(238)
EntityInfo(238).Name = "&icirc;"
EntityInfo(238).Asc127 = "i"

EntityInfo(239).Code = 239
EntityInfo(239).Char = Chr$(239)
EntityInfo(239).Name = "&iuml;"
EntityInfo(239).Asc127 = "i"

EntityInfo(240).Code = 240
EntityInfo(240).Char = Chr$(240)
EntityInfo(240).Name = "&eth;"
EntityInfo(240).Asc127 = ""

EntityInfo(241).Code = 241
EntityInfo(241).Char = Chr$(241)
EntityInfo(241).Name = "&ntilde;"
EntityInfo(241).Asc127 = "n"

EntityInfo(242).Code = 242
EntityInfo(242).Char = Chr$(242)
EntityInfo(242).Name = "&ograve;"
EntityInfo(242).Asc127 = "o"

EntityInfo(243).Code = 243
EntityInfo(243).Char = Chr$(243)
EntityInfo(243).Name = "&oacute;"
EntityInfo(243).Asc127 = "o"

EntityInfo(244).Code = 244
EntityInfo(244).Char = Chr$(244)
EntityInfo(244).Name = "&ocirc;"
EntityInfo(244).Asc127 = "o"

EntityInfo(245).Code = 245
EntityInfo(245).Char = Chr$(245)
EntityInfo(245).Name = "&otilde;"
EntityInfo(245).Asc127 = "o"

EntityInfo(246).Code = 246
EntityInfo(246).Char = Chr$(246)
EntityInfo(246).Name = "&ouml;"
EntityInfo(246).Asc127 = "o"

EntityInfo(247).Code = 247
EntityInfo(247).Char = Chr$(247)
EntityInfo(247).Name = "&divide;"
EntityInfo(247).Asc127 = "/"

EntityInfo(248).Code = 248
EntityInfo(248).Char = Chr$(248)
EntityInfo(248).Name = "&oslash;"
EntityInfo(248).Asc127 = "o"

EntityInfo(249).Code = 249
EntityInfo(249).Char = Chr$(249)
EntityInfo(249).Name = "&ugrave;"
EntityInfo(249).Asc127 = "u"

EntityInfo(250).Code = 250
EntityInfo(250).Char = Chr$(250)
EntityInfo(250).Name = "&uacute;"
EntityInfo(250).Asc127 = "u"

EntityInfo(251).Code = 251
EntityInfo(251).Char = Chr$(251)
EntityInfo(251).Name = "&ucirc;"
EntityInfo(251).Asc127 = "u"

EntityInfo(252).Code = 252
EntityInfo(252).Char = Chr$(252)
EntityInfo(252).Name = "&uuml;"
EntityInfo(252).Asc127 = "u"

EntityInfo(253).Code = 253
EntityInfo(253).Char = Chr$(253)
EntityInfo(253).Name = "&yacute;"
EntityInfo(253).Asc127 = "y"

EntityInfo(254).Code = 254
EntityInfo(254).Char = Chr$(254)
EntityInfo(254).Name = "&thorn;"
EntityInfo(254).Asc127 = ""

EntityInfo(255).Code = 255
EntityInfo(255).Char = Chr$(255)
EntityInfo(255).Name = "&yuml;"
EntityInfo(255).Asc127 = "y"

End Sub
