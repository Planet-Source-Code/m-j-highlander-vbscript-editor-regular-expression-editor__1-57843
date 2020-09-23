
'Remove UTF-8 BOM (Byte-Order-Marker, some sort of a header!)

Public Function Main ( ByVal Text )

'UTF_8_BOM = "﻿"
UTF_8_BOM = Chr(239) & Chr(187) & Chr(191)

Main = Replace( Text , UTF_8_BOM , "" )

End Function

