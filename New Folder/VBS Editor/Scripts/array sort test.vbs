Option UseEscapes

Public Function Main ( ByVal Text )

'Array of Variants:
'dim a(5) 
'a(0)="The"
'a(1)="And"
'a(2)="zoo"
'a(3)="max"
'a(4)="Help"
'a(5)="Inox"


'Variant containing an Array:
a = Array ("A","x","G","B","5","a")

'cause Type Mismatch Error:
'a = "XXX"

B = Sort(a,true,true)

Main = Join (B, "\n")

End Function

