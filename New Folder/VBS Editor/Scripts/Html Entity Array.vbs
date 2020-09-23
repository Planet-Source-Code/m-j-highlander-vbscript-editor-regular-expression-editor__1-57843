Function Main()

For i = 1 To 255
    A = "EntityInfo(" & i & ").Code =" & i & vbCrLf
    B = "EntityInfo(" & i & ").Char = Chr(" & i & ")" & vbCrLf
    C = "EntityInfo(" & i & ").Name = ""&;""" & vbCrLf
    D = "EntityInfo(" & i & ").Asc127 = """"" & vbCrLf & vbCrLf
    All = All & A & B & C & D
Next

Main = All

End function
