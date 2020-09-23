Function Add_Numbers_To_Right (Text)

Dim V,StartFrom,StepVal,Delimiter,IgnoreEmptyLines
Dim SF,SV,idx,cntr
        
        If Text="" then Exit Function
        
        

        
        V = gui.InputForm("Start From:|1","Step:|1","Delimiter:","Ignore Empty Lines (1=Yes / 0=No)|1")
        
        If Not IsArray(v) Then Exit Function
        
        StartFrom=V(0)
        StepVal=V(1)
        Delimiter =V(2)
        IgnoreEmptyLines = CBool(V(3))
        
        SF = CInt(StartFrom)
        SV=CInt(StepVal)
        
        V=Split(Text,vbCrLf)
        
        cntr=SF
        For idx=lbound(V) To ubound(V)
          If V(idx)="" and IgnoreEmptyLines = True Then
                  ' do nothing
          Else
                  V(idx)=V(idx) & Delimiter & CStr(cntr)
                  cntr = cntr + SV
          End If
        Next
        
        Add_Numbers_To_Right=Join(V,vbCrLf)

End Function 
