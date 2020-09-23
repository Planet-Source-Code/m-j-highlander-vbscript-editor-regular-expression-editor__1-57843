' CString Class is defined in "Autoload.vbs"
Option UseEscapes

Function Main

        Dim c
        Set c = New CString
        
        c.MaxLength=100
        c.Add "coco\n"
        c.Add "coco\t"
        c.Add "coco"
        
        c.CharAt(1)="X"
        
        Main = c.Value
        
        Set c=nothing
        
End Function
