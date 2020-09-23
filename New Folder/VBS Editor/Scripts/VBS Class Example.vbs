Function help

        dim c
        set c = new ctest
        
        help=c.x & c.coco
        set c=nothing

End Function


'=============
CLASS CTest
''''''''''''''''''''''''''''''
public function x()
        x="XXX"
end function
''''''''''''''''''''''''''''''
public property get coco()
        coco="COCO"
end property
''''''''''''''''''''''''''''''
End Class
'=============
