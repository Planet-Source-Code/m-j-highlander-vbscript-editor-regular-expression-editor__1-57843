Option UseEscapes

Public Function Main ( ByVal Text )

for ipage=74 to 11 step -1
        for ipic = 1 to 16
                s = s & "<a href=\qhttp://www.hot-galleries.com/3000/" & _
                CStr(ipage) & "/" & CStr(ipic) & ".jpg\q>" & CStr(ipic) & "</a><BR>\n"
        next
        
        s = "<html><head><title>" & cstr(ipage) & "</title></head><body>\n" & s & "\n</body></html>"
        SaveFile s,"G:\\2\\vg\\" & cstr(ipage) & ".htm"
        s = ""
next

Main = ""

End Function

