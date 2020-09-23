
Function OpenTextFileTest
Dim f
  Set f = OpenTextFile("c:\testfile.txt", ForAppending)
  f.Write "Hello world!"
  f.Close
End Function
