
Function Main(Text)

Main=RemovePath(Text)

'Comment out the follwing line if you don't want to remove file extensions

Main = RemoveExt(Main)

End Function



Function RemovePath(Text)

If Text="" Then Exit Function

v=split(Text,vbCrLf)

for idx=0 to ubound(v)
        pos=InstrRev(v(idx),"\")
        if pos>0 then v(idx)=Right(v(idx),len(v(idx))-pos)
next

RemovePath=Join(v,vbCrLf)

End Function

Function RemoveExt (Text)

If Text="" Then Exit Function

v=split(Text,vbCrLf)

for idx=0 to ubound(v)
        pos=InstrRev(v(idx),".")
        if pos>0 then v(idx)=Left(v(idx),pos-1)
next

RemoveExt = Join(v,vbCrLf)

End Function

