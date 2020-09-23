'affects only this file //DON'T CHANGE
Option UseEscapes

'Auto-Executed Code:
'On Error Resume Next
'	dim GUI , FSO
'	Set GUI = CreateObject("PAxForms.CAxForms")
'	Set FSO = CreateObject("Scripting.FileSystemObject")
'If Err Then
'	MsgBox "Cannot create object in 'Autoload.vbs'"
'	Err.Clear
'End If

'=====================================================

Function AxiomAutoLoadMain()
	AxiomAutoLoadMain=""   '//DON'T MOVE, REMOVE OR RENAME THIS FUNCTION!
End function
'=====================================================

Const vbKeyA=65
Const vbKeyB=66
Const vbKeyC=67
Const vbKeyD=68
Const vbKeyE=69
Const vbKeyF=70
Const vbKeyG=71
Const vbKeyH=72
Const vbKeyI=73
Const vbKeyJ=74
Const vbKeyK=75
Const vbKeyL=76
Const vbKeyM=77
Const vbKeyN=78
Const vbKeyO=79
Const vbKeyP=80
Const vbKeyQ=81
Const vbKeyR=82
Const vbKeyS=83
Const vbKeyT=84
Const vbKeyU=85
Const vbKeyV=86
Const vbKeyW=87
Const vbKeyX=88
Const vbKeyY=89
Const vbKeyZ=90

'--------------------------------------------------------------------
Function Swap (Var1,Var2)  'as VOID
        dim TempVar
        TempVar=Var1
        Var1=Var2
        Var2=TempVar
End Function

'--------------------------------------------------------------------
Function AddBackSlash(Str)
	Const BACK_SLASH = 092  ' \
	AddBackSlash = Str & "\\"  'Chr(BACK_SLASH)
End Function

'--------------------------------------------------------------------
Function AddCrLf(Str)
	AddCrLf = Str & vbCrLf
End Function

'--------------------------------------------------------------------
Function EnQuote(Str)
	EnQuote = "\q" & Str & "\q"
End Function

'--------------------------------------------------------------------
Function StrDel(sString, nIndex, nCount)
Dim sLeft, sRight, nLen
        nLen = Len(sString)
        If nIndex >= 0 And nIndex <= nLen Then
            If nIndex > 1 And nLen > 0 Then
                sLeft = Left(sString, nIndex - 1)
            Else
                sLeft = ""
            End If
            If (nIndex + nCount) <= nLen Then
                sRight = Mid(sString, nIndex + nCount)
            Else
                sRight = ""
            End If
            StrDel = sLeft & sRight
        End If
End Function
'-------------------------------------------------------------------
Function StrIns(sMainStr, sNewStr, nIndex)
Dim szLeft, szRight
    If nIndex > 1 And Len(sMainStr) > 0 Then
        szLeft = Left(sMainStr, nIndex - 1)
    Else
        szLeft = ""
    End If
    szRight = Right(sMainStr, Len(sMainStr) - nIndex + 1)
    StrIns = szLeft & sNewStr & szRight
End Function
'-------------------------------------------------------------------
Public Function NeatSplit(ByVal Expression , ByVal Delimiter )

Dim varItems , i

varItems = Split(Expression, Delimiter, -1, vbTextCompare)

For i = LBound(varItems) To UBound(varItems)

    If Len(varItems(i)) = 0 Then varItems(i) = Delimiter

Next

NeatSplit = Filter(varItems, Delimiter, False)
    
End Function




'**************** RegExp Functions **********************

Function RegExpExtractToArray (byval Text ,byval Pattern )
   Dim V
   V = RegExpExtract0 ( Text , Pattern )	' RegExpExtract0 is an undocumented function!
   V = Split (V,Chr(0))
   RegExpExtractToArray = V
End Function

'--------------------------------------------------------------------------
Function RegExpReplaceFunc(ByVal Text, ByVal Pattern , ByVal ReplacerFunc)
   dim oRegExp
   Set oRegExp = New RegExp
   oRegExp.Pattern = Pattern
   oRegExp.Global = True
   oRegExp.IgnoreCase = True
   RegExpReplaceFunc = oRegExp.Replace(Text ,GetRef(ReplacerFunc))
End Function


'*************** Text Files Wrapper Functions ***************
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Function OpenTextFile ( ByVal FileName , ByVal IOMode )
   Dim fso, tfile , Create
   Set fso = CreateObject("Scripting.FileSystemObject")
   If IOMode = ForReading Then Create = False Else Create = True
   Set OpenTextFile = fso.OpenTextFile(FileName, IOMode, True)
End Function


'***************** String Concat Class ********************
CLASS CString

        Private ms_BigStr
        Private ml_Pos
        Private ml_MaxLength
        
        Public Default Property Get Value()
               Value = Left(ms_BigStr, ml_Pos)
        End Property

        Public Property Let MaxLength(ByVal lNewValue)
            ms_BigStr = Space(lNewValue)
            ml_Pos = 0
        End Property
        
        Public Function Add(NewStr)
                'this function is fixed in the VB code (why not use that one ?)
                ms_BigStr=MidStr(ms_BigStr,NewStr, ml_Pos + 1,-1) 
                ml_Pos = ml_Pos + Len(NewStr)
        End Function

        Public Sub Clear()
                ms_BigStr = ""
                ml_MaxLength = 0
                ml_Pos = 0
        End Sub

        Public Property Get CharAt(Position)
            If Position > 0 Then
                CharAt = Mid(StrVal, Position, 1)
            End If
        End Property
        
        Public Property Let CharAt(Position , sNewValue)
            If Position > 0 Then
                ms_BigStr=MidStr(ms_BigStr,sNewValue, Position, 1)
            End If
        End Property

END CLASS

'=============================================
'FUNCTIONS IMPORTED FROM  "VBAFunctions.dll"
'///////////// Now moved to CAxiomFunction (built-in, no longer in a DLL)
'Function Format(String,Formatting)
'       Set objDLL = CreateObject("PVBAFunctions.CVBAFunctions")
'        Format = objDLL.FormatString(String,Formatting) 
'        Set objDLL = Nothing
'End Function
'--------------------------------------------------------------------
'Function MidStr(sString , sNewValue , lStart , lLength )
'        Set objDLL = CreateObject("PVBAFunctions.CVBAFunctions")
'        MidStr = objDLL.MidStr(sString , sNewValue , lStart , lLength )
'        Set objDLL = Nothing
'End Function
