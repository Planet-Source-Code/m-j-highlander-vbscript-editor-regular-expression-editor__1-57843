Attribute VB_Name = "TypeLibExplorer"
Option Explicit
Option Base 0

'// DON'T MOVE or REMOVE
Private tliTypeLibInfo As TypeLibInfo

Public Function ProduceDefaultValue(DefVal As Variant, ByVal tliTypeInfo As TypeInfo) As String
'This helper function adapted from Microsoft documentation
Dim lngTrackVal As Long
Dim MI As MemberInfo
Dim tliTypeKinds As TypeKinds
    
If tliTypeInfo Is Nothing Then
    Select Case VarType(DefVal)
        Case vbString
            If Len(DefVal) Then
                ProduceDefaultValue = """" & DefVal & """"
            End If
        Case vbBoolean 'Always show for Boolean
            ProduceDefaultValue = DefVal
        Case vbDate
            If DefVal Then
                ProduceDefaultValue = "#" & DefVal & "#"
            End If
        Case Else 'Numeric Values
            If DefVal <> 0 Then
                ProduceDefaultValue = DefVal
            End If
    End Select
Else
    'Resolve constants to their enums
    tliTypeKinds = tliTypeInfo.TypeKind
    Do While tliTypeKinds = TKIND_ALIAS
        tliTypeKinds = TKIND_MAX
        On Error Resume Next
        Set tliTypeInfo = tliTypeInfo.ResolvedType
        If Err = 0 Then
            tliTypeKinds = tliTypeInfo.TypeKind
        End If
        On Error GoTo 0
    Loop
    If tliTypeInfo.TypeKind = TKIND_ENUM Then
        lngTrackVal = DefVal
        For Each MI In tliTypeInfo.Members
            If MI.Value = lngTrackVal Then
                ProduceDefaultValue = " = " & MI.Name
                Exit For
            End If
        Next
    End If
End If
End Function
Public Function GetSearchType(ByVal SearchData As Long) As TliSearchTypes
    'This helper function adapted from Microsoft documentation
    If SearchData And &H80000000 Then
        GetSearchType = ((SearchData And &H7FFFFFFF) \ &H1000000 And &H7F&) Or &H80
    Else
        GetSearchType = SearchData \ &H1000000 And &HFF&
    End If
End Function
Public Function PrototypeMember(ByVal SearchData As Long, _
    ByVal InvokeKinds As InvokeKinds, _
    Optional ByVal MemberName As String, Optional ByRef HelpString As String) As String
    'This helper function adapted from Microsoft documentation
    
    Dim tliParameterInfo As ParameterInfo
    Dim bFirstParameter As Boolean
    Dim bIsConstant As Boolean
    Dim bByVal As Boolean
    Dim strReturn As String
    Dim ConstVal As Variant
    Dim strTypeName As String
    Dim intVarTypeCur As Integer
    Dim bDefault As Boolean
    Dim bOptional As Boolean
    Dim bParamArray As Boolean
    Dim tliTypeInfo As TypeInfo
    Dim tliResolvedTypeInfo As TypeInfo
    Dim tliTypeKinds As TypeKinds
  
    With tliTypeLibInfo
        'First, determine the type of member we're dealing with
        bIsConstant = GetSearchType(SearchData) And tliStConstants
        With .GetMemberInfo(SearchData, InvokeKinds, , MemberName)
            If bIsConstant Then
                strReturn = "Const "
            ElseIf InvokeKinds = INVOKE_FUNC Or InvokeKinds = INVOKE_EVENTFUNC Then
                Select Case .ReturnType.VarType
                    Case VT_VOID, VT_HRESULT
                        strReturn = "Sub "
                    Case Else
                        strReturn = "Function "
                End Select
            Else
                strReturn = "Property "
            End If
        
            'Now add the name of the member
            strReturn = strReturn & .Name
        
            'Process the member's paramters
            With .Parameters
                If .Count Then
                    strReturn = strReturn & " ("
                    bFirstParameter = True
                    bParamArray = .OptionalCount = -1
                    For Each tliParameterInfo In .Me
                        If Not bFirstParameter Then
                            strReturn = strReturn & ", "
                        End If
                        bFirstParameter = False
                        bDefault = tliParameterInfo.Default
                        bOptional = bDefault Or tliParameterInfo.Optional
                        If bOptional Then
                            If bParamArray Then
                                'This will be the only optional parameter
                                strReturn = strReturn & "[ParamArray "
                            Else
                                strReturn = strReturn & "["
                            End If
                        End If
                    
                        With tliParameterInfo.VarTypeInfo
                            Set tliTypeInfo = Nothing
                            Set tliResolvedTypeInfo = Nothing
                            tliTypeKinds = TKIND_MAX
                            intVarTypeCur = .VarType
                            If (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                                On Error Resume Next
                                Set tliTypeInfo = .TypeInfo
                                If Not tliTypeInfo Is Nothing Then
                                    Set tliResolvedTypeInfo = tliTypeInfo
                                    tliTypeKinds = tliResolvedTypeInfo.TypeKind
                                    Do While tliTypeKinds = TKIND_ALIAS
                                        tliTypeKinds = TKIND_MAX
                                        Set tliResolvedTypeInfo = tliResolvedTypeInfo.ResolvedType
                                        If Err Then
                                            Err.Clear
                                        Else
                                            tliTypeKinds = tliResolvedTypeInfo.TypeKind
                                        End If
                                    Loop
                                End If
                            
                                'Determine whether parameters are ByVal or ByRef
                                Select Case tliTypeKinds
                                    Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
                                        bByVal = .PointerLevel = 1
                                    Case TKIND_RECORD
                                        'Records not passed ByVal in VB
                                        bByVal = False
                                    Case Else
                                        bByVal = .PointerLevel = 0
                                End Select
                            
                                'Indicate ByVal
                                If bByVal Then
                                    strReturn = strReturn & "ByVal "
                                End If
                            
                                'Display the parameter name
                                strReturn = strReturn & tliParameterInfo.Name
                            
                                If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                    strReturn = strReturn & "()"
                                End If
                                
                                If tliTypeInfo Is Nothing Then 'Information not available
                                    strReturn = strReturn & " As ?"
                                Else
                                    If .IsExternalType Then
                                        strReturn = strReturn & " As " & .TypeLibInfoExternal.Name & "." & tliTypeInfo.Name
                                    Else
                                        strReturn = strReturn & " As " & tliTypeInfo.Name
                                    End If
                                End If
                            
                                'Reset error handling
                                On Error GoTo 0
                            Else
                                If .PointerLevel = 0 Then
                                    strReturn = strReturn & "ByVal "
                                End If
                                    
                                strReturn = strReturn & tliParameterInfo.Name
                                If intVarTypeCur <> vbVariant Then
                                    strTypeName = TypeName(.TypedVariant)
                                    If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                        strReturn = strReturn & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                    Else
                                        strReturn = strReturn & " As " & strTypeName
                                    End If
                                End If
                            End If
                                
                            If bOptional Then
                                If bDefault Then
                                    strReturn = strReturn & ProduceDefaultValue(tliParameterInfo.DefaultValue, tliResolvedTypeInfo)
                                    'strReturn = strReturn & " = " & tliParameterInfo.DefaultValue
                                End If
                                strReturn = strReturn & "]"
                            End If
                        End With
                    Next
                    strReturn = strReturn & ")"
                End If
            End With
        
            If bIsConstant Then
                ConstVal = .Value
                strReturn = strReturn & " = " & ConstVal
                Select Case VarType(ConstVal)
                    Case vbInteger, vbLong
                        If ConstVal < 0 Or ConstVal > 15 Then
                            strReturn = strReturn & " (&H" & Hex$(ConstVal) & ")"
                        End If
                End Select
            Else
                With .ReturnType
                    intVarTypeCur = .VarType
                    If intVarTypeCur = 0 Or (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                        On Error Resume Next
                        If Not .TypeInfo Is Nothing Then
                            If Err Then 'Information not available
                                strReturn = strReturn & " As ?"
                            Else
                                If .IsExternalType Then
                                    strReturn = strReturn & " As " & .TypeLibInfoExternal.Name & "." & .TypeInfo.Name
                                Else
                                    strReturn = strReturn & " As " & .TypeInfo.Name
                                End If
                            End If
                        End If
                        
                        If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                            strReturn = strReturn & "()"
                        End If
                        On Error GoTo 0
                    Else
                        Select Case intVarTypeCur
                            Case VT_VARIANT, VT_VOID, VT_HRESULT
                            Case Else
                                strTypeName = TypeName(.TypedVariant)
                                If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                    strReturn = strReturn & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                Else
                                    strReturn = strReturn & " As " & strTypeName
                                End If
                        End Select
                    End If
                End With
            End If
            
            PrototypeMember = strReturn & vbCrLf
            'MsgBox "Member of " & tliTypeLibInfo.Name & "." & tliTypeLibInfo.GetTypeInfo(SearchData And &HFFFF&).Name
            HelpString = .HelpString
        End With
    End With
End Function

Public Function ProcessTypeLibrary(ByVal LibraryPath As String, ClassName As String, FunctionName() As String, FunctionHeader() As String, FunctionDescription() As String) As Boolean
On Error GoTo Err_TypeLibrary

Dim i As Long, ClassSearchID As Long
Dim tliInvokeKinds As InvokeKinds
Dim tliTypeInfo As TypeInfo
Dim s As SearchResults

Set tliTypeLibInfo = New TypeLibInfo
tliTypeLibInfo.AppObjString = "<Global>"
tliTypeLibInfo.ContainingFile = LibraryPath
'MsgBox tliTypeLibInfo.Name               'project name
'MsgBox tliTypeLibInfo.ContainingFile     'file path
'MsgBox tliTypeLibInfo.HelpString         'description
'MsgBox tliTypeLibInfo.MajorVersion & "." & tliTypeLibInfo.MinorVersion 'version
'MsgBox tliTypeLibInfo.HelpFile           'Help File
'MsgBox tliTypeLibInfo.Guid                'GUID

Set s = tliTypeLibInfo.GetTypes(s, tliStAll, True)
'second class,since first is <GLOBAL>
ClassName = tliTypeLibInfo.Name & "." & s.Item(2)
ClassSearchID = s.Item(2).SearchData
Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(1) ' 0-based !!!
Set s = tliTypeLibInfo.GetMembers(s.Item(2).SearchData)
ReDim FunctionName(0 To s.Count - 1)
ReDim FunctionHeader(0 To s.Count - 1)
ReDim FunctionDescription(0 To s.Count - 1)
For i = 1 To s.Count
    tliInvokeKinds = s.Item(i).MemberId
    FunctionName(i - 1) = s.Item(i)
    FunctionHeader(i - 1) = PrototypeMember(ClassSearchID, tliInvokeKinds, s.Item(i), FunctionDescription(i - 1))
    If Right$(FunctionHeader(i - 1), 2) = vbCrLf Then FunctionHeader(i - 1) = Left$(FunctionHeader(i - 1), Len(FunctionHeader(i - 1)) - 2)
Next i

Set tliTypeLibInfo = Nothing

ProcessTypeLibrary = True

Exit Function
Err_TypeLibrary:
If Err Then
    ProcessTypeLibrary = False
    Exit Function
End If

End Function
