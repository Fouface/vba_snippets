Attribute VB_Name = "zZ_list_object_properties"
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type TTYPEDESC
    #If Win64 Then
        pTypeDesc As LongLong
    #Else
        pTypeDesc As Long
    #End If
    vt            As Integer
End Type

Private Type TPARAMDESC
    #If Win64 Then
        pPARAMDESCEX  As LongLong
    #Else
        pPARAMDESCEX  As Long
    #End If
    wParamFlags       As Integer
End Type

Private Type TELEMDESC
    tdesc  As TTYPEDESC
    pdesc  As TPARAMDESC
End Type

Type TYPEATTR
        aGUID As GUID
        LCID As Long
        dwReserved As Long
        memidConstructor As Long
        memidDestructor As Long
        #If Win64 Then
            lpstrSchema As LongLong
        #Else
            lpstrSchema As Long
        #End If
        cbSizeInstance As Integer
        typekind As Long
        cFuncs As Integer
        cVars As Integer
        cImplTypes As Integer
        cbSizeVft As Integer
        cbAlignment As Integer
        wTypeFlags As Integer
        wMajorVerNum As Integer
        wMinorVerNum As Integer
        tdescAlias As Long
        idldescType As Long
End Type


Type FUNCDESC
    memid As Long
    #If Win64 Then
        lReserved1 As LongLong
        lprgelemdescParam As LongLong
    #Else
        lReserved1 As Long
        lprgelemdescParam As Long
    #End If
    funckind As Long
    INVOKEKIND As Long
    CallConv As Long
    cParams As Integer
    cParamsOpt As Integer
    oVft As Integer
    cReserved2 As Integer
    elemdescFunc As TELEMDESC
    wFuncFlags As Integer
End Type

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
#Else
'    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
 '   Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
  '  Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
#End If

Private Type PropertyInfo
    Name As String
    Value As Variant
End Type





















Sub RunTest()

    Dim PropertiesInfo() As PropertyInfo, i As Long
 
    Range("A1:C1").CurrentRegion.ClearContents
    Range("A1") = "Property Name"
    Range("B1") = "Property Current Value"
    Range("C1") = "Total Properties Found"
    Range("A1:C1").Font.Bold = True
 
   'PropertiesInfo = GetPropertiesInfoOfObject(Obj:=Application) '<== Change Object as required.
    PropertiesInfo = GetPropertiesInfoOfObject(Obj:=Cells) '<== Change Object as required.
 
 
    For i = LBound(PropertiesInfo) To UBound(PropertiesInfo)
        Range("A" & i + 2) = PropertiesInfo(i).Name
        Range("B" & i + 2) = PropertiesInfo(i).Value
    Next i
 
    Range("C2") = UBound(PropertiesInfo)

End Sub































Public Function GetPropertiesInfoOfObject(ByVal Obj As Object) As PropertyInfo()

    Dim oFuncCol As New Collection, i As Long, j As Long
    Dim aArray() As PropertyInfo, vRet As Variant, lErr As Long

    Set oFuncCol = GetObjectFunctions(TheObject:=Obj, FuncType:=VbGet)
    For i = 1 To oFuncCol.Count
        On Error Resume Next
            Set vRet = CallByName(Obj, Split(oFuncCol.item(i), vbTab)(0), VbGet)
            lErr = Err.Number
            Err.Clear
            If lErr Then
                vRet = CallByName(Obj, Split(oFuncCol.item(i), vbTab)(0), VbGet)
                If IsArray(vRet) = False And IsObject(vRet) = False Then
                    ReDim Preserve aArray(j)
                    aArray(j).Name = Split(oFuncCol.item(i), vbTab)(0)
                    If IsError(vRet) Then
                        vRet = "Error"
                    ElseIf IsNull(vRet) Then
                        vRet = "Null"
                    ElseIf Len(vRet) = 0 Then
                        vRet = "Property Not Set"
                    End If
                    aArray(j).Value = vRet
                    j = j + 1
                End If
            End If
        On Error GoTo 0
    Next
    GetPropertiesInfoOfObject = aArray

End Function




Private Function GetObjectFunctions(ByVal TheObject As Object, Optional ByVal FuncType As VbCallType) As Collection

    Dim tTYPEATTR As TYPEATTR
    Dim tFUNCDESC As FUNCDESC

    Dim aGUID(0 To 11) As Long, lFuncsCount As Long
 
    #If Win64 Then
        Const vTblOffsetFac_32_64 = 2
        Dim aTYPEATTR() As LongLong, aFUNCDESC() As LongLong, farPtr As LongLong
    #Else
        Const vTblOffsetFac_32_64 = 1
        Dim aTYPEATTR() As Long, aFUNCDESC() As Long, farPtr As Long
    #End If
 
    Dim ITypeInfo As IUnknown
    Dim IDispatch As IUnknown
    Dim sName As String, oCol As New Collection
 
    Const CC_STDCALL As Long = 4
    Const IUNK_QueryInterface As Long = 0
    Const IDSP_GetTypeInfo As Long = 16 * vTblOffsetFac_32_64
    Const ITYP_GetTypeAttr As Long = 12 * vTblOffsetFac_32_64
    Const ITYP_GetFuncDesc As Long = 20 * vTblOffsetFac_32_64
    Const ITYP_GetDocument As Long = 48 * vTblOffsetFac_32_64

    Const ITYP_ReleaseTypeAttr As Long = 76 * vTblOffsetFac_32_64
    Const ITYP_ReleaseFuncDesc As Long = 80 * vTblOffsetFac_32_64


    aGUID(0) = &H20400: aGUID(2) = &HC0&: aGUID(3) = &H46000000
    CallFunction_COM ObjPtr(TheObject), IUNK_QueryInterface, vbLong, CC_STDCALL, VarPtr(aGUID(0)), VarPtr(IDispatch)
    If IDispatch Is Nothing Then MsgBox "error":   Exit Function

    CallFunction_COM ObjPtr(IDispatch), IDSP_GetTypeInfo, vbLong, CC_STDCALL, 0&, 0&, VarPtr(ITypeInfo)
    If ITypeInfo Is Nothing Then MsgBox "error": Exit Function
 
    CallFunction_COM ObjPtr(ITypeInfo), ITYP_GetTypeAttr, vbLong, CC_STDCALL, VarPtr(farPtr)
    If farPtr = 0& Then MsgBox "error": Exit Function

    CopyMemory ByVal VarPtr(tTYPEATTR), ByVal farPtr, LenB(tTYPEATTR)
    ReDim aTYPEATTR(LenB(tTYPEATTR))
    CopyMemory ByVal VarPtr(aTYPEATTR(0)), tTYPEATTR, UBound(aTYPEATTR)
    CallFunction_COM ObjPtr(ITypeInfo), ITYP_ReleaseTypeAttr, vbEmpty, CC_STDCALL, farPtr
 
    For lFuncsCount = 0 To tTYPEATTR.cFuncs - 1
        Call CallFunction_COM(ObjPtr(ITypeInfo), ITYP_GetFuncDesc, vbLong, CC_STDCALL, lFuncsCount, VarPtr(farPtr))
        If farPtr = 0 Then MsgBox "error": Exit For
        CopyMemory ByVal VarPtr(tFUNCDESC), ByVal farPtr, LenB(tFUNCDESC)
        ReDim aFUNCDESC(LenB(tFUNCDESC))
        CopyMemory ByVal VarPtr(aFUNCDESC(0)), tFUNCDESC, UBound(aFUNCDESC)
        Call CallFunction_COM(ObjPtr(ITypeInfo), ITYP_ReleaseFuncDesc, vbEmpty, CC_STDCALL, farPtr)
         Call CallFunction_COM(ObjPtr(ITypeInfo), ITYP_GetDocument, vbLong, CC_STDCALL, aFUNCDESC(0), VarPtr(sName), 0, 0, 0)
        Call CallFunction_COM(ObjPtr(ITypeInfo), ITYP_GetDocument, vbLong, CC_STDCALL, aFUNCDESC(0), VarPtr(sName), 0, 0, 0)

        With tFUNCDESC
            If FuncType Then
                If .INVOKEKIND = FuncType Then
                    'Debug.Print sName & vbTab & Switch(.INVOKEKIND = 1, "VbMethod", .INVOKEKIND = 2, "VbGet", .INVOKEKIND = 4, "VbLet", .INVOKEKIND = 8, "VbSet")
                    oCol.Add sName & vbTab & Switch(.INVOKEKIND = 1, "VbMethod", .INVOKEKIND = 2, "VbGet", .INVOKEKIND = 4, "VbLet", .INVOKEKIND = 8, "VbSet")
                End If
            Else
                'Debug.Print sName & vbTab & Switch(.INVOKEKIND = 1, "VbMethod", .INVOKEKIND = 2, "VbGet", .INVOKEKIND = 4, "VbLet", .INVOKEKIND = 8, "VbSet")
                oCol.Add sName & vbTab & Switch(.INVOKEKIND = 1, "VbMethod", .INVOKEKIND = 2, "VbGet", .INVOKEKIND = 4, "VbLet", .INVOKEKIND = 8, "VbSet")
            End If
        End With
        sName = vbNullString
    Next
 
    Set GetObjectFunctions = oCol

End Function



#If Win64 Then
    Private Function CallFunction_COM(ByVal InterfacePointer As LongLong, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant

    Dim vParamPtr() As LongLong
#Else
    Private Function CallFunction_COM(ByVal InterfacePointer As Long, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant

    Dim vParamPtr() As Long
#End If

    If InterfacePointer = 0& Or VTableOffset < 0& Then Exit Function
    If Not (FunctionReturnType And &HFFFF0000) = 0& Then Exit Function

    Dim pIndex As Long, pCount As Long
    Dim vParamType() As Integer
    Dim vRtn As Variant, vParams() As Variant

    vParams() = FunctionParameters()
    pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
    If pCount = 0& Then
        ReDim vParamPtr(0 To 0)
        ReDim vParamType(0 To 0)
    Else
        ReDim vParamPtr(0 To pCount - 1&)
        ReDim vParamType(0 To pCount - 1&)
        For pIndex = 0& To pCount - 1&
            vParamPtr(pIndex) = VarPtr(vParams(pIndex))
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next
    End If

    pIndex = DispCallFunc(InterfacePointer, VTableOffset, CallConvention, FunctionReturnType, pCount, _
    vParamType(0), vParamPtr(0), vRtn)
    If pIndex = 0& Then
        CallFunction_COM = vRtn
    Else
        SetLastError pIndex
    End If

End Function
































