 Function getPreviousScopeVars() as Scope
        Dim var as Variant
        'Pointer to first var in this scope
        Dim ptrThisScope as LongPtr: ptrThisScope = varptr(var)
        'Each scope in memory is seporated by 152 bytes (From testing)
        Dim firstVarPrevScope as longPtr: firstVarPrevScope = ptrThisScope - 152
        'Ultimately we now have the VarPtr for the first variable in the previous scope.
        'To coerce this pointer into the value and name we can use that link
        'Okay so how to we move onto the next pointer? 
        '    Dim a As Object  'Requires 6 bytes
        '    Dim b As Variant 'Requires 16 bytes
        '    Dim c As String  'Requires 4 bytes
        '    Dim d As Integer 'Requires 2 bytes
        '    Dim e As Long    'Requires 6 bytes
        typesAndNamesInOrder = obtainPrevSubScopeVars()
        
        
        Dim curPtr as LongPtr: curPtr = firstVarPrevScope()
        Dim iTN as long, tTN as TypeAndName
        For iTN = 0 to ubound(typesAndNamesInOrder)
           tTN = typesAndNamesInOrder(iTN)
           
           select case tTN.type
              case vbObject
                 Dim o as object
                 Call RtlMoveMemory(o, curPtr, 6)
                 Call retScope.add(tTN.Name, tTN.type, o)
                 curPtr = curPtr + 6
              case vbVariant
                 Dim v as variant
                 Call RtlMoveMemory(v, curPtr, 16)
                 Call retScope.add(tTN.Name, tTN.Type, v)
                 curPtr = curPtr + 16
              '... etc ...
           end select
        next
    End Function