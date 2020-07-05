#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.5
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    ITypeInfo wrapper.
; Requirements:   AutoItObject, COMHelpers, MemoryHelpers, TLIStructures, TLIInterfaces
; $Revision: 34 $
; $Date: 2012-01-04 22:05:39 +0100 (Mi, 04. Jan 2012) $
;
; ------------------------------------------------------------------------------
#include <TLIStructures.au3>
#include <TLIInterfaces.au3>
#include <COMHelpers.au3>
#include <MemoryHelpers.au3>
#include <AutoItObject.au3>

#Region ##Structure helpers##
Func _FUNCDESC_GetElemDescParam(ByRef $tFuncDesc, $index)
    Local $result = DllStructCreate($tagELEMDESC, _carray_ptr(DllStructGetData($tFuncDesc, "lprgElemDescParam"), $index, "tagELEMDESC"))
    Return SetError(@error, @extended, $result)
EndFunc

Func _FUNCDESC_GetElemDescFunc(ByRef $tFuncDesc)
    Local $result = DllStructCreate($tagELEMDESC, DllStructGetPtr($tFuncDesc, "lpItemDesc"))
    Return SetError(@error, @extended, $result)
EndFunc

Func _VARDESC_GetElemDesc(ByRef $tVarDesc)
    Local $result = DllStructCreate($tagELEMDESC, DllStructGetPtr($tVarDesc, "lpItemDesc"))
    Return SetError(@error, @extended, $result)
EndFunc

Func _ARRAYDESC_GetRGBounds(ByRef $tArrayDesc, $dimension)
    Local $result = DllStructCreate($__Au3Obj_tagSAFEARRAYBOUND, _carray_ptr(DllStructGetPtr($tArrayDesc, "cElements"), $dimension, "__Au3Obj_tagSAFEARRAYBOUND"))
    Return SetError(@error, @extended, $result)
EndFunc
#EndRegion ;Structure helpers

#Region ##ITypeInfo interface##
;===============================================================================
; Function Name:   _ITypeInfo_GetContainingTypeLib
; Description:     Retrieves the containing type library and the index of the
;                  type description within that type library.
; 
; Parameter(s):    $oTInfo - ITypeInfo wrapper (s. _IDispatch_GetTypeInfo).
;                  $tiIndex - On return, contains the index of the type description
;                  within the containing type library.
; 
; Return Value(s): The containing type library.  In case of failure 0 is returned
;                  and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetContainingTypeLib
;===============================================================================
Func _ITypeInfo_GetContainingTypeLib(ByRef $oTInfo, ByRef $tiIndex)
    Local $tp = DllStructCreate("ptr;uint")
    Local $hr = $oTInfo.GetContainingTypeLib(Number(DllStructGetPtr($tp, 1)), Number(DllStructGetPtr($tp, 2)))
    If _COM_Succeded($hr) Then
        $tiIndex = DllStructGetData($tp, 2)
        Local $pTLib = DllStructGetData($tp, 1)
        If $pTLib Then Return _AutoItObject_WrapperCreate($pTLib, $tagITypeLib)
    EndIf
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetTypeAttr
; Description:     Retrieves a TYPEATTR structure that contains the attributes
;                  of the type description.
; 
; Parameter(s):    $oTInfo - ITypeInfo wrapper (s. _IDispatch_GetTypeInfo).
; 
; Return Value(s): Structure that contains the attributes of this type description.
;                  In case of failure 0 is returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetTypeAttr
;===============================================================================
Func _ITypeInfo_GetTypeAttr(ByRef $oTInfo)
    Local $p = _ITypeInfo_GetTypeAttrPtr($oTInfo)
    If $p Then Return DllStructCreate($tagTYPEATTR, $p)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetTypeAttrPtr
; Description:     Retrieves a TYPEATTR structure but as pointer (s. _ITypeInfo_GetTypeAttr).
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetTypeAttr
;===============================================================================
Func _ITypeInfo_GetTypeAttrPtr(ByRef $oTInfo)
    Local $tp = DllStructCreate("ptr")
    Local $hr = $oTInfo.GetTypeAttr(Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) Then Return DllStructGetData($tp, 1)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetNames
; Description:     Retrieves the variable with the specified member ID (or the name
;                  of the property or method and its parameters) that correspond
;                  to the specified function ID.
; 
; Parameter(s):    $oTInfo - ITypeInfo wrapper (s. _IDispatch_GetTypeInfo).
;                  $memid - The ID of the member whose name (or names) is to be returned.
;                  $rgBstrNames - Caller-allocated array. On return, each of its
;                  elements contains the name (or names) associated with the member.
; 
; Return Value(s): Number of names actually retrieved. In case of failure 0 is
;                  returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetNames
;===============================================================================
Func _ITypeInfo_GetNames(ByRef $oTInfo, $memid, ByRef $rgBstrNames)
    If Not IsArray($rgBstrNames) Then Return SetError(-1, 0, 0)
    Local $cNames = Ubound($rgBstrNames)
    Local $tp = DllStructCreate("ptr[" & $cNames & "];uint")
    Local $hr = $oTInfo.GetNames($memid, Number(DllStructGetPtr($tp, 1)), $cNames, Number(DllStructGetPtr($tp, 2)))
    If _COM_Succeded($hr) Then
        $cNames = DllStructGetData($tp, 2)
        For $i = 0 To $cNames - 1
            $rgBstrNames[$i] = __Au3Obj_SysReadString(DllStructGetData($tp, 1, $i + 1))
            __Au3Obj_SysFreeString(DllStructGetData($tp, 1, $i + 1))
        Next
        Return $cNames
    EndIf
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetFuncDesc
; Description:     Retrieves the FUNCDESC structure that contains information
;                  about a specified function.
; 
; Parameter(s):    $oTInfo - ITypeInfo wrapper (s. _IDispatch_GetTypeInfo).
;                  $index - Index of the function whose description is to be returned.
;                  The index should be in the range of 0 to 1 less than the
;                  number of functions in this type. 
; 
; Return Value(s): FUNCDESC that describes the specified function. In case of
;                  failure 0 is returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetFuncDesc
;===============================================================================
Func _ITypeInfo_GetFuncDesc(ByRef $oTInfo, $index)
    Local $p = _ITypeInfo_GetFuncDescPtr($oTInfo, $index)
    If $p Then Return DllStructCreate($tagFUNCDESC, $p)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetFuncDescPtr
; Description:     Retrieves the FUNCDESC structure but as pointer (s. _ITypeInfo_GetFuncDesc).
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetFuncDesc
;===============================================================================
Func _ITypeInfo_GetFuncDescPtr(ByRef $oTInfo, $index)
    Local $tp = DllStructCreate("ptr")
    Local $hr = $oTInfo.GetFuncDesc($index, Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) Then Return DllStructGetData($tp, 1)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetGetVarDesc
; Description:     Retrieves a VARDESC structure that describes the specified variable.
; 
; Parameter(s):    $oTInfo - ITypeInfo wrapper (s. _IDispatch_GetTypeInfo).
;                  $index - Index of the variable whose description is to be returned.
;                  The index should be in the range of 0 to 1 less than the
;                  number of variables in this type. 
; 
; Return Value(s): VARDESC that describes the specified variable. In case of
;                  failure 0 is returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetGetVarDesc
;===============================================================================
Func _ITypeInfo_GetGetVarDesc(ByRef $oTInfo, $index)
    Local $p = _ITypeInfo_GetGetVarDescPtr($oTInfo, $index)
    If $p Then Return DllStructCreate($tagVARDESC, $p)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetGetVarDescPtr
; Description:     Retrieves a VARDESC structure but as pointer (s. _ITypeInfo_GetGetVarDesc).
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetGetVarDesc
;===============================================================================
Func _ITypeInfo_GetGetVarDescPtr(ByRef $oTInfo, $index)
    Local $tp = DllStructCreate("ptr")
    Local $hr = $oTInfo.GetVarDesc($index, Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) Then Return DllStructGetData($tp, 1)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetDocumentation
; Description:     Retrieves the documentation string, the complete Help file
;                  name and path, and the context ID for the Help topic for a
;                  specified type description.
; 
; Parameter(s):    $oTInfo - ITypeInfo wrapper (s. _IDispatch_GetTypeInfo).
;                  $memid - ID of the member whose documentation is to be
;                  returned. Set to -1 to retrieve documentation for the type itself.
;                  $desiredInfo - Caller-allocated array of 4 elements. On return,
;                  each element that contained 1 would be filled with following
;                  documentation data
;                      $desiredInfo[0] - the name of the specified item
;                      $desiredInfo[1] - the documentation string for the specified item
;                      $desiredInfo[2] - the Help context associated with the specified item
;                      $desiredInfo[3] - the fully qualified name of the Help file
; 
; Return Value(s): True or False if failed and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetDocumentation
;===============================================================================
Func _ITypeInfo_GetDocumentation(ByRef $oTInfo, $memid, ByRef $desiredInfo)
    If Not IsArray($desiredInfo) Or 4 > UBound($desiredInfo) Then Return SetError(-1, 0, False)
    Local $tp = DllStructCreate("ptr;ptr;ulong;ptr")
    For $i = 0 To 3
        If $desiredInfo[$i] Then $desiredInfo[$i] = Number(DllStructGetPtr($tp, $i + 1))
    Next
    Local $hr = $oTInfo.GetDocumentation($memid, $desiredInfo[0], $desiredInfo[1], $desiredInfo[2], $desiredInfo[3])
    If _COM_Succeded($hr) Then
        For $i = 0 To 3
            If $TDOC_HELPCONTEXT = $i Then
                $desiredInfo[$i] = DllStructGetData($tp, $i + 1)
            Else
                $desiredInfo[$i] = __Au3Obj_SysReadString(DllStructGetData($tp, $i + 1))
                __Au3Obj_SysFreeString(DllStructGetData($tp, $i + 1))
            EndIf
        Next
        Return True
    EndIf
    _COM_SetError($hr)
    Return SetError(@error, @extended, False)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetImplTypeFlags
; Description:     Retrieves the IMPLTYPEFLAGS enumeration for one implemented
;                  interface or base interface in a type description.
; 
; Parameter(s):    $oTInfo - ITypeInfo wrapper (s. _IDispatch_GetTypeInfo).
;                  $index - Index of the implemented interface or base interface
;                  for which to get the flags. 
; 
; Return Value(s): The IMPLTYPEFLAGS enumeration value. In case of
;                  failure 0 is returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetImplTypeFlags
;===============================================================================
Func _ITypeInfo_GetImplTypeFlags(ByRef $oTInfo, $index)
    Local $tp = DllStructCreate("int")
    Local $hr = $oTInfo.GetImplTypeFlags($index, Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) Then Return DllStructGetData($tp, 1)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetTypeOfImplType
; Description:     If a type description describes a COM class, it retrieves the
;                  type description of the implemented interface types. For an
;                  interface, returns the type information for inherited
;                  interfaces, if any exist.
; 
; Parameter(s):    $oTInfo - ITypeInfo wrapper (s. _IDispatch_GetTypeInfo).
;                  $index - Index of the implemented type whose handle is returned.
;                  The valid range is 0 to the cImplTypes field in the TYPEATTR
;                  structure. 
; 
; Return Value(s): The referenced type description. In case of failure 0 is
;                  returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetRefTypeOfImplType
;===============================================================================
Func _ITypeInfo_GetTypeOfImplType(ByRef $oTInfo, $index)
    Local $tp = DllStructCreate("ulong;ptr")
    Local $hr = $oTInfo.GetRefTypeOfImplType($index, Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) Then
        $hr = $oTInfo.GetRefTypeInfo(DllStructGetData($tp, 1), Number(DllStructGetPtr($tp, 2)))
        If _COM_Succeded($hr) And DllStructGetData($tp, 2) Then Return _AutoItObject_WrapperCreate(DllStructGetData($tp, 2), $tagITypeInfo)
    EndIf
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetRefTypeInfo
; Description:     If a type description references other type descriptions, it
;                  retrieves the referenced type descriptions.
; 
; Parameter(s):    $oTInfo - ITypeInfo wrapper (s. _IDispatch_GetTypeInfo).
;                  $hRefType - Handle to the referenced type description to be returned. 
; 
; Return Value(s): The referenced type description. In case of failure 0 is
;                  returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetRefTypeInfo
;===============================================================================
Func _ITypeInfo_GetRefTypeInfo(ByRef $oTInfo, $hRefType)
    Local $tp = DllStructCreate("ptr")
    Local $hr = $oTInfo.GetRefTypeInfo($hRefType, Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) And DllStructGetData($tp, 1) Then Return _AutoItObject_WrapperCreate(DllStructGetData($tp, 1), $tagITypeInfo)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeInfo_GetIDsOfNames
; Description:     Maps between member names and member IDs, and parameter names
;                  and parameter IDs.
; 
; Parameter(s):    $oTInfo - ITypeInfo wrapper (s. _IDispatch_GetTypeInfo).
;                  $rgszNames - Array of names to be mapped. 
; 
; Return Value(s): Array in which name mappings are placed. In case of failure 0 is
;                  returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeInfo::GetIDsOfNames
;===============================================================================
Func _ITypeInfo_GetIDsOfNames(ByRef $oTInfo, ByRef $rgszNames)
    If Not IsArray($rgszNames) Then Return SetError(-1, 0, 0)
    Local $cNames = Ubound($rgszNames)
    Local $tp = DllStructCreate("ptr[" & $cNames & "];" & $tMEMBERID & "[" & $cNames & "]")
    
    For $i = 0 To $cNames - 1
        Local $tBuff = DllStructCreate("wchar[" & StringLen($rgszNames[$i]) + 1 & "]")
        DllStructSetData($tBuff, 1, $rgszNames[$i])
        DllStructSetData($tp, 1, DllStructGetPtr($tBuff, 1), 1)
    Next
    Local $hr = $oTInfo.GetIDsOfNames(Number(DllStructGetPtr($tp, 1)), $cNames, Number(DllStructGetPtr($tp, 2)))
    If _COM_Succeded($hr) Then
        Local $result[$cNames]
        For $i = 0 To $cNames - 1
            $result[$i] = DllStructGetData($tp, 2, $i + 1)
        Next
        Return $result
    EndIf
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
#EndRegion ;ITypeInfo interface