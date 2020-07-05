#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.5
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    ITypeLib wrapper.
; Requirements:   WinAPI, AutoItObject, COMHelpers, TLIStructures, TLIInterfaces
; $Revision: 34 $
; $Date: 2012-01-04 22:05:39 +0100 (Mi, 04. Jan 2012) $
;
; ------------------------------------------------------------------------------
#include <TLIStructures.au3>
#include <TLIInterfaces.au3>
#include <COMHelpers.au3>
#include <AutoItObject.au3>
#include <WinAPI.au3>
;===============================================================================
; Function Name:   _ITypeLib_GetLibAttr
; Description:     Retrieves the structure that contains the library's attributes.
; 
; Parameter(s):    $oTLib - ITypeLib wrapper (s. _ITypeLib_Load).
; 
; Return Value(s): Structure that contains the library's attributes. In case of
;                  failure 0 is returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeLib::GetLibAttr
;===============================================================================
Func _ITypeLib_GetLibAttr(ByRef $oTLib)
    Local $p = _ITypeLib_GetLibAttrPtr($oTLib)
    If $p Then Return DllStructCreate($tagTLIBATTR, $p)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeLib_GetLibAttrPtr
; Description:     Retrieves the structure that contains the library's attributes
;                  but as pointer (s. _ITypeLib_GetLibAttr).
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeLib::GetLibAttr
;===============================================================================
Func _ITypeLib_GetLibAttrPtr(ByRef $oTLib)
    Local $tp = DllStructCreate("ptr")
    Local $hr = $oTLib.GetLibAttr(Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) Then Return DllStructGetData($tp, 1)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeLib_GetDocumentation
; Description:     Retrieves the library's documentation string, the complete
;                  Help file name and path, and the context identifier for the
;                  library Help topic in the Help file.
; 
; Parameter(s):    $oTLib - ITypeLib wrapper (s. _ITypeLib_Load).
;                  $index - Index of the type description whose documentation is
;                  to be returned. If index is–1, then the documentation for
;                  the library itself is returned.
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
; Link:            @@MsdnLink@@ ITypeLib::GetDocumentation
;===============================================================================
Func _ITypeLib_GetDocumentation(ByRef $oTLib, $index, ByRef $desiredInfo)
    If Not IsArray($desiredInfo) Or 4 > UBound($desiredInfo) Then Return SetError(-1, 0, False)
    Local $tp = DllStructCreate("ptr;ptr;ulong;ptr")
    For $i = 0 To 3
        If $desiredInfo[$i] Then $desiredInfo[$i] = Number(DllStructGetPtr($tp, $i + 1))
    Next
    Local $hr = $oTLib.GetDocumentation($index, $desiredInfo[0], $desiredInfo[1], $desiredInfo[2], $desiredInfo[3])
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
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeLib_GetTypeInfoCount
; Description:     Returns the number of type descriptions in the type library.
; 
; Parameter(s):    $oTLib - ITypeLib wrapper (s. _ITypeLib_Load).
; 
; Return Value(s): Number of type descriptions. In case of
;                  failure 0 is returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeLib::GetTypeInfoCount
;===============================================================================
Func _ITypeLib_GetTypeInfoCount(ByRef $oTLib)
    Local $hr = $oTLib.GetTypeInfoCount()
    If _COM_Succeded($hr) Then Return $hr[0]
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeLib_GetTypeInfoType
; Description:     Retrieves the type of a type description.
; 
; Parameter(s):    $oTLib - ITypeLib wrapper (s. _ITypeLib_Load).
;                  $index - The index of the type description within the type library.
; 
; Return Value(s): TYPEKIND enumeration value for the type description. In case of
;                  failure -1 is returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeLib::GetTypeInfoType
;===============================================================================
Func _ITypeLib_GetTypeInfoType(ByRef $oTLib, $index)
    Local $tp = DllStructCreate("int")
    Local $hr = $oTLib.GetTypeInfoType($index, Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) Then Return DllStructGetData($tp, 1)
    _COM_SetError($hr)
    Return SetError(@error, @extended, -1)
EndFunc
;===============================================================================
; Function Name:   _ITypeLib_GetTypeInfo
; Description:     Retrieves the specified type description in the library.
; 
; Parameter(s):    $oTLib - ITypeLib wrapper (s. _ITypeLib_Load).
;                  $index - The index of the type description within the type library.
; 
; Return Value(s): ITypeInfo wrapper. In case of failure 0 is returned and @error
;                  is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeLib::GetTypeInfo
;===============================================================================
Func _ITypeLib_GetTypeInfo(ByRef $oTLib, $index)
    Local $tp = DllStructCreate("ptr")
    Local $hr = $oTLib.GetTypeInfo($index, Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) And DllStructGetData($tp, 1) Then Return _AutoItObject_WrapperCreate(DllStructGetData($tp, 1), $tagITypeInfo)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeLib_GetTypeInfoOfGuid
; Description:     Retrieves the type description that corresponds to the specified GUID.
; 
; Parameter(s):    $oTLib - ITypeLib wrapper (s. _ITypeLib_Load).
;                  $guid - GUID of the type description. Can be binary or string value.
; 
; Return Value(s): ITypeInfo wrapper. In case of failure 0 is returned and @error
;                  is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ ITypeLib::GetTypeInfoOfGuid
;===============================================================================
Func _ITypeLib_GetTypeInfoOfGuid(ByRef $oTLib, Const ByRef $guid)
    Local $tp = DllStructCreate("ptr;byte[16]")
    If IsBinary($guid) Then
        DllStructSetData($tp, 2, $guid)
    Else
        _WinAPI_GUIDFromStringEx(String($guid), DllStructGetPtr($tp, 2))
    EndIf
    Local $hr = $oTLib.GetTypeInfoOfGuid(Number(DllStructGetPtr($tp, 2)), Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) And DllStructGetData($tp, 1) Then Return _AutoItObject_WrapperCreate(DllStructGetData($tp, 1), $tagITypeInfo)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeLib_Load
; Description:     Loads a type library and (optionally) registers it in the
;                  system registry.
; 
; Parameter(s):    $szFile - Specification for the type library file.
;                  $regkind - Identifies the kind of registration to perform for
;                  the type library ($REGKIND_DEFAULT, $REGKIND_REGISTER, or $REGKIND_NONE).
; 
; Return Value(s): ITypeLib wrapper being loaded. In case of failure 0 is returned
;                  and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ LoadTypeLibEx
;===============================================================================
Func _ITypeLib_Load($szFile, $regkind = $REGKIND_NONE)
    If Not IsDeclared("gh_AU3Obj_oleautdll") Then Assign("gh_AU3Obj_oleautdll", DllOpen("oleaut32.dll"), 2)
    Local $tp = DllStructCreate("ptr")
    Local $hr = DllCall($gh_AU3Obj_oleautdll, "ulong", "LoadTypeLibEx", "wstr", $szFile, "int", $regkind, "ptr", DllStructGetPtr($tp, 1))
    If _COM_Succeded($hr) And DllStructGetData($tp, 1) Then Return _AutoItObject_WrapperCreate(DllStructGetData($tp, 1), $tagITypeLib)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _ITypeLib_LoadReg
; Description:     Uses registry information to load a type library.
; 
; Parameter(s):    $guid - The GUID of the library being loaded. Can be binary or
;                  string value.
;                  $wVerMajor - Major version number of the library being loaded.
;                  $wVerMinor - Minor version number of the library being loaded.
;                  $lcid - National language code of the library being loaded.
; 
; Return Value(s): ITypeLib wrapper being loaded. In case of failure 0 is returned
;                  and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ LoadRegTypeLib
;===============================================================================
Func _ITypeLib_LoadReg(Const ByRef $guid, $wVerMajor = 1, $wVerMinor = 0, $lcid = $LOCALE_SYSTEM_DEFAULT)
    If Not IsDeclared("gh_AU3Obj_oleautdll") Then Assign("gh_AU3Obj_oleautdll", DllOpen("oleaut32.dll"), 2)
    Local $tp = DllStructCreate("ptr;byte[16]")
    If IsBinary($guid) Then
        DllStructSetData($tp, 2, $guid)
    Else
        _WinAPI_GUIDFromStringEx(String($guid), DllStructGetPtr($tp, 2))
    EndIf
    Local $hr = DllCall($gh_AU3Obj_oleautdll, "ulong", "LoadRegTypeLib", $tREFIID, DllStructGetPtr($tp, 2), "ushort", Int($wVerMajor), "ushort", Int($wVerMinor), $tLCID, $lcid, "ptr", DllStructGetPtr($tp, 1))
    If _COM_Succeded($hr) And DllStructGetData($tp, 1) Then Return _AutoItObject_WrapperCreate(DllStructGetData($tp, 1), $tagITypeLib)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
