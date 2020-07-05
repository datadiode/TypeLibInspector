#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.5
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    IDispatch wrapper.
; Requirements:   AutoItObject, COMHelpers, COMInterfaces
; $Revision: 34 $
; $Date: 2012-01-04 22:05:39 +0100 (Mi, 04. Jan 2012) $
;
; ------------------------------------------------------------------------------
#include <COMHelpers.au3>
#include <COMInterfaces.au3>
#include <AutoItObject.au3>
;===============================================================================
; Function Name:   _IDispatch_Wrap
; Description:     Provides access to IDispatch on regular AutoIt objects
;                  (s. ObjCreate and ObjGet).
;
; Parameter(s):    $obj - COM object retrived with ObjCreate() or ObjGet().
;
; Return Value(s): AutoItObject wrapper for IDispatch interface.
;
; Author(s):       doudou
; Link:            @@MsdnLink@@ IDispatch
;===============================================================================
Func _IDispatch_Wrap(ByRef $obj)
    Local $result = 0
    Local $pDisp = _AutoItObject_IDispatchToPtr($obj)
    If $pDisp Then
        $result = _AutoItObject_WrapperCreate($pDisp, $tagIDispatch)
        If IsObj($result) Then _AutoItObject_IUnknownAddRef($pDisp)
    EndIf
    Return $result
EndFunc
;===============================================================================
; Function Name:   _IDispatch_GetTypeInfoCount
; Description:     Retrieves the number of type information interfaces that
;                  an object provides (either 0 or 1).
;
; Parameter(s):    $oDisp - IDispatch wrapper (s. _IDispatch_Wrap).
;
; Return Value(s): 1 or 0 if the object does not provide any type information.
;
; Author(s):       doudou
; Link:            @@MsdnLink@@ IDispatch::GetTypeInfoCount
;===============================================================================
Func _IDispatch_GetTypeInfoCount(ByRef $oDisp)
    Local $tp = DllStructCreate("uint")
    Local $hr = $oDisp.GetTypeInfoCount(Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) Then Return DllStructGetData($tp, 1)
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _IDispatch_GetTypeInfo
; Description:     Retrieves the type information for an object, which can then
;                  be used to get the type information for an interface.
; Requirement(s):  TLIInterfaces.
; 
; Parameter(s):    $oDisp - IDispatch wrapper (s. _IDispatch_Wrap).
;                  $iTInfo - (optional) The type information to return. Pass 0 to
;                  retrieve type information for the IDispatch implementation.
;                  $lcid - (optional) The locale identifier for the type information.
;                  An object may be able to return different type information for
;                  different languages. This is important for classes that support
;                  localized member names.
; Return Value(s): The requested type information object. In case of failure 0 is
;                  returned and @error is set to HRESULT value.
; 
; Author(s):       doudou
; Link:            @@MsdnLink@@ IDispatch::GetTypeInfo
;===============================================================================
Func _IDispatch_GetTypeInfo(ByRef $oDisp, $iTInfo = 0, $lcid = $LOCALE_SYSTEM_DEFAULT)
    Local $tp = DllStructCreate("ptr")
    Local $hr = $oDisp.GetTypeInfo($iTInfo, $lcid, Number(DllStructGetPtr($tp, 1)))
    If _COM_Succeded($hr) Then
        Local $pTInfo = DllStructGetData($tp, 1)
        If $pTInfo Then Return _AutoItObject_WrapperCreate($pTInfo, $tagITypeInfo)
    EndIf
    _COM_SetError($hr)
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _IDispatch_GetSizeInstance
; Description:     Convenience method for quick access to object size (s. _IDispatch_GetTypeInfo
;                  , _ITypeInfo_GetTypeAttr).
; Requirement(s):  TLIInterfaces, ITypeInfo.
;
; Parameter(s):    $oDisp - IDispatch wrapper (s. _IDispatch_Wrap).
;
; Return Value(s): The size of the instance of this type.
;
; Author(s):       doudou
;===============================================================================
Func _IDispatch_GetSizeInstance(ByRef $oDisp)
    Local $oTInfo = _IDispatch_GetTypeInfo($oDisp)
    If IsObj($oTInfo) Then
        Local $tiAttr = _ITypeInfo_GetTypeAttr($oTInfo)
        If IsDllStruct($tiAttr) Then
            Local $result = DllStructGetData($tiAttr, "cbSizeInstance")
            $oTInfo.ReleaseTypeAttr(Number(DllStructGetPtr($tiAttr)))
            Return $result
        EndIf 
    EndIf
    Return SetError(@error, @extended, -1)
EndFunc