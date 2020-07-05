#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    DispInterfaceInfo (specialization of InterfaceInfo) definition for TLBINF.
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
#Region ###DispInterfaceInfo construction/destruction###
Func DispInterfaceInfo_New(ByRef $objITypeInfo, Const ByRef $objTypeLibInfo, $index = -1)
    Local $parent = TypeInfo_New($objITypeInfo, $objTypeLibInfo, $index)
    Return DispInterfaceInfo_Inherit($parent)
EndFunc

Func DispInterfaceInfo_Inherit(ByRef $objTypeInfo)
    Local $parent = InterfaceInfo_Inherit($objTypeInfo)
    Return DispInterfaceInfo_InheritInterface($parent)
EndFunc

Func DispInterfaceInfo_InheritInterface(ByRef $objInterfaceInfo)
    Local $result = _AutoItObject_Create($objInterfaceInfo)
    _AutoItObject_AddMethod($result, "VTableInterface", "DispInterfaceInfo_VTableInterface")
    _AutoItObject_AddMethod($result, "Properties", "TypeInfo_Properties")
;     _AutoItObject_AddDestructor($result, "DispInterfaceInfo_Release")
    Return $result
EndFunc

; Func DispInterfaceInfo_Release($oSelf)
;     ConsoleWrite("[DBG] DispInterfaceInfo_Release" & @LF)
; EndFunc
#EndRegion ;DispInterfaceInfo construction/destruction

#Region ###DispInterfaceInfo public property getters###
Func DispInterfaceInfo_VTableInterface($oSelf)
    If $TKIND_DISPATCH = $oSelf.TypeKind() And 0 < BitAnd($TYPEFLAG_FDUAL, $oSelf.AttributeMask()) Then
        Local $obj = $oSelf._objITInfo
        Local $objImplTInfo = _ITypeInfo_GetTypeOfImplType($obj, -1)
        If IsObj($objImplTInfo) Then Return _TLI_TypeInfoFromITypeInfo($objImplTInfo, 0)
    EndIf
    Return 0
EndFunc
#EndRegion ;DispInterfaceInfo public property getters
