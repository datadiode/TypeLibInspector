#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    IntrinsicAliasInfo (specialization of TypeInfo) definition for TLBINF.
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
#Region ###IntrinsicAliasInfo construction/destruction###
Func IntrinsicAliasInfo_New(ByRef $objITypeInfo, Const ByRef $objTypeLibInfo, $index = -1)
    Local $parent = TypeInfo_New($objITypeInfo, $objTypeLibInfo, $index)
    Return IntrinsicAliasInfo_Inherit($parent)
EndFunc

Func IntrinsicAliasInfo_Inherit(ByRef $objTypeInfo)
    Local $result = _AutoItObject_Create($objTypeInfo)
    _AutoItObject_AddMethod($result, "ResolvedType", "IntrinsicAliasInfo_ResolvedType")
;     _AutoItObject_AddDestructor($result, "IntrinsicAliasInfo_Release")
    Return $result
EndFunc

; Func IntrinsicAliasInfo_Release($oSelf)
;     ConsoleWrite("[DBG] IntrinsicAliasInfo_Release" & @LF)
; EndFunc
#EndRegion ;IntrinsicAliasInfo construction/destruction

#Region ###IntrinsicAliasInfo public property getters###
Func IntrinsicAliasInfo_ResolvedType($oSelf)
    Local $p = $oSelf._typeAttr.GetPtr("lpItemDesc")
    If $p Then
        Local $tElemDesc = DllStructCreate($tagELEMDESC, $p)
        Return VarTypeInfo_New($oSelf, $tElemDesc)
    EndIf
    Return 0
EndFunc
#EndRegion ;IntrinsicAliasInfo public property getters
