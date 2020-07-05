#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.2
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Definitions for TLBINF:
;                 PropertyInfo  (specialization of MemberInfo),
;                 VarDesc
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 19 $
; $Date: 2010-05-16 20:24:46 +0200 (Sun, 16 May 2010) $
;
; ------------------------------------------------------------------------------
#Region ###TypeInfo public helpers###
Func TypeInfo_Properties($oSelf)
    Local $c = $oSelf._typeAttr.GetData("cVars")
    Return MemberCollection_New($oSelf, $c, $DESCKIND_VARDESC)
EndFunc
#EndRegion ;TypeInfo public helpers

#Region ###VarDesc construction/destruction###
Func VarDesc_New(ByRef $objITypeInfo, $index)
    Local $result = _AutoItObject_Create(_TLI_StructWrapper_New($tagVARDESC))
    _AutoItObject_AddProperty($result, "_objITInfo", $ELSCOPE_PRIVATE, $objITypeInfo)
    _AutoItObject_AddProperty($result, "_index", $ELSCOPE_PRIVATE, $index)
    _AutoItObject_AddMethod($result, "GetPtr", "VarDesc_GetPtr")
    _AutoItObject_AddDestructor($result, "VarDesc_Release")
    Return $result
EndFunc

Func VarDesc_ForMember(ByRef $objMemberInfo)
    Local $objITInfo = $objMemberInfo._tInfo._objITInfo
    Return VarDesc_New($objITInfo, $objMemberInfo.MemberNumber)
EndFunc

Func VarDesc_Release($oSelf)
;     ConsoleWrite("[DBG] VarDesc_Release" & @LF)
    If $oSelf._StructPtr Then
        If $_TLI_debug Then ConsoleWrite("[DBG] Freeing VARDESC pointer" & @LF)
        $oSelf._objITInfo.ReleaseVarDesc(Number($oSelf._StructPtr))
    EndIf
    $oSelf._StructPtr = 0
    $oSelf._objITInfo = 0
EndFunc
#EndRegion ;VarDesc construction/destruction 

#Region ###VarDesc publc methods###
Func VarDesc_GetPtr($oSelf, $element = 0)
    If 0 = $oSelf._StructPtr Then
        Local $objITInfo = $oSelf._objITInfo
        $oSelf._StructPtr = Number(_ITypeInfo_GetGetVarDescPtr($objITInfo, $oSelf._index))
    EndIf
    Return _TLI_StructWrapper_GetPtr($oSelf, $element)
EndFunc
#EndRegion ;VarDesc public methods 

#Region ###PropertyInfo construction/destruction###
Func PropertyInfo_New(ByRef $objTypeInfo, $index)
    Local $parent = MemberInfo_New($objTypeInfo, $index)
    Return PropertyInfo_Inherit($parent)
EndFunc

Func PropertyInfo_Inherit(ByRef $objMemberInfo)
    Local $result = _AutoItObject_Create($objMemberInfo)
    _AutoItObject_AddProperty($result, "_desc", $ELSCOPE_PRIVATE, VarDesc_ForMember($objMemberInfo))
    _AutoItObject_AddProperty($result, "DescKind", $ELSCOPE_READONLY, $DESCKIND_VARDESC)
    _AutoItObject_AddMethod($result, "AttributeMask", "PropertyInfo_AttributeMask")
    _AutoItObject_AddMethod($result, "ReturnType", "PropertyInfo_ReturnType")
    _AutoItObject_AddMethod($result, "VarKind", "PropertyInfo_VarKind")
    _AutoItObject_AddMethod($result, "Value", "PropertyInfo_Value")
    _AutoItObject_AddDestructor($result, "PropertyInfo_Release")
    Return $result
EndFunc

Func PropertyInfo_Release($oSelf)
;     ConsoleWrite("[DBG] PropertyInfo_Release" & @LF)
    $oSelf._desc = 0
EndFunc
#EndRegion ;PropertyInfo construction/destruction

#Region ###PropertyInfo public property getters###
Func PropertyInfo_AttributeMask($oSelf)
    Return $oSelf._desc.GetData("wVarFlags")
EndFunc

Func PropertyInfo_ReturnType($oSelf)
    Local $p = $oSelf._desc.GetPtr("lpItemDesc")
    If $p Then
        Local $obj = $oSelf._tInfo
        Local $t = DllStructCreate($tagELEMDESC, $p)
        Return VarTypeInfo_New($obj, $t)
    EndIf
    Return 0
EndFunc

Func PropertyInfo_VarKind($oSelf)
    Return $oSelf._desc.GetData("varkind")
EndFunc

Func PropertyInfo_Value($oSelf)
    If $VAR_CONST = $oSelf.VarKind() Then
        Local $p = $oSelf._desc.GetData("lpVarValue")
        If $p Then
            Local $result = _AutoItObject_VariantRead($p)
            If IsPtr($result) Then $result = Number($result)
            Return $result
        EndIf
    EndIf
    Return Default
EndFunc
#EndRegion ;PropertyInfo public property getters
