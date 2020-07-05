#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.6
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    VarTypeInfo definition for TLBINF.
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 35 $
; $Date: 2012-01-04 23:14:30 +0100 (Mi, 04. Jan 2012) $
;
; ------------------------------------------------------------------------------
#include <ITypeInfo.au3>
#include <AutoItObject.au3>

#Region ###VarTypeInfo construction/destruction###
Func VarTypeInfo_New(ByRef $objTypeInfo, ByRef $tElemDesc)
    Local $result = _AutoItObject_Create()
    
    Local $vt = DllStructGetData($tElemDesc, "vt")
    Local $ptrLevel = 0
    Local $tDerefDesc = $tElemDesc
    While $__Au3Obj_VT_PTR = $vt
        $tDerefDesc = DllStructCreate($tagTYPEDESC, DllStructGetData($tDerefDesc, "lpItemDesc"))
        $vt = DllStructGetData($tDerefDesc, "vt")
        $ptrLevel += 1
    WEnd
;     If $__Au3Obj_VT_SAFEARRAY = $vt Then
;         $tDerefDesc = DllStructCreate($tagTYPEDESC, DllStructGetData($tDerefDesc, "lpItemDesc"))
;         $vt = DllStructGetData($tDerefDesc, "vt")
;     EndIf
    
    Local $dims = 0
    Local $href = -1
    If $__Au3Obj_VT_USERDEFINED = $vt Then
        $href = DllStructGetData($tDerefDesc, "lpItemDesc")
    ElseIf $__Au3Obj_VT_CARRAY = $vt Or BitAnd($__Au3Obj_VT_ARRAY, $vt) Or BitAnd($__Au3Obj_VT_VECTOR, $vt) Then
        Local $tArrayDesc = DllStructCreate($tagARRAYDESC, DllStructGetData($tDerefDesc, "lpItemDesc"))
        $dims = DllStructGetData($tArrayDesc, "cDims")
    EndIf
    _AutoItObject_AddProperty($result, "VarType", $ELSCOPE_READONLY, $vt)
    _AutoItObject_AddProperty($result, "PointerLevel", $ELSCOPE_READONLY, $ptrLevel)
    _AutoItObject_AddProperty($result, "_tInfo", $ELSCOPE_PRIVATE, $objTypeInfo)
    _AutoItObject_AddProperty($result, "_arrayDims", $ELSCOPE_PRIVATE, $dims)
    _AutoItObject_AddProperty($result, "_href", $ELSCOPE_PRIVATE, $href)
    _AutoItObject_AddProperty($result, "_pDesc", $ELSCOPE_PRIVATE, Number(DllStructGetPtr($tDerefDesc)))
    _AutoItObject_AddMethod($result, "ArrayBounds", "VarTypeInfo_ArrayBounds")
    _AutoItObject_AddMethod($result, "ArrayElementInfo", "VarTypeInfo_ArrayElementInfo")
    _AutoItObject_AddMethod($result, "TypeInfo", "VarTypeInfo_TypeInfo")
    _AutoItObject_AddMethod($result, "IsExternalType", "VarTypeInfo_IsExternalType")
    _AutoItObject_AddDestructor($result, "VarTypeInfo_Release")
    Return $result
EndFunc

Func VarTypeInfo_Release($oSelf)
;     ConsoleWrite("[DBG] VarTypeInfo_Release" & @LF)
    $oSelf._pDesc = 0
    $oSelf._tInfo = 0
EndFunc
#EndRegion ;VarTypeInfo construction/destruction

#Region ###VarTypeInfo public property getters###
Func VarTypeInfo_ArrayBounds($oSelf)
    Local $result[$oSelf._arrayDims + 1] = [$oSelf._arrayDims]
    If 0 < $oSelf._arrayDims Then
        Local $tArrayDesc = DllStructCreate($tagARRAYDESC, DllStructGetData(DllStructCreate($tagTYPEDESC, $oSelf._pDesc), "lpItemDesc"))
        For $i = 0 To $oSelf._arrayDims - 1
            Local $tb = _ARRAYDESC_GetRGBounds($tArrayDesc, $i)
            Local $bounds[2] = [DllStructGetData($tb, "lLbound"), DllStructGetData($tb, "cElements")]
            $result[$i + 1] = $bounds
        Next
    EndIf
    Return $result
EndFunc

Func VarTypeInfo_ArrayElementInfo($oSelf)
    Local $tDesc = 0
    If 0 <= $oSelf._arrayDims Then
        Local $tArrayDesc = DllStructCreate($tagARRAYDESC, DllStructGetData(DllStructCreate($tagTYPEDESC, $oSelf._pDesc), "lpItemDesc"))
        $tDesc = DllStructCreate($tagTYPEDESC, DllStructGetPtr($tArrayDesc, "lpItemDesc"))
    ElseIf $__Au3Obj_VT_SAFEARRAY = $oSelf.VarType Then
        $tDesc = DllStructCreate($tagTYPEDESC, DllStructGetData(DllStructCreate($tagTYPEDESC, $oSelf._pDesc), "lpItemDesc"))
    EndIf
    If IsDllStruct($tDesc) Then
        Local $obj = $oSelf._tInfo
        Return VarTypeInfo_New($obj, $tDesc)
    EndIf
    Return 0
EndFunc

Func VarTypeInfo_TypeInfo($oSelf)
    If $__Au3Obj_VT_USERDEFINED = $oSelf.VarType And IsObj($oSelf._tInfo._objITInfo) Then
        Local $obj = $oSelf._tInfo._objITInfo
        Local $objITInfoRef = _ITypeInfo_GetRefTypeInfo($obj, $oSelf._href)
        If IsObj($objITInfoRef) Then
            Local $NULL = 0
            Return _TLI_TypeInfoFromITypeInfo($objITInfoRef, $NULL)
        EndIf
    EndIf
    Return 0
EndFunc

Func VarTypeInfo_IsExternalType($oSelf)
    Local $tInfo = $oSelf.TypeInfo()
    If IsObj($tInfo) And IsObj($oSelf._tInfo) Then
        Local $tlInfo1 = $oSelf._tInfo.Parent()
        Local $tlInfo2 = $tInfo.Parent()
        Return ($tlInfo1.GUID <> $tlInfo2.GUID Or $tlInfo1.MajorVersion <> $tlInfo2.MajorVersion Or $tlInfo1.MinorVersion <> $tlInfo2.MinorVersion Or $tlInfo1.LCID <> $tlInfo2.LCID)
    EndIf
    Return False
EndFunc
#EndRegion ;VarTypeInfo public property getters
