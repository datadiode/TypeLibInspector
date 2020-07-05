#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.1
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    CoClassInfo (specialization of TypeInfo) definition for TLBINF.
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
#Region ###CoClassInfo construction/destruction###
Func CoClassInfo_New(ByRef $objITypeInfo, Const ByRef $objTypeLibInfo, $index = -1)
    Local $parent = TypeInfo_New($objITypeInfo, $objTypeLibInfo, $index)
    Return CoClassInfo_Inherit($parent)
EndFunc

Func CoClassInfo_Inherit(ByRef $objTypeInfo)
    Local $result = _AutoItObject_Create($objTypeInfo)
    _AutoItObject_AddProperty($result, "_iDefIface", $ELSCOPE_PRIVATE, -1)
    _AutoItObject_AddProperty($result, "_iDefEvt", $ELSCOPE_PRIVATE, -1)
    _AutoItObject_AddMethod($result, "Interfaces", "InterfaceInfo_Interfaces")
    _AutoItObject_AddMethod($result, "DefaultInterface", "CoClassInfo_DefaultInterface")
    _AutoItObject_AddMethod($result, "DefaultEventInterface", "CoClassInfo_DefaultEventInterface")
    _AutoItObject_AddMethod($result, "ImplTypeFlags", "CoClassInfo_ImplTypeFlags")
    _AutoItObject_AddMethod($result, "_FindImplInterface", "CoClassInfo__FindImplInterface", True)
;     _AutoItObject_AddDestructor($result, "CoClassInfo_Release")
    Return $result
EndFunc

; Func CoClassInfo_Release($oSelf)
;     ConsoleWrite("[DBG] CoClassInfo_Release" & @LF)
; EndFunc
#EndRegion ;CoClassInfo construction/destruction

#Region ###CoClassInfo public property getters###
Func CoClassInfo_DefaultInterface($oSelf)
    If -1 = $oSelf._iDefIface Then
        $oSelf._iDefIface = $oSelf._FindImplInterface($IMPLTYPEFLAG_FDEFAULT, $IMPLTYPEFLAG_FSOURCE)
        If 0 > $oSelf._iDefIface Then $oSelf._iDefIface = -2
    EndIf
    If -1 < $oSelf._iDefIface Then
        Local $obj = $oSelf._objITInfo
        Local $objImplTInfo = _ITypeInfo_GetTypeOfImplType($obj, $oSelf._iDefIface)
        Return _TLI_TypeInfoFromITypeInfo($objImplTInfo, 0)
    EndIf
    Return 0
EndFunc

Func CoClassInfo_DefaultEventInterface($oSelf)
    If -1 = $oSelf._iDefEvt Then
        $oSelf._iDefEvt = $oSelf._FindImplInterface($IMPLTYPEFLAG_FSOURCE, 0)
;         $oSelf._iDefEvt = $oSelf._FindImplInterface(BitOr($IMPLTYPEFLAG_FDEFAULT, $IMPLTYPEFLAG_FSOURCE), 0)
        If 0 > $oSelf._iDefEvt Then $oSelf._iDefEvt = -2
    EndIf
    If -1 < $oSelf._iDefEvt Then
        Local $obj = $oSelf._objITInfo
        Local $objImplTInfo = _ITypeInfo_GetTypeOfImplType($obj, $oSelf._iDefEvt)
        Return _TLI_TypeInfoFromITypeInfo($objImplTInfo, 0)
    EndIf
    Return 0
EndFunc

Func CoClassInfo_ImplTypeFlags($oSelf, $implIndex)
    Local $result = 0
    If 0 < $oSelf._typeAttr.GetData("cImplTypes") Then
        Local $obj = $oSelf._objITInfo
        $result = _ITypeInfo_GetImplTypeFlags($obj, $implIndex)
    EndIf
    Return $result
EndFunc
#EndRegion ;CoClassInfo public property getters

#Region ###CoClassInfo private methods###
Func CoClassInfo__FindImplInterface($oSelf, $flags, $notFlags)
    Local $c = $oSelf._typeAttr.GetData("cImplTypes")
    If 0 < $c Then
        Local $obj = $oSelf._objITInfo
        For $i = 0 To $c - 1
            Local $f = _ITypeInfo_GetImplTypeFlags($obj, $i)
            If 0 < BitAnd($flags, $f) And 0 = BitAnd($notFlags, $f) Then Return $i
        Next
    EndIf
    Return -1
EndFunc
#EndRegion ;CoClassInfo private methods

