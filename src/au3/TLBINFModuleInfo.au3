#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    ModuleInfo (specialization of InterfaceInfo) definition for TLBINF.
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
#Region ###ModuleInfo construction/destruction###
Func ModuleInfo_New(ByRef $objITypeInfo, Const ByRef $objTypeLibInfo, $index = -1)
    Local $parent = TypeInfo_New($objITypeInfo, $objTypeLibInfo, $index)
    Return ModuleInfo_Inherit($parent)
EndFunc

Func ModuleInfo_Inherit(ByRef $objTypeInfo)
    Local $parent = InterfaceInfo_Inherit($objTypeInfo)
    Return ModuleInfo_InheritInterface($parent)
EndFunc

Func ModuleInfo_InheritInterface(ByRef $objInterfaceInfo)
    Local $result = _AutoItObject_Create($objInterfaceInfo)
    _AutoItObject_AddMethod($result, "Properties", "TypeInfo_Properties")
;     _AutoItObject_AddDestructor($result, "ModuleInfo_Release")
    Return $result
EndFunc

; Func ModuleInfo_Release($oSelf)
;     ConsoleWrite("[DBG] ModuleInfo_Release" & @LF)
; EndFunc
#EndRegion ;ModuleInfo construction/destruction
