#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    RecordInfo (specialization of TypeInfo) definition for TLBINF.
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
#Region ###RecordInfo construction/destruction###
Func RecordInfo_New(ByRef $objITypeInfo, Const ByRef $objTypeLibInfo, $index = -1)
    Local $parent = TypeInfo_New($objITypeInfo, $objTypeLibInfo, $index)
    Return RecordInfo_Inherit($parent)
EndFunc

Func RecordInfo_Inherit(ByRef $objTypeInfo)
    Local $result = _AutoItObject_Create($objTypeInfo)
    _AutoItObject_AddMethod($result, "Properties", "TypeInfo_Properties")
;     _AutoItObject_AddDestructor($result, "RecordInfo_Release")
    Return $result
EndFunc

; Func RecordInfo_Release($oSelf)
;     ConsoleWrite("[DBG] RecordInfo_Release" & @LF)
; EndFunc
#EndRegion ;RecordInfo construction/destruction
