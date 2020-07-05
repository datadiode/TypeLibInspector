#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Enum interface for simple AutoItObject collections.
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
#Region ###SimpleCollection enum###
Func SimpleCollection_EnumNext(ByRef $oSelf, ByRef $iterator)
    If Not IsNumber($iterator) Then $iterator = 0
    If $iterator >= $oSelf.Count Then Return SetError(1, 0, 0)
    Local $i = $iterator
    $iterator += 1
    Return $oSelf.Item($i)
EndFunc

Func SimpleCollection_EnumReset(ByRef $oSelf, ByRef $iterator)
    $iterator = 0
EndFunc
#EndRegion ;SimpleCollection enum
